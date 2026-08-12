import os
import re
import sqlite3
import unicodedata
import uuid
import zipfile
from datetime import datetime, timezone
from difflib import SequenceMatcher
from io import BytesIO
from pathlib import Path
from collections.abc import Mapping

import pandas as pd
import streamlit as st
import streamlit_authenticator as stauth
from google import genai
from google.genai import types
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, Field

try:
    import psycopg
    from psycopg.rows import dict_row
except Exception:
    psycopg = None
    dict_row = None


# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================

st.set_page_config(
    page_title="Sistema de Presupuestación Asistida",
    page_icon=None,
    layout="wide",
)


# =========================================================
# UTILIDADES
# =========================================================


def ahora_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def get_secret(nombre: str, default=None):
    try:
        if nombre in st.secrets:
            return st.secrets[nombre]
    except Exception:
        pass
    return os.getenv(nombre, default)


def _secrets_to_dict(value):
    """Convierte estructuras de st.secrets a dict/list normales."""
    if isinstance(value, Mapping):
        return {k: _secrets_to_dict(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_secrets_to_dict(v) for v in value]
    return value


def crear_autenticador():
    """Crea Streamlit-Authenticator usando exclusivamente Streamlit Secrets."""
    try:
        credentials = _secrets_to_dict(st.secrets["credenciales_app"])
        cookie_cfg = _secrets_to_dict(st.secrets["auth_cookie"])
    except Exception as exc:
        st.error(
            "Falta la configuración de acceso. Configure [credenciales_app] "
            "y [auth_cookie] en Streamlit Secrets."
        )
        st.caption(f"Detalle técnico: {exc}")
        st.stop()

    if not isinstance(credentials, dict) or "usernames" not in credentials:
        st.error("[credenciales_app] debe contener la sección 'usernames'.")
        st.stop()

    required_cookie = {"name", "key", "expiry_days"}
    missing = required_cookie - set(cookie_cfg or {})
    if missing:
        st.error(f"Faltan valores en [auth_cookie]: {', '.join(sorted(missing))}")
        st.stop()

    return stauth.Authenticate(
        credentials,
        str(cookie_cfg["name"]),
        str(cookie_cfg["key"]),
        float(cookie_cfg["expiry_days"]),
        auto_hash=True,
    )


def autenticar_usuario():
    """Bloquea toda la aplicación hasta que exista una sesión válida."""
    authenticator = crear_autenticador()

    st.title("Sistema de Presupuestación Asistida")
    if st.session_state.get("authentication_status") is not True:
        st.caption("Acceso interno")

    try:
        authenticator.login(
            location="main",
            max_login_attempts=5,
            fields={
                "Form name": "Acceso",
                "Username": "Usuario",
                "Password": "Contraseña",
                "Login": "Ingresar",
            },
        )
    except Exception as exc:
        st.error(f"No fue posible iniciar el módulo de autenticación: {exc}")
        st.stop()

    status = st.session_state.get("authentication_status")
    if status is False:
        st.error("Usuario o contraseña incorrectos.")
        st.stop()
    if status is None:
        st.stop()

    roles = st.session_state.get("roles") or []
    if isinstance(roles, str):
        roles = [roles]
    roles = [str(r).strip().lower() for r in roles]

    return authenticator, roles


def normalizar_texto(texto: str) -> str:
    texto = (texto or "").strip().lower()
    texto = "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )
    texto = re.sub(r"[^a-z0-9\s]", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def abreviar(texto: str, longitud: int = 3) -> str:
    limpio = normalizar_texto(texto).upper()
    palabras = [p for p in limpio.split() if p not in {"DE", "DEL", "LA", "EL", "EN", "Y"}]
    if not palabras:
        return "PRY"
    base = palabras[0]
    if len(base) >= longitud:
        return base[:longitud]
    return (base + "XXX")[:longitud]


def abreviacion_tipo(tipo: str) -> str:
    mapa = {
        "Baño": "BAN",
        "Cocina": "COC",
        "Recámara": "REC",
        "Sala / comedor": "SAL",
        "Local comercial": "LOC",
        "Oficina": "OFI",
        "Remodelación interior general": "REM",
        "Caseta / acceso": "CAS",
        "Otro": "OTR",
    }
    return mapa.get(tipo, abreviar(tipo))


def formato_moneda(valor: float) -> str:
    return f"${valor:,.2f}"


def score_similitud(a: str, b: str) -> float:
    a_n = normalizar_texto(a)
    b_n = normalizar_texto(b)
    if not a_n or not b_n:
        return 0.0

    secuencia = SequenceMatcher(None, a_n, b_n).ratio()
    ta = set(a_n.split())
    tb = set(b_n.split())
    union = ta | tb
    jaccard = len(ta & tb) / len(union) if union else 0.0
    return 0.65 * secuencia + 0.35 * jaccard


def limpiar_codigo(codigo: str, fallback: str) -> str:
    codigo = re.sub(r"[^A-Za-z0-9\-]", "", (codigo or "").upper())
    return codigo[:24] if codigo else fallback


# =========================================================
# BASE DE DATOS
# =========================================================


class Database:
    """Base relacional con PostgreSQL opcional y SQLite como respaldo local."""

    def __init__(self, database_url: str | None = None):
        self.database_url = (database_url or "").strip() or None
        self.kind = "postgres" if self.database_url else "sqlite"
        self.sqlite_path = Path(__file__).with_name("presupuestador_empresa.db")
        self._init_schema()

    @property
    def persistent(self) -> bool:
        return self.kind == "postgres"

    def _connect(self):
        if self.kind == "postgres":
            if psycopg is None:
                raise RuntimeError("psycopg no está instalado.")
            return psycopg.connect(
                self.database_url,
                row_factory=dict_row,
                prepare_threshold=None,
                connect_timeout=15,
            )

        conn = sqlite3.connect(self.sqlite_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        return conn

    def _adapt(self, sql: str) -> str:
        return sql.replace("?", "%s") if self.kind == "postgres" else sql

    def execute(self, sql: str, params=()):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(self._adapt(sql), params)
            conn.commit()

    def executemany(self, sql: str, seq_params):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.executemany(self._adapt(sql), seq_params)
            conn.commit()

    def fetchone(self, sql: str, params=()):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(self._adapt(sql), params)
            row = cur.fetchone()
            return dict(row) if row is not None else None

    def fetchall(self, sql: str, params=()):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(self._adapt(sql), params)
            rows = cur.fetchall()
            return [dict(r) for r in rows]

    def _init_schema(self):
        schema = [
            """
            CREATE TABLE IF NOT EXISTS projects (
                id TEXT PRIMARY KEY,
                code TEXT UNIQUE NOT NULL,
                name TEXT NOT NULL,
                project_type TEXT NOT NULL,
                location TEXT,
                dimension_mode TEXT,
                dimensions_text TEXT,
                description TEXT,
                guide_text TEXT,
                main_activity TEXT,
                created_at TEXT NOT NULL
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS budgets (
                id TEXT PRIMARY KEY,
                project_id TEXT NOT NULL,
                version INTEGER NOT NULL DEFAULT 1,
                status TEXT NOT NULL,
                indirect_pct REAL NOT NULL,
                profit_pct REAL NOT NULL,
                iva_pct REAL NOT NULL,
                waste_pct REAL NOT NULL,
                direct_cost REAL NOT NULL,
                indirect_cost REAL NOT NULL,
                profit REAL NOT NULL,
                sale_before_tax REAL NOT NULL,
                iva_amount REAL NOT NULL,
                total REAL NOT NULL,
                scope_summary TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS concepts (
                id TEXT PRIMARY KEY,
                code TEXT,
                category TEXT,
                subcategory TEXT,
                description TEXT NOT NULL,
                unit TEXT NOT NULL,
                normalized_description TEXT NOT NULL,
                created_budget_id TEXT,
                created_at TEXT NOT NULL
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS price_history (
                id TEXT PRIMARY KEY,
                concept_id TEXT NOT NULL,
                unit_cost REAL NOT NULL,
                source TEXT NOT NULL,
                source_detail TEXT,
                status TEXT NOT NULL,
                confidence TEXT,
                project_id TEXT,
                budget_id TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(concept_id) REFERENCES concepts(id) ON DELETE CASCADE,
                FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE SET NULL,
                FOREIGN KEY(budget_id) REFERENCES budgets(id) ON DELETE CASCADE
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS budget_items (
                id TEXT PRIMARY KEY,
                budget_id TEXT NOT NULL,
                concept_id TEXT,
                category TEXT,
                subcategory TEXT,
                code TEXT,
                description TEXT NOT NULL,
                unit TEXT NOT NULL,
                quantity REAL NOT NULL,
                unit_cost REAL NOT NULL,
                direct_amount REAL NOT NULL,
                unit_indirect REAL NOT NULL,
                unit_profit REAL NOT NULL,
                unit_sale REAL NOT NULL,
                sale_amount REAL NOT NULL,
                sale_margin_pct REAL NOT NULL,
                benefit_amount REAL NOT NULL,
                price_source TEXT,
                price_source_detail TEXT,
                price_confidence TEXT,
                quantity_criterion TEXT,
                inclusion_basis TEXT,
                considerations TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(budget_id) REFERENCES budgets(id) ON DELETE CASCADE,
                FOREIGN KEY(concept_id) REFERENCES concepts(id) ON DELETE SET NULL
            )
            """,
            "CREATE INDEX IF NOT EXISTS idx_concepts_norm ON concepts(normalized_description)",
            "CREATE INDEX IF NOT EXISTS idx_price_concept ON price_history(concept_id)",
            "CREATE INDEX IF NOT EXISTS idx_items_budget ON budget_items(budget_id)",
        ]
        for statement in schema:
            self.execute(statement)

    def stats(self) -> dict:
        return {
            "projects": self.fetchone("SELECT COUNT(*) AS n FROM projects")["n"],
            "budgets": self.fetchone("SELECT COUNT(*) AS n FROM budgets")["n"],
            "concepts": self.fetchone("SELECT COUNT(*) AS n FROM concepts")["n"],
        }

    def next_project_code(self, location: str, project_type: str) -> str:
        prefix = f"{abreviar(location or 'Proyecto')}-{abreviacion_tipo(project_type)}"
        rows = self.fetchall("SELECT code FROM projects WHERE code LIKE ?", (f"{prefix}-%",))
        max_num = 0
        for row in rows:
            m = re.search(r"-(\d+)$", row["code"] or "")
            if m:
                max_num = max(max_num, int(m.group(1)))
        return f"{prefix}-{max_num + 1:04d}"

    def price_candidates(self, unit: str, limit: int = 600) -> list[dict]:
        rows = self.fetchall(
            """
            SELECT
                c.id AS concept_id,
                c.code,
                c.category,
                c.subcategory,
                c.description,
                c.unit,
                c.normalized_description,
                ph.unit_cost,
                ph.source,
                ph.source_detail,
                ph.status,
                ph.confidence,
                ph.created_at
            FROM concepts c
            JOIN price_history ph ON ph.concept_id = c.id
            WHERE UPPER(c.unit) = UPPER(?)
            ORDER BY ph.created_at DESC
            """,
            (unit,),
        )

        # Conserva solamente el precio más reciente de cada concepto.
        unique = []
        seen = set()
        for row in rows:
            if row["concept_id"] in seen:
                continue
            seen.add(row["concept_id"])
            unique.append(row)
            if len(unique) >= limit:
                break
        return unique

    def save_generation(
        self,
        project_code: str,
        project_data: dict,
        result,
        items: list[dict],
        params: dict,
        financials: dict,
    ) -> tuple[str, str]:
        project_id = str(uuid.uuid4())
        budget_id = str(uuid.uuid4())
        created = ahora_iso()

        self.execute(
            """
            INSERT INTO projects (
                id, code, name, project_type, location, dimension_mode,
                dimensions_text, description, guide_text, main_activity, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                project_id,
                project_code,
                project_data["name"],
                project_data["project_type"],
                project_data["location"],
                project_data["dimension_mode"],
                project_data["dimensions_text"],
                project_data["description"],
                project_data["guide_text"],
                result.actividad_principal,
                created,
            ),
        )

        self.execute(
            """
            INSERT INTO budgets (
                id, project_id, version, status, indirect_pct, profit_pct,
                iva_pct, waste_pct, direct_cost, indirect_cost, profit,
                sale_before_tax, iva_amount, total, scope_summary, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                budget_id,
                project_id,
                1,
                "GENERADO",
                params["indirect_pct"],
                params["profit_pct"],
                params["iva_pct"],
                params["waste_pct"],
                financials["direct_cost"],
                financials["indirect_cost"],
                financials["profit"],
                financials["sale_before_tax"],
                financials["iva_amount"],
                financials["total"],
                result.alcance_resumido,
                created,
            ),
        )

        for item in items:
            concept_id = item.get("concept_id")

            if not concept_id:
                concept_id = str(uuid.uuid4())
                item["concept_id"] = concept_id
                self.execute(
                    """
                    INSERT INTO concepts (
                        id, code, category, subcategory, description, unit,
                        normalized_description, created_budget_id, created_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        concept_id,
                        item["code"],
                        item["category"],
                        item["subcategory"],
                        item["description"],
                        item["unit"],
                        normalizar_texto(item["description"]),
                        budget_id,
                        created,
                    ),
                )

                # Solo se crea historial nuevo cuando el sistema generó o encontró
                # una referencia externa nueva. Un precio interno reutilizado ya
                # cuenta con historial propio.
                if item["price_source"] not in {"BASE_INTERNA", "HISTORICO_IA"}:
                    self.execute(
                        """
                        INSERT INTO price_history (
                            id, concept_id, unit_cost, source, source_detail,
                            status, confidence, project_id, budget_id, created_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            str(uuid.uuid4()),
                            concept_id,
                            item["unit_cost"],
                            item["price_source"],
                            item["price_source_detail"],
                            item["price_status"],
                            item["price_confidence"],
                            project_id,
                            budget_id,
                            created,
                        ),
                    )

            self.execute(
                """
                INSERT INTO budget_items (
                    id, budget_id, concept_id, category, subcategory, code,
                    description, unit, quantity, unit_cost, direct_amount,
                    unit_indirect, unit_profit, unit_sale, sale_amount,
                    sale_margin_pct, benefit_amount, price_source,
                    price_source_detail, price_confidence, quantity_criterion,
                    inclusion_basis, considerations, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    str(uuid.uuid4()),
                    budget_id,
                    concept_id,
                    item["category"],
                    item["subcategory"],
                    item["code"],
                    item["description"],
                    item["unit"],
                    item["quantity"],
                    item["unit_cost"],
                    item["direct_amount"],
                    item["unit_indirect"],
                    item["unit_profit"],
                    item["unit_sale"],
                    item["sale_amount"],
                    item["sale_margin_pct"],
                    item["benefit_amount"],
                    item["price_source"],
                    item["price_source_detail"],
                    item["price_confidence"],
                    item["quantity_criterion"],
                    item["inclusion_basis"],
                    item["considerations"],
                    created,
                ),
            )

        return project_id, budget_id

    def delete_generation(self, project_id: str, budget_id: str):
        # La eliminación del presupuesto borra partidas e historial vinculado
        # mediante ON DELETE CASCADE.
        self.execute("DELETE FROM budgets WHERE id = ?", (budget_id,))

        # Elimina conceptos creados exclusivamente por el presupuesto descartado.
        self.execute(
            """
            DELETE FROM concepts
            WHERE created_budget_id = ?
              AND NOT EXISTS (
                  SELECT 1 FROM budget_items bi WHERE bi.concept_id = concepts.id
              )
              AND NOT EXISTS (
                  SELECT 1 FROM price_history ph WHERE ph.concept_id = concepts.id
              )
            """,
            (budget_id,),
        )

        remaining = self.fetchone(
            "SELECT COUNT(*) AS n FROM budgets WHERE project_id = ?",
            (project_id,),
        )["n"]
        if remaining == 0:
            self.execute("DELETE FROM projects WHERE id = ?", (project_id,))


    # -----------------------------------------------------
    # Administración de la base interna
    # -----------------------------------------------------

    def list_concepts(self, search: str = "", limit: int = 500) -> list[dict]:
        search_n = normalizar_texto(search)
        if search_n:
            rows = self.fetchall(
                """
                SELECT c.*,
                       (SELECT ph.unit_cost FROM price_history ph
                        WHERE ph.concept_id = c.id
                        ORDER BY ph.created_at DESC LIMIT 1) AS latest_cost,
                       (SELECT ph.source FROM price_history ph
                        WHERE ph.concept_id = c.id
                        ORDER BY ph.created_at DESC LIMIT 1) AS latest_source,
                       (SELECT ph.status FROM price_history ph
                        WHERE ph.concept_id = c.id
                        ORDER BY ph.created_at DESC LIMIT 1) AS latest_status,
                       (SELECT COUNT(*) FROM budget_items bi
                        WHERE bi.concept_id = c.id) AS usage_count
                FROM concepts c
                WHERE c.normalized_description LIKE ?
                   OR LOWER(COALESCE(c.code, '')) LIKE ?
                   OR LOWER(COALESCE(c.category, '')) LIKE ?
                   OR LOWER(COALESCE(c.subcategory, '')) LIKE ?
                ORDER BY c.description
                LIMIT ?
                """,
                (f"%{search_n}%", f"%{search.lower()}%", f"%{search.lower()}%", f"%{search.lower()}%", limit),
            )
        else:
            rows = self.fetchall(
                """
                SELECT c.*,
                       (SELECT ph.unit_cost FROM price_history ph
                        WHERE ph.concept_id = c.id
                        ORDER BY ph.created_at DESC LIMIT 1) AS latest_cost,
                       (SELECT ph.source FROM price_history ph
                        WHERE ph.concept_id = c.id
                        ORDER BY ph.created_at DESC LIMIT 1) AS latest_source,
                       (SELECT ph.status FROM price_history ph
                        WHERE ph.concept_id = c.id
                        ORDER BY ph.created_at DESC LIMIT 1) AS latest_status,
                       (SELECT COUNT(*) FROM budget_items bi
                        WHERE bi.concept_id = c.id) AS usage_count
                FROM concepts c
                ORDER BY c.created_at DESC
                LIMIT ?
                """,
                (limit,),
            )
        return rows

    def get_concept(self, concept_id: str) -> dict | None:
        return self.fetchone("SELECT * FROM concepts WHERE id = ?", (concept_id,))

    def create_concept(self, code: str, category: str, subcategory: str, description: str, unit: str) -> str:
        concept_id = str(uuid.uuid4())
        self.execute(
            """
            INSERT INTO concepts (
                id, code, category, subcategory, description, unit,
                normalized_description, created_budget_id, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                concept_id,
                limpiar_codigo(code, "MAN-001"),
                category.strip(),
                subcategory.strip(),
                description.strip(),
                unit.strip().upper(),
                normalizar_texto(description),
                None,
                ahora_iso(),
            ),
        )
        return concept_id

    def update_concept(self, concept_id: str, code: str, category: str, subcategory: str, description: str, unit: str):
        self.execute(
            """
            UPDATE concepts
            SET code = ?, category = ?, subcategory = ?, description = ?,
                unit = ?, normalized_description = ?
            WHERE id = ?
            """,
            (
                limpiar_codigo(code, "CON-001"),
                category.strip(),
                subcategory.strip(),
                description.strip(),
                unit.strip().upper(),
                normalizar_texto(description),
                concept_id,
            ),
        )

    def concept_usage(self, concept_id: str) -> dict:
        return {
            "budget_items": self.fetchone(
                "SELECT COUNT(*) AS n FROM budget_items WHERE concept_id = ?", (concept_id,)
            )["n"],
            "prices": self.fetchone(
                "SELECT COUNT(*) AS n FROM price_history WHERE concept_id = ?", (concept_id,)
            )["n"],
        }

    def delete_concept(self, concept_id: str):
        self.execute("DELETE FROM concepts WHERE id = ?", (concept_id,))

    def list_prices(self, concept_id: str) -> list[dict]:
        return self.fetchall(
            """
            SELECT * FROM price_history
            WHERE concept_id = ?
            ORDER BY created_at DESC
            """,
            (concept_id,),
        )

    def add_price(
        self,
        concept_id: str,
        unit_cost: float,
        source: str,
        source_detail: str,
        status: str,
        confidence: str,
    ) -> str:
        price_id = str(uuid.uuid4())
        self.execute(
            """
            INSERT INTO price_history (
                id, concept_id, unit_cost, source, source_detail,
                status, confidence, project_id, budget_id, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                price_id,
                concept_id,
                float(unit_cost),
                source.strip().upper(),
                source_detail.strip(),
                status.strip().upper(),
                confidence.strip(),
                None,
                None,
                ahora_iso(),
            ),
        )
        return price_id

    def delete_price(self, price_id: str):
        self.execute("DELETE FROM price_history WHERE id = ?", (price_id,))

    def list_projects(self, limit: int = 500) -> list[dict]:
        return self.fetchall(
            """
            SELECT p.*,
                   (SELECT COUNT(*) FROM budgets b WHERE b.project_id = p.id) AS budget_count,
                   (SELECT MAX(b.total) FROM budgets b WHERE b.project_id = p.id) AS latest_total
            FROM projects p
            ORDER BY p.created_at DESC
            LIMIT ?
            """,
            (limit,),
        )

    def get_project(self, project_id: str) -> dict | None:
        return self.fetchone("SELECT * FROM projects WHERE id = ?", (project_id,))

    def update_project(
        self,
        project_id: str,
        name: str,
        project_type: str,
        location: str,
        main_activity: str,
        dimensions_text: str,
        description: str,
        guide_text: str,
    ):
        self.execute(
            """
            UPDATE projects
            SET name = ?, project_type = ?, location = ?, main_activity = ?,
                dimensions_text = ?, description = ?, guide_text = ?
            WHERE id = ?
            """,
            (
                name.strip(), project_type.strip(), location.strip(), main_activity.strip(),
                dimensions_text.strip(), description.strip(), guide_text.strip(), project_id,
            ),
        )

    def delete_project(self, project_id: str):
        budget_rows = self.fetchall("SELECT id FROM budgets WHERE project_id = ?", (project_id,))
        budget_ids = [r["id"] for r in budget_rows]
        self.execute("DELETE FROM projects WHERE id = ?", (project_id,))

        # Limpieza de conceptos que nacieron en presupuestos del proyecto eliminado
        # y que ya no cuentan con uso ni historial.
        for budget_id in budget_ids:
            self.execute(
                """
                DELETE FROM concepts
                WHERE created_budget_id = ?
                  AND NOT EXISTS (
                      SELECT 1 FROM budget_items bi WHERE bi.concept_id = concepts.id
                  )
                  AND NOT EXISTS (
                      SELECT 1 FROM price_history ph WHERE ph.concept_id = concepts.id
                  )
                """,
                (budget_id,),
            )

    def list_budgets(self, project_id: str | None = None, limit: int = 500) -> list[dict]:
        if project_id:
            return self.fetchall(
                """
                SELECT b.*, p.code AS project_code, p.name AS project_name,
                       p.location AS project_location
                FROM budgets b
                JOIN projects p ON p.id = b.project_id
                WHERE b.project_id = ?
                ORDER BY b.created_at DESC
                LIMIT ?
                """,
                (project_id, limit),
            )
        return self.fetchall(
            """
            SELECT b.*, p.code AS project_code, p.name AS project_name,
                   p.location AS project_location
            FROM budgets b
            JOIN projects p ON p.id = b.project_id
            ORDER BY b.created_at DESC
            LIMIT ?
            """,
            (limit,),
        )

    def get_budget(self, budget_id: str) -> dict | None:
        return self.fetchone(
            """
            SELECT b.*, p.code AS project_code, p.name AS project_name
            FROM budgets b
            JOIN projects p ON p.id = b.project_id
            WHERE b.id = ?
            """,
            (budget_id,),
        )

    def list_budget_items(self, budget_id: str) -> list[dict]:
        return self.fetchall(
            """
            SELECT * FROM budget_items
            WHERE budget_id = ?
            ORDER BY category, subcategory, code, created_at
            """,
            (budget_id,),
        )

    def delete_budget(self, budget_id: str):
        budget = self.fetchone("SELECT project_id FROM budgets WHERE id = ?", (budget_id,))
        if not budget:
            return
        self.delete_generation(budget["project_id"], budget_id)

    def export_table(self, table_name: str) -> list[dict]:
        allowed = {"projects", "budgets", "concepts", "price_history", "budget_items"}
        if table_name not in allowed:
            raise ValueError("Tabla no permitida.")
        return self.fetchall(f"SELECT * FROM {table_name}")


@st.cache_resource(show_spinner=False)
def get_database(database_url: str | None):
    """Usa PostgreSQL si DATABASE_URL existe; SQLite solo para desarrollo sin URL."""
    if database_url:
        # En producción no se hace fallback silencioso: si Supabase falla,
        # se detiene para evitar guardar datos en un SQLite efímero por accidente.
        return Database(database_url), None
    return Database(None), None


# =========================================================
# MODELOS DE RESPUESTA ESTRUCTURADA
# =========================================================


class ActividadIA(BaseModel):
    partida: str = Field(description="Partida general, por ejemplo PREPARACIÓN Y DEMOLICIONES")
    subpartida: str = Field(description="Subpartida breve, por ejemplo Preliminares")
    codigo_sugerido: str = Field(description="Código breve como PRE-01 o CAR-03")
    descripcion_tecnica: str = Field(description="Descripción técnica completa orientada a presupuesto")
    unidad: str = Field(description="Unidad: LOTE, PZA, M2, M3, ML, PTO, JGO, etc.")
    cantidad: float = Field(ge=0, description="Cantidad justificable con la información disponible")
    costo_unitario_estimado: float = Field(
        ge=0,
        description="Costo unitario integrado estimado del subcontratista, en MXN, antes de indirectos y utilidad"
    )
    criterio_cantidad: str = Field(description="Criterio verificable usado para determinar la cantidad")
    fundamento_inclusion: str = Field(description="Razón breve para incluir la actividad en el alcance")
    nivel_confianza_cantidad: str = Field(description="Alta, Media o Baja")
    nivel_confianza_precio: str = Field(description="Alta, Media o Baja")
    requiere_cotizacion: bool = Field(description="True si el costo debería confirmarse con proveedor especializado")
    consideraciones: str = Field(description="Supuestos, exclusiones o condiciones relevantes")


class PresupuestoIA(BaseModel):
    nombre_proyecto: str
    actividad_principal: str
    alcance_resumido: str
    consideraciones_generales: list[str]
    datos_faltantes: list[str]
    actividades: list[ActividadIA]


# =========================================================
# GEMINI
# =========================================================


def get_api_key() -> str | None:
    return get_secret("GEMINI_API_KEY")


def generar_presupuesto_ia(
    api_key: str,
    model_name: str,
    project_data: dict,
    params: dict,
) -> PresupuestoIA:
    client = genai.Client(api_key=api_key)
    year = datetime.now().year

    prompt = f"""
Actúa como especialista en presupuestación para una empresa de interiorismo y
remodelación en México. La empresa SUBCONTRATA prácticamente todas las
actividades. El objetivo no es desarrollar análisis de precios unitarios de
materiales y mano de obra, sino generar actividades contratables, claras y
útiles para solicitar o comparar cotizaciones.

OBJETIVO
Genera un presupuesto inicial consistente, estructurado por PARTIDA,
SUBPARTIDA y ACTIVIDAD. Debe servir como base para un archivo Excel que después
será revisado y corregido por la empresa.

DATOS DEL PROYECTO
Nombre: {project_data['name']}
Tipo: {project_data['project_type']}
Ubicación: {project_data['location'] or 'No indicada'}
Modo de dimensiones: {project_data['dimension_mode']}
Dimensiones / referencias: {project_data['dimensions_text']}

DESCRIPCIÓN GENERAL DE LOS TRABAJOS
{project_data['description']}

TEXTO GUÍA / CRITERIOS ADICIONALES
{project_data['guide_text'] or 'Sin texto guía adicional.'}

PARÁMETROS COMERCIALES
Indirectos: {params['indirect_pct']:.2f}%
Utilidad: {params['profit_pct']:.2f}%
IVA: {params['iva_pct']:.2f}%
Desperdicio de referencia: {params['waste_pct']:.2f}%

REGLAS DE GENERACIÓN
1. Trabaja con actividades generales subcontratables. No desgloses materiales,
   cuadrillas, rendimientos o herramienta salvo que sea indispensable para
   describir correctamente el alcance.
2. El costo_unitario_estimado representa el COSTO PARA LA EMPRESA del servicio
   completo del subcontratista. NO incluye indirectos corporativos, utilidad ni
   IVA; esos conceptos se calculan posteriormente en Python.
3. Los costos se expresan en MXN y deben representar una referencia razonable
   para {project_data['location'] or 'México'} en {year}, con nivel comercial
   medio-medio salvo que el texto guía indique otro nivel.
4. No aparentes precisión inexistente. En actividades poco comunes, especializadas
   o con alta variación de proveedor, utiliza una estimación prudente, marca
   requiere_cotizacion=True y asigna confianza de precio Baja.
5. Si la actividad es más apropiada como paquete integral, usa LOTE, PZA, JGO o
   una unidad equivalente en lugar de inventar cuantificaciones detalladas.
6. Solo calcula M2, ML, M3 u otras cantidades cuando los datos proporcionados lo
   permitan razonablemente. Si las dimensiones son variables, utiliza el texto
   descriptivo y evita asumir dimensiones no indicadas.
7. El desperdicio es una referencia técnica. En un costo integrado de
   subcontratista se considera dentro de la estimación cuando aplique; NO lo
   devuelvas como una partida separada ni agregues un porcentaje global.
8. Agrupa las actividades en partidas y subpartidas coherentes. Ejemplos de
   partidas: PREPARACIÓN Y DEMOLICIONES, ALBAÑILERÍA, ACABADOS, CARPINTERÍA,
   HERRERÍA Y CANCELERÍA, INSTALACIONES, MOBILIARIO, LIMPIEZA Y ENTREGA.
9. Incluye trámites o licencias únicamente cuando el alcance realmente los haga
   previsibles o el usuario los solicite.
10. No dupliques actividades. No agregues trabajos que no se desprendan del
    alcance salvo complementos técnicos indispensables, y en ese caso indícalo
    claramente en consideraciones.
11. Para cada actividad explica de forma breve el criterio de cantidad y el
    fundamento de inclusión. No expongas cadenas de pensamiento ni razonamiento
    interno.
12. Los datos faltantes deben concentrarse en datos_faltantes, pero no deben
    impedir generar un presupuesto inicial cuando sea posible trabajar con LOTE
    o con una estimación razonable.
13. La descripción técnica debe ser suficientemente completa para poder copiarse
    a una plataforma de presupuestación o solicitar una cotización a proveedor.
14. No calcules indirectos, utilidad, venta, beneficio, margen ni IVA. Python los
    calculará de forma determinista.
"""

    modelos = []
    for model in [
        model_name,
        "gemini-3.6-flash",
        "gemini-3.5-flash",
        "gemini-3.5-flash-lite",
    ]:
        if model and model not in modelos:
            modelos.append(model)

    last_error = None
    for model in modelos:
        try:
            response = client.models.generate_content(
                model=model,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=PresupuestoIA,
                ),
            )
            if not response.text:
                raise RuntimeError(f"Gemini ({model}) devolvió una respuesta vacía.")
            return PresupuestoIA.model_validate_json(response.text)
        except Exception as exc:
            last_error = exc
            msg = str(exc).lower()
            model_error = (
                "404" in msg
                or "not_found" in msg
                or "no longer available" in msg
                or ("model" in msg and "not available" in msg)
            )
            if not model_error:
                raise

    raise RuntimeError(f"No fue posible usar un modelo Gemini disponible. Último error: {last_error}")


# =========================================================
# MOTOR DE PRECIOS
# =========================================================


def buscar_precio_interno(db: Database, actividad: ActividadIA) -> dict | None:
    candidatos = db.price_candidates(actividad.unidad)
    best = None
    best_score = 0.0

    for row in candidatos:
        score = score_similitud(actividad.descripcion_tecnica, row["description"])
        if score > best_score:
            best_score = score
            best = row

    if best is None or best_score < 0.82:
        return None

    original_source = (best.get("source") or "").upper()
    original_status = (best.get("status") or "").upper()

    if original_status in {"VALIDADO", "COSTO_REAL", "COTIZADO_PROVEEDOR"}:
        source = "BASE_INTERNA"
        confidence = "Alta"
    elif original_source == "IA_ESTIMADO" or original_status == "ESTIMADO_IA":
        source = "HISTORICO_IA"
        confidence = "Media" if best_score >= 0.9 else "Baja"
    else:
        source = "BASE_INTERNA"
        confidence = best.get("confidence") or "Media"

    return {
        "concept_id": best["concept_id"],
        "unit_cost": float(best["unit_cost"]),
        "source": source,
        "source_detail": (
            f"Coincidencia {best_score:.0%} con: {best['description']} "
            f"| origen previo: {best.get('source') or 'sin dato'}"
        ),
        "status": original_status or "HISTORICO",
        "confidence": confidence,
        "match_score": best_score,
    }


def buscar_precio_externo(actividad: ActividadIA, project_data: dict) -> dict | None:
    """
    Punto de integración para un catálogo externo estructurado.

    En una versión posterior puede conectarse aquí un catálogo de referencia
    institucional, un export de OPUS u otra base aprobada por la empresa.
    """
    return None


def resolver_items(
    db: Database,
    result: PresupuestoIA,
    project_data: dict,
    params: dict,
) -> list[dict]:
    items = []

    for idx, act in enumerate(result.actividades, start=1):
        internal = buscar_precio_interno(db, act)
        external = None if internal else buscar_precio_externo(act, project_data)

        if internal:
            unit_cost = internal["unit_cost"]
            concept_id = internal["concept_id"]
            source = internal["source"]
            source_detail = internal["source_detail"]
            price_status = internal["status"]
            price_confidence = internal["confidence"]
        elif external:
            unit_cost = float(external["unit_cost"])
            concept_id = None
            source = "REFERENCIA_EXTERNA"
            source_detail = external.get("source_detail", "Catálogo externo")
            price_status = "REFERENCIA_EXTERNA"
            price_confidence = external.get("confidence", "Media")
        else:
            unit_cost = float(act.costo_unitario_estimado)
            concept_id = None
            source = "IA_ESTIMADO"
            source_detail = "Estimación inicial de Gemini; requiere validación comercial."
            price_status = "ESTIMADO_IA"
            price_confidence = act.nivel_confianza_precio

        quantity = max(float(act.cantidad), 0.0)
        indirect_unit = unit_cost * params["indirect_pct"] / 100.0
        profit_unit = (unit_cost + indirect_unit) * params["profit_pct"] / 100.0
        sale_unit = unit_cost + indirect_unit + profit_unit
        direct_amount = quantity * unit_cost
        sale_amount = quantity * sale_unit
        benefit_amount = sale_amount - direct_amount
        sale_margin_pct = (benefit_amount / sale_amount * 100.0) if sale_amount else 0.0

        fallback = f"CON-{idx:03d}"
        code = limpiar_codigo(act.codigo_sugerido, fallback)

        considerations = act.consideraciones.strip()
        if act.requiere_cotizacion:
            considerations = (considerations + " | " if considerations else "") + "Requiere cotización de proveedor."

        items.append(
            {
                "concept_id": concept_id,
                "category": act.partida.strip().upper(),
                "subcategory": act.subpartida.strip(),
                "code": code,
                "description": act.descripcion_tecnica.strip(),
                "unit": act.unidad.strip().upper(),
                "quantity": quantity,
                "unit_cost": unit_cost,
                "direct_amount": direct_amount,
                "unit_indirect": indirect_unit,
                "unit_profit": profit_unit,
                "unit_sale": sale_unit,
                "sale_amount": sale_amount,
                "benefit_amount": benefit_amount,
                "sale_margin_pct": sale_margin_pct,
                "price_source": source,
                "price_source_detail": source_detail,
                "price_status": price_status,
                "price_confidence": price_confidence,
                "quantity_confidence": act.nivel_confianza_cantidad,
                "quantity_criterion": act.criterio_cantidad.strip(),
                "inclusion_basis": act.fundamento_inclusion.strip(),
                "considerations": considerations,
            }
        )

    return items


def calcular_financieros(items: list[dict], params: dict) -> dict:
    direct_cost = sum(x["direct_amount"] for x in items)
    indirect_cost = direct_cost * params["indirect_pct"] / 100.0
    profit = (direct_cost + indirect_cost) * params["profit_pct"] / 100.0
    sale_before_tax = direct_cost + indirect_cost + profit
    iva_amount = sale_before_tax * params["iva_pct"] / 100.0
    total = sale_before_tax + iva_amount

    return {
        "direct_cost": direct_cost,
        "indirect_cost": indirect_cost,
        "profit": profit,
        "sale_before_tax": sale_before_tax,
        "iva_amount": iva_amount,
        "total": total,
    }


# =========================================================
# EXCEL
# =========================================================


def crear_excel(
    project_code: str,
    project_data: dict,
    result: PresupuestoIA,
    items: list[dict],
    params: dict,
) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    navy = "1F4E78"
    dark = "203864"
    gray = "D9E1F2"
    light = "EEF3F8"
    green = "E2F0D9"
    orange = "FCE4D6"
    white = "FFFFFF"
    border_side = Side(style="thin", color="D9E0E7")
    border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)

    # -----------------------------------------------------
    # FACTORES Y PARÁMETROS
    # -----------------------------------------------------
    wp = wb.create_sheet("Factores y Parámetros")
    wp.merge_cells("A2:C2")
    wp["A2"] = "Factores de Sobrecosto Corporativos"
    wp["A2"].font = Font(bold=True, size=14, color=white)
    wp["A2"].fill = PatternFill("solid", fgColor=dark)
    wp["A2"].alignment = Alignment(horizontal="center")

    wp["A4"] = "Concepto de Factor"
    wp["B4"] = "Nomenclatura"
    wp["C4"] = "Porcentaje"
    for cell in wp[4]:
        cell.font = Font(bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.border = border

    factors = [
        ("Costos Indirectos", "IND", params["indirect_pct"] / 100),
        ("Utilidad de la Empresa", "UTIL", params["profit_pct"] / 100),
        ("I.V.A.", "IVA", params["iva_pct"] / 100),
        ("Desperdicio de referencia", "DESP", params["waste_pct"] / 100),
    ]
    for r, row in enumerate(factors, start=5):
        for c, val in enumerate(row, start=1):
            wp.cell(r, c, val).border = border
        wp.cell(r, 3).number_format = "0.00%"

    wp.column_dimensions["A"].width = 42
    wp.column_dimensions["B"].width = 18
    wp.column_dimensions["C"].width = 18

    # -----------------------------------------------------
    # DESGLOSE DETALLADO
    # -----------------------------------------------------
    wd = wb.create_sheet("Desglose Detallado")
    wd.merge_cells("A2:L2")
    wd["A2"] = f"PRESUPUESTO BASE DE OBRA - {project_data['name'].upper()}"
    wd["A2"].font = Font(bold=True, size=15, color=white)
    wd["A2"].fill = PatternFill("solid", fgColor=dark)
    wd["A2"].alignment = Alignment(horizontal="center")

    wd.merge_cells("A3:L3")
    wd["A3"] = (
        f"Código: {project_code} | Ubicación: {project_data['location'] or 'No indicada'} | "
        f"Tipo: {project_data['project_type']}"
    )
    wd["A3"].font = Font(italic=True, color="4F5B66")

    headers = [
        "Partida",
        "Subpartida",
        "Código",
        "Descripción Técnica del Concepto",
        "Unidad",
        "Cantidad",
        "Costo Unitario ($)",
        "Costo Directo ($)",
        "Indirecto Unit. ($)",
        "Utilidad Unit. ($)",
        "P.U. Venta ($)",
        "Importe Venta ($)",
    ]
    header_row = 5
    for c, h in enumerate(headers, start=1):
        cell = wd.cell(header_row, c, h)
        cell.font = Font(bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    data_rows = []
    category_subtotals = {}
    row = 6
    last_category = None

    for item in items:
        category = item["category"]
        if category != last_category:
            if last_category is not None:
                row += 1
            wd.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
            c = wd.cell(row, 1, category)
            c.font = Font(bold=True, color=white)
            c.fill = PatternFill("solid", fgColor=dark)
            c.alignment = Alignment(vertical="center")
            category_subtotals[category] = []
            row += 1
            last_category = category

        r = row
        data_rows.append(r)
        category_subtotals[category].append(r)

        values = [
            category,
            item["subcategory"],
            item["code"],
            item["description"],
            item["unit"],
            item["quantity"],
            item["unit_cost"],
        ]
        for c, val in enumerate(values, start=1):
            cell = wd.cell(r, c, val)
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)

        # Fórmulas editables: al corregir cantidad o costo en Excel, el resto se actualiza.
        wd.cell(r, 8, f"=F{r}*G{r}")
        wd.cell(r, 9, f"=G{r}*'Factores y Parámetros'!$C$5")
        wd.cell(r, 10, f"=(G{r}+I{r})*'Factores y Parámetros'!$C$6")
        wd.cell(r, 11, f"=G{r}+I{r}+J{r}")
        wd.cell(r, 12, f"=F{r}*K{r}")
        for c in range(8, 13):
            wd.cell(r, c).border = border
            wd.cell(r, c).alignment = Alignment(vertical="top")

        wd.cell(r, 6).number_format = "0.000"
        for c in range(7, 13):
            wd.cell(r, c).number_format = '$#,##0.00'

        row += 1

    # Subtotales por partida.
    subtotal_cells = {}
    for category, rows in category_subtotals.items():
        subtotal_row = row
        wd.merge_cells(start_row=subtotal_row, start_column=1, end_row=subtotal_row, end_column=10)
        wd.cell(subtotal_row, 1, f"SUBTOTAL {category}")
        wd.cell(subtotal_row, 1).font = Font(bold=True)
        wd.cell(subtotal_row, 1).fill = PatternFill("solid", fgColor=gray)
        wd.cell(subtotal_row, 11, "Venta")
        wd.cell(subtotal_row, 11).font = Font(bold=True)
        wd.cell(subtotal_row, 11).fill = PatternFill("solid", fgColor=gray)
        formula = "+".join(f"L{x}" for x in rows) if rows else "0"
        wd.cell(subtotal_row, 12, f"={formula}")
        wd.cell(subtotal_row, 12).number_format = '$#,##0.00'
        wd.cell(subtotal_row, 12).font = Font(bold=True)
        wd.cell(subtotal_row, 12).fill = PatternFill("solid", fgColor=gray)
        subtotal_cells[category] = subtotal_row
        row += 1

    row += 1
    direct_formula = "+".join(f"H{x}" for x in data_rows) if data_rows else "0"
    sale_formula = "+".join(f"L{x}" for x in data_rows) if data_rows else "0"

    summary_rows = [
        ("Costo directo", f"={direct_formula}"),
        ("Indirectos", f"=L{row}*'Factores y Parámetros'!$C$5"),
        ("Utilidad", f"=(L{row}+L{row+1})*'Factores y Parámetros'!$C$6"),
        ("Venta antes de IVA", f"={sale_formula}"),
        ("IVA", f"=L{row+3}*'Factores y Parámetros'!$C$7"),
        ("TOTAL GENERAL", f"=L{row+3}+L{row+4}"),
    ]

    financial_summary_rows = {}
    for idx, (label, formula) in enumerate(summary_rows):
        rr = row + idx
        financial_summary_rows[label] = rr
        wd.cell(rr, 10, label)
        wd.cell(rr, 10).font = Font(bold=True)
        wd.cell(rr, 11, "")
        wd.cell(rr, 12, formula)
        wd.cell(rr, 12).number_format = '$#,##0.00'
        wd.cell(rr, 10).fill = PatternFill("solid", fgColor=green if label == "TOTAL GENERAL" else light)
        wd.cell(rr, 11).fill = PatternFill("solid", fgColor=green if label == "TOTAL GENERAL" else light)
        wd.cell(rr, 12).fill = PatternFill("solid", fgColor=green if label == "TOTAL GENERAL" else light)
        if label == "TOTAL GENERAL":
            wd.cell(rr, 12).font = Font(bold=True, size=12)

    widths = {
        "A": 24,
        "B": 20,
        "C": 14,
        "D": 72,
        "E": 10,
        "F": 11,
        "G": 17,
        "H": 17,
        "I": 17,
        "J": 17,
        "K": 17,
        "L": 18,
    }
    for col, width in widths.items():
        wd.column_dimensions[col].width = width
    wd.freeze_panes = "A6"

    # -----------------------------------------------------
    # RESUMEN EJECUTIVO
    # -----------------------------------------------------
    wr = wb.create_sheet("Resumen Ejecutivo", 0)
    wr.merge_cells("B2:H2")
    wr["B2"] = f"RESUMEN DE PRESUPUESTO - {project_data['name'].upper()}"
    wr["B2"].font = Font(bold=True, size=16, color=white)
    wr["B2"].fill = PatternFill("solid", fgColor=dark)
    wr["B2"].alignment = Alignment(horizontal="center")

    wr.merge_cells("B3:H3")
    wr["B3"] = f"{project_code} | {project_data['location'] or 'Ubicación no indicada'} | {project_data['project_type']}"
    wr["B3"].font = Font(italic=True, color="4F5B66")
    wr["B3"].alignment = Alignment(horizontal="center")

    wr["B5"] = "Partida"
    wr["C5"] = "Descripción"
    wr["D5"] = "Monto de Venta (MXN)"
    for cell in wr[5][1:4]:
        cell.font = Font(bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.border = border

    rr = 6
    for n, (category, subtotal_row) in enumerate(subtotal_cells.items(), start=1):
        wr.cell(rr, 2, f"{n}.0")
        wr.cell(rr, 3, category)
        wr.cell(rr, 4, f"='Desglose Detallado'!L{subtotal_row}")
        wr.cell(rr, 4).number_format = '$#,##0.00'
        for c in range(2, 5):
            wr.cell(rr, c).border = border
        rr += 1

    rr += 1
    wr.cell(rr, 3, "VENTA ANTES DE IVA")
    wr.cell(rr, 4, f"='Desglose Detallado'!L{financial_summary_rows['Venta antes de IVA']}")
    rr += 1
    wr.cell(rr, 3, "IVA")
    wr.cell(rr, 4, f"='Desglose Detallado'!L{financial_summary_rows['IVA']}")
    rr += 1
    wr.cell(rr, 3, "TOTAL DE INVERSIÓN")
    wr.cell(rr, 4, f"='Desglose Detallado'!L{financial_summary_rows['TOTAL GENERAL']}")
    for row_i in range(rr - 2, rr + 1):
        wr.cell(row_i, 3).font = Font(bold=True)
        wr.cell(row_i, 4).font = Font(bold=True)
        wr.cell(row_i, 4).number_format = '$#,##0.00'
        fill = green if row_i == rr else light
        wr.cell(row_i, 3).fill = PatternFill("solid", fgColor=fill)
        wr.cell(row_i, 4).fill = PatternFill("solid", fgColor=fill)

    wr.merge_cells(start_row=rr + 3, start_column=2, end_row=rr + 3, end_column=8)
    wr.cell(rr + 3, 2, "Alcance resumido")
    wr.cell(rr + 3, 2).font = Font(bold=True, color=white)
    wr.cell(rr + 3, 2).fill = PatternFill("solid", fgColor=navy)
    wr.merge_cells(start_row=rr + 4, start_column=2, end_row=rr + 6, end_column=8)
    wr.cell(rr + 4, 2, result.alcance_resumido)
    wr.cell(rr + 4, 2).alignment = Alignment(wrap_text=True, vertical="top")

    wr.column_dimensions["B"].width = 15
    wr.column_dimensions["C"].width = 58
    wr.column_dimensions["D"].width = 22
    for col in ["E", "F", "G", "H"]:
        wr.column_dimensions[col].width = 12

    # -----------------------------------------------------
    # TRAZABILIDAD INTERNA
    # -----------------------------------------------------
    wt = wb.create_sheet("Trazabilidad")
    wt.merge_cells("A1:J1")
    wt["A1"] = "Trazabilidad de conceptos y precios"
    wt["A1"].font = Font(bold=True, size=14, color=white)
    wt["A1"].fill = PatternFill("solid", fgColor=dark)

    trace_headers = [
        "Código",
        "Concepto",
        "Unidad",
        "Cantidad",
        "Costo unitario",
        "Fuente precio",
        "Confianza precio",
        "Criterio cantidad",
        "Fundamento",
        "Consideraciones",
    ]
    for c, h in enumerate(trace_headers, 1):
        cell = wt.cell(3, c, h)
        cell.font = Font(bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.border = border
        cell.alignment = Alignment(wrap_text=True)

    for r, item in enumerate(items, start=4):
        vals = [
            item["code"],
            item["description"],
            item["unit"],
            item["quantity"],
            item["unit_cost"],
            f"{item['price_source']} - {item['price_source_detail']}",
            item["price_confidence"],
            item["quantity_criterion"],
            item["inclusion_basis"],
            item["considerations"],
        ]
        for c, val in enumerate(vals, 1):
            cell = wt.cell(r, c, val)
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        wt.cell(r, 5).number_format = '$#,##0.00'
        if item["price_source"] in {"IA_ESTIMADO", "HISTORICO_IA"}:
            wt.cell(r, 6).fill = PatternFill("solid", fgColor=orange)

    trace_widths = [14, 70, 10, 11, 16, 48, 17, 50, 50, 55]
    for idx, width in enumerate(trace_widths, 1):
        wt.column_dimensions[get_column_letter(idx)].width = width
    wt.freeze_panes = "A4"

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


# =========================================================
# TXT PARA CAPTURA MANUAL EN PLATAFORMA
# =========================================================


def crear_txt(
    project_code: str,
    project_data: dict,
    result: PresupuestoIA,
    items: list[dict],
    params: dict,
    financials: dict,
) -> bytes:
    lines = []
    lines.append(f"PROYECTO: {project_data['name']}")
    lines.append(f"CÓDIGO: {project_code}")
    lines.append(f"TIPO: {project_data['project_type']}")
    lines.append(f"UBICACIÓN: {project_data['location'] or 'No indicada'}")
    lines.append("")
    lines.append("ACTIVIDADES PARA CAPTURA MANUAL")
    lines.append("=" * 72)

    current_category = None
    for item in items:
        if item["category"] != current_category:
            current_category = item["category"]
            lines.append("")
            lines.append(current_category)
            lines.append("-" * len(current_category))

        lines.append("")
        lines.append(f"SUBPARTIDA: {item['subcategory']}")
        lines.append(f"ACTIVIDAD: {item['description']}")
        lines.append(f"CANTIDAD: {item['quantity']:.3f}")
        lines.append(f"UNIDAD: {item['unit']}")
        lines.append(f"MARGEN DE VENTA (%): {item['sale_margin_pct']:.2f}")
        lines.append(f"BENEFICIO ($): {item['benefit_amount']:.2f}")
        lines.append(f"COSTO ($): {item['unit_cost']:.2f}")
        lines.append(f"VENTA ($): {item['unit_sale']:.2f}")
        lines.append(f"COSTO TOTAL ACTIVIDAD ($): {item['direct_amount']:.2f}")
        lines.append(f"VENTA TOTAL ACTIVIDAD ($): {item['sale_amount']:.2f}")

    lines.append("")
    lines.append("=" * 72)
    lines.append("RESUMEN")
    lines.append(f"COSTO DIRECTO: {financials['direct_cost']:.2f}")
    lines.append(f"INDIRECTOS ({params['indirect_pct']:.2f}%): {financials['indirect_cost']:.2f}")
    lines.append(f"UTILIDAD ({params['profit_pct']:.2f}%): {financials['profit']:.2f}")
    lines.append(f"VENTA ANTES DE IVA: {financials['sale_before_tax']:.2f}")
    lines.append(f"IVA ({params['iva_pct']:.2f}%): {financials['iva_amount']:.2f}")
    lines.append(f"TOTAL: {financials['total']:.2f}")

    if result.consideraciones_generales:
        lines.append("")
        lines.append("CONSIDERACIONES")
        for x in result.consideraciones_generales:
            lines.append(f"- {x}")

    return "\n".join(lines).encode("utf-8-sig")


def crear_paquete_zip(project_code: str, excel_bytes: bytes, txt_bytes: bytes) -> bytes:
    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        base = project_code
        zf.writestr(f"{base}/{project_code}_Presupuesto.xlsx", excel_bytes)
        zf.writestr(f"{base}/{project_code}_Captura_Plataforma.txt", txt_bytes)
    out.seek(0)
    return out.getvalue()


# =========================================================
# DATAFRAMES DE PRESENTACIÓN
# =========================================================


def dataframe_resumen(items: list[dict]) -> pd.DataFrame:
    rows = []
    for x in items:
        rows.append(
            {
                "Partida": x["category"],
                "Subpartida": x["subcategory"],
                "Concepto": x["description"],
                "Unidad": x["unit"],
                "Cantidad": x["quantity"],
                "Costo unitario": x["unit_cost"],
                "Venta unitaria": x["unit_sale"],
                "Importe venta": x["sale_amount"],
                "Fuente": x["price_source"],
                "Confianza": x["price_confidence"],
            }
        )
    return pd.DataFrame(rows)



# =========================================================
# API DE GEMINI
# =========================================================


def get_api_key_runtime() -> str | None:
    """La clave de empresa vive únicamente en Streamlit Secrets."""
    value = get_secret("GEMINI_API_KEY")
    return str(value).strip() if value else None


# =========================================================
# PANEL DE ADMINISTRACIÓN DE BASE INTERNA
# =========================================================


def _df_or_empty(rows: list[dict], columns: list[str] | None = None) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame(columns=columns or [])
    return pd.DataFrame(rows)


def render_admin_database(db: Database):
    st.header("Base interna")
    st.caption("Consulta y administración del catálogo, historial de precios, proyectos y presupuestos.")

    if db.persistent:
        st.success("Base persistente PostgreSQL activa.")

    if not db.persistent:
        st.warning(
            "La app está usando SQLite local. En Streamlit Community Cloud este archivo no debe considerarse "
            "almacenamiento empresarial permanente. Para producción configure DATABASE_URL con PostgreSQL."
        )

    try:
        stats = db.stats()
        m1, m2, m3 = st.columns(3)
        m1.metric("Proyectos", stats["projects"])
        m2.metric("Presupuestos", stats["budgets"])
        m3.metric("Conceptos", stats["concepts"])
    except Exception as exc:
        st.error(f"No fue posible consultar la base: {exc}")
        return

    tab_concepts, tab_projects, tab_budgets, tab_export = st.tabs(
        ["Conceptos y precios", "Proyectos", "Presupuestos", "Exportar datos"]
    )

    # -----------------------------------------------------
    # Conceptos y precios
    # -----------------------------------------------------
    with tab_concepts:
        st.subheader("Catálogo interno")
        search = st.text_input(
            "Buscar concepto",
            placeholder="Descripción, código, partida o subpartida",
            key="admin_concept_search",
        )
        concepts = db.list_concepts(search)

        if concepts:
            display_rows = []
            for r in concepts:
                display_rows.append({
                    "Código": r.get("code"),
                    "Partida": r.get("category"),
                    "Subpartida": r.get("subcategory"),
                    "Concepto": r.get("description"),
                    "Unidad": r.get("unit"),
                    "Último costo": r.get("latest_cost"),
                    "Fuente": r.get("latest_source"),
                    "Estado": r.get("latest_status"),
                    "Usos": r.get("usage_count", 0),
                })
            st.dataframe(
                pd.DataFrame(display_rows),
                use_container_width=True,
                hide_index=True,
                column_config={"Último costo": st.column_config.NumberColumn(format="$ %.2f")},
            )

            concept_map = {
                r["id"]: f"{r.get('code') or 'SIN-COD'} | {r.get('description') or ''}"
                for r in concepts
            }
            selected_concept_id = st.selectbox(
                "Seleccionar concepto",
                options=list(concept_map.keys()),
                format_func=lambda x: concept_map[x],
                key="admin_selected_concept",
            )
            concept = db.get_concept(selected_concept_id)

            if concept:
                st.markdown("#### Editar concepto")
                with st.form("admin_edit_concept"):
                    e1, e2 = st.columns(2)
                    with e1:
                        c_code = st.text_input("Código", value=concept.get("code") or "")
                        c_category = st.text_input("Partida", value=concept.get("category") or "")
                        c_subcategory = st.text_input("Subpartida", value=concept.get("subcategory") or "")
                    with e2:
                        c_unit = st.text_input("Unidad", value=concept.get("unit") or "")
                        c_description = st.text_area(
                            "Descripción",
                            value=concept.get("description") or "",
                            height=130,
                        )
                    update_concept_btn = st.form_submit_button("Guardar cambios")

                if update_concept_btn:
                    if not c_description.strip() or not c_unit.strip():
                        st.error("Descripción y unidad son obligatorias.")
                    else:
                        db.update_concept(
                            selected_concept_id, c_code, c_category, c_subcategory,
                            c_description, c_unit,
                        )
                        st.success("Concepto actualizado.")
                        st.rerun()

                st.markdown("#### Historial de precios")
                prices = db.list_prices(selected_concept_id)
                if prices:
                    price_df = pd.DataFrame([
                        {
                            "ID": p["id"],
                            "Costo": p["unit_cost"],
                            "Fuente": p["source"],
                            "Detalle": p.get("source_detail"),
                            "Estado": p["status"],
                            "Confianza": p.get("confidence"),
                            "Fecha": p["created_at"],
                        }
                        for p in prices
                    ])
                    st.dataframe(
                        price_df,
                        use_container_width=True,
                        hide_index=True,
                        column_config={"Costo": st.column_config.NumberColumn(format="$ %.2f")},
                    )
                else:
                    st.caption("El concepto todavía no tiene historial de precios.")

                with st.form("admin_add_price"):
                    p1, p2, p3 = st.columns(3)
                    with p1:
                        new_cost = st.number_input("Nuevo costo unitario", min_value=0.0, step=100.0)
                        new_source = st.selectbox(
                            "Fuente",
                            ["COTIZACION_PROVEEDOR", "COSTO_REAL", "REFERENCIA_EXTERNA", "IA_ESTIMADO", "MANUAL"],
                        )
                    with p2:
                        new_status = st.selectbox(
                            "Estado",
                            ["VALIDADO", "COTIZADO_PROVEEDOR", "COSTO_REAL", "REFERENCIA_EXTERNA", "ESTIMADO_IA"],
                        )
                        new_confidence = st.selectbox("Confianza", ["Alta", "Media", "Baja"])
                    with p3:
                        new_detail = st.text_area("Detalle / proveedor / referencia", height=105)
                    add_price_btn = st.form_submit_button("Agregar precio")

                if add_price_btn:
                    if new_cost <= 0:
                        st.error("El costo debe ser mayor que cero.")
                    else:
                        db.add_price(
                            selected_concept_id,
                            new_cost,
                            new_source,
                            new_detail,
                            new_status,
                            new_confidence,
                        )
                        st.success("Precio agregado al historial.")
                        st.rerun()

                if prices:
                    with st.expander("Eliminar un precio"):
                        price_map = {
                            p["id"]: f"{formato_moneda(float(p['unit_cost']))} | {p['source']} | {p['created_at']}"
                            for p in prices
                        }
                        price_to_delete = st.selectbox(
                            "Precio",
                            options=list(price_map.keys()),
                            format_func=lambda x: price_map[x],
                            key="admin_price_delete_select",
                        )
                        confirm_price = st.text_input(
                            "Escriba ELIMINAR PRECIO",
                            key="admin_confirm_delete_price",
                        )
                        if st.button("Eliminar precio seleccionado", key="admin_delete_price_btn"):
                            if confirm_price.strip().upper() != "ELIMINAR PRECIO":
                                st.error("Confirmación incorrecta.")
                            else:
                                db.delete_price(price_to_delete)
                                st.success("Precio eliminado.")
                                st.rerun()

                with st.expander("Eliminar concepto"):
                    usage = db.concept_usage(selected_concept_id)
                    st.write(
                        f"Usos en presupuestos: {usage['budget_items']} | Registros de precio: {usage['prices']}"
                    )
                    if usage["budget_items"]:
                        st.warning(
                            "El concepto ya fue utilizado en presupuestos. Al eliminarlo, los presupuestos históricos "
                            "conservarán sus datos, pero perderán el vínculo al concepto y se eliminará su historial de precios."
                        )
                    confirm_concept = st.text_input(
                        "Escriba ELIMINAR CONCEPTO",
                        key="admin_confirm_delete_concept",
                    )
                    if st.button("Eliminar concepto", key="admin_delete_concept_btn"):
                        if confirm_concept.strip().upper() != "ELIMINAR CONCEPTO":
                            st.error("Confirmación incorrecta.")
                        else:
                            db.delete_concept(selected_concept_id)
                            st.success("Concepto eliminado.")
                            st.rerun()
        else:
            st.info("No se encontraron conceptos.")

        st.divider()
        st.markdown("#### Alta manual de concepto")
        with st.form("admin_create_concept"):
            a1, a2 = st.columns(2)
            with a1:
                a_code = st.text_input("Código sugerido", value="MAN-001")
                a_category = st.text_input("Partida", key="admin_new_category")
                a_subcategory = st.text_input("Subpartida", key="admin_new_subcategory")
            with a2:
                a_unit = st.text_input("Unidad", value="LOTE", key="admin_new_unit")
                a_description = st.text_area("Descripción", height=120, key="admin_new_description")
            create_concept_btn = st.form_submit_button("Crear concepto")
        if create_concept_btn:
            if not a_description.strip() or not a_unit.strip():
                st.error("Descripción y unidad son obligatorias.")
            else:
                new_id = db.create_concept(a_code, a_category, a_subcategory, a_description, a_unit)
                st.success("Concepto creado. Puede seleccionarlo para agregar precios.")
                st.rerun()

    # -----------------------------------------------------
    # Proyectos
    # -----------------------------------------------------
    with tab_projects:
        st.subheader("Proyectos registrados")
        projects = db.list_projects()
        if not projects:
            st.info("No hay proyectos registrados.")
        else:
            project_df = pd.DataFrame([
                {
                    "Código": p["code"],
                    "Nombre": p["name"],
                    "Tipo": p["project_type"],
                    "Ubicación": p.get("location"),
                    "Presupuestos": p.get("budget_count", 0),
                    "Último total": p.get("latest_total"),
                    "Fecha": p["created_at"],
                }
                for p in projects
            ])
            st.dataframe(
                project_df,
                use_container_width=True,
                hide_index=True,
                column_config={"Último total": st.column_config.NumberColumn(format="$ %.2f")},
            )

            project_map = {p["id"]: f"{p['code']} | {p['name']}" for p in projects}
            selected_project_id = st.selectbox(
                "Seleccionar proyecto",
                options=list(project_map.keys()),
                format_func=lambda x: project_map[x],
                key="admin_selected_project",
            )
            project = db.get_project(selected_project_id)

            if project:
                with st.form("admin_edit_project"):
                    p1, p2 = st.columns(2)
                    with p1:
                        pr_name = st.text_input("Nombre", value=project.get("name") or "")
                        pr_type = st.text_input("Tipo", value=project.get("project_type") or "")
                        pr_location = st.text_input("Ubicación", value=project.get("location") or "")
                        pr_activity = st.text_input("Actividad principal", value=project.get("main_activity") or "")
                    with p2:
                        pr_dimensions = st.text_area("Dimensiones", value=project.get("dimensions_text") or "", height=90)
                        pr_description = st.text_area("Descripción", value=project.get("description") or "", height=130)
                        pr_guide = st.text_area("Texto guía", value=project.get("guide_text") or "", height=90)
                    save_project_btn = st.form_submit_button("Guardar cambios")

                if save_project_btn:
                    if not pr_name.strip():
                        st.error("El nombre es obligatorio.")
                    else:
                        db.update_project(
                            selected_project_id,
                            pr_name,
                            pr_type,
                            pr_location,
                            pr_activity,
                            pr_dimensions,
                            pr_description,
                            pr_guide,
                        )
                        st.success("Proyecto actualizado.")
                        st.rerun()

                project_budgets = db.list_budgets(selected_project_id)
                if project_budgets:
                    st.markdown("#### Presupuestos del proyecto")
                    st.dataframe(
                        pd.DataFrame([
                            {
                                "Versión": b["version"],
                                "Estado": b["status"],
                                "Costo directo": b["direct_cost"],
                                "Venta sin IVA": b["sale_before_tax"],
                                "Total": b["total"],
                                "Fecha": b["created_at"],
                            }
                            for b in project_budgets
                        ]),
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "Costo directo": st.column_config.NumberColumn(format="$ %.2f"),
                            "Venta sin IVA": st.column_config.NumberColumn(format="$ %.2f"),
                            "Total": st.column_config.NumberColumn(format="$ %.2f"),
                        },
                    )

                with st.expander("Eliminar proyecto"):
                    st.warning("Esta acción elimina también los presupuestos asociados al proyecto.")
                    confirm_project = st.text_input(
                        "Escriba ELIMINAR PROYECTO",
                        key="admin_confirm_delete_project",
                    )
                    if st.button("Eliminar proyecto", key="admin_delete_project_btn"):
                        if confirm_project.strip().upper() != "ELIMINAR PROYECTO":
                            st.error("Confirmación incorrecta.")
                        else:
                            db.delete_project(selected_project_id)
                            st.success("Proyecto eliminado.")
                            st.rerun()

    # -----------------------------------------------------
    # Presupuestos
    # -----------------------------------------------------
    with tab_budgets:
        st.subheader("Presupuestos registrados")
        budgets = db.list_budgets()
        if not budgets:
            st.info("No hay presupuestos registrados.")
        else:
            budget_df = pd.DataFrame([
                {
                    "Proyecto": b["project_code"],
                    "Nombre": b["project_name"],
                    "Estado": b["status"],
                    "Costo directo": b["direct_cost"],
                    "Indirectos": b["indirect_cost"],
                    "Utilidad": b["profit"],
                    "Venta sin IVA": b["sale_before_tax"],
                    "Total": b["total"],
                    "Fecha": b["created_at"],
                }
                for b in budgets
            ])
            st.dataframe(
                budget_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Costo directo": st.column_config.NumberColumn(format="$ %.2f"),
                    "Indirectos": st.column_config.NumberColumn(format="$ %.2f"),
                    "Utilidad": st.column_config.NumberColumn(format="$ %.2f"),
                    "Venta sin IVA": st.column_config.NumberColumn(format="$ %.2f"),
                    "Total": st.column_config.NumberColumn(format="$ %.2f"),
                },
            )

            budget_map = {
                b["id"]: f"{b['project_code']} | {b['project_name']} | {formato_moneda(float(b['total']))}"
                for b in budgets
            }
            selected_budget_id = st.selectbox(
                "Seleccionar presupuesto",
                options=list(budget_map.keys()),
                format_func=lambda x: budget_map[x],
                key="admin_selected_budget",
            )
            budget = db.get_budget(selected_budget_id)
            if budget:
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Costo directo", formato_moneda(float(budget["direct_cost"])))
                m2.metric("Indirectos", formato_moneda(float(budget["indirect_cost"])))
                m3.metric("Utilidad", formato_moneda(float(budget["profit"])))
                m4.metric("Total", formato_moneda(float(budget["total"])))

                budget_items = db.list_budget_items(selected_budget_id)
                if budget_items:
                    st.dataframe(
                        pd.DataFrame([
                            {
                                "Partida": i["category"],
                                "Subpartida": i["subcategory"],
                                "Código": i["code"],
                                "Concepto": i["description"],
                                "Unidad": i["unit"],
                                "Cantidad": i["quantity"],
                                "Costo unitario": i["unit_cost"],
                                "Venta unitaria": i["unit_sale"],
                                "Venta": i["sale_amount"],
                                "Fuente": i.get("price_source"),
                            }
                            for i in budget_items
                        ]),
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "Costo unitario": st.column_config.NumberColumn(format="$ %.2f"),
                            "Venta unitaria": st.column_config.NumberColumn(format="$ %.2f"),
                            "Venta": st.column_config.NumberColumn(format="$ %.2f"),
                        },
                    )

                with st.expander("Eliminar presupuesto"):
                    st.warning(
                        "Se eliminarán las partidas e historial vinculados a este presupuesto. "
                        "Si es el único presupuesto del proyecto, también se eliminará el proyecto."
                    )
                    confirm_budget = st.text_input(
                        "Escriba ELIMINAR PRESUPUESTO",
                        key="admin_confirm_delete_budget",
                    )
                    if st.button("Eliminar presupuesto", key="admin_delete_budget_btn"):
                        if confirm_budget.strip().upper() != "ELIMINAR PRESUPUESTO":
                            st.error("Confirmación incorrecta.")
                        else:
                            db.delete_budget(selected_budget_id)
                            st.success("Presupuesto eliminado.")
                            st.rerun()

    # -----------------------------------------------------
    # Exportación de base
    # -----------------------------------------------------
    with tab_export:
        st.subheader("Exportar base interna")
        st.caption("Exportación administrativa en CSV. No modifica la base de datos.")
        for table_name, label in [
            ("concepts", "Conceptos"),
            ("price_history", "Historial de precios"),
            ("projects", "Proyectos"),
            ("budgets", "Presupuestos"),
            ("budget_items", "Partidas de presupuestos"),
        ]:
            rows = db.export_table(table_name)
            csv_bytes = pd.DataFrame(rows).to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                f"Descargar {label} CSV",
                data=csv_bytes,
                file_name=f"{table_name}.csv",
                mime="text/csv",
                key=f"download_{table_name}",
                use_container_width=True,
            )


# =========================================================
# AUTENTICACIÓN Y ESTADO DE APLICACIÓN
# =========================================================


authenticator, user_roles = autenticar_usuario()
is_admin = "admin" in user_roles

st.caption("Generación inicial de presupuestos para remodelación e interiorismo.")

database_url = get_secret("DATABASE_URL")
try:
    db, db_error = get_database(database_url)
except Exception as exc:
    st.error("No fue posible conectar con la base PostgreSQL configurada.")
    st.caption(
        "Revise DATABASE_URL en Streamlit Secrets y la cadena de conexión de Supabase. "
        "La aplicación no usará SQLite como respaldo cuando exista una DATABASE_URL."
    )
    st.exception(exc)
    st.stop()

with st.sidebar:
    st.header("Sesión")
    display_name = st.session_state.get("name") or st.session_state.get("username") or "Usuario"
    st.write(str(display_name))
    role_label = "Administrador" if is_admin else "Usuario"
    st.caption(role_label)
    authenticator.logout(
        button_name="Cerrar sesión",
        location="sidebar",
        key="logout_main",
        use_container_width=True,
    )

    st.divider()
    st.header("Navegación")
    sections = ["Generar presupuesto"]
    if is_admin:
        sections.append("Base interna")

    section = st.radio(
        "Sección",
        sections,
        key="main_section",
        label_visibility="collapsed",
    )

    if section == "Generar presupuesto":
        st.divider()
        st.header("Parámetros")

        model_name = st.text_input(
            "Modelo",
            value="gemini-3.6-flash",
            help="Modelo principal. Se prueban alternativas si no está disponible.",
            key="model_name",
        )

        indirect_pct = st.number_input(
            "Indirectos (%)",
            min_value=0.0,
            max_value=100.0,
            value=12.0,
            step=0.5,
            key="indirect_pct",
        )
        profit_pct = st.number_input(
            "Utilidad (%)",
            min_value=0.0,
            max_value=100.0,
            value=10.0,
            step=0.5,
            key="profit_pct",
        )
        iva_pct = st.number_input(
            "IVA (%)",
            min_value=0.0,
            max_value=100.0,
            value=16.0,
            step=1.0,
            key="iva_pct",
        )
        waste_pct = st.number_input(
            "Desperdicio de referencia (%)",
            min_value=0.0,
            max_value=50.0,
            value=5.0,
            step=0.5,
            key="waste_pct",
        )
        st.caption(
            "El desperdicio es una referencia para la estimación cuando aplique; "
            "no se suma de forma global."
        )

        st.divider()
        st.subheader("Estado de base interna")
        try:
            stats = db.stats()
            st.write(f"Proyectos: {stats['projects']}")
            st.write(f"Presupuestos: {stats['budgets']}")
            st.write(f"Conceptos: {stats['concepts']}")
        except Exception:
            st.write("Sin estadísticas disponibles.")

        if db.persistent:
            st.success("PostgreSQL conectado.")
        else:
            st.warning(
                "SQLite local activo. Úselo solo para desarrollo; en Streamlit Community Cloud "
                "configure DATABASE_URL para persistencia."
            )

        if not get_api_key_runtime():
            st.error("Falta GEMINI_API_KEY en Streamlit Secrets.")

# =========================================================
# BASE INTERNA
# =========================================================


if section == "Base interna":
    if not is_admin:
        st.error("No tiene permisos para administrar la base interna.")
        st.stop()
    render_admin_database(db)
    st.stop()


# =========================================================
# FORMULARIO INICIAL
# =========================================================


if "generated" not in st.session_state:
    with st.form("form_proyecto"):
        simulation_mode = st.toggle(
            "Simulación",
            value=False,
            help="Genera Excel, TXT y resultados, pero no guarda proyectos, presupuestos, conceptos ni precios en la base interna.",
            key="simulation_mode",
        )
        if simulation_mode:
            st.info("Modo simulación activo: esta corrida no modificará la base interna.")

        c1, c2 = st.columns(2)

        with c1:
            name = st.text_input(
                "Nombre del proyecto",
                value=st.session_state.get("project_name", ""),
                key="project_name",
            )
            project_type = st.selectbox(
                "Tipo de obra",
                [
                    "Remodelación interior general",
                    "Baño",
                    "Cocina",
                    "Recámara",
                    "Sala / comedor",
                    "Local comercial",
                    "Oficina",
                    "Caseta / acceso",
                    "Otro",
                ],
                key="project_type",
            )
            location = st.text_input(
                "Ubicación",
                placeholder="Ciudad, alcaldía o zona de referencia",
                key="project_location",
            )

        with c2:
            dimension_mode = st.radio(
                "Dimensiones",
                ["Definidas", "Variables"],
                horizontal=True,
                key="dimension_mode",
            )

            if dimension_mode == "Definidas":
                d1, d2, d3 = st.columns(3)
                length = d1.number_input("Largo (m)", min_value=0.0, value=2.0, step=0.10, key="length")
                width = d2.number_input("Ancho (m)", min_value=0.0, value=2.0, step=0.10, key="width")
                height = d3.number_input("Altura (m)", min_value=0.0, value=2.5, step=0.10, key="height")
                dimension_notes = st.text_input(
                    "Referencias adicionales de medida",
                    placeholder="Opcional",
                    key="dimension_notes",
                )
                dimensions_text = (
                    f"Largo {length:.3f} m; ancho {width:.3f} m; altura {height:.3f} m. "
                    f"{dimension_notes}".strip()
                )
            else:
                dimensions_text = st.text_area(
                    "Descripción de dimensiones",
                    placeholder="Indique las medidas por zona, elemento o actividad cuando se conozcan.",
                    height=115,
                    key="variable_dimensions",
                )

        description = st.text_area(
            "Descripción general de trabajos",
            placeholder="Describa las actividades, elementos a retirar, suministrar, instalar, modificar y cualquier condición relevante.",
            height=230,
            key="project_description",
        )

        guide_text = st.text_area(
            "Texto guía",
            placeholder="Opcional: nivel de acabado, criterios de costo, exclusiones, restricciones o instrucciones especiales.",
            height=110,
            key="guide_text",
        )

        generate = st.form_submit_button(
            "Generar presupuesto",
            type="primary",
            use_container_width=True,
        )

    if generate:
        api_key = get_api_key_runtime()
        if not api_key:
            st.error("Falta GEMINI_API_KEY en Streamlit Secrets.")
            st.stop()

        if not name.strip():
            st.error("Ingrese el nombre del proyecto.")
            st.stop()

        if not description.strip():
            st.error("Ingrese la descripción general de los trabajos.")
            st.stop()

        if dimension_mode == "Definidas" and (length <= 0 or width <= 0 or height <= 0):
            st.error("Las dimensiones definidas deben ser mayores que cero.")
            st.stop()

        if dimension_mode == "Variables" and not dimensions_text.strip():
            dimensions_text = "Dimensiones variables; consultar descripción general de trabajos."

        params = {
            "indirect_pct": float(indirect_pct),
            "profit_pct": float(profit_pct),
            "iva_pct": float(iva_pct),
            "waste_pct": float(waste_pct),
        }
        project_data = {
            "name": name.strip(),
            "project_type": project_type,
            "location": location.strip(),
            "dimension_mode": dimension_mode,
            "dimensions_text": dimensions_text.strip(),
            "description": description.strip(),
            "guide_text": guide_text.strip(),
        }

        with st.spinner("Generando presupuesto inicial..."):
            try:
                result = generar_presupuesto_ia(
                    api_key=api_key,
                    model_name=model_name,
                    project_data=project_data,
                    params=params,
                )

                # La base interna puede consultarse en simulación, pero nunca se escribe.
                items = resolver_items(db, result, project_data, params)
                if not items:
                    raise RuntimeError("La IA no generó actividades utilizables.")

                financials = calcular_financieros(items, params)
                base_code = db.next_project_code(
                    project_data["location"],
                    project_data["project_type"],
                )
                project_code = f"SIM-{base_code}" if simulation_mode else base_code

                excel_bytes = crear_excel(
                    project_code=project_code,
                    project_data=project_data,
                    result=result,
                    items=items,
                    params=params,
                )
                txt_bytes = crear_txt(
                    project_code=project_code,
                    project_data=project_data,
                    result=result,
                    items=items,
                    params=params,
                    financials=financials,
                )
                zip_bytes = crear_paquete_zip(project_code, excel_bytes, txt_bytes)

                project_id = None
                budget_id = None
                if not simulation_mode:
                    project_id, budget_id = db.save_generation(
                        project_code=project_code,
                        project_data=project_data,
                        result=result,
                        items=items,
                        params=params,
                        financials=financials,
                    )

                st.session_state["generated"] = {
                    "project_id": project_id,
                    "budget_id": budget_id,
                    "simulation": bool(simulation_mode),
                    "project_code": project_code,
                    "project_data": project_data,
                    "params": params,
                    "result": result.model_dump(),
                    "items": items,
                    "financials": financials,
                    "excel_bytes": excel_bytes,
                    "txt_bytes": txt_bytes,
                    "zip_bytes": zip_bytes,
                }
                st.rerun()

            except Exception as exc:
                st.exception(exc)


# =========================================================
# RESULTADO FINAL - SIN EDICIÓN EN LA APP
# =========================================================


else:
    g = st.session_state["generated"]
    result = PresupuestoIA.model_validate(g["result"])
    items = g["items"]
    financials = g["financials"]

    if g.get("simulation"):
        st.warning("SIMULACIÓN: esta generación no fue guardada en la base interna.")

    st.subheader(g["project_code"])
    st.write(result.alcance_resumido)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Costo directo", formato_moneda(financials["direct_cost"]))
    m2.metric("Indirectos", formato_moneda(financials["indirect_cost"]))
    m3.metric("Utilidad", formato_moneda(financials["profit"]))
    m4.metric("Total con IVA", formato_moneda(financials["total"]))

    st.subheader("Presupuesto generado")
    df = dataframe_resumen(items)
    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Cantidad": st.column_config.NumberColumn(format="%.3f"),
            "Costo unitario": st.column_config.NumberColumn(format="$ %.2f"),
            "Venta unitaria": st.column_config.NumberColumn(format="$ %.2f"),
            "Importe venta": st.column_config.NumberColumn(format="$ %.2f"),
        },
    )

    source_counts = df["Fuente"].value_counts().to_dict() if not df.empty else {}
    source_text = " | ".join(f"{k}: {v}" for k, v in source_counts.items())
    if source_text:
        st.caption(f"Origen de precios: {source_text}")

    if result.consideraciones_generales or result.datos_faltantes:
        with st.expander("Consideraciones"):
            if result.consideraciones_generales:
                st.markdown("**Consideraciones generales**")
                for item in result.consideraciones_generales:
                    st.write(f"- {item}")
            if result.datos_faltantes:
                st.markdown("**Datos no confirmados**")
                for item in result.datos_faltantes:
                    st.write(f"- {item}")

    st.subheader("Archivos")
    st.download_button(
        "Descargar paquete",
        data=g["zip_bytes"],
        file_name=f"{g['project_code']}.zip",
        mime="application/zip",
        use_container_width=True,
        type="primary",
    )

    with st.expander("Descargas individuales"):
        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "Descargar Excel",
                data=g["excel_bytes"],
                file_name=f"{g['project_code']}_Presupuesto.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with c2:
            st.download_button(
                "Descargar TXT",
                data=g["txt_bytes"],
                file_name=f"{g['project_code']}_Captura_Plataforma.txt",
                mime="text/plain",
                use_container_width=True,
            )

    st.divider()
    if st.button("Modificar datos iniciales", use_container_width=True):
        try:
            if not g.get("simulation") and g.get("project_id") and g.get("budget_id"):
                db.delete_generation(g["project_id"], g["budget_id"])
        finally:
            del st.session_state["generated"]
            st.rerun()
