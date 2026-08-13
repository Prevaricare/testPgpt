import os
import re
import hmac
import json
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


def get_delete_key() -> str | None:
    """Clave privada para operaciones destructivas, almacenada en Secrets."""
    value = get_secret("DELETE_KEY")
    if value is None:
        return None
    return str(value).strip()


def validar_clave_borrado(valor: str) -> bool:
    esperada = get_delete_key()
    if not esperada:
        return False
    return hmac.compare_digest(str(valor or "").strip(), esperada)


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


def abreviar_cliente(texto: str, max_len: int = 4) -> str:
    """
    Genera una abreviatura estable para el cliente.
    Si hay varias palabras utiliza sus iniciales; si hay una sola, usa sus
    primeras letras. Ej.: "Desarrollos de la Vega" -> "DDV".
    """
    limpio = normalizar_texto(texto).upper()
    palabras = [
        p for p in limpio.split()
        if p not in {"DE", "DEL", "LA", "LAS", "EL", "LOS", "EN", "Y", "SA", "CV"}
    ]
    if not palabras:
        return "CLI"
    if len(palabras) >= 2:
        iniciales = "".join(p[0] for p in palabras if p)
        return iniciales[:max_len] or "CLI"
    return (palabras[0] + "XXXX")[:3]


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

    def _ensure_column(self, table: str, column: str, sql_type: str):
        """Agrega una columna a una base existente sin destruir información."""
        allowed_tables = {"projects", "budgets", "concepts", "price_history", "budget_items"}
        allowed_columns = {"parent_budget_id", "revision_instruction"}
        if table not in allowed_tables or column not in allowed_columns:
            raise ValueError("Migración de columna no permitida.")

        if self.kind == "postgres":
            exists = self.fetchone(
                """
                SELECT 1 AS ok
                FROM information_schema.columns
                WHERE table_schema = 'public'
                  AND table_name = ?
                  AND column_name = ?
                """,
                (table, column),
            )
        else:
            rows = self.fetchall(f"PRAGMA table_info({table})")
            exists = any(str(r.get("name")) == column for r in rows)

        if not exists:
            self.execute(f"ALTER TABLE {table} ADD COLUMN {column} {sql_type}")

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

        # Migraciones no destructivas para instalaciones creadas con versiones
        # anteriores de la app.
        self._ensure_column("budgets", "parent_budget_id", "TEXT")
        self._ensure_column("budgets", "revision_instruction", "TEXT")

    def stats(self) -> dict:
        return {
            "projects": self.fetchone("SELECT COUNT(*) AS n FROM projects")["n"],
            "budgets": self.fetchone("SELECT COUNT(*) AS n FROM budgets")["n"],
            "concepts": self.fetchone("SELECT COUNT(*) AS n FROM concepts")["n"],
        }

    def get_latest_project_record(self) -> dict | None:
        """Devuelve el último proyecto guardado y su último total conocido."""
        return self.fetchone(
            """
            SELECT p.*,
                   (SELECT COUNT(*) FROM budgets b WHERE b.project_id = p.id) AS budget_count,
                   (SELECT b.total FROM budgets b
                    WHERE b.project_id = p.id
                    ORDER BY b.created_at DESC LIMIT 1) AS latest_total
            FROM projects p
            ORDER BY p.created_at DESC
            LIMIT 1
            """
        )

    def clear_all_data(self):
        """
        Elimina todos los datos empresariales de las tablas de la aplicación.
        Conserva el esquema para que la app siga funcionando inmediatamente.
        """
        # El orden evita conflictos de claves foráneas tanto en PostgreSQL como SQLite.
        for table in [
            "budget_items",
            "price_history",
            "budgets",
            "concepts",
            "projects",
        ]:
            self.execute(f"DELETE FROM {table}")

    def next_project_code(self, client_name: str, location: str) -> str:
        """
        Código corporativo: abreviatura del cliente + ubicación + consecutivo.
        Ejemplo: Desarrollos de la Vega / Farallón -> DDV-FAR-0001.
        """
        prefix = f"{abreviar_cliente(client_name or 'Cliente')}-{abreviar(location or 'Ubicacion')}"
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

    def save_revision(
        self,
        project_id: str,
        parent_budget_id: str,
        result,
        items: list[dict],
        params: dict,
        financials: dict,
        revision_instruction: str,
    ) -> tuple[str, int]:
        """
        Guarda una revisión como NUEVO presupuesto del mismo proyecto.
        El presupuesto anterior se conserva para mantener trazabilidad.
        """
        row = self.fetchone(
            "SELECT COALESCE(MAX(version), 0) AS max_version FROM budgets WHERE project_id = ?",
            (project_id,),
        )
        version = int(row["max_version"] or 0) + 1
        budget_id = str(uuid.uuid4())
        created = ahora_iso()

        self.execute(
            """
            INSERT INTO budgets (
                id, project_id, version, status, indirect_pct, profit_pct,
                iva_pct, waste_pct, direct_cost, indirect_cost, profit,
                sale_before_tax, iva_amount, total, scope_summary, created_at,
                parent_budget_id, revision_instruction
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                budget_id,
                project_id,
                version,
                "REVISION",
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
                parent_budget_id,
                revision_instruction.strip(),
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

        return budget_id, version

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


DATABASE_CACHE_VERSION = "2026-08-12-v7.0"


@st.cache_resource(show_spinner=False)
def get_database(database_url: str | None, cache_version: str):
    """
    Usa PostgreSQL si DATABASE_URL existe; SQLite solo para desarrollo sin URL.

    cache_version forma parte deliberadamente de la clave de caché. Esto evita
    reutilizar una instancia de Database creada con una versión anterior de la
    clase después de actualizar app.py en Streamlit Community Cloud.
    """
    _ = cache_version
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


class OperacionRevisionIA(BaseModel):
    accion: str = Field(description="AGREGAR, MODIFICAR o ELIMINAR")
    codigo_objetivo: str = Field(
        default="",
        description="Código exacto de la actividad actual para MODIFICAR o ELIMINAR"
    )
    actividad: ActividadIA | None = Field(
        default=None,
        description="Actividad completa resultante para AGREGAR o MODIFICAR; null para ELIMINAR"
    )
    recalcular_precio: bool = Field(
        default=True,
        description="True si el cambio altera materialmente el alcance o la base del costo unitario"
    )
    motivo: str = Field(description="Resumen técnico breve del cambio solicitado")


class RevisionPresupuestoIA(BaseModel):
    resumen_revision: str
    actividad_principal_actualizada: str
    alcance_resumido_actualizado: str
    consideraciones_generales_actualizadas: list[str]
    datos_faltantes_actualizados: list[str]
    operaciones: list[OperacionRevisionIA]


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
Cliente: {project_data['name']}
Ubicación: {project_data['location'] or 'No indicada'}
Tipo de obra: {project_data['project_type']}

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
6. Solo calcula M2, ML, M3 u otras cantidades cuando las medidas incluidas en
   la descripción general permitan hacerlo razonablemente. Las dimensiones pueden
   aparecer por planta, zona o elemento. Evita asumir medidas no indicadas.
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


def revisar_presupuesto_ia(
    api_key: str,
    model_name: str,
    project_data: dict,
    params: dict,
    current_result: PresupuestoIA,
    current_items: list[dict],
    revision_request: str,
) -> RevisionPresupuestoIA:
    """
    Segundo modo de uso de Gemini.
    No vuelve a generar el presupuesto completo: devuelve exclusivamente
    operaciones estructuradas sobre las actividades existentes.
    """
    client = genai.Client(api_key=api_key)

    presupuesto_actual = [
        {
            "codigo": x["code"],
            "partida": x["category"],
            "subpartida": x["subcategory"],
            "descripcion": x["description"],
            "unidad": x["unit"],
            "cantidad": x["quantity"],
            "costo_unitario_actual": x["unit_cost"],
            "fuente_precio": x["price_source"],
            "criterio_cantidad": x["quantity_criterion"],
            "consideraciones": x["considerations"],
        }
        for x in current_items
    ]

    prompt = f"""
Actúa como revisor técnico de un presupuesto de remodelación e interiorismo.
NO debes volver a generar el presupuesto desde cero.

Tu trabajo consiste en interpretar UNA SOLICITUD DE CAMBIO ESTRUCTURAL y
devolver solamente las operaciones necesarias para transformar el presupuesto
actual en una nueva versión.

DATOS ORIGINALES DEL PROYECTO — SON INMUTABLES
Cliente: {project_data['name']}
Ubicación: {project_data['location']}
Tipo de obra: {project_data['project_type']}

DESCRIPCIÓN GENERAL ORIGINAL
{project_data['description']}

TEXTO GUÍA / CONSIDERACIONES ORIGINALES
{project_data['guide_text'] or 'Sin texto guía adicional.'}

PRESUPUESTO ACTUAL
{json.dumps(presupuesto_actual, ensure_ascii=False, separators=(',', ':'))}

ALCANCE RESUMIDO ACTUAL
{current_result.alcance_resumido}

SOLICITUD DE CAMBIO
{revision_request}

REGLAS OBLIGATORIAS
1. Modifica EXCLUSIVAMENTE lo que sea necesario para cumplir la solicitud.
2. Todo concepto que no esté relacionado con la solicitud debe permanecer
   exactamente sin cambios; Python conservará esas actividades sin regenerarlas.
3. Utiliza únicamente las acciones AGREGAR, MODIFICAR y ELIMINAR.
4. Para MODIFICAR o ELIMINAR debes indicar codigo_objetivo usando exactamente
   uno de los códigos existentes en PRESUPUESTO ACTUAL.
5. Para AGREGAR y MODIFICAR devuelve actividad completa.
6. Para ELIMINAR, actividad debe ser null.
7. Si una modificación solo mejora redacción o detalle y no cambia la base del
   costo unitario, usa recalcular_precio=False.
8. Si cambia materialmente alcance, especificación, unidad, calidad, dimensiones
   relevantes o naturaleza del servicio, usa recalcular_precio=True.
9. Para actividades nuevas usa recalcular_precio=True.
10. No hagas correcciones menores de precios por iniciativa propia.
11. Conserva la lógica de actividades generales subcontratables; no desarrolles
    APUs de material/mano de obra salvo que sea indispensable.
12. No calcules indirectos, utilidad, venta ni IVA.
13. Devuelve un alcance_resumido_actualizado que refleje la nueva versión.
14. Devuelve las listas completas y actualizadas de consideraciones y datos
    faltantes, no únicamente lo nuevo.
15. No expongas razonamiento interno. En motivo escribe solo la justificación
    técnica breve y verificable de cada operación.
16. Si la solicitud no requiere cambiar una actividad, no generes una operación
    para ella.
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
                    response_schema=RevisionPresupuestoIA,
                ),
            )
            if not response.text:
                raise RuntimeError(f"Gemini ({model}) devolvió una revisión vacía.")
            return RevisionPresupuestoIA.model_validate_json(response.text)
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

    raise RuntimeError(
        f"No fue posible usar un modelo Gemini para la revisión. Último error: {last_error}"
    )


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
    force_new_price_codes: set[str] | None = None,
) -> list[dict]:
    items = []
    force_new_price_codes = {
        str(x).strip().upper() for x in (force_new_price_codes or set())
    }

    for idx, act in enumerate(result.actividades, start=1):
        fallback = f"CON-{idx:03d}"
        requested_code = limpiar_codigo(act.codigo_sugerido, fallback)
        force_new_price = requested_code.upper() in force_new_price_codes

        internal = None if force_new_price else buscar_precio_interno(db, act)
        external = None if internal or force_new_price else buscar_precio_externo(act, project_data)

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

        code = requested_code

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


def recalcular_item_financiero(item: dict, params: dict) -> dict:
    """Recalcula importes de un item sin pedir operaciones matemáticas a Gemini."""
    out = dict(item)
    unit_cost = float(out["unit_cost"])
    quantity = max(float(out["quantity"]), 0.0)

    indirect_unit = unit_cost * params["indirect_pct"] / 100.0
    profit_unit = (unit_cost + indirect_unit) * params["profit_pct"] / 100.0
    sale_unit = unit_cost + indirect_unit + profit_unit
    direct_amount = quantity * unit_cost
    sale_amount = quantity * sale_unit
    benefit_amount = sale_amount - direct_amount
    sale_margin_pct = (benefit_amount / sale_amount * 100.0) if sale_amount else 0.0

    out.update(
        {
            "quantity": quantity,
            "unit_cost": unit_cost,
            "direct_amount": direct_amount,
            "unit_indirect": indirect_unit,
            "unit_profit": profit_unit,
            "unit_sale": sale_unit,
            "sale_amount": sale_amount,
            "benefit_amount": benefit_amount,
            "sale_margin_pct": sale_margin_pct,
        }
    )
    return out


def item_a_actividad(item: dict) -> ActividadIA:
    return ActividadIA(
        partida=item["category"],
        subpartida=item["subcategory"],
        codigo_sugerido=item["code"],
        descripcion_tecnica=item["description"],
        unidad=item["unit"],
        cantidad=float(item["quantity"]),
        costo_unitario_estimado=float(item["unit_cost"]),
        criterio_cantidad=item.get("quantity_criterion") or "Cantidad de la versión vigente.",
        fundamento_inclusion=item.get("inclusion_basis") or "Actividad incluida en el alcance vigente.",
        nivel_confianza_cantidad=item.get("quantity_confidence") or "Media",
        nivel_confianza_precio=item.get("price_confidence") or "Media",
        requiere_cotizacion="requiere cotización" in (item.get("considerations") or "").lower(),
        consideraciones=item.get("considerations") or "",
    )


def aplicar_revision_estructural(
    db: Database,
    current_result: PresupuestoIA,
    current_items: list[dict],
    revision: RevisionPresupuestoIA,
    project_data: dict,
    params: dict,
) -> tuple[PresupuestoIA, list[dict], list[str]]:
    """
    Aplica las operaciones devueltas por Gemini. Las actividades no mencionadas
    se copian literalmente desde la versión anterior.
    """
    if not revision.operaciones:
        raise RuntimeError(
            "Gemini no identificó operaciones estructurales. "
            "Revise que la solicitud describa un cambio importante de alcance."
        )

    operations = []
    activities_to_resolve = []
    force_codes = set()

    for op in revision.operaciones:
        action = (op.accion or "").strip().upper()
        if action not in {"AGREGAR", "MODIFICAR", "ELIMINAR"}:
            raise RuntimeError(f"Acción de revisión no válida: {op.accion}")

        if action in {"AGREGAR", "MODIFICAR"}:
            if op.actividad is None:
                raise RuntimeError(f"La acción {action} no contiene una actividad completa.")
            activities_to_resolve.append(op.actividad)
            if action == "AGREGAR" or op.recalcular_precio:
                force_codes.add(
                    limpiar_codigo(
                        op.actividad.codigo_sugerido,
                        f"REV-{len(activities_to_resolve):03d}",
                    )
                )

        operations.append((action, op))

    resolved_changes = []
    if activities_to_resolve:
        temp_result = PresupuestoIA(
            nombre_proyecto=current_result.nombre_proyecto,
            actividad_principal=revision.actividad_principal_actualizada,
            alcance_resumido=revision.alcance_resumido_actualizado,
            consideraciones_generales=revision.consideraciones_generales_actualizadas,
            datos_faltantes=revision.datos_faltantes_actualizados,
            actividades=activities_to_resolve,
        )
        resolved_changes = resolver_items(
            db,
            temp_result,
            project_data,
            params,
            force_new_price_codes=force_codes,
        )

    items = [dict(x) for x in current_items]
    resolved_index = 0
    change_log = []

    def find_index(code: str) -> int:
        code_n = (code or "").strip().upper()
        for i, item in enumerate(items):
            if str(item.get("code") or "").strip().upper() == code_n:
                return i
        raise RuntimeError(
            f"La revisión hizo referencia al código '{code}', "
            "pero ese código no existe en el presupuesto actual."
        )

    for action, op in operations:
        if action == "ELIMINAR":
            idx = find_index(op.codigo_objetivo)
            removed = items.pop(idx)
            change_log.append(f"ELIMINADO {removed['code']}: {removed['description']}")
            continue

        new_item = dict(resolved_changes[resolved_index])
        resolved_index += 1

        if action == "MODIFICAR":
            idx = find_index(op.codigo_objetivo)
            old_item = items[idx]

            # Si Gemini declara que el cambio no altera la base de costo, se
            # conserva exactamente el precio histórico utilizado y solo se
            # recalculan importes con la nueva cantidad.
            if not op.recalcular_precio:
                for key in [
                    "concept_id",
                    "unit_cost",
                    "price_source",
                    "price_source_detail",
                    "price_status",
                    "price_confidence",
                ]:
                    new_item[key] = old_item.get(key)
                new_item = recalcular_item_financiero(new_item, params)

            # Evita que una modificación cambie el código a uno que ya pertenece
            # a otra actividad.
            proposed = str(new_item["code"]).strip().upper()
            conflicts = {
                str(x.get("code") or "").strip().upper()
                for j, x in enumerate(items)
                if j != idx
            }
            if proposed in conflicts:
                new_item["code"] = old_item["code"]

            items[idx] = new_item
            change_log.append(
                f"MODIFICADO {old_item['code']}: {op.motivo.strip() or 'Cambio de alcance'}"
            )
            continue

        # AGREGAR
        existing_codes = {
            str(x.get("code") or "").strip().upper() for x in items
        }
        base_code = str(new_item["code"]).strip().upper() or "REV"
        candidate = base_code
        counter = 2
        while candidate in existing_codes:
            candidate = f"{base_code}-{counter}"
            counter += 1
        new_item["code"] = candidate

        # Inserta después de la última actividad de la misma partida para
        # mantener el Excel ordenado por bloques.
        insert_at = len(items)
        same_category = [
            i for i, x in enumerate(items)
            if str(x.get("category") or "").strip().upper()
            == str(new_item.get("category") or "").strip().upper()
        ]
        if same_category:
            insert_at = max(same_category) + 1
        items.insert(insert_at, new_item)
        change_log.append(
            f"AGREGADO {new_item['code']}: {new_item['description']}"
        )

    revised_result = PresupuestoIA(
        nombre_proyecto=current_result.nombre_proyecto,
        actividad_principal=revision.actividad_principal_actualizada,
        alcance_resumido=revision.alcance_resumido_actualizado,
        consideraciones_generales=revision.consideraciones_generales_actualizadas,
        datos_faltantes=revision.datos_faltantes_actualizados,
        actividades=[item_a_actividad(x) for x in items],
    )

    return revised_result, items, change_log


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
    version: int = 1,
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
        f"Código: {project_code} | Versión: V{version:02d} | "
        f"Ubicación: {project_data['location'] or 'No indicada'} | Tipo: {project_data['project_type']}"
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
    wr["B3"] = (
        f"{project_code} | V{version:02d} | "
        f"{project_data['location'] or 'Ubicación no indicada'} | {project_data['project_type']}"
    )
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
    version: int = 1,
) -> bytes:
    lines = []
    lines.append(f"CLIENTE: {project_data['name']}")
    lines.append(f"CÓDIGO: {project_code}")
    lines.append(f"VERSIÓN: V{version:02d}")
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
    st.header("Catálogo e historial")
    st.caption(
        "Consulta y administra conceptos, precios históricos, proyectos y presupuestos "
        "sin editar directamente las tablas de la base de datos."
    )

    if db.persistent:
        st.success("Base persistente PostgreSQL conectada.")
    else:
        st.warning(
            "SQLite local activo. Úselo únicamente para desarrollo. "
            "En Streamlit Community Cloud configure DATABASE_URL para persistencia."
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

    st.divider()

    tab_concepts, tab_prices, tab_projects, tab_budgets, tab_maintenance, tab_export = st.tabs(
        ["Conceptos", "Precios", "Proyectos", "Presupuestos", "Mantenimiento", "Exportar"]
    )

    source_labels = {
        "COTIZACION_PROVEEDOR": "Cotización de proveedor",
        "COSTO_REAL": "Costo real de obra",
        "REFERENCIA_EXTERNA": "Referencia externa",
        "IA_ESTIMADO": "Estimación de IA",
        "MANUAL": "Registro manual",
        "BASE_INTERNA": "Base interna",
        "HISTORICO_IA": "Histórico generado por IA",
    }
    status_labels = {
        "VALIDADO": "Validado",
        "COTIZADO_PROVEEDOR": "Cotizado por proveedor",
        "COSTO_REAL": "Costo real",
        "REFERENCIA_EXTERNA": "Referencia externa",
        "ESTIMADO_IA": "Estimado por IA",
    }

    def friendly_source(value):
        return source_labels.get(str(value or "").upper(), value or "Sin fuente")

    def friendly_status(value):
        return status_labels.get(str(value or "").upper(), value or "Sin estado")

    # =====================================================
    # CONCEPTOS
    # =====================================================
    with tab_concepts:
        st.subheader("Catálogo de conceptos")
        st.caption(
            "Busque un concepto, abra su ficha y, si es necesario, modifique sus datos. "
            "Los precios se administran por separado en la pestaña Precios."
        )

        all_concepts = db.list_concepts("", limit=1000)
        categories = sorted(
            {str(c.get("category") or "").strip() for c in all_concepts if str(c.get("category") or "").strip()}
        )

        f1, f2 = st.columns([2, 1])
        with f1:
            search = st.text_input(
                "Buscar",
                placeholder="Ej. cancelería, demolición, pintura, PRE-01",
                key="catalog_search",
            )
        with f2:
            category_filter = st.selectbox(
                "Partida",
                ["Todas"] + categories,
                key="catalog_category",
            )

        concepts = db.list_concepts(search, limit=1000)
        if category_filter != "Todas":
            concepts = [
                c for c in concepts
                if str(c.get("category") or "").strip() == category_filter
            ]

        st.caption(f"{len(concepts)} concepto(s) encontrados.")

        if concepts:
            catalog_df = pd.DataFrame([
                {
                    "Código": c.get("code") or "",
                    "Concepto": c.get("description") or "",
                    "Partida": c.get("category") or "",
                    "Subpartida": c.get("subcategory") or "",
                    "Unidad": c.get("unit") or "",
                    "Último costo": c.get("latest_cost"),
                    "Fuente": friendly_source(c.get("latest_source")),
                    "Usos": int(c.get("usage_count") or 0),
                }
                for c in concepts
            ])
            st.dataframe(
                catalog_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Último costo": st.column_config.NumberColumn(format="$ %.2f"),
                    "Usos": st.column_config.NumberColumn(format="%d"),
                },
            )

            concept_map = {
                c["id"]: f"{c.get('code') or 'SIN-COD'} — {c.get('description') or ''}"
                for c in concepts
            }
            selected_id = st.selectbox(
                "Abrir concepto",
                options=list(concept_map.keys()),
                format_func=lambda x: concept_map[x],
                key="catalog_open_concept",
            )

            concept = db.get_concept(selected_id)
            if concept:
                usage = db.concept_usage(selected_id)
                prices = db.list_prices(selected_id)

                with st.container(border=True):
                    st.markdown(f"### {concept.get('description') or 'Concepto'}")
                    st.caption(f"Código: {concept.get('code') or 'Sin código'}")

                    d1, d2, d3 = st.columns(3)
                    d1.metric("Unidad", concept.get("unit") or "—")
                    d2.metric("Usos en presupuestos", int(usage.get("budget_items") or 0))
                    d3.metric("Precios registrados", int(usage.get("prices") or 0))

                    st.markdown("**Partida**")
                    st.write(concept.get("category") or "Sin partida")
                    st.markdown("**Subpartida**")
                    st.write(concept.get("subcategory") or "Sin subpartida")

                    if prices:
                        last = prices[0]
                        st.markdown("**Precio más reciente**")
                        st.write(
                            f"{formato_moneda(float(last['unit_cost']))} · "
                            f"{friendly_source(last.get('source'))} · "
                            f"{friendly_status(last.get('status'))}"
                        )
                    else:
                        st.info("Este concepto todavía no tiene precios registrados.")

                edit_key = f"editing_concept_{selected_id}"
                if edit_key not in st.session_state:
                    st.session_state[edit_key] = False

                b1, b2 = st.columns([1, 3])
                with b1:
                    if st.button(
                        "Editar concepto",
                        key=f"open_edit_{selected_id}",
                        use_container_width=True,
                    ):
                        st.session_state[edit_key] = True
                        st.rerun()
                with b2:
                    st.caption(
                        "Para agregar, revisar o eliminar costos históricos use la pestaña Precios."
                    )

                if st.session_state.get(edit_key):
                    with st.container(border=True):
                        st.markdown("#### Editar concepto")
                        st.caption(
                            "Modificar estos datos no altera los precios históricos ni los presupuestos ya generados."
                        )
                        with st.form(f"edit_concept_form_{selected_id}"):
                            ec1, ec2 = st.columns(2)
                            with ec1:
                                c_code = st.text_input(
                                    "Código",
                                    value=concept.get("code") or "",
                                )
                                c_category = st.text_input(
                                    "Partida",
                                    value=concept.get("category") or "",
                                )
                                c_subcategory = st.text_input(
                                    "Subpartida",
                                    value=concept.get("subcategory") or "",
                                )
                            with ec2:
                                c_unit = st.text_input(
                                    "Unidad",
                                    value=concept.get("unit") or "",
                                )
                                c_description = st.text_area(
                                    "Descripción",
                                    value=concept.get("description") or "",
                                    height=145,
                                )
                            save_col, cancel_col = st.columns(2)
                            save_edit = save_col.form_submit_button(
                                "Guardar cambios",
                                type="primary",
                                use_container_width=True,
                            )
                            cancel_edit = cancel_col.form_submit_button(
                                "Cancelar",
                                use_container_width=True,
                            )

                        if cancel_edit:
                            st.session_state[edit_key] = False
                            st.rerun()

                        if save_edit:
                            if not c_description.strip() or not c_unit.strip():
                                st.error("Descripción y unidad son obligatorias.")
                            else:
                                db.update_concept(
                                    selected_id,
                                    c_code,
                                    c_category,
                                    c_subcategory,
                                    c_description,
                                    c_unit,
                                )
                                st.session_state[edit_key] = False
                                st.success("Concepto actualizado.")
                                st.rerun()

                if prices:
                    st.markdown("#### Últimos precios")
                    recent_prices = prices[:5]
                    recent_df = pd.DataFrame([
                        {
                            "Costo": p["unit_cost"],
                            "Fuente": friendly_source(p.get("source")),
                            "Estado": friendly_status(p.get("status")),
                            "Confianza": p.get("confidence") or "",
                            "Fecha": p.get("created_at") or "",
                        }
                        for p in recent_prices
                    ])
                    st.dataframe(
                        recent_df,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "Costo": st.column_config.NumberColumn(format="$ %.2f")
                        },
                    )

                with st.expander("Opciones avanzadas"):
                    st.warning(
                        "Eliminar un concepto es una acción administrativa. "
                        "Los presupuestos históricos conservan los datos de la actividad, "
                        "pero pueden perder el vínculo con el catálogo."
                    )
                    st.write(
                        f"Usos en presupuestos: {usage['budget_items']} · "
                        f"Precios históricos: {usage['prices']}"
                    )
                    confirm_concept = st.checkbox(
                        "Confirmo que deseo eliminar este concepto.",
                        key=f"confirm_delete_concept_{selected_id}",
                    )
                    if st.button(
                        "Eliminar concepto",
                        key=f"delete_concept_{selected_id}",
                        disabled=not confirm_concept,
                    ):
                        db.delete_concept(selected_id)
                        st.success("Concepto eliminado.")
                        st.rerun()
        else:
            st.info("No se encontraron conceptos con esos filtros.")

        st.divider()

        create_flag = "admin_create_concept_open"
        if create_flag not in st.session_state:
            st.session_state[create_flag] = False

        if not st.session_state[create_flag]:
            if st.button("Crear concepto nuevo", use_container_width=True):
                st.session_state[create_flag] = True
                st.rerun()
        else:
            with st.container(border=True):
                st.markdown("### Nuevo concepto")
                st.caption(
                    "Úselo para registrar manualmente una actividad que todavía no existe en el catálogo."
                )
                with st.form("catalog_create_concept"):
                    nc1, nc2 = st.columns(2)
                    with nc1:
                        a_code = st.text_input("Código", value="MAN-001")
                        a_category = st.text_input("Partida")
                        a_subcategory = st.text_input("Subpartida")
                    with nc2:
                        a_unit = st.text_input("Unidad", value="LOTE")
                        a_description = st.text_area("Descripción", height=145)

                    add_col, cancel_col = st.columns(2)
                    create_btn = add_col.form_submit_button(
                        "Crear concepto",
                        type="primary",
                        use_container_width=True,
                    )
                    cancel_create = cancel_col.form_submit_button(
                        "Cancelar",
                        use_container_width=True,
                    )

                if cancel_create:
                    st.session_state[create_flag] = False
                    st.rerun()

                if create_btn:
                    if not a_description.strip() or not a_unit.strip():
                        st.error("Descripción y unidad son obligatorias.")
                    else:
                        new_id = db.create_concept(
                            a_code,
                            a_category,
                            a_subcategory,
                            a_description,
                            a_unit,
                        )
                        st.session_state[create_flag] = False
                        st.success("Concepto creado. Puede agregarle un precio desde la pestaña Precios.")
                        st.rerun()

    # =====================================================
    # PRECIOS
    # =====================================================
    with tab_prices:
        st.subheader("Historial de precios")
        st.caption(
            "Los precios se guardan como registros históricos. Agregar un precio nuevo "
            "no elimina ni reemplaza automáticamente los anteriores."
        )

        concepts_for_prices = db.list_concepts("", limit=1500)

        if not concepts_for_prices:
            st.info("Primero debe existir al menos un concepto en el catálogo.")
        else:
            price_search = st.text_input(
                "Buscar concepto para administrar sus precios",
                placeholder="Ej. pintura, carpintería, cancel",
                key="price_concept_search",
            )
            if price_search.strip():
                filtered_price_concepts = db.list_concepts(price_search, limit=500)
            else:
                filtered_price_concepts = concepts_for_prices

            if not filtered_price_concepts:
                st.info("No se encontraron conceptos.")
            else:
                price_concept_map = {
                    c["id"]: f"{c.get('code') or 'SIN-COD'} — {c.get('description') or ''} [{c.get('unit') or ''}]"
                    for c in filtered_price_concepts
                }
                price_concept_id = st.selectbox(
                    "Concepto",
                    options=list(price_concept_map.keys()),
                    format_func=lambda x: price_concept_map[x],
                    key="price_selected_concept",
                )

                concept = db.get_concept(price_concept_id)
                prices = db.list_prices(price_concept_id)

                with st.container(border=True):
                    st.markdown(f"### {concept.get('description') or 'Concepto'}")
                    pinfo1, pinfo2, pinfo3 = st.columns(3)
                    pinfo1.metric("Código", concept.get("code") or "—")
                    pinfo2.metric("Unidad", concept.get("unit") or "—")
                    pinfo3.metric("Registros", len(prices))

                    st.caption(
                        f"{concept.get('category') or 'Sin partida'} / "
                        f"{concept.get('subcategory') or 'Sin subpartida'}"
                    )

                if prices:
                    price_df = pd.DataFrame([
                        {
                            "Costo unitario": p["unit_cost"],
                            "Fuente": friendly_source(p.get("source")),
                            "Estado": friendly_status(p.get("status")),
                            "Confianza": p.get("confidence") or "",
                            "Detalle": p.get("source_detail") or "",
                            "Fecha": p.get("created_at") or "",
                        }
                        for p in prices
                    ])
                    st.dataframe(
                        price_df,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "Costo unitario": st.column_config.NumberColumn(format="$ %.2f")
                        },
                    )
                else:
                    st.info("No hay precios registrados para este concepto.")

                st.markdown("#### Agregar precio")
                st.caption(
                    "Registre una nueva referencia. El historial anterior se conserva."
                )

                source_options = [
                    "COTIZACION_PROVEEDOR",
                    "COSTO_REAL",
                    "REFERENCIA_EXTERNA",
                    "IA_ESTIMADO",
                    "MANUAL",
                ]
                status_options = [
                    "VALIDADO",
                    "COTIZADO_PROVEEDOR",
                    "COSTO_REAL",
                    "REFERENCIA_EXTERNA",
                    "ESTIMADO_IA",
                ]

                with st.form(f"add_price_form_{price_concept_id}"):
                    ap1, ap2 = st.columns(2)
                    with ap1:
                        new_cost = st.number_input(
                            "Costo unitario",
                            min_value=0.0,
                            step=100.0,
                            format="%.2f",
                        )
                        source_label_choice = st.selectbox(
                            "Origen del precio",
                            options=source_options,
                            format_func=lambda x: source_labels[x],
                        )
                        confidence = st.selectbox(
                            "Nivel de confianza",
                            ["Alta", "Media", "Baja"],
                        )
                    with ap2:
                        status_choice = st.selectbox(
                            "Estado",
                            options=status_options,
                            format_func=lambda x: status_labels[x],
                        )
                        detail = st.text_area(
                            "Referencia / proveedor / nota",
                            height=125,
                            placeholder="Ej. Cotización Proveedor X, agosto 2026.",
                        )
                    add_price = st.form_submit_button(
                        "Agregar al historial",
                        type="primary",
                        use_container_width=True,
                    )

                if add_price:
                    if new_cost <= 0:
                        st.error("El costo debe ser mayor que cero.")
                    else:
                        db.add_price(
                            price_concept_id,
                            new_cost,
                            source_label_choice,
                            detail,
                            status_choice,
                            confidence,
                        )
                        st.success("Precio agregado al historial.")
                        st.rerun()

                if prices:
                    with st.expander("Eliminar un precio registrado"):
                        st.caption(
                            "Utilícelo únicamente para eliminar registros incorrectos. "
                            "No es necesario borrar precios antiguos."
                        )
                        price_map = {
                            p["id"]: (
                                f"{formato_moneda(float(p['unit_cost']))} — "
                                f"{friendly_source(p.get('source'))} — "
                                f"{p.get('created_at') or ''}"
                            )
                            for p in prices
                        }
                        price_to_delete = st.selectbox(
                            "Registro",
                            options=list(price_map.keys()),
                            format_func=lambda x: price_map[x],
                            key=f"delete_price_select_{price_concept_id}",
                        )
                        confirm_price = st.text_input(
                            "Para eliminar, escriba ELIMINAR PRECIO",
                            key=f"delete_price_confirm_{price_concept_id}",
                        )
                        if st.button(
                            "Eliminar registro",
                            key=f"delete_price_button_{price_concept_id}",
                        ):
                            if confirm_price.strip().upper() != "ELIMINAR PRECIO":
                                st.error("Confirmación incorrecta.")
                            else:
                                db.delete_price(price_to_delete)
                                st.success("Precio eliminado.")
                                st.rerun()

    # =====================================================
    # PROYECTOS
    # =====================================================
    with tab_projects:
        st.subheader("Historial de proyectos")
        st.caption(
            "Consulte los proyectos que han generado presupuestos reales. "
            "Las simulaciones no aparecen aquí porque no se guardan."
        )

        projects = db.list_projects(limit=1000)
        project_search = st.text_input(
            "Buscar proyecto",
            placeholder="Código, cliente, ubicación o tipo de obra",
            key="project_history_search",
        )

        if project_search.strip():
            q = normalizar_texto(project_search)
            projects = [
                p for p in projects
                if q in normalizar_texto(
                    " ".join([
                        str(p.get("code") or ""),
                        str(p.get("name") or ""),
                        str(p.get("location") or ""),
                        str(p.get("project_type") or ""),
                    ])
                )
            ]

        if not projects:
            st.info("No se encontraron proyectos.")
        else:
            project_df = pd.DataFrame([
                {
                    "Código": p.get("code") or "",
                    "Cliente": p.get("name") or "",
                    "Tipo": p.get("project_type") or "",
                    "Ubicación": p.get("location") or "",
                    "Presupuestos": int(p.get("budget_count") or 0),
                    "Último total": p.get("latest_total"),
                    "Fecha": p.get("created_at") or "",
                }
                for p in projects
            ])
            st.dataframe(
                project_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Último total": st.column_config.NumberColumn(format="$ %.2f")
                },
            )

            project_map = {
                p["id"]: f"{p.get('code') or ''} — {p.get('name') or ''}"
                for p in projects
            }
            project_id = st.selectbox(
                "Abrir proyecto",
                options=list(project_map.keys()),
                format_func=lambda x: project_map[x],
                key="open_project_history",
            )
            project = db.get_project(project_id)

            if project:
                project_budgets = db.list_budgets(project_id=project_id)

                with st.container(border=True):
                    st.markdown(f"### {project.get('name') or 'Cliente'}")
                    st.caption(project.get("code") or "")

                    pr1, pr2, pr3 = st.columns(3)
                    pr1.metric("Tipo", project.get("project_type") or "—")
                    pr2.metric("Ubicación", project.get("location") or "—")
                    pr3.metric("Presupuestos", len(project_budgets))

                    if project.get("main_activity"):
                        st.markdown("**Actividad principal**")
                        st.write(project.get("main_activity"))
                    if project.get("dimensions_text"):
                        st.markdown("**Dimensiones / referencias**")
                        st.write(project.get("dimensions_text"))
                    if project.get("description"):
                        st.markdown("**Descripción inicial**")
                        st.write(project.get("description"))

                if project_budgets:
                    st.markdown("#### Presupuestos del proyecto")
                    pb_df = pd.DataFrame([
                        {
                            "Versión": b.get("version"),
                            "Estado": b.get("status"),
                            "Costo directo": b.get("direct_cost"),
                            "Venta sin IVA": b.get("sale_before_tax"),
                            "Total": b.get("total"),
                            "Fecha": b.get("created_at"),
                        }
                        for b in project_budgets
                    ])
                    st.dataframe(
                        pb_df,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "Costo directo": st.column_config.NumberColumn(format="$ %.2f"),
                            "Venta sin IVA": st.column_config.NumberColumn(format="$ %.2f"),
                            "Total": st.column_config.NumberColumn(format="$ %.2f"),
                        },
                    )

                edit_project_key = f"editing_project_{project_id}"
                if edit_project_key not in st.session_state:
                    st.session_state[edit_project_key] = False

                if st.button(
                    "Editar datos del proyecto",
                    key=f"edit_project_button_{project_id}",
                ):
                    st.session_state[edit_project_key] = True
                    st.rerun()

                if st.session_state.get(edit_project_key):
                    with st.container(border=True):
                        st.markdown("#### Editar proyecto")
                        st.caption(
                            "Esta edición corrige los datos descriptivos del proyecto. "
                            "No recalcula presupuestos existentes."
                        )
                        with st.form(f"edit_project_form_{project_id}"):
                            ep1, ep2 = st.columns(2)
                            with ep1:
                                p_name = st.text_input(
                                    "Nombre del cliente",
                                    value=project.get("name") or "",
                                )
                                p_type = st.text_input(
                                    "Tipo de obra",
                                    value=project.get("project_type") or "",
                                )
                                p_location = st.text_input(
                                    "Ubicación",
                                    value=project.get("location") or "",
                                )
                                p_main_activity = st.text_input(
                                    "Actividad principal",
                                    value=project.get("main_activity") or "",
                                )
                            with ep2:
                                p_dimensions = project.get("dimensions_text") or ""
                                p_description = st.text_area(
                                    "Descripción",
                                    value=project.get("description") or "",
                                    height=120,
                                )
                                p_guide = st.text_area(
                                    "Texto guía",
                                    value=project.get("guide_text") or "",
                                    height=90,
                                )
                            save_pc, cancel_pc = st.columns(2)
                            save_project = save_pc.form_submit_button(
                                "Guardar cambios",
                                type="primary",
                                use_container_width=True,
                            )
                            cancel_project = cancel_pc.form_submit_button(
                                "Cancelar",
                                use_container_width=True,
                            )

                        if cancel_project:
                            st.session_state[edit_project_key] = False
                            st.rerun()

                        if save_project:
                            if not p_name.strip():
                                st.error("El nombre del proyecto es obligatorio.")
                            else:
                                db.update_project(
                                    project_id,
                                    p_name,
                                    p_type,
                                    p_location,
                                    p_main_activity,
                                    p_dimensions,
                                    p_description,
                                    p_guide,
                                )
                                st.session_state[edit_project_key] = False
                                st.success("Proyecto actualizado.")
                                st.rerun()

                with st.expander("Opciones avanzadas"):
                    st.warning(
                        "Eliminar el proyecto borra sus presupuestos, partidas y los conceptos/precios "
                        "creados exclusivamente por ese proyecto."
                    )
                    project_delete_key = st.text_input(
                        "Clave de eliminación",
                        type="password",
                        key=f"confirm_delete_project_{project_id}",
                    )
                    if st.button(
                        "Eliminar proyecto completo",
                        key=f"delete_project_button_{project_id}",
                    ):
                        if not validar_clave_borrado(project_delete_key):
                            st.error("Clave de eliminación incorrecta o no configurada.")
                        else:
                            db.delete_project(project_id)
                            st.success("Proyecto y su trazabilidad asociada fueron eliminados.")
                            st.rerun()

    # =====================================================
    # PRESUPUESTOS
    # =====================================================
    with tab_budgets:
        st.subheader("Historial de presupuestos")
        st.caption(
            "Consulte presupuestos ya guardados. Esta sección es de consulta; "
            "las correcciones comerciales siguen realizándose en el Excel generado."
        )

        budgets = db.list_budgets(limit=1000)

        budget_search = st.text_input(
            "Buscar presupuesto",
            placeholder="Código de proyecto, nombre o ubicación",
            key="budget_history_search",
        )
        if budget_search.strip():
            q = normalizar_texto(budget_search)
            budgets = [
                b for b in budgets
                if q in normalizar_texto(
                    " ".join([
                        str(b.get("project_code") or ""),
                        str(b.get("project_name") or ""),
                        str(b.get("project_location") or ""),
                    ])
                )
            ]

        if not budgets:
            st.info("No se encontraron presupuestos.")
        else:
            budget_df = pd.DataFrame([
                {
                    "Código": b.get("project_code") or "",
                    "Cliente": b.get("project_name") or "",
                    "Estado": b.get("status") or "",
                    "Costo directo": b.get("direct_cost"),
                    "Indirectos": b.get("indirect_cost"),
                    "Utilidad": b.get("profit"),
                    "Venta sin IVA": b.get("sale_before_tax"),
                    "Total": b.get("total"),
                    "Fecha": b.get("created_at") or "",
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
                b["id"]: (
                    f"{b.get('project_code') or ''} — "
                    f"{b.get('project_name') or ''} — "
                    f"{formato_moneda(float(b.get('total') or 0))}"
                )
                for b in budgets
            }
            budget_id = st.selectbox(
                "Abrir presupuesto",
                options=list(budget_map.keys()),
                format_func=lambda x: budget_map[x],
                key="open_budget_history",
            )

            budget = db.get_budget(budget_id)
            items = db.list_budget_items(budget_id)

            if budget:
                with st.container(border=True):
                    st.markdown(
                        f"### {budget.get('project_code') or ''} — "
                        f"{budget.get('project_name') or ''}"
                    )
                    bm1, bm2, bm3, bm4 = st.columns(4)
                    bm1.metric("Costo directo", formato_moneda(float(budget.get("direct_cost") or 0)))
                    bm2.metric("Indirectos", formato_moneda(float(budget.get("indirect_cost") or 0)))
                    bm3.metric("Utilidad", formato_moneda(float(budget.get("profit") or 0)))
                    bm4.metric("Total", formato_moneda(float(budget.get("total") or 0)))

                    st.caption(
                        f"Indirectos: {float(budget.get('indirect_pct') or 0):.2f}% · "
                        f"Utilidad: {float(budget.get('profit_pct') or 0):.2f}% · "
                        f"IVA: {float(budget.get('iva_pct') or 0):.2f}% · "
                        f"Estado: {budget.get('status') or ''}"
                    )

                if items:
                    st.markdown("#### Actividades")
                    item_df = pd.DataFrame([
                        {
                            "Partida": i.get("category") or "",
                            "Subpartida": i.get("subcategory") or "",
                            "Código": i.get("code") or "",
                            "Actividad": i.get("description") or "",
                            "Unidad": i.get("unit") or "",
                            "Cantidad": i.get("quantity"),
                            "Costo unitario": i.get("unit_cost"),
                            "Venta unitaria": i.get("unit_sale"),
                            "Venta": i.get("sale_amount"),
                            "Fuente": friendly_source(i.get("price_source")),
                        }
                        for i in items
                    ])
                    st.dataframe(
                        item_df,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "Costo unitario": st.column_config.NumberColumn(format="$ %.2f"),
                            "Venta unitaria": st.column_config.NumberColumn(format="$ %.2f"),
                            "Venta": st.column_config.NumberColumn(format="$ %.2f"),
                        },
                    )
                else:
                    st.info("Este presupuesto no contiene actividades registradas.")

                with st.expander("Opciones avanzadas"):
                    st.warning(
                        "Eliminar un presupuesto elimina sus actividades e historial vinculado. "
                        "Si es el único presupuesto del proyecto, también puede eliminarse el proyecto."
                    )
                    confirm_budget = st.text_input(
                        "Para eliminar, escriba ELIMINAR PRESUPUESTO",
                        key=f"confirm_delete_budget_{budget_id}",
                    )
                    if st.button(
                        "Eliminar presupuesto",
                        key=f"delete_budget_button_{budget_id}",
                    ):
                        if confirm_budget.strip().upper() != "ELIMINAR PRESUPUESTO":
                            st.error("Confirmación incorrecta.")
                        else:
                            db.delete_budget(budget_id)
                            st.success("Presupuesto eliminado.")
                            st.rerun()

    # =====================================================
    # MANTENIMIENTO
    # =====================================================
    with tab_maintenance:
        st.subheader("Mantenimiento")
        st.caption(
            "Herramientas para la etapa de pruebas. Las acciones de esta sección "
            "afectan datos persistentes y requieren la clave de eliminación."
        )

        if not get_delete_key():
            st.error(
                "Falta DELETE_KEY en Streamlit Secrets. "
                "Las operaciones destructivas permanecerán bloqueadas."
            )

        stats_now = db.stats()
        with st.container(border=True):
            st.markdown("### Estado actual")
            sm1, sm2, sm3 = st.columns(3)
            sm1.metric("Proyectos", stats_now["projects"])
            sm2.metric("Presupuestos", stats_now["budgets"])
            sm3.metric("Conceptos", stats_now["concepts"])

        try:
            latest = db.get_latest_project_record()
        except Exception as exc:
            latest = None
            st.error(f"No fue posible consultar el último proyecto: {exc}")

        with st.container(border=True):
            st.markdown("### Último proyecto guardado")
            if latest:
                st.write(f"**{latest.get('code') or ''} — {latest.get('name') or ''}**")
                st.caption(
                    f"{latest.get('project_type') or ''} · "
                    f"{latest.get('location') or 'Sin ubicación'} · "
                    f"{latest.get('created_at') or ''}"
                )
                if latest.get("latest_total") is not None:
                    st.write(f"Último total: **{formato_moneda(float(latest['latest_total']))}**")
                st.caption(
                    "Para borrar este proyecto sin ser administrador también existe "
                    "la herramienta de corrección en la sección Generar presupuesto."
                )
            else:
                st.info("La base no contiene proyectos.")

        with st.container(border=True):
            st.markdown("### Eliminar todos los datos de la aplicación")
            st.error(
                "Esta acción borra proyectos, presupuestos, actividades, conceptos e historial "
                "de precios. La estructura de las tablas se conserva para que la aplicación "
                "pueda seguir funcionando."
            )
            wipe_key = st.text_input(
                "Clave de eliminación",
                type="password",
                key="maintenance_wipe_key",
            )
            wipe_confirm = st.checkbox(
                "Confirmo que deseo vaciar toda la base de datos de la aplicación.",
                key="maintenance_wipe_confirm",
            )
            if st.button(
                "Eliminar todos los datos",
                type="primary",
                use_container_width=True,
                disabled=not wipe_confirm,
                key="maintenance_wipe_database",
            ):
                if not validar_clave_borrado(wipe_key):
                    st.error("Clave de eliminación incorrecta o no configurada.")
                else:
                    db.clear_all_data()
                    st.session_state.pop("generated", None)
                    st.success("Todos los datos de la aplicación fueron eliminados.")
                    st.rerun()

    # =====================================================
    # EXPORTAR
    # =====================================================
    with tab_export:
        st.subheader("Exportar información")
        st.caption(
            "Descarga copias CSV para revisión, respaldo o análisis. "
            "Estas descargas no modifican la base de datos."
        )

        export_items = [
            ("concepts", "Catálogo de conceptos", "conceptos.csv"),
            ("price_history", "Historial de precios", "historial_precios.csv"),
            ("projects", "Proyectos", "proyectos.csv"),
            ("budgets", "Presupuestos", "presupuestos.csv"),
            ("budget_items", "Actividades de presupuestos", "actividades_presupuestos.csv"),
        ]

        for table_name, label, filename in export_items:
            with st.container(border=True):
                c1, c2 = st.columns([3, 1])
                with c1:
                    st.markdown(f"**{label}**")
                    st.caption(filename)
                rows = db.export_table(table_name)
                csv_bytes = pd.DataFrame(rows).to_csv(index=False).encode("utf-8-sig")
                with c2:
                    st.download_button(
                        "Descargar CSV",
                        data=csv_bytes,
                        file_name=filename,
                        mime="text/csv",
                        key=f"friendly_download_{table_name}",
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
    db, db_error = get_database(database_url, DATABASE_CACHE_VERSION)

    # Protección adicional para despliegues en caliente:
    # si Streamlit llegara a conservar un objeto Database de una versión
    # anterior, se limpia la caché de recursos y se crea uno nuevo.
    required_database_methods = (
        "get_latest_project_record",
        "clear_all_data",
        "delete_generation",
        "save_generation",
        "save_revision",
    )
    if any(not hasattr(db, method) for method in required_database_methods):
        st.cache_resource.clear()
        db, db_error = get_database(database_url, DATABASE_CACHE_VERSION)
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
        sections.append("Catálogo e historial")

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

        st.divider()
        with st.expander("Corrección de última carga"):
            st.caption(
                "Permite retirar el último proyecto guardado si una prueba se cargó "
                "por error. No requiere rol administrador, pero sí la clave de eliminación."
            )
            try:
                latest_sidebar = db.get_latest_project_record()
            except Exception as exc:
                latest_sidebar = None
                st.error(f"No fue posible consultar el último proyecto: {exc}")

            if latest_sidebar:
                st.write(
                    f"**{latest_sidebar.get('code') or ''} — "
                    f"{latest_sidebar.get('name') or ''}**"
                )
                if latest_sidebar.get("latest_total") is not None:
                    st.caption(
                        f"Total: {formato_moneda(float(latest_sidebar['latest_total']))}"
                    )
                last_delete_key = st.text_input(
                    "Clave",
                    type="password",
                    key="sidebar_last_delete_key",
                )
                if st.button(
                    "Borrar último proyecto y trazabilidad",
                    key="sidebar_delete_last_project",
                    use_container_width=True,
                ):
                    if not validar_clave_borrado(last_delete_key):
                        st.error("Clave incorrecta o no configurada.")
                    else:
                        current = st.session_state.get("generated")
                        db.delete_project(latest_sidebar["id"])
                        if (
                            current
                            and current.get("project_id") == latest_sidebar["id"]
                        ):
                            st.session_state.pop("generated", None)
                        st.success("Último proyecto eliminado por completo.")
                        st.rerun()
            else:
                st.caption("No hay proyectos guardados.")

# =========================================================
# BASE INTERNA
# =========================================================


if section == "Catálogo e historial":
    if not is_admin:
        st.error("No tiene permisos para administrar el catálogo e historial.")
        st.stop()
    render_admin_database(db)
    st.stop()


# =========================================================
# FORMULARIO INICIAL
# =========================================================


DESCRIPTION_EXAMPLE = """Ejemplo de formato esperado:

Casa habitación.

SEGUNDA PLANTA
Recámara principal de aproximadamente 4.20 x 3.80 m.
- Retiro de piso laminado existente.
- Colocación de piso nuevo.
- Reparación y pintura de muros.
- Sustitución de luminarias.

AZOTEA
Área aproximada de 9 x 4 m.
- Retiro de material suelto y limpieza.
- Preparación de superficie.
- Impermeabilización completa.
- Revisión de bajadas pluviales existentes.

ESCALERA
- Reparación de acabados dañados.
- Pintura de muros y plafón.
"""

GUIDE_EXAMPLE = """Ejemplo de criterios generales:

- Considerar acabados de gama media.
- Incluir protección con plástico y cartón engomado en las zonas de tránsito.
- Considerar retiro de desperdicios y limpieza final.
- El edificio permite trabajos de lunes a viernes de 09:00 a 18:00.
- No considerar jardinería.
- Cuando no exista una especificación definitiva, utilizar una alternativa
  comercial de gama media y señalarla como consideración.
"""

REVISION_EXAMPLE = """Ejemplo:

El presupuesto no contempló la impermeabilización completa de la azotea.
Agregar la preparación de superficie, reparación puntual de fisuras y el
sistema de impermeabilización para el área de 9 x 4 m indicada en la
descripción original.

Además, la cancelería de la segunda planta debe contemplarse en aluminio
línea pesada y cristal templado, por lo que esa actividad debe actualizarse
de forma completa.
"""


if "generated" not in st.session_state:
    with st.form("form_proyecto"):
        simulation_mode = st.toggle(
            "Simulación",
            value=False,
            help=(
                "Genera resultados y archivos normalmente, pero no guarda nuevos "
                "proyectos, presupuestos, conceptos ni precios."
            ),
            key="simulation_mode",
        )
        if simulation_mode:
            st.info("Modo simulación activo. Esta corrida no escribirá en la base interna.")

        st.subheader("Datos del proyecto")

        f1, f2 = st.columns(2)
        with f1:
            client_name = st.text_input(
                "Nombre del cliente",
                value=st.session_state.get("client_name", ""),
                placeholder="Ej. Desarrollos de la Vega",
                key="client_name",
            )
        with f2:
            location = st.text_input(
                "Ubicación",
                value=st.session_state.get("project_location", ""),
                placeholder="Ej. Farallón, Álvaro Obregón, CDMX",
                key="project_location",
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

        st.markdown("**Descripción general de trabajos**")
        st.caption(
            "Describa zonas, medidas conocidas, estado actual, actividades solicitadas "
            "y condiciones relevantes. Las dimensiones se escriben directamente aquí."
        )
        description = st.text_area(
            "Descripción general de trabajos",
            placeholder=DESCRIPTION_EXAMPLE,
            height=430,
            key="project_description",
            label_visibility="collapsed",
        )

        st.markdown("**Texto guía**")
        st.caption(
            "Criterios que deben aplicarse al presupuesto en general: nivel de acabado, "
            "exclusiones, restricciones, protecciones, horarios o instrucciones especiales."
        )
        guide_text = st.text_area(
            "Texto guía",
            placeholder=GUIDE_EXAMPLE,
            height=300,
            key="guide_text",
            label_visibility="collapsed",
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

        if not client_name.strip():
            st.error("Ingrese el nombre del cliente.")
            st.stop()

        if not location.strip():
            st.error("Ingrese la ubicación.")
            st.stop()

        if not description.strip():
            st.error("Ingrese la descripción general de los trabajos.")
            st.stop()

        params = {
            "indirect_pct": float(indirect_pct),
            "profit_pct": float(profit_pct),
            "iva_pct": float(iva_pct),
            "waste_pct": float(waste_pct),
        }
        project_data = {
            # Se conserva la llave "name" por compatibilidad con la base y el Excel.
            # A partir de V7 representa el nombre del cliente.
            "name": client_name.strip(),
            "project_type": project_type,
            "location": location.strip(),
            "dimension_mode": "Integradas en descripción",
            "dimensions_text": "",
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

                items = resolver_items(db, result, project_data, params)
                if not items:
                    raise RuntimeError("La IA no generó actividades utilizables.")

                financials = calcular_financieros(items, params)
                base_code = db.next_project_code(
                    project_data["name"],
                    project_data["location"],
                )
                project_code = f"SIM-{base_code}" if simulation_mode else base_code
                version = 1
                file_code = f"{project_code}-V{version:02d}"

                excel_bytes = crear_excel(
                    project_code=project_code,
                    project_data=project_data,
                    result=result,
                    items=items,
                    params=params,
                    version=version,
                )
                txt_bytes = crear_txt(
                    project_code=project_code,
                    project_data=project_data,
                    result=result,
                    items=items,
                    params=params,
                    financials=financials,
                    version=version,
                )
                zip_bytes = crear_paquete_zip(file_code, excel_bytes, txt_bytes)

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
                    "file_code": file_code,
                    "version": version,
                    "project_data": project_data,
                    "params": params,
                    "result": result.model_dump(),
                    "items": items,
                    "financials": financials,
                    "excel_bytes": excel_bytes,
                    "txt_bytes": txt_bytes,
                    "zip_bytes": zip_bytes,
                    "revision_history": [],
                }
                st.rerun()

            except Exception as exc:
                st.exception(exc)


# =========================================================
# RESULTADO FINAL Y REVISIONES ESTRUCTURALES
# =========================================================


else:
    g = st.session_state["generated"]
    result = PresupuestoIA.model_validate(g["result"])
    items = g["items"]
    financials = g["financials"]
    version = int(g.get("version") or 1)

    if g.get("simulation"):
        st.warning("SIMULACIÓN: esta generación no fue guardada en la base interna.")

    st.subheader(g["project_code"])
    st.caption(f"Versión V{version:02d}")
    st.write(result.alcance_resumido)

    # -----------------------------------------------------
    # Datos originales siempre visibles y bloqueados
    # -----------------------------------------------------
    st.subheader("Datos originales del proyecto")
    st.caption(
        "Estos datos permanecen bloqueados durante las revisiones para evitar que "
        "una nueva versión cambie accidentalmente la definición original del proyecto."
    )

    p1, p2, p3 = st.columns(3)
    with p1:
        st.text_input(
            "Cliente",
            value=g["project_data"]["name"],
            disabled=True,
            key=f"locked_client_{version}",
        )
    with p2:
        st.text_input(
            "Ubicación",
            value=g["project_data"]["location"],
            disabled=True,
            key=f"locked_location_{version}",
        )
    with p3:
        st.text_input(
            "Tipo de obra",
            value=g["project_data"]["project_type"],
            disabled=True,
            key=f"locked_type_{version}",
        )

    st.text_area(
        "Descripción general de trabajos",
        value=g["project_data"]["description"],
        height=260,
        disabled=True,
        key=f"locked_description_{version}",
    )
    st.text_area(
        "Texto guía",
        value=g["project_data"]["guide_text"] or "Sin texto guía adicional.",
        height=170,
        disabled=True,
        key=f"locked_guide_{version}",
    )

    # -----------------------------------------------------
    # Resumen económico
    # -----------------------------------------------------
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

    # -----------------------------------------------------
    # Archivos
    # -----------------------------------------------------
    st.subheader("Archivos")
    file_code = g.get("file_code") or f"{g['project_code']}-V{version:02d}"
    st.download_button(
        "Descargar paquete",
        data=g["zip_bytes"],
        file_name=f"{file_code}.zip",
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
                file_name=f"{file_code}_Presupuesto.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with c2:
            st.download_button(
                "Descargar TXT",
                data=g["txt_bytes"],
                file_name=f"{file_code}_Captura_Plataforma.txt",
                mime="text/plain",
                use_container_width=True,
            )

    # -----------------------------------------------------
    # Historial breve de revisiones de la sesión
    # -----------------------------------------------------
    revision_history = g.get("revision_history") or []
    if revision_history:
        with st.expander("Historial de revisiones de esta sesión"):
            for rev in revision_history:
                st.markdown(f"**V{int(rev['version']):02d} — {rev['summary']}**")
                st.caption(rev["request"])
                for change in rev.get("changes") or []:
                    st.write(f"- {change}")

    # -----------------------------------------------------
    # Revisión estructural
    # -----------------------------------------------------
    st.divider()
    st.subheader("Revisión estructural del presupuesto")
    st.warning(
        "Utilice esta opción únicamente para cambios importantes de alcance: "
        "actividades omitidas, especificaciones que cambian una partida completa, "
        "nuevas zonas o trabajos que requieren agregar, sustituir o eliminar conceptos."
    )
    st.markdown(
        "**No utilice esta herramienta para cambios menores de precio, cantidad o redacción. "
        "Esas correcciones deben realizarse directamente en el Excel.**"
    )

    revision_request = st.text_area(
        "Cambios estructurales solicitados",
        placeholder=REVISION_EXAMPLE,
        height=260,
        key=f"revision_request_v{version}",
    )

    if st.button(
        "Generar nueva versión",
        type="primary",
        use_container_width=True,
        key=f"generate_revision_v{version}",
    ):
        if not revision_request.strip():
            st.error("Describa los cambios estructurales que desea realizar.")
        else:
            api_key = get_api_key_runtime()
            if not api_key:
                st.error("Falta GEMINI_API_KEY en Streamlit Secrets.")
            else:
                with st.spinner("Analizando cambios y generando nueva versión..."):
                    try:
                        revision = revisar_presupuesto_ia(
                            api_key=api_key,
                            model_name=model_name,
                            project_data=g["project_data"],
                            params=g["params"],
                            current_result=result,
                            current_items=items,
                            revision_request=revision_request.strip(),
                        )

                        revised_result, revised_items, change_log = aplicar_revision_estructural(
                            db=db,
                            current_result=result,
                            current_items=items,
                            revision=revision,
                            project_data=g["project_data"],
                            params=g["params"],
                        )

                        revised_financials = calcular_financieros(
                            revised_items,
                            g["params"],
                        )

                        if g.get("simulation"):
                            new_budget_id = None
                            new_version = version + 1
                        else:
                            if not g.get("project_id") or not g.get("budget_id"):
                                raise RuntimeError(
                                    "La generación actual no tiene referencias persistentes válidas."
                                )
                            new_budget_id, new_version = db.save_revision(
                                project_id=g["project_id"],
                                parent_budget_id=g["budget_id"],
                                result=revised_result,
                                items=revised_items,
                                params=g["params"],
                                financials=revised_financials,
                                revision_instruction=revision_request.strip(),
                            )

                        new_file_code = f"{g['project_code']}-V{new_version:02d}"
                        excel_bytes = crear_excel(
                            project_code=g["project_code"],
                            project_data=g["project_data"],
                            result=revised_result,
                            items=revised_items,
                            params=g["params"],
                            version=new_version,
                        )
                        txt_bytes = crear_txt(
                            project_code=g["project_code"],
                            project_data=g["project_data"],
                            result=revised_result,
                            items=revised_items,
                            params=g["params"],
                            financials=revised_financials,
                            version=new_version,
                        )
                        zip_bytes = crear_paquete_zip(
                            new_file_code,
                            excel_bytes,
                            txt_bytes,
                        )

                        history = list(g.get("revision_history") or [])
                        history.append(
                            {
                                "version": new_version,
                                "request": revision_request.strip(),
                                "summary": revision.resumen_revision,
                                "changes": change_log,
                            }
                        )

                        g.update(
                            {
                                "budget_id": new_budget_id if not g.get("simulation") else g.get("budget_id"),
                                "version": new_version,
                                "file_code": new_file_code,
                                "result": revised_result.model_dump(),
                                "items": revised_items,
                                "financials": revised_financials,
                                "excel_bytes": excel_bytes,
                                "txt_bytes": txt_bytes,
                                "zip_bytes": zip_bytes,
                                "revision_history": history,
                            }
                        )
                        st.session_state["generated"] = g
                        st.success(f"Nueva versión V{new_version:02d} generada.")
                        st.rerun()

                    except Exception as exc:
                        st.exception(exc)

    # -----------------------------------------------------
    # Acciones persistentes / destructivas
    # -----------------------------------------------------
    st.divider()
    st.subheader("Acciones sobre esta generación")

    if g.get("simulation"):
        st.caption(
            "Esta corrida sigue siendo una simulación. Puede guardar la versión actual "
            "como proyecto real sin volver a ejecutar Gemini."
        )

        ac1, ac2 = st.columns(2)
        with ac1:
            if st.button(
                "Guardar versión actual en la base de datos",
                type="primary",
                use_container_width=True,
                key=f"promote_simulation_v{version}",
            ):
                try:
                    real_code = db.next_project_code(
                        g["project_data"]["name"],
                        g["project_data"]["location"],
                    )
                    project_id, budget_id = db.save_generation(
                        project_code=real_code,
                        project_data=g["project_data"],
                        result=result,
                        items=items,
                        params=g["params"],
                        financials=financials,
                    )

                    # Al persistir una simulación se convierte en V01 real. Las
                    # revisiones previas de simulación no habían sido persistidas.
                    real_version = 1
                    real_file_code = f"{real_code}-V{real_version:02d}"
                    excel_bytes = crear_excel(
                        project_code=real_code,
                        project_data=g["project_data"],
                        result=result,
                        items=items,
                        params=g["params"],
                        version=real_version,
                    )
                    txt_bytes = crear_txt(
                        project_code=real_code,
                        project_data=g["project_data"],
                        result=result,
                        items=items,
                        params=g["params"],
                        financials=financials,
                        version=real_version,
                    )
                    zip_bytes = crear_paquete_zip(
                        real_file_code,
                        excel_bytes,
                        txt_bytes,
                    )

                    g.update(
                        {
                            "project_id": project_id,
                            "budget_id": budget_id,
                            "simulation": False,
                            "project_code": real_code,
                            "file_code": real_file_code,
                            "version": real_version,
                            "excel_bytes": excel_bytes,
                            "txt_bytes": txt_bytes,
                            "zip_bytes": zip_bytes,
                            "revision_history": [],
                        }
                    )
                    st.session_state["generated"] = g
                    st.success("La versión actual fue guardada como proyecto real.")
                    st.rerun()
                except Exception as exc:
                    st.exception(exc)

        with ac2:
            if st.button(
                "Descartar simulación",
                use_container_width=True,
                key=f"discard_simulation_v{version}",
            ):
                st.session_state.pop("generated", None)
                st.rerun()

    else:
        st.caption(
            "Las revisiones crean nuevas versiones y conservan las anteriores. "
            "La siguiente acción elimina únicamente la versión actualmente mostrada."
        )
        current_delete_key = st.text_input(
            "Clave de eliminación",
            type="password",
            key=f"current_generation_delete_key_v{version}",
        )

        if st.button(
            "Eliminar esta versión de la base",
            use_container_width=True,
            key=f"delete_current_generation_v{version}",
        ):
            if not validar_clave_borrado(current_delete_key):
                st.error("Clave incorrecta o no configurada.")
            elif not g.get("project_id") or not g.get("budget_id"):
                st.error("No se encontró la referencia guardada de esta versión.")
            else:
                db.delete_generation(g["project_id"], g["budget_id"])
                st.session_state.pop("generated", None)
                st.success("La versión actual fue eliminada.")
                st.rerun()
