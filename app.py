import os
import re
import json
import time
import random
import sqlite3
import unicodedata
import uuid
import zipfile
import traceback
from html import escape
from concurrent.futures import ThreadPoolExecutor, wait
from datetime import datetime, timezone
from difflib import SequenceMatcher
from io import BytesIO, StringIO
from pathlib import Path
from collections.abc import Mapping

import pandas as pd
import streamlit as st
from google import genai
from google.genai import types
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from pydantic import BaseModel, Field

try:
    import psycopg
    from psycopg.rows import dict_row
except Exception:
    psycopg = None
    dict_row = None

try:
    from streamlit_local_storage import LocalStorage
except Exception:
    LocalStorage = None


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


def normalizar_texto(texto) -> str:
    """
    Normaliza cualquier valor recibido, no solamente cadenas.

    Al leer archivos Excel antiguos pueden aparecer números, booleanos, fechas,
    resultados de fórmulas o valores vacíos en las mismas filas que se recorren
    buscando encabezados. Convertir primero a texto evita errores como:
    AttributeError: 'float' object has no attribute 'strip'
    """
    if texto is None:
        texto = ""
    elif not isinstance(texto, str):
        texto = str(texto)

    texto = texto.strip().lower()
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


SECCIONES_COMERCIALES_PREFERENTES = [
    # Orden de presentación basado en la secuencia normal de preparación y obra.
    "PROYECTO Y TRÁMITES",
    "PRELIMINARES Y PROTECCIONES",
    "DESMONTAJES Y DEMOLICIONES",
    "ALBAÑILERÍA Y ESTRUCTURA",
    "INSTALACIONES HIDROSANITARIAS",
    "INSTALACIONES ELÉCTRICAS",
    "ACABADOS Y RECUBRIMIENTOS",
    "CARPINTERÍA",
    "CANCELERÍA Y HERRERÍA",
    "EXTERIORES Y AMENIDADES",
    "OTROS TRABAJOS",
    "LIMPIEZA Y ENTREGA",
]

SECCION_ALIASES = {
    "PROYECTO Y TRAMITES": "PROYECTO Y TRÁMITES",
    "DEMOLICION Y DESMONTAJE": "DESMONTAJES Y DEMOLICIONES",
    "DEMOLICION Y DESMONTAJES": "DESMONTAJES Y DEMOLICIONES",
    "CARPINTERIA Y MOBILIARIO": "CARPINTERÍA",
    "PROYECTO": "PROYECTO Y TRÁMITES",
    "TRAMITES": "PROYECTO Y TRÁMITES",
    "PRELIMINARES": "PRELIMINARES Y PROTECCIONES",
    "PRELIMINARES Y PROTECCIONES": "PRELIMINARES Y PROTECCIONES",
    "PROTECCIONES": "PRELIMINARES Y PROTECCIONES",
    "PREPARACION": "PRELIMINARES Y PROTECCIONES",
    "PREPARACION Y DEMOLICIONES": "DESMONTAJES Y DEMOLICIONES",
    "PRELIMINARES Y DEMOLICIONES": "DESMONTAJES Y DEMOLICIONES",
    "DEMOLICIONES": "DESMONTAJES Y DEMOLICIONES",
    "DESMONTAJES": "DESMONTAJES Y DEMOLICIONES",
    "ALBANILERIA": "ALBAÑILERÍA Y ESTRUCTURA",
    "ESTRUCTURA": "ALBAÑILERÍA Y ESTRUCTURA",
    "ALBANILERIA Y ESTRUCTURA": "ALBAÑILERÍA Y ESTRUCTURA",
    "ELECTRICA": "INSTALACIONES ELÉCTRICAS",
    "INSTALACION ELECTRICA": "INSTALACIONES ELÉCTRICAS",
    "INSTALACIONES ELECTRICAS": "INSTALACIONES ELÉCTRICAS",
    "HIDROSANITARIA": "INSTALACIONES HIDROSANITARIAS",
    "INSTALACIONES HIDROSANITARIAS": "INSTALACIONES HIDROSANITARIAS",
    "PLOMERIA": "INSTALACIONES HIDROSANITARIAS",
    "ACABADOS": "ACABADOS Y RECUBRIMIENTOS",
    "ACABADOS Y RECUBRIMIENTOS": "ACABADOS Y RECUBRIMIENTOS",
    "CARPINTERIA": "CARPINTERÍA",
    "HERRERIA Y CANCELERIA": "CANCELERÍA Y HERRERÍA",
    "CANCELERIA Y HERRERIA": "CANCELERÍA Y HERRERÍA",
    "EXTERIORES": "EXTERIORES Y AMENIDADES",
    "EXTERIORES Y AMENIDADES": "EXTERIORES Y AMENIDADES",
    "LIMPIEZA": "LIMPIEZA Y ENTREGA",
    "LIMPIEZA FINAL": "LIMPIEZA Y ENTREGA",
    "LIMPIEZA Y ENTREGA": "LIMPIEZA Y ENTREGA",
}


def normalizar_seccion_comercial(value: str) -> str:
    raw = str(value or "").strip()
    if not raw:
        return "OTROS TRABAJOS"
    key = normalizar_texto(raw).upper()
    return SECCION_ALIASES.get(key, raw.upper())


def titulo_comercial_item(item: dict) -> str:
    value = str(item.get("commercial_title") or "").strip()
    if value:
        return value
    fallback = str(item.get("subcategory") or "").strip()
    if fallback:
        return fallback
    description = str(item.get("description") or "").strip()
    if not description:
        return "Concepto"
    first = re.split(r"[.;:]", description, maxsplit=1)[0].strip()
    return first[:80] or description[:80]


def seccion_ejecucion_item(item: dict) -> str:
    """
    Clasificación final de seguridad.

    La clasificación principal la revisa Gemini viendo el presupuesto completo.
    Python solo corrige dos casos inequívocos de secuencia:
    - protección TEMPORAL de áreas de obra -> preliminares;
    - limpieza FINA/FINAL de entrega -> cierre.

    No se usan palabras aisladas como "protección", porque pueden describir una
    propiedad técnica de un elemento permanente y provocar falsos positivos.
    """
    category = normalizar_seccion_comercial(item.get("category"))

    title_context = normalizar_texto(
        " ".join(
            str(item.get(k) or "")
            for k in ("subcategory", "commercial_title")
        )
    ).upper()

    preliminary_phrases = (
        "PROTECCION DE AREAS",
        "PROTECCION DE AREA",
        "PROTECCION DE PISOS",
        "PROTECCION DE PISO",
        "PROTECCION DE ACCESOS",
        "PROTECCION DE MOBILIARIO",
        "PROTECCION TEMPORAL",
        "PROTECCIONES TEMPORALES",
        "CUBRIR AREAS",
        "CUBRIR PISOS",
        "TAPIAL DE OBRA",
        "TAPIALES DE OBRA",
        "TRAZO Y REPLANTEO",
        "REPLANTEO",
    )
    if any(phrase in title_context for phrase in preliminary_phrases):
        return "PRELIMINARES Y PROTECCIONES"

    closing_phrases = (
        "LIMPIEZA FINA",
        "LIMPIEZA FINAL",
        "LIMPIEZA DE ENTREGA",
        "LIMPIEZA Y ENTREGA",
        "ASEO FINAL",
        "ENTREGA FINAL",
        "CIERRE DE OBRA",
    )
    if any(phrase in title_context for phrase in closing_phrases):
        return "LIMPIEZA Y ENTREGA"

    return category


def ordenar_items_comercialmente(items: list[dict]) -> list[dict]:
    """
    Ordena por fase macro de obra y después por orden_ejecucion auditado.
    """
    preferred = {
        name: index for index, name in enumerate(SECCIONES_COMERCIALES_PREFERENTES)
    }

    def key(pair):
        idx, item = pair
        section = seccion_ejecucion_item(item)
        try:
            execution_order = int(item.get("execution_order") or 500)
        except (TypeError, ValueError):
            execution_order = 500
        phase = preferred.get(
            section,
            preferred.get("OTROS TRABAJOS", 10),
        )
        return (phase, execution_order, idx)

    ordered = [dict(item) for _, item in sorted(enumerate(items), key=key)]
    for item in ordered:
        item["category"] = seccion_ejecucion_item(item)
    return ordered


def nombre_partida_excel(section: str) -> str:
    """Nombre breve de partida para el Excel, en orden de ejecución."""
    section = normalizar_seccion_comercial(section)
    preferred_names = {
        "PROYECTO Y TRÁMITES": "Trámites",
        "PRELIMINARES Y PROTECCIONES": "Preliminares y Protecciones",
        "DESMONTAJES Y DEMOLICIONES": "Desmontajes y Demoliciones",
        "ALBAÑILERÍA Y ESTRUCTURA": "Albañilería y Estructura",
        "INSTALACIONES HIDROSANITARIAS": "Instalaciones Hidrosanitarias",
        "INSTALACIONES ELÉCTRICAS": "Instalaciones Eléctricas",
        "ACABADOS Y RECUBRIMIENTOS": "Acabados",
        "CARPINTERÍA": "Carpintería",
        "CANCELERÍA Y HERRERÍA": "Cancelería y Herrería",
        "EXTERIORES Y AMENIDADES": "Exteriores y Amenidades",
        "LIMPIEZA Y ENTREGA": "Limpieza y Entrega",
        "OTROS TRABAJOS": "Otros Trabajos",
    }
    if section in preferred_names:
        return preferred_names[section]
    return section.title()


def nombre_subpartida_excel(item: dict) -> str:
    """
    La subpartida del Excel debe ser corta: Pisos, Muros, Frentes, Barra, etc.
    Para datos históricos se usa título comercial como respaldo.
    """
    value = str(item.get("subcategory") or "").strip()
    if value:
        return value
    return titulo_comercial_item(item)


def estructura_partidas_excel(items: list[dict]) -> list[dict]:
    """
    Asigna numeración jerárquica estable según el orden comercial:
      1. Trámites       / 1.1 Licencias
      2. Acabados       / 2.1 Pisos
      3. Carpintería    / 3.1 Frentes
    La numeración se genera en Python y no se deja a criterio de Gemini.
    """
    ordered = ordenar_items_comercialmente(items)
    section_numbers = {}
    section_counts = {}
    output = []

    for item in ordered:
        section = seccion_ejecucion_item(item)

        if section not in section_numbers:
            section_numbers[section] = len(section_numbers) + 1
            section_counts[section] = 0

        part_num = section_numbers[section]
        section_counts[section] += 1
        sub_num = section_counts[section]

        enriched = dict(item)
        enriched["part_number"] = part_num
        enriched["subpart_number"] = sub_num
        enriched["partida_excel"] = f"{part_num}. {nombre_partida_excel(section)}"
        enriched["subpartida_excel"] = (
            f"{part_num}.{sub_num} {nombre_subpartida_excel(item)}"
        )
        output.append(enriched)

    return output


def codigo_capitulo_cliente(part_number: int) -> int:
    """Código numérico de capítulo (capítulo * 1000), esquema del ejemplo."""
    return int(part_number) * 1000


def codigo_partida_cliente(part_number: int, subpart_number: int) -> int:
    """Código numérico de partida (capítulo * 1000 + subpartida)."""
    return int(part_number) * 1000 + int(subpart_number)


def asignar_codigos_jerarquicos(items: list[dict]) -> list[dict]:
    """
    Sustituye el código de cada item (antes alfanumérico: CON-001, IMP-001,
    MAN-001, etc.) por el nuevo esquema numérico jerárquico -- capítulo * 1000
    + subpartida --, el mismo que usa la columna Código del Excel formato
    cliente. Se calcula una sola vez a partir del orden comercial (el mismo
    que ya usan ambos Excels vía estructura_partidas_excel) y queda grabado
    en item["code"], de modo que también es lo que se guarda en la base de
    datos (concepts.code, budget_items.code) de aquí en adelante.

    No se basa en la posición dentro de la lista `items`, sino en el código
    previo de cada item, así que el orden original de `items` se conserva.
    """
    structured = estructura_partidas_excel(items)
    nuevo_codigo = {
        it["code"]: str(codigo_partida_cliente(it["part_number"], it["subpart_number"]))
        for it in structured
    }
    actualizados = []
    for it in items:
        nuevo = dict(it)
        original = it.get("code")
        if original in nuevo_codigo:
            nuevo["code"] = nuevo_codigo[original]
        actualizados.append(nuevo)
    return actualizados


def descripcion_excel_item(item: dict) -> str:
    """
    El ejemplo de la empresa coloca una descripción técnica amplia en una sola
    celda. Anteponemos el título comercial cuando aporta contexto y no está ya
    incluido al inicio de la descripción.
    """
    title = titulo_comercial_item(item).strip()
    description = str(item.get("description") or "").strip()

    if not title:
        return description
    if not description:
        return title

    norm_title = normalizar_texto(title).upper()
    norm_desc = normalizar_texto(description).upper()
    if norm_desc.startswith(norm_title):
        return description
    return f"{title}. {description}"


AREA_GENERAL = "General"

# Margen fijo que se aplica sobre el Importe interno para calcular el Precio
# del Excel formato cliente (hoja Partidas). Ya no es configurable desde la
# interfaz: es una constante de negocio.
MARGEN_PRESUPUESTO_CLIENTE_PCT = 30.0


PATRONES_AREAS_EXPLICITAS = [
    r"\broof\s*garden\b",
    r"\b(?:cuarto|área|area)\s+de\s+lavado\b",
    r"\b(?:walk[\s-]*in\s+closet|walking\s+closet|vestidor)\b",
    r"\bbañ(?:o|os)(?:\s+(?:principal|de\s+visitas|visitas|social|secundario|[1-9]))?\b",
    r"\bcocina\b",
    r"\bsala(?:\s*(?:/|y)\s*comedor)?\b",
    r"\bcomedor\b",
    r"\brec[aá]mara(?:\s+(?:principal|secundaria|[1-9]))?\b",
    r"\bhabitaci[oó]n(?:\s+(?:principal|secundaria|[1-9]))?\b",
    r"\bestudio(?:\s+[1-9])?\b",
    r"\bpatio(?:\s+(?:trasero|posterior|frontal|delantero))?\b",
    r"\bazotea\b",
    r"\bterraza\b",
    r"\bbalc[oó]n\b",
    r"\bfachada\b",
    r"\bescalera\b",
    r"\blavander[ií]a\b",
    r"\bbodega\b",
    r"\bestacionamiento(?:\s+[1-9])?\b",
    r"\bcochera(?:\s+[1-9])?\b",
    r"\bjard[ií]n\b",
    r"\bpasillo\b",
]


def normalizar_nombre_area(value: str) -> str:
    raw = re.sub(r"\s+", " ", str(value or "").strip())
    if not raw:
        return AREA_GENERAL

    key = normalizar_texto(raw).upper()
    if key in {
        "GENERAL", "GENERALES", "AREA GENERAL", "AREAS GENERALES",
        "TODA LA OBRA", "TODO EL PROYECTO",
    }:
        return AREA_GENERAL

    display = raw.replace("area ", "Área ").replace("Area ", "Área ")
    words = []
    for word in display.split():
        if any(ch.isdigit() for ch in word):
            words.append(word)
        elif word.upper() in {"UV"}:
            words.append(word.upper())
        else:
            words.append(word[:1].upper() + word[1:].lower())
    return " ".join(words)


def tipo_base_area(area: str) -> str:
    key = normalizar_texto(area).upper()
    equivalencias = [
        ("BANO", "BANO"), ("COCINA", "COCINA"), ("SALA", "SALA"),
        ("COMEDOR", "COMEDOR"), ("RECAMARA", "DORMITORIO"),
        ("HABITACION", "DORMITORIO"), ("ESTUDIO", "ESTUDIO"),
        ("PATIO", "PATIO"), ("AZOTEA", "AZOTEA"),
        ("ROOF GARDEN", "ROOF GARDEN"), ("TERRAZA", "TERRAZA"),
        ("BALCON", "BALCON"), ("FACHADA", "FACHADA"),
        ("ESCALERA", "ESCALERA"), ("CUARTO DE LAVADO", "LAVADO"),
        ("AREA DE LAVADO", "LAVADO"), ("LAVANDERIA", "LAVADO"),
        ("BODEGA", "BODEGA"), ("WALK IN CLOSET", "VESTIDOR"),
        ("WALKING CLOSET", "VESTIDOR"), ("VESTIDOR", "VESTIDOR"),
        ("ESTACIONAMIENTO", "ESTACIONAMIENTO"), ("COCHERA", "ESTACIONAMIENTO"),
        ("JARDIN", "JARDIN"), ("PASILLO", "PASILLO"),
    ]
    for prefix, base in equivalencias:
        if key.startswith(prefix):
            return base
    return key


def detectar_areas_explicitas(project_data: dict) -> list[str]:
    """Detecta áreas solo en la descripción original escrita por el usuario."""
    source = str(project_data.get("description") or "")
    if not source.strip():
        return []

    matches = []
    occupied = []
    for pattern in PATRONES_AREAS_EXPLICITAS:
        for match in re.finditer(pattern, source, flags=re.I):
            span = match.span()
            if any(not (span[1] <= a or span[0] >= b) for a, b in occupied):
                continue
            occupied.append(span)
            matches.append((span[0], normalizar_nombre_area(match.group(0))))

    matches.sort(key=lambda x: x[0])
    output, seen = [], set()
    for _, area in matches:
        key = normalizar_texto(area).upper()
        if key not in seen:
            seen.add(key)
            output.append(area)
    return output


def _texto_item_para_area(item: dict) -> str:
    return " ".join(
        str(item.get(key) or "")
        for key in (
            "subcategory", "commercial_title", "description",
            "quantity_criterion", "inclusion_basis",
        )
    )


def _areas_mencionadas_en_item(project_data: dict, item: dict) -> list[str]:
    areas = detectar_areas_explicitas(project_data)
    if not areas:
        return []

    item_key = normalizar_texto(_texto_item_para_area(item)).upper()
    exact = []
    for area in areas:
        area_key = normalizar_texto(area).upper()
        if area_key and area_key in item_key:
            exact.append(area)
    if exact:
        return exact

    by_type = {}
    for area in areas:
        by_type.setdefault(tipo_base_area(area), []).append(area)

    generic = []
    tokens_by_type = {
        "BANO": ("BANO",), "COCINA": ("COCINA",), "SALA": ("SALA",),
        "COMEDOR": ("COMEDOR",), "DORMITORIO": ("RECAMARA", "HABITACION"),
        "ESTUDIO": ("ESTUDIO",), "PATIO": ("PATIO",), "AZOTEA": ("AZOTEA",),
        "ROOF GARDEN": ("ROOF GARDEN",), "TERRAZA": ("TERRAZA",),
        "BALCON": ("BALCON",), "FACHADA": ("FACHADA",),
        "ESCALERA": ("ESCALERA",), "LAVADO": ("LAVADO", "LAVANDERIA"),
        "BODEGA": ("BODEGA",), "VESTIDOR": ("VESTIDOR", "CLOSET"),
        "ESTACIONAMIENTO": ("ESTACIONAMIENTO", "COCHERA"),
        "JARDIN": ("JARDIN",), "PASILLO": ("PASILLO",),
    }
    for base, candidates in by_type.items():
        if len(candidates) != 1:
            continue
        if any(token in item_key for token in tokens_by_type.get(base, (base,))):
            generic.append(candidates[0])
    return generic


def _m2_explicitos_cerca_de_area(project_data: dict, area: str) -> float | None:
    source = str(project_data.get("description") or "")
    source_norm = normalizar_texto(source).upper()
    area_norm = normalizar_texto(area).upper()
    start = source_norm.find(area_norm)
    if start < 0:
        return None

    # Cortar la búsqueda antes de la siguiente área explícita. Así una medida de
    # Cocina no puede terminar asignándose accidentalmente a Baño 1, por ejemplo.
    next_starts = []
    for other in detectar_areas_explicitas(project_data):
        other_norm = normalizar_texto(other).upper()
        pos = source_norm.find(other_norm, start + len(area_norm))
        if pos > start:
            next_starts.append(pos)
    end = min(next_starts) if next_starts else min(len(source), start + 120)
    end = min(end, start + 120)
    window = source[start:end]

    direct = re.search(
        r"(\d+(?:[.,]\d+)?)\s*(?:m2|m²|metros?\s+cuadrados?)",
        window, flags=re.I,
    )
    if direct:
        return float(direct.group(1).replace(",", "."))

    dims = re.search(
        r"(\d+(?:[.,]\d+)?)\s*[x×]\s*(\d+(?:[.,]\d+)?)\s*m\b",
        window, flags=re.I,
    )
    if dims:
        return float(dims.group(1).replace(",", ".")) * float(dims.group(2).replace(",", "."))
    return None


def asignar_areas_deterministicamente(project_data: dict, item: dict) -> list[dict]:
    """
    1) una sola área explícita -> 100 %;
    2) varias áreas + concepto M2 + m² explícitos conciliables -> proporcional;
    3) cualquier caso ambiguo -> General 100 %.
    """
    matches = _areas_mencionadas_en_item(project_data, item)

    if len(matches) == 1:
        return [{
            "area": matches[0], "porcentaje": 100.0,
            "cantidad_referencia": float(item.get("quantity") or 0.0),
            "criterio": "Área indicada explícitamente en el alcance del concepto.",
            "confianza": "Alta",
        }]

    if len(matches) > 1 and normalizar_unidad(item.get("unit")) == "M2":
        measures = {area: _m2_explicitos_cerca_de_area(project_data, area) for area in matches}
        if all(v is not None and v > 0 for v in measures.values()):
            total_measure = sum(measures.values())
            item_qty = float(item.get("quantity") or 0.0)
            tolerance = max(1.0, item_qty * 0.15)
            if item_qty > 0 and abs(total_measure - item_qty) <= tolerance:
                result, accumulated = [], 0.0
                for idx, area in enumerate(matches):
                    if idx == len(matches) - 1:
                        pct = max(100.0 - accumulated, 0.0)
                    else:
                        pct = round(measures[area] / total_measure * 100.0, 6)
                        accumulated += pct
                    result.append({
                        "area": area, "porcentaje": pct,
                        "cantidad_referencia": measures[area],
                        "criterio": "Reparto calculado con m² explícitos del texto inicial.",
                        "confianza": "Alta",
                    })
                return result

    return [{
        "area": AREA_GENERAL, "porcentaje": 100.0,
        "cantidad_referencia": None,
        "criterio": "No existe una asignación por área verificable con los datos iniciales.",
        "confianza": "Alta",
    }]


def normalizar_asignaciones_area(asignaciones) -> list[dict]:
    if isinstance(asignaciones, str):
        try:
            asignaciones = json.loads(asignaciones)
        except Exception:
            asignaciones = []

    rows = []
    for raw in asignaciones or []:
        if hasattr(raw, "model_dump"):
            raw = raw.model_dump()
        if not isinstance(raw, dict):
            continue
        try:
            pct = max(float(raw.get("porcentaje") or 0.0), 0.0)
        except (TypeError, ValueError):
            pct = 0.0
        if pct <= 0:
            continue
        rows.append({
            "area": normalizar_nombre_area(raw.get("area")),
            "porcentaje": pct,
            "cantidad_referencia": raw.get("cantidad_referencia"),
            "criterio": str(raw.get("criterio") or ""),
            "confianza": str(raw.get("confianza") or "Alta"),
        })

    if not rows:
        return [{
            "area": AREA_GENERAL, "porcentaje": 100.0,
            "cantidad_referencia": None,
            "criterio": "Sin asignación verificable.", "confianza": "Alta",
        }]

    total, running = sum(x["porcentaje"] for x in rows), 0.0
    for idx, row in enumerate(rows):
        if idx == len(rows) - 1:
            row["porcentaje"] = max(100.0 - running, 0.0)
        else:
            row["porcentaje"] = round(row["porcentaje"] / total * 100.0, 6)
            running += row["porcentaje"]
    return rows


def obtener_asignaciones_area_item(item: dict) -> list[dict]:
    if item.get("area_allocations"):
        return normalizar_asignaciones_area(item.get("area_allocations"))
    if item.get("area_allocations_json"):
        return normalizar_asignaciones_area(item.get("area_allocations_json"))
    return normalizar_asignaciones_area([])


def recalcular_areas_items(project_data: dict, items: list[dict]) -> list[dict]:
    output = []
    for item in items:
        out = dict(item)
        raw_hint = str(out.get("area_hint") or "").strip()
        if raw_hint:
            out["area_allocations"] = [{
                "area": normalizar_nombre_area(raw_hint),
                "porcentaje": 100.0,
                "cantidad_referencia": float(out.get("quantity") or 0.0),
                "criterio": "Área específica indicada en la actividad.",
                "confianza": "Alta",
            }]
        elif out.get("area_allocations") or out.get("area_allocations_json"):
            # Preserva el área ya guardada en BD o reconstruida desde Excel.
            out["area_allocations"] = obtener_asignaciones_area_item(out)
        else:
            # Compatibilidad con presupuestos antiguos que todavía no tenían Área.
            out["area_allocations"] = asignar_areas_deterministicamente(project_data, out)
        output.append(out)
    return output


def area_excel_item(item: dict) -> str:
    asignaciones = obtener_asignaciones_area_item(item)
    if len(asignaciones) == 1:
        return normalizar_nombre_area(asignaciones[0].get("area"))
    return AREA_GENERAL


def descripcion_areas_item(item: dict) -> str:
    return " · ".join(x["area"] for x in obtener_asignaciones_area_item(item))


def item_esta_incluido(item: dict) -> bool:
    """Compatibilidad entre presupuestos nuevos, antiguos y recargados desde Excel."""
    value = item.get("included", True)
    if value is None:
        return True
    if isinstance(value, str):
        return normalizar_texto(value).upper() not in {"NO", "0", "FALSE", "FALSO"}
    return bool(value)


NIVELES_PRESUPUESTO = ["Económico", "Medio", "Medio-alto", "Alto"]


def criterio_nivel_presupuesto(nivel: str) -> str:
    criterios = {
        "Económico": (
            "Selecciona materiales y acabados comerciales de costo contenido, funcionales y durables. "
            "El nivel NO debe abaratar artificialmente demolición, albañilería, instalaciones, trámites, "
            "limpieza, acarreos ni otros trabajos cuyo costo dependa principalmente de mano de obra, "
            "cantidad, permisos o condiciones de ejecución. Solo reduce especificaciones de acabado "
            "cuando el proyecto lo permita."
        ),
        "Medio": (
            "Selecciona materiales, acabados y soluciones comerciales de calidad media. "
            "Mantén los costos de trabajos base relativamente independientes del nivel y modifica "
            "principalmente acabados, accesorios, herrajes y especificaciones que sí cambian con la calidad."
        ),
        "Medio-alto": (
            "Selecciona acabados, materiales, herrajes y soluciones de buena calidad, superiores al promedio. "
            "No apliques un incremento general al presupuesto: el nivel debe reflejarse principalmente "
            "en los elementos cuyo costo realmente cambia por calidad, especificación o complejidad."
        ),
        "Alto": (
            "Selecciona acabados, materiales, herrajes y soluciones de gama alta cuando sean coherentes "
            "con el proyecto. Puede requerir proveedores especializados. No incrementes automáticamente "
            "demoliciones, albañilería, trámites, limpieza, acarreos o trabajos base si la especificación "
            "no cambia su costo real."
        ),
    }
    return criterios.get(nivel, criterios["Medio-alto"])


def normalizar_composicion_costos(
    materiales_pct: float,
    mano_obra_pct: float,
    otros_pct: float,
) -> tuple[float, float, float]:
    """Normaliza la composición estimada para que sume 100 %."""
    material = max(float(materiales_pct or 0), 0.0)
    labor = max(float(mano_obra_pct or 0), 0.0)
    other = max(float(otros_pct or 0), 0.0)
    total = material + labor + other

    if total <= 0:
        # Para datos históricos sin desglose preferimos no inventar.
        return 0.0, 0.0, 100.0

    return (
        material / total * 100.0,
        labor / total * 100.0,
        other / total * 100.0,
    )


def aplicar_composicion_costo(item: dict) -> dict:
    """
    Calcula una descomposición informativa del costo integrado.
    No modifica el costo ni agrega nuevamente el desperdicio.
    """
    out = dict(item)
    material_pct, labor_pct, other_pct = normalizar_composicion_costos(
        out.get("material_share_pct", 0.0),
        out.get("labor_share_pct", 0.0),
        out.get("other_share_pct", 0.0),
    )
    waste_pct = max(float(out.get("waste_reference_pct", 0.0) or 0.0), 0.0)
    unit_cost = max(float(out.get("unit_cost", 0.0) or 0.0), 0.0)

    out["material_share_pct"] = material_pct
    out["labor_share_pct"] = labor_pct
    out["other_share_pct"] = other_pct
    out["waste_reference_pct"] = waste_pct
    out["material_unit_est"] = unit_cost * material_pct / 100.0
    out["labor_unit_est"] = unit_cost * labor_pct / 100.0
    out["other_unit_est"] = unit_cost * other_pct / 100.0
    out["waste_reference_unit"] = out["material_unit_est"] * waste_pct / 100.0
    return out


def normalizar_unidad(unidad: str) -> str:
    # Los superíndices deben convertirse antes de normalizar texto; de lo
    # contrario "m²" podría quedar reducido a "m".
    raw = str(unidad or "").strip()
    raw = raw.replace("²", "2").replace("³", "3")
    value = normalizar_texto(raw).upper()
    value = (
        value.replace("M²", "M2")
        .replace("M³", "M3")
        .replace("MTS2", "M2")
        .replace("MTS3", "M3")
        .replace("METROS CUADRADOS", "M2")
        .replace("METRO CUADRADO", "M2")
        .replace("METROS CUBICOS", "M3")
        .replace("METRO CUBICO", "M3")
        .replace("METROS LINEALES", "ML")
        .replace("METRO LINEAL", "ML")
        .replace("PIEZAS", "PZA")
        .replace("PIEZA", "PZA")
        .replace("PZAS", "PZA")
        .replace("PZ", "PZA")
        .replace("KILOGRAMOS", "KG")
        .replace("KILOGRAMO", "KG")
        .replace("TONELADAS", "TON")
        .replace("TONELADA", "TON")
        .replace("LITROS", "L")
        .replace("LITRO", "L")
    )
    value = re.sub(r"[^A-Z0-9%]+", "", value)
    aliases = {
        "M2": "M2",
        "M3": "M3",
        "ML": "ML",
        "M": "M",
        "PZA": "PZA",
        "PZA.": "PZA",
        "KG": "KG",
        "TON": "TON",
        "L": "L",
        "LT": "L",
        "H": "H",
        "HR": "H",
        "HRA": "H",
        "HORA": "H",
        "DIA": "DIA",
        "MES": "MES",
        "LOTE": "LOTE",
        "JGO": "JGO",
        "JUEGO": "JGO",
        "PTO": "PTO",
        "PUNTO": "PTO",
        "SERV": "SERV",
        "SERVICIO": "SERV",
        "%": "%",
    }
    return aliases.get(value, value[:16])


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
        allowed_columns = {
            "parent_budget_id",
            "revision_instruction",
            "commercial_title",
            "budget_level",
            "material_share_pct",
            "labor_share_pct",
            "other_share_pct",
            "waste_reference_pct",
            "execution_order",
            "area_allocations_json",
            "included",
            "contract_lot",
        }
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
                budget_level TEXT,
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
                commercial_title TEXT,
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
                material_share_pct REAL,
                labor_share_pct REAL,
                other_share_pct REAL,
                waste_reference_pct REAL,
                execution_order INTEGER,
                area_allocations_json TEXT,
                included INTEGER NOT NULL DEFAULT 1,
                contract_lot TEXT,
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
        self._ensure_column("budget_items", "commercial_title", "TEXT")
        self._ensure_column("projects", "budget_level", "TEXT")
        self._ensure_column("budget_items", "material_share_pct", "REAL")
        self._ensure_column("budget_items", "labor_share_pct", "REAL")
        self._ensure_column("budget_items", "other_share_pct", "REAL")
        self._ensure_column("budget_items", "waste_reference_pct", "REAL")
        self._ensure_column("budget_items", "execution_order", "INTEGER")
        self._ensure_column("budget_items", "area_allocations_json", "TEXT")
        self._ensure_column("budget_items", "included", "INTEGER NOT NULL DEFAULT 1")
        self._ensure_column("budget_items", "contract_lot", "TEXT")

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
                id, code, name, project_type, budget_level, location, dimension_mode,
                dimensions_text, description, guide_text, main_activity, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                project_id,
                project_code,
                project_data["name"],
                project_data["project_type"],
                project_data.get("budget_level", "Medio-alto"),
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
            concept_was_existing = bool(concept_id)

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
                        item.get("concepto_base", item["description"]),
                        item["unit"],
                        normalizar_texto(item.get("concepto_base", item["description"])),
                        budget_id,
                        created,
                    ),
                )

                # Solo se crea historial nuevo cuando el precio fue generado/estimado
                # por la valuación actual. Un precio interno reutilizado ya cuenta
                # con historial propio.
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

            if concept_was_existing and item.get("record_new_price"):
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
                    commercial_title, description, unit, quantity, unit_cost, direct_amount,
                    unit_indirect, unit_profit, unit_sale, sale_amount,
                    sale_margin_pct, benefit_amount, price_source,
                    price_source_detail, price_confidence,
                    material_share_pct, labor_share_pct, other_share_pct, waste_reference_pct,
                    execution_order, area_allocations_json, included, contract_lot, quantity_criterion,
                    inclusion_basis, considerations, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    str(uuid.uuid4()),
                    budget_id,
                    concept_id,
                    item["category"],
                    item["subcategory"],
                    item["code"],
                    titulo_comercial_item(item),
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
                    item.get("material_share_pct", 0.0),
                    item.get("labor_share_pct", 0.0),
                    item.get("other_share_pct", 100.0),
                    item.get("waste_reference_pct", 0.0),
                    int(item.get("execution_order") or 500),
                    json.dumps(
                        obtener_asignaciones_area_item(item),
                        ensure_ascii=False,
                        separators=(",", ":"),
                    ),
                    1 if item_esta_incluido(item) else 0,
                    str(item.get("contract_lot") or "1"),
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
            concept_was_existing = bool(concept_id)

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
                        item.get("concepto_base", item["description"]),
                        item["unit"],
                        normalizar_texto(item.get("concepto_base", item["description"])),
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

            if concept_was_existing and item.get("record_new_price"):
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
                    commercial_title, description, unit, quantity, unit_cost, direct_amount,
                    unit_indirect, unit_profit, unit_sale, sale_amount,
                    sale_margin_pct, benefit_amount, price_source,
                    price_source_detail, price_confidence,
                    material_share_pct, labor_share_pct, other_share_pct, waste_reference_pct,
                    execution_order, area_allocations_json, included, contract_lot, quantity_criterion,
                    inclusion_basis, considerations, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    str(uuid.uuid4()),
                    budget_id,
                    concept_id,
                    item["category"],
                    item["subcategory"],
                    item["code"],
                    titulo_comercial_item(item),
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
                    item.get("material_share_pct", 0.0),
                    item.get("labor_share_pct", 0.0),
                    item.get("other_share_pct", 100.0),
                    item.get("waste_reference_pct", 0.0),
                    int(item.get("execution_order") or 500),
                    json.dumps(
                        obtener_asignaciones_area_item(item),
                        ensure_ascii=False,
                        separators=(",", ":"),
                    ),
                    1 if item_esta_incluido(item) else 0,
                    str(item.get("contract_lot") or "1"),
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
            SELECT bi.*, c.description AS concepto_base
            FROM budget_items bi
            LEFT JOIN concepts c ON c.id = bi.concept_id
            WHERE bi.budget_id = ?
            ORDER BY bi.category, bi.subcategory, bi.code, bi.created_at
            """,
            (budget_id,),
        )

    def delete_budget(self, budget_id: str):
        budget = self.fetchone("SELECT project_id FROM budgets WHERE id = ?", (budget_id,))
        if not budget:
            return
        self.delete_generation(budget["project_id"], budget_id)

    def export_table(self, table_name: str) -> list[dict]:
        allowed = {
            "projects", "budgets", "concepts", "price_history", "budget_items",
        }
        if table_name not in allowed:
            raise ValueError("Tabla no permitida.")
        return self.fetchall(f"SELECT * FROM {table_name}")


DATABASE_CACHE_VERSION = "2026-09-21-v25-estilo-presupuesto-empresa"


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
    area: str = Field(
        description=(
            "Área física específica donde se ejecuta la actividad, por ejemplo Cocina, "
            "Baño 1, Baño 2, Recámara 1 o Fachada. Usa General únicamente para trabajos "
            "que realmente aplican al conjunto de la obra y no a un espacio particular."
        )
    )
    partida: str = Field(
        description="Sección comercial amplia del presupuesto, por ejemplo ACABADOS Y RECUBRIMIENTOS"
    )
    subpartida: str = Field(
        description=(
            "Subpartida breve que sí se mostrará en el Excel, sin numeración. "
            "Debe ser concreta y normalmente de 1 a 5 palabras, por ejemplo "
            "Licencias, Pisos, Muros, Frentes, Módulo Refri, Barra o Retiros."
        )
    )
    codigo_sugerido: str = Field(description="Código interno breve como PRE-01 o CAR-03")
    orden_ejecucion: int = Field(
        ge=1,
        le=999,
        description=(
            "Orden relativo de ejecución dentro de la obra. Menor significa antes. "
            "Debe responder a dependencias constructivas y no al orden del texto del usuario."
        ),
    )
    titulo_comercial: str = Field(
        description="Título corto y legible para el cliente, por ejemplo Pintura general o Demolición de muros"
    )
    concepto_base: str = Field(
        default="Concepto genérico",
        description=(
            "Nombre GENÉRICO y estandarizado para el catálogo histórico de la base de datos "
            "(ej. 'Mueble a medida acabados altos', 'Piso de duela', 'Muro de tabique'). "
            "NO incluyas dimensiones ni ubicaciones aquí."
        ),
    )
    descripcion_tecnica: str = Field(
        description="Descripción hiperespecífica del alcance que aparecerá debajo del título comercial en el Excel"
    )
    unidad: str = Field(description="Unidad: LOTE, PZA, M2, M3, ML, PTO, JGO, etc.")
    cantidad: float = Field(ge=0, description="Cantidad justificable con la información disponible")
    costo_unitario_estimado: float = Field(
        ge=0,
        description=(
            "Costo unitario integrado estimado del subcontratista, en MXN, antes de "
            "indirectos y utilidad. Es respaldo si no existe una referencia más confiable."
        )
    )
    porcentaje_materiales: float = Field(
        ge=0,
        le=100,
        description="Participación estimada de materiales dentro del costo integrado; solo informativa"
    )
    porcentaje_mano_obra: float = Field(
        ge=0,
        le=100,
        description="Participación estimada de mano de obra dentro del costo integrado; solo informativa"
    )
    porcentaje_otros: float = Field(
        ge=0,
        le=100,
        description="Participación estimada de equipo, proveedor, transporte u otros dentro del costo integrado"
    )
    desperdicio_materiales_pct: float = Field(
        ge=0,
        le=50,
        description=(
            "Desperdicio de referencia aplicable a materiales; informativo y no aditivo"
        )
    )
    criterio_cantidad: str = Field(description="Criterio verificable usado para determinar la cantidad")
    fundamento_inclusion: str = Field(description="Razón breve para incluir la actividad en el alcance")
    nivel_confianza_cantidad: str = Field(description="Alta, Media o Baja")
    nivel_confianza_precio: str = Field(description="Alta, Media o Baja")
    requiere_cotizacion: bool = Field(description="True si el costo debería confirmarse con proveedor especializado")
    consideraciones: str = Field(description="Supuestos, exclusiones o condiciones relevantes")


class AreaNecesidadesIA(BaseModel):
    area: str = Field(description="Área física identificada en el documento, normalizada y reconocible")
    superficie_m2: float | None = Field(default=None, ge=0, description="Superficie explícita si el documento la proporciona")
    necesidades: list[str] = Field(default_factory=list, description="Necesidades explícitas de esa área")
    restricciones: list[str] = Field(default_factory=list, description="Restricciones, preferencias o condiciones que afectan la solución")
    intervenciones_probables: list[str] = Field(default_factory=list, description="Trabajos o sistemas que probablemente deben presupuestarse para resolver las necesidades")


class PaqueteAlcanceIA(BaseModel):
    area: str = Field(description="Área física principal donde se ejecuta el paquete")
    nombre: str = Field(description="Nombre corto del paquete de trabajo que podría convertirse en una o varias partidas")
    disciplina: str = Field(description="Disciplina principal: CARPINTERIA, ACABADOS, ELECTRICA, HIDROSANITARIA, ALBANILERIA, etc.")
    objetivo: str = Field(description="Qué necesidad del cliente resuelve")
    trabajos_implicitos: list[str] = Field(default_factory=list, description="Trabajos necesarios para entregar el objetivo aunque no estén escritos literalmente")
    entregables: list[str] = Field(default_factory=list, description="Elementos concretos que deben quedar terminados")
    unidad_sugerida: str = Field(default="", description="Unidad natural de cotización si puede determinarse")
    base_de_cantidad: str = Field(default="", description="Cómo debería metarse la cantidad con los datos disponibles")
    depende_de: list[str] = Field(default_factory=list, description="Trabajos previos que deben existir antes")
    nivel_certeza: str = Field(default="Media", description="Alta, Media o Baja")


class MapaAlcanceIA(BaseModel):
    objetivo_general: str
    criterios_transversales: list[str] = Field(default_factory=list)
    areas: list[AreaNecesidadesIA] = Field(default_factory=list)
    paquetes: list[PaqueteAlcanceIA] = Field(default_factory=list)
    datos_faltantes: list[str] = Field(default_factory=list)


class PresupuestoIA(BaseModel):
    nombre_proyecto: str
    actividad_principal: str
    alcance_resumido: str
    consideraciones_generales: list[str]
    datos_faltantes: list[str]
    actividades: list[ActividadIA]


class ValuacionPrecioIA(BaseModel):
    codigo: str = Field(description="Código exacto de la actividad valuada")
    costo_unitario_final: float = Field(
        ge=0,
        description=(
            "Costo unitario final recomendado de SUBCONTRATACIÓN, integrado y en MXN, "
            "antes de los indirectos, utilidad de la empresa e IVA. Debe considerar "
            "alcance, especificaciones, acabados, complejidad, logística y condiciones "
            "razonables de contratación en CDMX."
        ),
    )
    nivel_confianza: str = Field(description="Alta, Media o Baja")
    requiere_cotizacion: bool = Field(
        description="True cuando la variabilidad o especialización hace recomendable confirmar con proveedor."
    )
    fundamento_precio: str = Field(
        description="Explicación breve de los factores principales considerados para el precio."
    )


class ValuacionPreciosIA(BaseModel):
    valuaciones: list[ValuacionPrecioIA]


class RecursoCosteoIA(BaseModel):
    categoria: str = Field(
        description=(
            "Familia del recurso: MATERIAL, HERRAJE, MANO_OBRA, CONSUMIBLE, EQUIPO, "
            "TRANSPORTE, DESPERDICIO, SUBCONTRATO u OTROS."
        )
    )
    concepto: str = Field(
        description=(
            "Recurso concreto que se presupuestaría para una unidad del concepto. "
            "Ejemplos: tablero melamínico 18 mm, canto PVC, bisagra cierre suave, "
            "taquete y tornillo, oficial de carpintería, ayudante, traslado."
        )
    )
    unidad: str = Field(description="Unidad de compra o consumo del recurso: PZA, ML, M2, H, JGO, L, KG, etc.")
    cantidad: float = Field(ge=0, description="Cantidad del recurso necesaria para UNA unidad de la actividad principal")
    costo_unitario: float = Field(ge=0, description="Costo estimado en MXN por unidad del recurso, antes de indirectos y utilidad de la empresa")
    obligatorio: bool = Field(description="True si el recurso forma parte normal del paquete para entregar correctamente el concepto")
    criterio: str = Field(description="Criterio breve de metrado, consumo o estimación del recurso")


class CosteoActividadIA(BaseModel):
    codigo: str = Field(description="Código exacto de la actividad")
    recursos: list[RecursoCosteoIA] = Field(description="Desglose interno completo de recursos para una unidad de actividad")
    confianza: str = Field(description="Alta, Media o Baja")
    requiere_cotizacion: bool = Field(description="True cuando el costo tiene alta variabilidad o requiere proveedor especializado")
    advertencias: list[str] = Field(default_factory=list, description="Advertencias técnicas o supuestos relevantes")


class CosteoPresupuestoIA(BaseModel):
    actividades: list[CosteoActividadIA]


class PrecioCompactoIA(BaseModel):
    codigo: str
    costo_unitario: float = Field(gt=0, description="MXN sin IVA por unidad comercial; solo para conceptos simples")
    confianza: str
    requiere_cotizacion: bool
    fundamento: str = Field(description="Base del precio y supuestos; breve")
    fuentes: list[str] = Field(default_factory=list, description="URLs de referencias consultadas; nunca inventarlas")
    recursos: list[RecursoCosteoIA] = Field(default_factory=list, description="Solo conceptos complejos; cantidades por UNA unidad comercial")


class PreciosCompactosIA(BaseModel):
    precios: list[PrecioCompactoIA]


def actividad_precio_complejo(actividad) -> bool:
    texto = normalizar_texto(" ".join(str(getattr(actividad, field, '') or '')
                                     for field in ('partida', 'subpartida', 'titulo_comercial', 'descripcion_tecnica')))
    return any(word in texto for word in (
        'carpinter', 'mueble', 'closet', 'vestidor', 'repisa', 'canceler',
        'herreria', 'estructura', 'a medida', 'especial'))


def buscar_precio_validado_exacto(db, actividad):
    """Reutiliza solo costos positivos, recientes, con unidad/alcance idénticos."""
    for row in db.price_candidates(actividad.unidad):
        if (str(row.get('status') or '').upper() not in
                {'VALIDADO', 'COSTO_REAL', 'COTIZADO_PROVEEDOR'}):
            continue
        if normalizar_unidad(row.get('unit')) != normalizar_unidad(actividad.unidad):
            continue
        if normalizar_texto(row.get('description')) != normalizar_texto(actividad.descripcion_tecnica):
            continue
        try:
            recorded = datetime.fromisoformat(str(row.get('created_at') or '').replace('Z', '+00:00'))
            if recorded.tzinfo is None:
                recorded = recorded.replace(tzinfo=timezone.utc)
            age = (datetime.now(timezone.utc) - recorded).days
            cost = float(row['unit_cost'])
            if not 0 <= age <= 180 or not 0 < cost < float('inf'):
                continue
        except (ValueError, TypeError, KeyError):
            continue
        return row
    return None


def revisar_precios_compactos_ia(api_key, model_name, project_data, packets,
                                progress_callback=None):
    """Revisa hasta cinco conceptos sin repetir el documento original ni auditar todo."""
    prompt = f"""
Fija costos unitarios de SUBCONTRATACIÓN para estos conceptos de remodelación.
Mercado: {project_data.get('location') or 'CDMX'}, México, {datetime.now().year}.
Nivel: {project_data.get('budget_level', 'Medio-alto')}.
Costos en MXN sin IVA, indirectos ni utilidad de nuestra empresa.
Devuelve exactamente un precio por código. No cambies unidades ni cantidades.
Respeta especificaciones e inclusiones. El costo inicial y referencias IA no son evidencia.
Usa Google Search para referencias actuales de proveedores; indica URLs reales.
Comprueba unidad de compra, presentación, IVA, fecha, suministro e instalación.
No uses precio de material como costo de servicio instalado. Si no puedes verificarlo,
marca cotización y confianza baja, explica supuestos, sin inventar fuentes.
Para desglose_requerido=true devuelve recursos esenciales (materiales, herrajes,
mano de obra, consumibles y logística cuando apliquen) POR UNA unidad comercial.
Python sumará cantidad por costo de recursos; evita doble conteo de desperdicio.
Para conceptos simples basta el costo integrado y su fundamento, recursos=[].
No inventes geometría como dato confirmado: declara las hipótesis y pide cotización.
Responde compacto, sin razonamiento extenso.
CONCEPTOS
{json.dumps(packets, ensure_ascii=False, separators=(',', ':'))}
"""
    client = genai.Client(api_key=api_key)
    for model in _modelos_gemini_disponibles(model_name):
        try:
            response = generar_con_gemini_resistente(
                client, model, prompt,
                configuracion_gemini_razonada(PreciosCompactosIA, thinking_level="low",
                                              max_output_tokens=12288, ground_with_search=True),
                progress_callback=progress_callback, etapa="2/3 · Revisando precios de mercado",
            )
            result = PreciosCompactosIA.model_validate_json(response.text)
            expected = {p['codigo'].upper() for p in packets}
            received = [p.codigo.strip().upper() for p in result.precios]
            if len(received) != len(set(received)) or set(received) != expected:
                raise RuntimeError("La revisión de precios devolvió códigos omitidos, duplicados o desconocidos.")
            # Guarda también las fuentes efectivamente devueltas por Google Search.
            grounded = set()
            for candidate in getattr(response, 'candidates', None) or []:
                metadata = getattr(candidate, 'grounding_metadata', None)
                for chunk in getattr(metadata, 'grounding_chunks', None) or []:
                    web = getattr(chunk, 'web', None)
                    uri = getattr(web, 'uri', None)
                    if uri:
                        grounded.add(uri)
            return result, sorted(grounded)
        except Exception as exc:
            if not error_gemini_modelo_no_disponible(exc):
                raise
    raise RuntimeError("Ningún modelo configurado está disponible para revisar precios.")


class AuditoriaCosteoActividadIA(BaseModel):
    codigo: str = Field(description="Código exacto de la actividad auditada")
    recursos_corregidos: list[RecursoCosteoIA] = Field(description="Hoja de costeo corregida; sustituye completamente la anterior")
    confianza: str = Field(description="Alta, Media o Baja")
    requiere_cotizacion: bool = Field(description="True cuando debe confirmarse con proveedor")
    hallazgos: list[str] = Field(default_factory=list, description="Errores u omisiones encontrados y corregidos")


class AuditoriaCosteoPresupuestoIA(BaseModel):
    actividades: list[AuditoriaCosteoActividadIA]


class ClasificacionActividadIA(BaseModel):
    codigo: str = Field(description="Código exacto de la actividad recibida")
    partida: str = Field(description="Partida comercial corregida")
    subpartida: str = Field(description="Subpartida breve corregida, sin numeración")
    titulo_comercial: str = Field(description="Título corto, claro y comercial corregido")
    descripcion_tecnica: str = Field(description="Descripción técnica/comercial corregida, breve pero completa")
    concepto_base: str = Field(description="Nombre genérico y reutilizable para el histórico")
    orden_ejecucion: int = Field(
        ge=1,
        le=999,
        description="Orden relativo corregido según la secuencia constructiva",
    )


class AuditoriaEstructuraIA(BaseModel):
    actividades: list[ClasificacionActividadIA]


class CambiosActividadIA(BaseModel):
    # PATCH parcial: Gemini solo rellena lo que realmente cambia.
    area: str | None = Field(default=None, description="Nueva Área; null = conservar")
    partida: str | None = Field(default=None, description="Nueva Partida; null = conservar")
    subpartida: str | None = Field(default=None, description="Nueva Subpartida; null = conservar")
    titulo_comercial: str | None = Field(default=None, description="Nuevo título comercial; null = conservar")
    concepto_base: str | None = Field(default=None, description="Nuevo concepto base; null = conservar")
    descripcion_tecnica: str | None = Field(default=None, description="Nueva descripción técnica; null = conservar")
    unidad: str | None = Field(default=None, description="Nueva unidad; null = conservar")
    cantidad: float | None = Field(default=None, ge=0, description="Nueva cantidad; null = conservar")
    costo_unitario_estimado: float | None = Field(default=None, ge=0, description="Nuevo costo unitario interno; null = conservar")
    porcentaje_materiales: float | None = Field(default=None, ge=0, le=100, description="Nuevo porcentaje de materiales; null = conservar")
    porcentaje_mano_obra: float | None = Field(default=None, ge=0, le=100, description="Nuevo porcentaje de mano de obra; null = conservar")
    porcentaje_otros: float | None = Field(default=None, ge=0, le=100, description="Nuevo porcentaje de otros; null = conservar")
    desperdicio_materiales_pct: float | None = Field(default=None, ge=0, le=50, description="Nuevo desperdicio; null = conservar")
    orden_ejecucion: int | None = Field(default=None, ge=1, le=999, description="Nuevo orden constructivo; null = conservar")
    requiere_cotizacion: bool | None = Field(default=None, description="Nueva marca de cotización; null = conservar")
    consideraciones: str | None = Field(default=None, description="Nuevas consideraciones; null = conservar")
    included: bool | None = Field(default=None, description="Activar/desactivar; null = conservar")
    contract_lot: str | None = Field(default=None, description="Lote interno; null = conservar")


class OperacionRevisionIA(BaseModel):
    accion: str = Field(description="AGREGAR, MODIFICAR, ELIMINAR o MOVER")
    codigo_objetivo: str = Field(default="", description="Código exacto actual cuando sea conocido")
    fila_objetivo: int | None = Field(default=None, ge=1, description="Número 1-based de la fila del presupuesto mostrado en el contexto; respaldo del código")
    texto_objetivo: str = Field(default="", description="Descripción breve usada para identificar la actividad si no hay código")
    actividad: ActividadIA | None = Field(default=None, description="Actividad completa para AGREGAR; opcional en MODIFICAR si se usa cambios")
    cambios: CambiosActividadIA | None = Field(default=None, description="Cambios parciales para MODIFICAR; solo incluir campos que cambian")
    posicion: int | None = Field(default=None, ge=1, description="Nueva posición 1-based para MOVER; opcional")
    recalcular_precio: bool = Field(default=False, description="True solo si el costo debe reevaluarse")
    motivo: str = Field(default="", description="Resumen técnico breve del cambio solicitado")


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


def actualizar_progreso(progress_callback, porcentaje: int, mensaje: str):
    if progress_callback is None:
        return
    try:
        progress_callback(max(0, min(int(porcentaje), 100)), str(mensaje))
    except Exception:
        pass


def limpiar_log_generacion(value) -> str:
    """Oculta la credencial de Gemini en mensajes y trazas visibles."""
    text = str(value)
    key = get_api_key_runtime()
    if key:
        text = text.replace(str(key), "[API_KEY_OCULTA]")
    return re.sub(r"AIza[0-9A-Za-z_-]{20,}", "[API_KEY_OCULTA]", text)


def mostrar_log_generacion(placeholder, entries):
    # Fondo claro propio: el registro sigue legible también con el tema oscuro.
    body = escape("\n".join(entries))
    placeholder.markdown(
        '<div role="log" aria-live="polite" style="background:#f8fafc;'
        'color:#172033;border:1px solid #cbd5e1;border-radius:8px;'
        'padding:16px;max-height:360px;overflow:auto;white-space:pre-wrap;'
        'overflow-wrap:anywhere;font:13px/1.6 monospace;">'
        + body + '</div>',
        unsafe_allow_html=True,
    )


def error_gemini_modelo_no_disponible(exc: Exception) -> bool:
    """Detecta errores que indican que el modelo solicitado no está disponible."""
    msg = str(exc).upper()
    return any(marker in msg for marker in (
        "404", "NOT_FOUND", "MODEL_NOT_FOUND", "MODEL IS NOT AVAILABLE",
        "MODEL NOT AVAILABLE", "NO LONGER AVAILABLE", "UNKNOWN MODEL",
    ))


def error_gemini_transitorio(exc: Exception) -> bool:
    """Indica que conviene repetir la misma solicitud Gemini."""
    msg = str(exc).upper()
    return any(marker in msg for marker in (
        "503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED", "500", "INTERNAL",
        "502", "BAD GATEWAY", "504", "GATEWAY TIMEOUT", "TIMEOUT",
        "DEADLINE_EXCEEDED", "SERVICE UNAVAILABLE", "TEMPORARILY UNAVAILABLE",
        "HIGH DEMAND",
    ))


MAX_REINTENTOS_GEMINI = 6
DELAY_REINTENTO_GEMINI_SEG = 20


def configuracion_gemini_razonada(
    response_schema=None,
    thinking_level: str = "high",
    max_output_tokens: int = 32768,
    ground_with_search: bool = False,
):
    """Configura Gemini para presupuestación con razonamiento y búsqueda de mercado opcional."""
    kwargs = {
        "max_output_tokens": max_output_tokens,
    }
    if response_schema is not None:
        kwargs.update({
            "response_mime_type": "application/json",
            "response_schema": response_schema,
        })
    try:
        kwargs["thinking_config"] = types.ThinkingConfig(thinking_level=thinking_level)
    except Exception:
        # Compatibilidad con una versión antigua del SDK. El servidor mantiene
        # razonamiento dinámico en modelos Gemini compatibles aunque no forcemos el nivel.
        pass
    if ground_with_search:
        try:
            kwargs["tools"] = [types.Tool(google_search=types.GoogleSearch())]
        except Exception:
            # Si una versión vieja del SDK no reconoce GoogleSearch, seguimos sin grounding.
            pass
    return types.GenerateContentConfig(**kwargs)


def generar_con_gemini_resistente(
    client,
    model: str,
    contents: str,
    config,
    progress_callback=None,
    etapa: str = "Procesando",
    max_reintentos_transitorios: int = MAX_REINTENTOS_GEMINI,
):
    """
    Ejecuta una etapa Gemini sin abandonarla por saturación temporal.

    Ante un 503/UNAVAILABLE, alterna inmediatamente entre Flash 3.8 y 3.7.
    Si también falla el respaldo, espera 20 s y vuelve al modelo preferido.
    Los 429 y otros errores transitorios esperan sin cambiar de modelo.
    El límite cuenta llamadas totales, incluidos los intentos de respaldo.
    Los errores definitivos se propagan al llamador.
    """
    ultimo_error = None
    modelo_preferido = model
    modelo_respaldo = {
        "gemini-3.8-flash": "gemini-3.7-flash",
        "gemini-3.7-flash": "gemini-3.8-flash",
    }.get(model)

    for intento in range(1, max_reintentos_transitorios + 1):
        try:
            actualizar_progreso(
                progress_callback,
                0,
                f"{etapa} · {model} · intento {intento}/{max_reintentos_transitorios}",
            )
            started = time.monotonic()
            if progress_callback is None:
                response = client.models.generate_content(
                    model=model, contents=contents, config=config
                )
            else:
                # Solo la llamada HTTP trabaja en otro hilo. Streamlit y sus
                # callbacks se ejecutan siempre en el hilo principal.
                with ThreadPoolExecutor(max_workers=1) as executor:
                    pending = executor.submit(
                        client.models.generate_content,
                        model=model, contents=contents, config=config,
                    )
                    while not wait([pending], timeout=5).done:
                        elapsed = int(time.monotonic() - started)
                        actualizar_progreso(
                            progress_callback, 0,
                            f"{etapa} · {model} · esperando respuesta de Gemini ({elapsed} s).",
                        )
                    response = pending.result()
            if not getattr(response, "text", None):
                raise RuntimeError(
                    f"Gemini ({model}) devolvió una respuesta vacía."
                )
            actualizar_progreso(
                progress_callback, 0,
                f"{etapa} · {model} · respuesta recibida en {time.monotonic() - started:.1f} s; validando datos.",
            )
            return response
        except Exception as exc:
            ultimo_error = exc
            actualizar_progreso(
                progress_callback, 0,
                f"ERROR Gemini · {etapa} · {model} · {type(exc).__name__}: {exc}",
            )
            if error_gemini_modelo_no_disponible(exc) or not error_gemini_transitorio(exc):
                raise
            if intento >= max_reintentos_transitorios:
                raise
            siguiente = intento + 1
            # Un 503 es saturación del servicio; un 429 es cuota y no activa
            # el cambio inmediato. Se conserva el mismo prompt y configuración
            # (incluida Google Search) al cambiar entre estos dos modelos.
            error_msg = str(exc).upper()
            error_code = str(getattr(exc, "code", ""))
            saturado = (
                error_code == "503" or "503" in error_msg
                or "UNAVAILABLE" in error_msg
            ) and not (error_code == "429" or "429" in error_msg
                       or "RESOURCE_EXHAUSTED" in error_msg)
            if modelo_respaldo and saturado and model == modelo_preferido:
                model = modelo_respaldo
                actualizar_progreso(
                    progress_callback, 0,
                    f"{etapa} · servicio saturado; probando {model} inmediatamente "
                    f"(intento {siguiente}/{max_reintentos_transitorios}).",
                )
                continue
            if modelo_respaldo and saturado and model == modelo_respaldo:
                model = modelo_preferido
            actualizar_progreso(
                progress_callback,
                0,
                (
                    f"{etapa} · Gemini no respondió correctamente. Próximo modelo: {model}. "
                    f"Esperando {DELAY_REINTENTO_GEMINI_SEG} s antes del intento "
                    f"{siguiente}/{max_reintentos_transitorios}..."
                ),
            )
            deadline = time.monotonic() + DELAY_REINTENTO_GEMINI_SEG
            while time.monotonic() < deadline:
                remaining = deadline - time.monotonic()
                time.sleep(min(5, remaining))
                remaining = max(0, int(deadline - time.monotonic() + 0.999))
                actualizar_progreso(
                    progress_callback, 0,
                    f"{etapa} · reintento {siguiente}/{max_reintentos_transitorios} en {remaining} s.",
                )

    raise ultimo_error if ultimo_error else RuntimeError("Error desconocido de Gemini.")




def generar_presupuesto_ia(api_key, model_name, project_data, params,
                          progress_callback=None) -> PresupuestoIA:
    """Una sola lectura del alcance; las referencias de precios se revisan después."""
    client = genai.Client(api_key=api_key)
    prompt = f"""
Eres presupuestista de remodelación e interiorismo en México. La empresa subcontrata.
Genera conceptos comerciales claros a partir del alcance, no una fila por recurso.
Ubicación: {project_data.get('location') or 'CDMX'}. Año: {datetime.now().year}.
Cliente: {project_data['name']}. Tipo: {project_data['project_type']}.
Nivel: {project_data.get('budget_level', 'Medio-alto')}.
{criterio_nivel_presupuesto(project_data.get('budget_level', 'Medio-alto'))}

ALCANCE ORIGINAL
{project_data['description']}
CONDICIONES
{project_data.get('guide_text') or 'Sin condiciones adicionales.'}

REGLAS
1. Incluye trabajos pedidos y complementarios indispensables; señala lo inferido.
   No agregues decoración, trámites ni trabajos opcionales sin justificación.
2. Separa áreas y muebles distintos. Conserva materiales, dimensiones, acabados,
   herrajes, accesos y condiciones importantes dentro de cada descripción:
   la revisión de precios recibirá únicamente esos conceptos, no este documento.
3. Usa partidas por fase de obra, subpartidas breves y códigos únicos.
   Protecciones antes de demoliciones; limpieza final al cierre. Evita duplicados.
4. Cantidades justificadas: M2, ML, M3, PZA, PTO, JGO o LOTE. No inventes medidas.
   Si falta información, declara el supuesto y marca confianza baja/cotización.
5. Estima costo UNITARIO integrado de subcontratación en MXN sin IVA, indirectos
   ni utilidad de nuestra empresa. Incluye suministro/instalación solo si aplican.
   No confundas precio de material con servicio instalado o costo por unidad con total.
6. El nivel afecta especificaciones; no multipliques arbitrariamente los precios.
   Porcentajes de materiales/mano de obra/otros son informativos y suman 100.
   El desperdicio ya está incluido y no se suma otra vez.
7. Criterios y advertencias breves. Devuelve solo el objeto estructurado.
"""
    for model in _modelos_gemini_disponibles(model_name):
        try:
            response = generar_con_gemini_resistente(
                client, model, prompt,
                configuracion_gemini_razonada(PresupuestoIA, thinking_level="low",
                                              max_output_tokens=16384),
                progress_callback=progress_callback, etapa="1/3 · Generando partidas",
            )
            result = PresupuestoIA.model_validate_json(response.text)
            if not result.actividades:
                raise RuntimeError("Gemini devolvió un presupuesto sin actividades.")
            codes = [a.codigo_sugerido.strip().upper() for a in result.actividades]
            if len(codes) != len(set(codes)) or not all(codes):
                raise RuntimeError("Gemini devolvió códigos vacíos o duplicados.")
            return result
        except Exception as exc:
            if not error_gemini_modelo_no_disponible(exc):
                raise
    raise RuntimeError("Ningún modelo configurado está disponible para generar partidas.")





def sincronizar_items_con_estructura(
    result: PresupuestoIA,
    items: list[dict],
) -> list[dict]:
    """Sincroniza partida/subpartida/orden sin modificar costos."""
    by_code = {
        str(act.codigo_sugerido or "").strip().upper(): act
        for act in result.actividades
    }
    output = []

    for item in items:
        out = dict(item)
        act = by_code.get(str(out.get("code") or "").strip().upper())
        if act is not None:
            out["category"] = normalizar_seccion_comercial(act.partida)
            out["subcategory"] = act.subpartida.strip()
            out["execution_order"] = int(act.orden_ejecucion)
        output.append(out)

    return output

def revisar_presupuesto_ia(
    api_key: str,
    model_name: str,
    project_data: dict,
    params: dict,
    current_result: PresupuestoIA,
    current_items: list[dict],
    revision_request: str,
    progress_callback=None,
) -> RevisionPresupuestoIA:
    """Interpreta una petición libre como operaciones granulares sobre el presupuesto vigente."""
    client = genai.Client(api_key=api_key)

    presupuesto_actual = []
    for idx, x in enumerate(current_items, start=1):
        presupuesto_actual.append({
            "fila": idx,
            "codigo": x.get("code", ""),
            "area": area_excel_item(x),
            "partida": x.get("category", ""),
            "subpartida": x.get("subcategory", ""),
            "titulo_comercial": titulo_comercial_item(x),
            "descripcion": x.get("description", ""),
            "unidad": x.get("unit", ""),
            "cantidad": x.get("quantity", 0),
            "precio_venta_unitario": x.get("unit_sale", 0),
            "costo_unitario_actual": x.get("unit_cost", 0),
            "orden_ejecucion": x.get("execution_order", 500),
            "incluido": item_esta_incluido(x),
            "fuente_precio": x.get("price_source", ""),
            "consideraciones": x.get("considerations", ""),
        })

    prompt = f"""
Actúa como EDITOR EXPERTO de un presupuesto de remodelación e interiorismo.
Tu trabajo NO es regenerar el presupuesto. Debes interpretar la instrucción del usuario
como cambios concretos sobre la tabla existente y devolver SOLO las operaciones necesarias.

DATOS DEL PROYECTO
Cliente: {project_data['name']}
Ubicación: {project_data['location']}
Tipo de obra: {project_data['project_type']}
Nivel: {project_data.get('budget_level', 'Medio-alto')}

PRESUPUESTO ACTUAL (fila es 1-based y código es el identificador principal)
{json.dumps(presupuesto_actual, ensure_ascii=False, separators=(',', ':'))}

INSTRUCCIÓN DEL USUARIO
{revision_request}

REGLAS DEL EDITOR
1. Conserva TODO lo que el usuario no pida cambiar.
2. Puedes AGREGAR, MODIFICAR, ELIMINAR y MOVER.
3. Para MODIFICAR, usa codigo_objetivo. Si no es evidente, usa fila_objetivo o texto_objetivo para identificarla.
4. Para MODIFICAR NO devuelvas la actividad completa: utiliza cambios y rellena únicamente los campos que cambian.
5. Para AGREGAR sí devuelve actividad completa. Puedes generar una actividad nueva desde cero.
6. Para ELIMINAR no necesitas actividad ni cambios.
7. Para MOVER usa posicion para la nueva posición si el usuario indica dónde quiere colocarla.
8. "Agrega X después de Y", "pon X antes de Y", "mueve X al final de Acabados" y solicitudes similares deben convertirse en MOVER/AGREGAR con una posición coherente.
9. Puedes cambiar libremente Área, Partida, Subpartida, Título, descripción, unidad, cantidad,
   costo unitario, composición, desperdicio, orden, cotización, consideraciones, activar/desactivar y lote.
10. Si el usuario da un PRECIO DE VENTA concreto, no lo guardes como costo_unitario_estimado sin analizar el contexto:
    este módulo trabaja principalmente con costo interno. Solo cambia costo_unitario_estimado cuando la petición hable de costo/precio de subcontratación o pida recalcularlo.
11. Si el usuario dice que algo está caro/barato, revisa costo_unitario_estimado y usa recalcular_precio=True.
12. Si solo cambia cantidad, texto, área o clasificación, recalcular_precio=False.
13. Si cambia unidad, especificación/material, complejidad o naturaleza, recalcular_precio=True.
14. Si el usuario aporta una cifra exacta de costo unitario, úsala como costo_unitario_estimado y recalcular_precio=True.
15. No cambies precios no mencionados.
16. No agrupes áreas físicas diferentes en una sola actividad.
17. No agrupes muebles diferentes en una sola actividad cuando tengan función, modelo o especificación diferente.
18. Mantén concepto_base genérico, sin dimensiones ni ubicación.
19. No calcules indirectos, utilidad, IVA ni importes finales; Python los recalcula.
20. La respuesta debe ser ejecutable. No expliques razonamientos internos.
21. Si la petición es ambigua, haz el cambio más conservador posible y deja una consideración breve.
"""

    modelos = []
    for model in _modelos_gemini_disponibles(model_name):
        if model and model not in modelos:
            modelos.append(model)

    last_error = None
    for model in modelos:
        try:
            response = generar_con_gemini_resistente(
                client=client,
                model=model,
                contents=prompt,
                config=configuracion_gemini_razonada(
                    RevisionPresupuestoIA, thinking_level="low", max_output_tokens=16384
                ),
                progress_callback=progress_callback,
                etapa="2/4 · Interpretando cambios",
            )
            return RevisionPresupuestoIA.model_validate_json(response.text)
        except Exception as exc:
            last_error = exc
            msg = str(exc).lower()
            model_error = (
                "404" in msg or "not_found" in msg or
                "no longer available" in msg or
                ("model" in msg and "not available" in msg)
            )
            if not model_error:
                raise

    raise RuntimeError(f"No fue posible usar un modelo Gemini para la revisión. Último error: {last_error}")



# =========================================================
# MOTOR DE PRECIOS
# =========================================================


def buscar_precio_interno(db: Database, actividad: ActividadIA) -> dict | None:
    candidatos = db.price_candidates(actividad.unidad)
    best = None
    best_score = 0.0
    best_priority = -1

    for row in candidatos:
        score = score_similitud(actividad.descripcion_tecnica, row["description"])
        priority = 2 if (row.get("status") or "").upper() == "VALIDADO" else 1
        if (
            score >= 0.82
            and (priority > best_priority or (priority == best_priority and score > best_score))
        ):
            best_priority = priority
            best_score = score
            best = row

    if best is None:
        return None

    original_source = (best.get("source") or "").upper()
    original_status = (best.get("status") or "").upper()

    if original_status in {"VALIDADO", "COSTO_REAL", "COTIZADO_PROVEEDOR"}:
        source = "BASE_INTERNA"
        confidence = "Alta"
    elif original_source == "IA_ESTIMADO" or original_status == "ESTIMADO_IA":
        source = "HISTORICO_IA"
        confidence = "Media" if best_score >= 0.9 else "Baja"
    elif original_status in {"REFERENCIA_EXTERNA", "REFERENCIA_CDMX"} or original_source in {"REFERENCIA_EXTERNA", "REFERENCIA_CDMX", "HISTORICO_EXTERNO"}:
        # Compatibilidad con registros históricos antiguos: ya no se consideran
        # una fuente externa operativa y se tratan como evidencia interna heredada.
        source = "BASE_INTERNA"
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










def _modelos_gemini_disponibles(model_name: str | None) -> list[str]:
    modelos = []
    for model in [
        model_name,
        "gemini-3.8-flash",
        "gemini-3.7-flash",
        "gemini-3.6-flash",
        "gemini-3.5-flash",
    ]:
        if model and model not in modelos:
            modelos.append(model)
    return modelos






def normalizar_recursos_costeo(resources: list[RecursoCosteoIA]) -> tuple[list[dict], float]:
    """Normaliza recursos y calcula en Python el costo unitario, sin delegar la aritmética a Gemini."""
    if not resources:
        raise RuntimeError("La hoja de costeo llegó sin recursos; no se permite regresar a un precio por ML/M2/PZA.")

    rows = []
    total = 0.0
    categories = set()
    for resource in resources:
        qty = max(float(resource.cantidad), 0.0)
        unit_cost = max(float(resource.costo_unitario), 0.0)
        amount = qty * unit_cost
        total += amount
        category = str(resource.categoria or "OTROS").strip().upper()
        categories.add(category)
        rows.append({
            "categoria": category,
            "concepto": str(resource.concepto or "Recurso").strip(),
            "unidad": normalizar_unidad(resource.unidad),
            "cantidad": qty,
            "costo_unitario": unit_cost,
            "importe": round(amount, 2),
            "obligatorio": bool(resource.obligatorio),
            "criterio": str(resource.criterio or "").strip(),
        })

    # Para muebles/carpintería exigimos como mínimo una base material/herraje y mano de obra.
    # En otros conceptos dejamos que el tipo de trabajo determine las categorías.
    hay_carpinteria = any(
        token in normalizar_texto(f"{row['concepto']} {row['categoria']}").upper()
        for row in rows
        for token in ("MUEBLE", "CARPINTER", "TABLERO", "MELAMINA", "MDF", "REPISA", "CLOSET")
    )
    if hay_carpinteria:
        if not any(cat in categories for cat in {"MATERIAL", "HERRAJE", "CONSUMIBLE"}):
            raise RuntimeError("El costeo de carpintería no contiene materiales/herrajes/consumibles identificables.")
        if "MANO_OBRA" not in categories:
            raise RuntimeError("El costeo de carpintería no contiene mano de obra identificable.")

    if total <= 0:
        raise RuntimeError("El costeo detallado produjo un costo unitario cero.")
    return rows, round(total, 2)


def resolver_items(db, result, project_data, params, force_new_price_codes=None,
                  api_key=None, model_name=None, progress_callback=None) -> list[dict]:
    """Costos exactos validados primero; los faltantes se revisan en lotes de cinco."""
    forced = {str(code).strip().upper() for code in (force_new_price_codes or set())}
    prices, pending = {}, []
    actualizar_progreso(progress_callback, 50, "2/3 · Buscando costos históricos validados")
    for idx, act in enumerate(result.actividades, 1):
        code = limpiar_codigo(act.codigo_sugerido, f"CON-{idx:03d}")
        if code in prices or any(p['codigo'] == code for p in pending):
            raise RuntimeError(f"Código repetido después de normalizar: {code}")
        reference = None if code.upper() in forced else buscar_precio_validado_exacto(db, act)
        if reference:
            prices[code] = dict(cost=float(reference['unit_cost']),
                concept_id=reference['concept_id'], source='BASE_INTERNA',
                status=reference['status'], confidence='Alta', resources=[],
                quote=False, detail=f"Costo validado de alcance y unidad idénticos; fecha {reference['created_at']}.")
            actualizar_progreso(progress_callback, 52, f"{code} · costo histórico validado reutilizado")
        else:
            hint = None if code.upper() in forced else buscar_precio_interno(db, act)
            pending.append(dict(codigo=code, descripcion=act.descripcion_tecnica,
                titulo=act.titulo_comercial, area=act.area, unidad=act.unidad,
                cantidad=float(act.cantidad), criterio_cantidad=act.criterio_cantidad,
                supuestos=act.consideraciones, desglose_requerido=actividad_precio_complejo(act),
                costo_inicial=float(act.costo_unitario_estimado),
                referencia=(dict(costo=hint['unit_cost'], estado=hint['status'],
                                 detalle=hint['source_detail']) if hint else None)))

    if pending:
        api_key = api_key or get_api_key_runtime()
        if not api_key:
            raise RuntimeError("Falta GEMINI_API_KEY para revisar los precios pendientes.")
    for start in range(0, len(pending), 5):
        batch = pending[start:start + 5]
        pct = 55 + int(start / max(len(pending), 1) * 35)
        actualizar_progreso(progress_callback, pct,
            f"2/3 · Revisando lote {start // 5 + 1}/{(len(pending) + 4) // 5}: {len(batch)} conceptos")
        checked, grounded = revisar_precios_compactos_ia(
            api_key, model_name or 'gemini-3.8-flash', project_data, batch,
            progress_callback=(lambda _pct, msg: actualizar_progreso(progress_callback, pct, msg)))
        packets = {p['codigo'].upper(): p for p in batch}
        for price in checked.precios:
            code = price.codigo.strip().upper()
            packet = packets[code]
            if packet['desglose_requerido'] and not price.recursos:
                raise RuntimeError(f"{code}: falta desglose esencial para comprobar el precio complejo.")
            resources, cost = normalizar_recursos_costeo(price.recursos) if price.recursos else ([], float(price.costo_unitario))
            if not 0 < cost < float('inf'):
                raise RuntimeError(f"{code}: costo unitario inválido.")
            # Las URLs escritas por el modelo son declaradas; las de grounding
            # sí constan en la respuesta de búsqueda, sin garantizar comparabilidad.
            declared = [url for url in price.fuentes if url.startswith(('https://', 'http://'))]
            detail = price.fundamento
            if declared:
                detail += ' | Referencias declaradas: ' + ', '.join(declared)
            if grounded:
                detail += ' | Fuentes de búsqueda del lote: ' + ', '.join(grounded)
            else:
                detail += ' | Sin evidencia de búsqueda devuelta por la API; confirmar con proveedor.'
            prices[code] = dict(cost=cost, concept_id=None, source='IA_ESTIMADO',
                status='ESTIMADO_IA', confidence=(price.confianza if grounded else 'Baja'),
                quote=price.requiere_cotizacion or not grounded, resources=resources, detail=detail)
            actualizar_progreso(progress_callback, pct, f"{code} · costo revisado: ${cost:,.2f}/{packet['unidad']}")

    items = []
    for idx, act in enumerate(result.actividades, 1):
        code = limpiar_codigo(act.codigo_sugerido, f"CON-{idx:03d}")
        price = prices[code.upper()]
        considerations = act.consideraciones.strip()
        if price['quote'] or act.requiere_cotizacion:
            considerations += (' | ' if considerations else '') + 'Requiere cotización de proveedor.'
        item = dict(concept_id=price['concept_id'], area_hint=normalizar_nombre_area(act.area),
            category=act.partida, subcategory=act.subpartida, code=code,
            execution_order=act.orden_ejecucion, commercial_title=act.titulo_comercial,
            concepto_base=act.concepto_base, description=act.descripcion_tecnica,
            unit=normalizar_unidad(act.unidad), quantity=float(act.cantidad), unit_cost=price['cost'],
            price_source=price['source'], price_source_detail=price['detail'],
            price_status=price['status'], price_confidence=price['confidence'],
            material_share_pct=act.porcentaje_materiales, labor_share_pct=act.porcentaje_mano_obra,
            other_share_pct=act.porcentaje_otros, waste_reference_pct=act.desperdicio_materiales_pct,
            included=True, contract_lot='1', quantity_confidence=act.nivel_confianza_cantidad,
            quantity_criterion=act.criterio_cantidad, inclusion_basis=act.fundamento_inclusion,
            considerations=considerations, costing_breakdown=price['resources'],
            area_allocations=[dict(area=normalizar_nombre_area(act.area), porcentaje=100.0,
                cantidad_referencia=float(act.cantidad), criterio='Área indicada en el alcance.')])
        items.append(recalcular_item_financiero(item, params))
    actualizar_progreso(progress_callback, 92, "2/3 · Precios revisados y totales calculados en Python")
    return ordenar_items_comercialmente(items)



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
    return aplicar_composicion_costo(out)


def item_a_actividad(item: dict) -> ActividadIA:
    return ActividadIA(
        area=area_excel_item(item),
        partida=item["category"],
        subpartida=item["subcategory"],
        codigo_sugerido=item["code"],
        orden_ejecucion=int(item.get("execution_order") or 500),
        titulo_comercial=titulo_comercial_item(item),
        concepto_base=item.get("concepto_base") or item["description"],
        descripcion_tecnica=item["description"],
        unidad=item["unit"],
        cantidad=float(item["quantity"]),
        costo_unitario_estimado=float(item["unit_cost"]),
        porcentaje_materiales=float(item.get("material_share_pct") or 0.0),
        porcentaje_mano_obra=float(item.get("labor_share_pct") or 0.0),
        porcentaje_otros=float(
            item.get("other_share_pct")
            if item.get("other_share_pct") is not None
            else 100.0
        ),
        desperdicio_materiales_pct=float(item.get("waste_reference_pct") or 0.0),
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
    api_key: str | None = None,
    model_name: str | None = None,
) -> tuple[PresupuestoIA, list[dict], list[str]]:
    """Aplica operaciones granulares; MODIFICAR usa parches y nunca reemplaza campos no pedidos."""
    if not revision.operaciones:
        raise RuntimeError("Gemini no identificó cambios aplicables.")

    items = [dict(x) for x in current_items]
    change_log: list[str] = []

    def find_index(op: OperacionRevisionIA) -> int:
        code = str(op.codigo_objetivo or "").strip().upper()
        if code:
            for i, item in enumerate(items):
                if str(item.get("code") or "").strip().upper() == code:
                    return i
        if op.fila_objetivo is not None:
            idx = int(op.fila_objetivo) - 1
            if 0 <= idx < len(items):
                return idx
        target = normalizar_texto(op.texto_objetivo or "")
        if target:
            best_idx, best_score = None, 0.0
            for i, item in enumerate(items):
                hay = normalizar_texto(" ".join([
                    str(item.get("subcategory") or ""),
                    titulo_comercial_item(item),
                    str(item.get("description") or ""),
                ]))
                score = score_similitud(target, hay)
                if score > best_score:
                    best_score, best_idx = score, i
            if best_idx is not None and best_score >= 0.50:
                return best_idx
        raise RuntimeError(
            f"No pude identificar la actividad objetivo para la operación '{op.accion}'. "
            f"Código='{op.codigo_objetivo}', fila='{op.fila_objetivo}', texto='{op.texto_objetivo}'."
        )

    def resolve_new_activity(act: ActividadIA, force_price: bool = True) -> dict:
        temp = PresupuestoIA(
            nombre_proyecto=current_result.nombre_proyecto,
            actividad_principal=revision.actividad_principal_actualizada,
            alcance_resumido=revision.alcance_resumido_actualizado,
            consideraciones_generales=revision.consideraciones_generales_actualizadas,
            datos_faltantes=revision.datos_faltantes_actualizados,
            actividades=[act],
        )
        forced = {limpiar_codigo(act.codigo_sugerido, "REV-001")} if force_price else set()
        resolved = resolver_items(
            db, temp, project_data, params,
            force_new_price_codes=forced,
            api_key=api_key,
            model_name=model_name,
        )
        if not resolved:
            raise RuntimeError("No fue posible resolver la actividad nueva.")
        return dict(resolved[0])

    for op in revision.operaciones:
        action = (op.accion or "").strip().upper()
        if action not in {"AGREGAR", "MODIFICAR", "ELIMINAR", "MOVER"}:
            raise RuntimeError(f"Acción de revisión no válida: {op.accion}")

        if action == "AGREGAR":
            if op.actividad is None:
                raise RuntimeError("Una operación AGREGAR necesita actividad completa.")
            new_item = resolve_new_activity(op.actividad, force_price=True)
            if op.posicion:
                insert_at = max(0, min(int(op.posicion) - 1, len(items)))
                items.insert(insert_at, new_item)
                change_log.append(f"AGREGADO {new_item.get('code')}: {titulo_comercial_item(new_item)} en posición {insert_at + 1}")
            else:
                items.append(new_item)
                change_log.append(f"AGREGADO {new_item.get('code')}: {titulo_comercial_item(new_item)}")
            continue

        idx = find_index(op)
        item = dict(items[idx])

        if action == "ELIMINAR":
            removed = items.pop(idx)
            change_log.append(f"ELIMINADO {removed.get('code')}: {titulo_comercial_item(removed)}")
            continue

        if action == "MOVER":
            if op.posicion is None:
                raise RuntimeError("MOVER necesita posicion.")
            moved = items.pop(idx)
            new_idx = max(0, min(int(op.posicion) - 1, len(items)))
            items.insert(new_idx, moved)
            moved["execution_order"] = (new_idx + 1) * 10
            change_log.append(f"MOVIDO {moved.get('code')}: posición {idx + 1} → {new_idx + 1}")
            continue

        # MODIFICAR: parche campo por campo.
        patch = op.cambios.model_dump(exclude_none=True) if op.cambios is not None else {}
        if op.actividad is not None and not patch:
            # Compatibilidad con revisiones antiguas: convertir actividad completa en cambios,
            # pero solo sobre los campos estructurales/cantidad/precio relevantes.
            full = op.actividad.model_dump()
            patch = {
                "area": full["area"],
                "partida": full["partida"],
                "subpartida": full["subpartida"],
                "titulo_comercial": full["titulo_comercial"],
                "concepto_base": full["concepto_base"],
                "descripcion_tecnica": full["descripcion_tecnica"],
                "unidad": full["unidad"],
                "cantidad": full["cantidad"],
                "costo_unitario_estimado": full["costo_unitario_estimado"],
                "porcentaje_materiales": full["porcentaje_materiales"],
                "porcentaje_mano_obra": full["porcentaje_mano_obra"],
                "porcentaje_otros": full["porcentaje_otros"],
                "desperdicio_materiales_pct": full["desperdicio_materiales_pct"],
                "orden_ejecucion": full["orden_ejecucion"],
                "consideraciones": full["consideraciones"],
            }

        if not patch:
            continue

        # Mapear nombres IA -> claves internas.
        field_map = {
            "area": "area_hint",
            "partida": "category",
            "subpartida": "subcategory",
            "titulo_comercial": "commercial_title",
            "concepto_base": "concepto_base",
            "descripcion_tecnica": "description",
            "unidad": "unit",
            "cantidad": "quantity",
            "costo_unitario_estimado": "unit_cost",
            "porcentaje_materiales": "material_share_pct",
            "porcentaje_mano_obra": "labor_share_pct",
            "porcentaje_otros": "other_share_pct",
            "desperdicio_materiales_pct": "waste_reference_pct",
            "orden_ejecucion": "execution_order",
            "requiere_cotizacion": "requires_quote",
            "consideraciones": "considerations",
            "included": "included",
            "contract_lot": "contract_lot",
        }
        old_desc = item.get("description")
        for src, value in patch.items():
            dst = field_map[src]
            if dst == "category":
                value = normalizar_seccion_comercial(value)
            elif dst == "unit":
                value = normalizar_unidad(value)
            elif dst == "area_hint":
                value = normalizar_nombre_area(value)
            item[dst] = value

        # Recalcular costos solo cuando corresponde. Si cambia el costo explícitamente,
        # el costo dado por el usuario/IA prevalece. Si no cambia costo, se conserva.
        recalcular_precio = bool(op.recalcular_precio)
        if "unidad" in patch and str(item.get("unit") or "") != str(current_items[idx].get("unit") or ""):
            recalcular_precio = True
        if recalcular_precio and "costo_unitario_estimado" not in patch:
            # Resolver una nueva valuación usando la actividad ya modificada.
            act = item_a_actividad(item)
            temp = PresupuestoIA(
                nombre_proyecto=current_result.nombre_proyecto,
                actividad_principal=revision.actividad_principal_actualizada,
                alcance_resumido=revision.alcance_resumido_actualizado,
                consideraciones_generales=revision.consideraciones_generales_actualizadas,
                datos_faltantes=revision.datos_faltantes_actualizados,
                actividades=[act],
            )
            resolved = resolver_items(
                db, temp, project_data, params,
                force_new_price_codes={str(item.get("code"))},
                api_key=api_key,
                model_name=model_name,
            )
            if resolved:
                item = dict(resolved[0])
                item["included"] = item_esta_incluido(items[idx])
                item["contract_lot"] = str(items[idx].get("contract_lot") or "1")

        item = recalcular_item_financiero(item, params)
        item = aplicar_composicion_costo(item)
        item["included"] = item.get("included", True)
        items[idx] = item
        change_log.append(f"MODIFICADO {item.get('code')}: " + ", ".join(patch.keys()))

    # La posición real de la lista es la fuente de verdad del orden.
    # Esto evita que Excel vuelva a ordenar un elemento movido por un execution_order antiguo.
    for i, item in enumerate(items, start=1):
        item["execution_order"] = i * 10

    revised_result = PresupuestoIA(
        nombre_proyecto=current_result.nombre_proyecto,
        actividad_principal=revision.actividad_principal_actualizada or current_result.actividad_principal,
        alcance_resumido=revision.alcance_resumido_actualizado or current_result.alcance_resumido,
        consideraciones_generales=revision.consideraciones_generales_actualizadas or current_result.consideraciones_generales,
        datos_faltantes=revision.datos_faltantes_actualizados or current_result.datos_faltantes,
        actividades=[item_a_actividad(x) for x in items],
    )
    return revised_result, items, change_log


def calcular_financieros(items: list[dict], params: dict) -> dict:
    active_items = [x for x in items if item_esta_incluido(x)]
    direct_cost = sum(float(x.get("direct_amount") or 0.0) for x in active_items)
    indirect_cost = direct_cost * params["indirect_pct"] / 100.0
    profit = (direct_cost + indirect_cost) * params["profit_pct"] / 100.0

    # El presupuesto interno visible respeta el importe vigente, pero
    # solamente suma las actividades marcadas como incluidas.
    sale_before_tax = sum(float(x.get("sale_amount") or 0.0) for x in active_items)

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
# IMPORTAR PRESUPUESTO DESDE EXCEL
# =========================================================


def _valor_float_excel(value, default: float = 0.0) -> float:
    if value is None or value == "":
        return default
    if isinstance(value, (int, float)):
        return float(value)
    raw = str(value).strip()
    raw = raw.replace("$", "").replace(",", "")
    raw = raw.replace("%", "")
    try:
        return float(raw)
    except Exception:
        return default


def _pct_excel_a_porcentaje(value, default: float) -> float:
    number = _valor_float_excel(value, default)
    # Excel suele guardar 12 % como 0.12.
    if 0 <= number <= 1:
        return number * 100.0
    return number


def _quitar_numeracion_excel(value: str) -> str:
    raw = str(value or "").strip()
    raw = re.sub(r"^\s*\d+(?:\.\d+)*\.?\s*", "", raw)
    return raw.strip()


def _buscar_encabezado_presupuesto(ws) -> int | None:
    """Busca la fila de Partida/Subpartida/... sin depender de una versión."""
    expected = {
        "PARTIDA",
        "SUBPARTIDA",
        "DESCRIPCION TECNICA",
        "UNIDAD",
    }

    for row in range(1, min(ws.max_row, 150) + 1):
        values = {
            normalizar_texto(ws.cell(row, col).value).upper()
            for col in range(1, min(ws.max_column, 12) + 1)
        }
        if expected.issubset(values):
            return row
    return None


def _buscar_hoja_presupuesto(workbook):
    preferred = [
        "01 Presupuesto",
        "Presupuesto",
        "PRESUPUESTO",
    ]
    for name in preferred:
        if name in workbook.sheetnames:
            ws = workbook[name]
            header = _buscar_encabezado_presupuesto(ws)
            if header:
                return ws, header

    for ws in workbook.worksheets:
        header = _buscar_encabezado_presupuesto(ws)
        if header:
            return ws, header

    raise RuntimeError(
        "No encontré una tabla de presupuesto reconocible. "
        "Se requieren las columnas Partida, Subpartida, Descripción Técnica y Unidad."
    )


def _mapear_columnas_excel(ws, header_row: int) -> dict:
    aliases = {
        "area": {"AREA", "ÁREA"},
        "partida": {"PARTIDA"},
        "subpartida": {"SUBPARTIDA"},
        "description": {
            "DESCRIPCION TECNICA",
            "DESCRIPCION",
            "CONCEPTO",
        },
        "unit": {"UNIDAD"},
        "quantity": {"CANT", "CANTIDAD", "CANT."},
        "unit_price": {
            "PRECIO UNITARIO MXN",
            "PRECIO UNITARIO",
            "P U",
            "P U VENTA",
        },
        "amount": {
            "IMPORTE INTERNO MXN",
            "IMPORTE INTERNO",
            "IMPORTE TOTAL MXN",
            "IMPORTE TOTAL",
            "IMPORTE",
        },
        "final_amount": {
            "IMPORTE FINAL MXN",
            "IMPORTE FINAL",
        },
        "included": {
            "CONSIDERAR",
            "INCLUIR",
            "ACTIVA",
            "ACTIVO",
        },
    }

    mapping = {}
    for col in range(1, ws.max_column + 1):
        key = normalizar_texto(ws.cell(header_row, col).value).upper()
        for field, options in aliases.items():
            if key in options and field not in mapping:
                mapping[field] = col

    required = {"partida", "subpartida", "description", "unit"}
    missing = required - set(mapping)
    if missing:
        raise RuntimeError(
            "Faltan columnas necesarias en el presupuesto: "
            + ", ".join(sorted(missing))
        )

    return mapping


def _leer_parametros_control(workbook, fallback_params: dict) -> tuple[dict, list[dict]]:
    """
    Recupera parámetros y filas de Control Interno cuando existen.
    Si la hoja no existe, conserva los parámetros actuales de la app.
    """
    params = dict(fallback_params)
    control_rows = []

    if "02 Control Interno" not in workbook.sheetnames:
        return params, control_rows

    ws = workbook["02 Control Interno"]

    # Parámetros por etiqueta, no por número de fila.
    for row in range(1, min(ws.max_row, 30) + 1):
        label = normalizar_texto(ws.cell(row, 1).value).upper()
        value = ws.cell(row, 2).value

        if label == "INDIRECTOS":
            params["indirect_pct"] = _pct_excel_a_porcentaje(
                value, params["indirect_pct"]
            )
        elif label.startswith("UTILIDAD"):
            params["profit_pct"] = _pct_excel_a_porcentaje(
                value, params["profit_pct"]
            )
        elif label == "IVA":
            params["iva_pct"] = _pct_excel_a_porcentaje(
                value, params["iva_pct"]
            )
        elif "DESPERDICIO" in label:
            params["waste_pct"] = _pct_excel_a_porcentaje(
                value, params["waste_pct"]
            )

    # Buscar encabezados de la tabla.
    header_row = None
    headers = {}
    for row in range(1, min(ws.max_row, 80) + 1):
        current = {}
        for col in range(1, ws.max_column + 1):
            value = normalizar_texto(ws.cell(row, col).value).upper()
            if value:
                current[value] = col
        has_cost = any(
            name in current
            for name in {
                "COSTO BASE UNIT",
                "COSTO BASE ESTIMADO UNIT",
                "PRECIO SUBCONTRATISTA UNIT",
            }
        )
        if "CODIGO" in current and has_cost:
            header_row = row
            headers = current
            break

    if header_row is None:
        return params, control_rows

    def col(*names):
        for name in names:
            if name in headers:
                return headers[name]
        return None

    c_code = col("CODIGO")
    c_title = col("TITULO COMERCIAL")
    # En V16 el precio negociable con subcontratista es la fuente principal
    # del costo directo al recargar el Excel. Se conserva compatibilidad V15.
    c_cost = col(
        "PRECIO SUBCONTRATISTA UNIT",
        "COSTO BASE UNIT",
        "COSTO BASE ESTIMADO UNIT",
    )
    c_mat = col("MATERIALES EST UNIT")
    c_labor = col("M O EST UNIT")
    c_other = col("OTROS INTEGRADO EST UNIT")
    c_waste_pct = col("DESPERDICIO MATERIALES REF")
    c_waste_unit = col("DESPERDICIO REF UNIT")
    c_lot = col("LOTE")
    c_order = None

    for row in range(header_row + 1, ws.max_row + 1):
        code = str(ws.cell(row, c_code).value or "").strip() if c_code else ""
        title = str(ws.cell(row, c_title).value or "").strip() if c_title else ""
        cost = _valor_float_excel(ws.cell(row, c_cost).value) if c_cost else 0.0

        if not code and not title and cost == 0:
            continue

        material_unit = (
            _valor_float_excel(ws.cell(row, c_mat).value) if c_mat else 0.0
        )
        labor_unit = (
            _valor_float_excel(ws.cell(row, c_labor).value) if c_labor else 0.0
        )
        other_unit = (
            _valor_float_excel(ws.cell(row, c_other).value) if c_other else 0.0
        )
        if c_waste_pct:
            waste_pct = _pct_excel_a_porcentaje(
                ws.cell(row, c_waste_pct).value, 0.0
            )
        elif c_waste_unit and material_unit > 0:
            waste_unit = _valor_float_excel(ws.cell(row, c_waste_unit).value, 0.0)
            waste_pct = max(waste_unit / material_unit * 100.0, 0.0)
        else:
            waste_pct = 0.0

        control_rows.append(
            {
                "code": code,
                "title": title,
                "unit_cost": cost,
                "material_unit": material_unit,
                "labor_unit": labor_unit,
                "other_unit": other_unit,
                "waste_pct": waste_pct,
                "contract_lot": (
                    str(ws.cell(row, c_lot).value or "1").strip()
                    if c_lot else "1"
                ),
                "execution_order": c_order,
            }
        )

    return params, control_rows


def _leer_metadatos_excel(ws, fallback_name: str = "") -> dict:
    """
    Lee el encabezado propio de la app cuando está disponible.

    Si el archivo es un presupuesto anterior o externo cuya primera fila ya es
    la tabla de conceptos, NO interpreta las primeras actividades como nombre,
    tipo de obra o ubicación.
    """
    safe_name = Path(str(fallback_name or "")).stem.strip() or "Proyecto importado"

    project_type = "Remodelación interior general"
    budget_level = "Medio-alto"
    location = "Ubicación por confirmar"
    project_code = "IMPORTADO"
    version = 1
    name = safe_name

    first_cell = normalizar_texto(ws["A1"].value or "").upper()
    meta = str(ws["A3"].value or "").strip()

    # Solo interpretar A2/A3 como metadatos si realmente existe el encabezado
    # producido por esta aplicación.
    has_app_header = (
        first_cell == "PRESUPUESTO"
        and "·" in meta
    )

    if has_app_header:
        name = str(ws["A2"].value or safe_name).strip() or safe_name
        parts = [x.strip() for x in meta.split("·") if x.strip()]

        if len(parts) >= 1:
            project_type = parts[0]
        if len(parts) >= 2 and parts[1] in NIVELES_PRESUPUESTO:
            budget_level = parts[1]

        version_index = None
        for idx, value in enumerate(parts):
            match = re.fullmatch(r"V(\d+)", value, flags=re.I)
            if match:
                version_index = idx
                version = max(int(match.group(1)), 1)
                if idx >= 1:
                    project_code = parts[idx - 1]
                break

        if version_index is not None:
            start_location = (
                2
                if len(parts) >= 2 and parts[1] in NIVELES_PRESUPUESTO
                else 1
            )
            end_location = max(version_index - 1, start_location)
            location_parts = parts[start_location:end_location]
            if location_parts:
                location = " · ".join(location_parts)
        elif len(parts) >= 3:
            location = parts[2]

    return {
        "name": name,
        "project_type": project_type,
        "budget_level": budget_level,
        "location": location,
        "project_code": project_code,
        "version": version,
    }


def importar_presupuesto_excel(
    excel_bytes: bytes,
    fallback_params: dict,
    file_name: str = "",
) -> dict:
    """
    Reconstruye un presupuesto editable desde un .xlsx.

    Prioridades:
    1. 01 Presupuesto = alcance, cantidades e importes vigentes.
    2. 02 Control Interno = costos internos y parámetros, si existe.
    3. Si Control Interno no existe, se reconstruye el costo base a partir del
       importe interno y los porcentajes actuales.
    """
    if not excel_bytes:
        raise RuntimeError("El archivo está vacío.")

    try:
        wb_values = load_workbook(
            filename=BytesIO(excel_bytes),
            data_only=True,
            read_only=False,
        )
    except Exception as exc:
        raise RuntimeError("No fue posible leer el archivo Excel.") from exc

    ws, header_row = _buscar_hoja_presupuesto(wb_values)
    columns = _mapear_columnas_excel(ws, header_row)
    metadata = _leer_metadatos_excel(
        ws,
        fallback_name=file_name,
    )

    params, control_rows = _leer_parametros_control(
        wb_values,
        fallback_params,
    )

    raw_rows = []
    blank_streak = 0

    for row in range(header_row + 1, ws.max_row + 1):
        part_raw = ws.cell(row, columns["partida"]).value
        sub_raw = ws.cell(row, columns["subpartida"]).value
        desc_raw = ws.cell(row, columns["description"]).value
        unit_raw = ws.cell(row, columns["unit"]).value

        part_text = str(part_raw or "").strip()
        if "PRESUPUESTO INTERNO" in normalizar_texto(part_text).upper():
            break

        has_content = any(
            str(value or "").strip()
            for value in (part_raw, sub_raw, desc_raw, unit_raw)
        )
        if not has_content:
            blank_streak += 1
            if blank_streak >= 3 and raw_rows:
                break
            continue
        blank_streak = 0

        description = str(desc_raw or "").strip()
        unit = str(unit_raw or "").strip().upper()
        if not description or not unit:
            continue

        quantity = (
            _valor_float_excel(ws.cell(row, columns["quantity"]).value, 1.0)
            if columns.get("quantity")
            else 1.0
        )
        quantity = max(quantity, 0.0)

        amount = (
            _valor_float_excel(ws.cell(row, columns["amount"]).value, 0.0)
            if columns.get("amount")
            else 0.0
        )
        unit_price = (
            _valor_float_excel(ws.cell(row, columns["unit_price"]).value, 0.0)
            if columns.get("unit_price")
            else 0.0
        )

        if amount <= 0 and unit_price > 0 and quantity > 0:
            amount = unit_price * quantity
        if unit_price <= 0 and amount > 0 and quantity > 0:
            unit_price = amount / quantity

        area_value = (
            normalizar_nombre_area(ws.cell(row, columns["area"]).value)
            if columns.get("area")
            else ""
        )
        included_value = (
            ws.cell(row, columns["included"]).value
            if columns.get("included")
            else "Sí"
        )
        included = normalizar_texto(included_value).upper() not in {
            "NO", "0", "FALSE", "FALSO"
        }

        raw_rows.append(
            {
                "area": area_value,
                "included": included,
                "category": normalizar_seccion_comercial(
                    _quitar_numeracion_excel(part_raw)
                ),
                "subcategory": _quitar_numeracion_excel(sub_raw),
                "description": description,
                "unit": unit,
                "quantity": quantity,
                "unit_sale": unit_price,
                "sale_amount": amount,
            }
        )

    if not raw_rows:
        raise RuntimeError(
            "No encontré conceptos utilizables dentro de la tabla del presupuesto."
        )

    # Descripción reconstruida: no inventa información; simplemente convierte
    # los renglones actuales en contexto para los ajustes posteriores con Gemini.
    description_lines = ["PRESUPUESTO RECARGADO DESDE EXCEL."]
    for row in raw_rows:
        description_lines.append(
            f"- [{row['area'] or AREA_GENERAL} / {row['category']} / {row['subcategory']}] "
            f"{row['description']} | {row['quantity']:g} {row['unit']}"
        )

    if metadata["project_code"] == "IMPORTADO":
        metadata["project_code"] = (
            f"{abreviar_cliente(metadata['name'])}-IMP-0001"
        )

    project_data = {
        "name": metadata["name"],
        "project_type": metadata["project_type"],
        "budget_level": metadata["budget_level"],
        "location": metadata["location"],
        "dimension_mode": "Recuperadas de Excel",
        "dimensions_text": "",
        "description": "\n".join(description_lines),
        "guide_text": DEFAULT_GUIDE_TEXT,
    }

    factor = (
        (1.0 + float(params["indirect_pct"]) / 100.0)
        * (1.0 + float(params["profit_pct"]) / 100.0)
    )
    if factor <= 0:
        factor = 1.0

    items = []
    for idx, row in enumerate(raw_rows):
        control = control_rows[idx] if idx < len(control_rows) else {}

        quantity = float(row["quantity"])
        sale_amount = float(row["sale_amount"])
        commercial_unit = (
            sale_amount / quantity
            if quantity > 0
            else float(row["unit_sale"] or 0.0)
        )

        unit_cost = float(control.get("unit_cost") or 0.0)
        if unit_cost <= 0:
            unit_cost = commercial_unit / factor if factor else commercial_unit

        code = str(control.get("code") or "").strip()
        if not code:
            code = f"IMP-{idx + 1:03d}"

        title = str(control.get("title") or "").strip()
        if not title:
            title = row["subcategory"] or re.split(
                r"[.;:]",
                row["description"],
                maxsplit=1,
            )[0][:80]

        direct_amount = quantity * unit_cost
        indirect_unit = unit_cost * params["indirect_pct"] / 100.0
        profit_unit = (
            unit_cost + indirect_unit
        ) * params["profit_pct"] / 100.0

        material_unit = float(control.get("material_unit") or 0.0)
        labor_unit = float(control.get("labor_unit") or 0.0)
        other_unit = float(control.get("other_unit") or 0.0)
        split_total = material_unit + labor_unit + other_unit

        if split_total > 0 and unit_cost > 0:
            material_share = material_unit / split_total * 100.0
            labor_share = labor_unit / split_total * 100.0
            other_share = other_unit / split_total * 100.0
        else:
            material_share = 0.0
            labor_share = 0.0
            other_share = 100.0

        benefit_amount = sale_amount - direct_amount
        margin = (
            benefit_amount / sale_amount * 100.0
            if sale_amount else 0.0
        )

        item = {
            "concept_id": None,
            "area_hint": row.get("area") or "",
            "category": row["category"],
            "subcategory": row["subcategory"],
            "code": limpiar_codigo(code, f"IMP-{idx + 1:03d}"),
            "execution_order": (idx + 1) * 10,
            "commercial_title": title,
            "description": row["description"],
            "unit": row["unit"],
            "quantity": quantity,
            "unit_cost": unit_cost,
            "direct_amount": direct_amount,
            "unit_indirect": indirect_unit,
            "unit_profit": profit_unit,
            "unit_sale": commercial_unit,
            "sale_amount": sale_amount,
            "benefit_amount": benefit_amount,
            "sale_margin_pct": margin,
            "price_source": "EXCEL_IMPORTADO",
            "price_source_detail": (
                "Precio recuperado del archivo Excel cargado por el usuario."
            ),
            "price_status": "IMPORTADO",
            "price_confidence": "Media",
            "material_share_pct": material_share,
            "labor_share_pct": labor_share,
            "other_share_pct": other_share,
            "waste_reference_pct": float(control.get("waste_pct") or 0.0),
            "included": bool(row.get("included", True)),
            "contract_lot": str(control.get("contract_lot") or "1"),
            "quantity_confidence": "Alta",
            "quantity_criterion": "Cantidad recuperada del archivo Excel.",
            "inclusion_basis": "Concepto existente en el presupuesto importado.",
            "considerations": "",
        }
        item = aplicar_composicion_costo(item)
        items.append(item)

    items = recalcular_areas_items(project_data, items)
    items = asignar_codigos_jerarquicos(items)

    activities = [item_a_actividad(item) for item in items]
    result = PresupuestoIA(
        nombre_proyecto=project_data["name"],
        actividad_principal=project_data["project_type"],
        alcance_resumido=(
            "Presupuesto reconstruido desde el archivo Excel cargado. "
            "Los conceptos existentes se conservan como alcance vigente."
        ),
        consideraciones_generales=[
            "Presupuesto recargado desde Excel para continuar con modificaciones."
        ],
        datos_faltantes=[],
        actividades=activities,
    )

    financials = calcular_financieros(items, params)

    excel_bytes_rebuilt = crear_paquete_excels(
        project_code=metadata["project_code"],
        project_data=project_data,
        result=result,
        items=items,
        params=params,
        version=metadata["version"],
    )

    return {
        "project_code": metadata["project_code"],
        "version": metadata["version"],
        "project_data": project_data,
        "params": params,
        "result": result,
        "items": items,
        "financials": financials,
        "excel_bytes": excel_bytes_rebuilt,
    }



def preparar_items_desde_editor_excel(
    editor_df: pd.DataFrame,
    current_items: list[dict],
    params: dict,
    project_data: dict,
) -> list[dict]:
    """Aplica cambios hechos en el editor rápido de Excel sin usar Gemini."""
    required = [
        "Área", "Partida", "Subpartida", "Descripción Técnica",
        "Unidad", "Cant.", "Precio Unitario (MXN)",
    ]
    if not isinstance(editor_df, pd.DataFrame):
        raise ValueError("El editor no devolvió una tabla válida.")
    for col in required:
        if col not in editor_df.columns:
            raise ValueError(f"Falta la columna requerida: {col}")

    def num(value, default=0.0):
        try:
            if pd.isna(value):
                return float(default)
        except Exception:
            pass
        try:
            return float(value)
        except (TypeError, ValueError):
            return float(default)

    factor = ((1.0 + float(params.get("indirect_pct", 0.0)) / 100.0)
              * (1.0 + float(params.get("profit_pct", 0.0)) / 100.0))
    if factor <= 0:
        factor = 1.0

    updated = []
    used_codes = {str(item.get("code") or "").strip().upper() for item in current_items}
    current_by_code = {str(item.get("code") or "").strip().upper(): dict(item) for item in current_items}

    for position, (_, row) in enumerate(editor_df.iterrows()):
        area = normalizar_nombre_area(str(row.get("Área") or "").strip())
        category = str(row.get("Partida") or "").strip()
        subcategory = str(row.get("Subpartida") or "").strip()
        description = str(row.get("Descripción Técnica") or "").strip()
        unit = str(row.get("Unidad") or "").strip().upper()
        quantity = max(num(row.get("Cant."), 0.0), 0.0)
        unit_sale = max(num(row.get("Precio Unitario (MXN)"), 0.0), 0.0)

        if not any((area, category, subcategory, description, unit)) and quantity == 0 and unit_sale == 0:
            continue
        if not description or not unit:
            raise ValueError(f"La fila {position + 1} necesita Descripción Técnica y Unidad.")
        if quantity <= 0:
            raise ValueError(f"La fila {position + 1} debe tener una Cantidad mayor a 0.")

        internal_code = str(row.get("__code") or "").strip().upper()
        if internal_code and internal_code in current_by_code:
            item = dict(current_by_code[internal_code])
        else:
            idx = 1
            while f"MAN-{idx:03d}" in used_codes:
                idx += 1
            code = f"MAN-{idx:03d}"
            used_codes.add(code)
            unit_cost = unit_sale / factor if factor else unit_sale
            indirect_unit = unit_cost * float(params.get("indirect_pct", 0.0)) / 100.0
            profit_unit = (unit_cost + indirect_unit) * float(params.get("profit_pct", 0.0)) / 100.0
            item = {
                "concept_id": None,
                "area_hint": area,
                "category": category or "General",
                "subcategory": subcategory or description[:80],
                "code": code,
                "execution_order": (len(current_items) + len(updated) + 1) * 10,
                "commercial_title": subcategory or description[:80],
                "concepto_base": description,
                "description": description,
                "unit": unit,
                "quantity": quantity,
                "unit_cost": unit_cost,
                "direct_amount": quantity * unit_cost,
                "unit_indirect": indirect_unit,
                "unit_profit": profit_unit,
                "unit_sale": unit_sale,
                "sale_amount": quantity * unit_sale,
                "benefit_amount": quantity * (unit_sale - unit_cost),
                "sale_margin_pct": ((unit_sale - unit_cost) / unit_sale * 100.0) if unit_sale else 0.0,
                "price_source": "EDITOR_EXCEL",
                "price_source_detail": "Actividad agregada manualmente desde el editor de Excel.",
                "price_status": "AGREGADO_MANUAL",
                "price_confidence": "Alta",
                "material_share_pct": 0.0,
                "labor_share_pct": 0.0,
                "other_share_pct": 100.0,
                "waste_reference_pct": 0.0,
                "included": True,
                "contract_lot": "1",
                "quantity_confidence": "Alta",
                "quantity_criterion": "Cantidad capturada manualmente en el editor de Excel.",
                "inclusion_basis": "Actividad agregada manualmente en el editor de Excel.",
                "considerations": "",
            }

        item["area_hint"] = area
        item["category"] = category or item.get("category") or "General"
        item["subcategory"] = subcategory or item.get("subcategory") or description[:80]
        item["description"] = description
        item["unit"] = unit
        item["quantity"] = quantity
        item["unit_sale"] = unit_sale
        item["sale_amount"] = quantity * unit_sale

        # El precio unitario capturado por el usuario es el precio interno de salida.
        # Solo reconstruimos el desglose oculto necesario para que el resto del app siga consistente.
        recalculated = recalcular_item_financiero(item, params)
        recalculated["unit_sale"] = unit_sale
        recalculated["sale_amount"] = quantity * unit_sale
        recalculated["benefit_amount"] = recalculated["sale_amount"] - float(recalculated.get("direct_amount") or 0.0)
        recalculated["sale_margin_pct"] = (
            recalculated["benefit_amount"] / recalculated["sale_amount"] * 100.0
            if recalculated["sale_amount"] else 0.0
        )
        recalculated["price_source"] = "EDITOR_EXCEL"
        recalculated["price_source_detail"] = "Precio actualizado manualmente desde el editor de Excel."
        recalculated["price_status"] = "EDITADO_MANUAL"
        updated.append(recalculated)

    if not updated:
        raise ValueError("El editor no contiene actividades con información válida.")
    return recalcular_areas_items(project_data, updated)


def _numero_excel_flexible(value, default=0.0) -> float:
    """Lee números pegados desde Excel tanto con punto como con coma decimal."""
    if value is None:
        return float(default)
    if isinstance(value, (int, float)):
        return float(value)
    raw = str(value).strip().replace("$", "").replace("%", "").replace(" ", "")
    if not raw:
        return float(default)
    try:
        if "," in raw and "." in raw:
            # 1,234.56 -> 1234.56
            raw = raw.replace(",", "")
        elif "," in raw:
            tail = raw.rsplit(",", 1)[-1]
            if len(tail) <= 2:
                raw = raw.replace(".", "").replace(",", ".")
            else:
                raw = raw.replace(",", "")
        return float(raw)
    except Exception:
        return float(default)


def _normalizar_encabezado_pegado(value) -> str:
    return normalizar_texto(value).upper()


def parsear_actividades_pegadas(texto: str) -> pd.DataFrame:
    """Convierte un bloque pegado desde Excel en las 7 columnas del editor.

    Acepta encabezados en cualquier orden, alias habituales y bloques sin encabezado.
    Excel normalmente pega TSV, pero también se toleran ; y ,.
    """
    if not str(texto or "").strip():
        raise ValueError("No hay datos para pegar.")

    raw = str(texto).replace("\r\n", "\n").replace("\r", "\n").strip()
    lines = [line for line in raw.split("\n") if line.strip()]
    if not lines:
        raise ValueError("No hay filas utilizables.")

    sample = lines[0]
    if "\t" in sample:
        sep = "\t"
    elif ";" in sample:
        sep = ";"
    else:
        sep = ","

    rows = [line.split(sep) for line in lines]
    target_cols = ["Área", "Partida", "Subpartida", "Descripción Técnica", "Unidad", "Cant.", "Precio Unitario (MXN)"]
    aliases = {
        "area": {"AREA", "ÁREA"},
        "partida": {"PARTIDA"},
        "subpartida": {"SUBPARTIDA"},
        "description": {"DESCRIPCION TECNICA", "DESCRIPCION", "CONCEPTO", "DESCRIPCION DEL CONCEPTO", "CONCEPTO / DESCRIPCION"},
        "unit": {"UNIDAD", "U", "UN."},
        "quantity": {"CANT", "CANTIDAD", "CANT.", "CANTIDAD FISICA"},
        "unit_price": {"PRECIO UNITARIO MXN", "PRECIO UNITARIO", "P U", "P U VENTA", "P.U.", "PRECIO"},
    }
    normalized_aliases = {_normalizar_encabezado_pegado(x) for vals in aliases.values() for x in vals}
    first_norm = [_normalizar_encabezado_pegado(x) for x in rows[0]]
    has_header = len(set(first_norm) & normalized_aliases) >= 3

    if has_header:
        positions = {}
        for i, h in enumerate(first_norm):
            for field, opts in aliases.items():
                if h in {_normalizar_encabezado_pegado(x) for x in opts} and field not in positions:
                    positions[field] = i
        data_rows = rows[1:]
    else:
        positions = {field: i for i, field in enumerate(["area", "partida", "subcategory", "description", "unit", "quantity", "unit_price"])}
        # Corrección para el orden estándar solicitado: Área, Partida, Subpartida...
        positions = {"area":0, "partida":1, "subcategory":2, "description":3, "unit":4, "quantity":5, "unit_price":6}
        data_rows = rows

    missing = [f for f in ["description", "unit"] if f not in positions]
    if missing:
        raise ValueError("No pude detectar las columnas mínimas: Descripción Técnica y Unidad.")

    out = []
    last_area = ""
    last_partida = ""
    last_subpartida = ""
    for row in data_rows:
        def cell(field):
            i = positions.get(field)
            return row[i].strip() if i is not None and i < len(row) else ""

        area = cell("area")
        partida = cell("partida")
        subpartida = cell("subpartida")
        description = cell("description")
        unit = cell("unit")
        qty_raw = cell("quantity")
        price_raw = cell("unit_price")

        # Excel puede traer celdas combinadas como vacías: heredamos la última clasificación.
        if area:
            last_area = area
        else:
            area = last_area
        if partida:
            last_partida = _quitar_numeracion_excel(partida)
        else:
            partida = last_partida
        if subpartida:
            last_subpartida = _quitar_numeracion_excel(subpartida)
        else:
            subpartida = last_subpartida

        description = description.strip()
        unit = normalizar_unidad(unit)
        if not description and not qty_raw and not price_raw:
            continue
        if not description or not unit:
            raise ValueError("Cada fila pegada necesita al menos Descripción Técnica y Unidad.")

        out.append({
            "Área": normalizar_nombre_area(area),
            "Partida": normalizar_seccion_comercial(_quitar_numeracion_excel(partida)) if partida else "OTROS TRABAJOS",
            "Subpartida": _quitar_numeracion_excel(subpartida) if subpartida else description[:80],
            "Descripción Técnica": description,
            "Unidad": unit,
            "Cant.": max(_numero_excel_flexible(qty_raw, 1.0), 0.0),
            "Precio Unitario (MXN)": max(_numero_excel_flexible(price_raw, 0.0), 0.0),
        })

    if not out:
        raise ValueError("No encontré actividades utilizables en el bloque pegado.")
    return pd.DataFrame(out, columns=target_cols)


def insertar_actividades_pegadas(
    current_items: list[dict],
    pasted_df: pd.DataFrame,
    params: dict,
    project_data: dict,
    mode: str = "Automático",
    anchor_code: str = "",
) -> list[dict]:
    """Inserta actividades por Partida/Subpartida o en una posición exacta."""
    if pasted_df.empty:
        return list(current_items)

    # Convertir las filas pegadas usando el mismo motor financiero del editor.
    base_df = pasted_df.copy()
    base_rows = []
    for _, row in base_df.iterrows():
        base_rows.append({
            "Área": row.get("Área", ""),
            "Partida": row.get("Partida", ""),
            "Subpartida": row.get("Subpartida", ""),
            "Descripción Técnica": row.get("Descripción Técnica", ""),
            "Unidad": row.get("Unidad", ""),
            "Cant.": row.get("Cant.", 0),
            "Precio Unitario (MXN)": row.get("Precio Unitario (MXN)", 0),
        })

    # Preparar primero como nuevas filas, sin depender de posiciones actuales.
    empty = []
    prepared = preparar_items_desde_editor_excel(pd.DataFrame(base_rows, columns=base_df.columns), empty, params, project_data)

    result = [dict(x) for x in current_items]
    if mode.startswith("Después de") and anchor_code:
        idx = next((i for i,x in enumerate(result) if str(x.get("code")) == anchor_code), len(result)-1)
        result[idx+1:idx+1] = prepared
    elif mode.startswith("Antes de") and anchor_code:
        idx = next((i for i,x in enumerate(result) if str(x.get("code")) == anchor_code), len(result))
        result[idx:idx] = prepared
    elif mode == "Al final":
        result.extend(prepared)
    else:
        # Automático: mismo Partida + Subpartida -> después del último de ese grupo;
        # si solo coincide Partida -> después de la última actividad de esa Partida;
        # si es una Partida nueva -> se coloca donde corresponde a la secuencia comercial.
        for new_item in prepared:
            category = normalizar_seccion_comercial(new_item.get("category"))
            sub = normalizar_texto(new_item.get("subcategory") or "").upper()
            exact = [i for i,x in enumerate(result) if normalizar_seccion_comercial(x.get("category")) == category and normalizar_texto(x.get("subcategory") or "").upper() == sub and sub]
            same = [i for i,x in enumerate(result) if normalizar_seccion_comercial(x.get("category")) == category]
            if exact:
                insert_at = exact[-1] + 1
            elif same:
                insert_at = same[-1] + 1
            else:
                preferred = {name:i for i,name in enumerate(SECCIONES_COMERCIALES_PREFERENTES)}
                target_phase = preferred.get(category, preferred.get("OTROS TRABAJOS", 10))
                insert_at = len(result)
                for i,x in enumerate(result):
                    phase = preferred.get(normalizar_seccion_comercial(x.get("category")), 10)
                    if phase > target_phase:
                        insert_at = i
                        break
            result.insert(insert_at, new_item)

    for i, item in enumerate(result, start=1):
        item["execution_order"] = i * 10
    return result


# =========================================================
# EXCEL
# =========================================================



EDITOR_EXCEL_COLUMNS = [
    "Área", "Partida", "Subpartida", "Descripción Técnica",
    "Unidad", "Cant.", "Precio Unitario (MXN)",
]


def crear_excel_desde_editor(df: pd.DataFrame) -> bytes:
    """Crea un Excel limpio con únicamente las 7 columnas del editor."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Presupuesto"

    data = df.copy()
    for col in EDITOR_EXCEL_COLUMNS:
        if col not in data.columns:
            data[col] = ""
    data = data[EDITOR_EXCEL_COLUMNS]

    mask = data.apply(
        lambda row: any(str(v).strip() not in {"", "nan", "None"} for v in row),
        axis=1,
    )
    data = data.loc[mask].copy()

    header_fill = PatternFill("solid", fgColor="4A342B")
    header_font = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="D4D4D4")

    for col_idx, name in enumerate(EDITOR_EXCEL_COLUMNS, start=1):
        cell = ws.cell(1, col_idx, name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=thin)

    for row_idx, (_, row) in enumerate(data.iterrows(), start=2):
        for col_idx, name in enumerate(EDITOR_EXCEL_COLUMNS, start=1):
            value = row[name]
            if name in {"Cant.", "Precio Unitario (MXN)"}:
                try:
                    value = float(value) if str(value).strip() else None
                except (TypeError, ValueError):
                    value = None
            else:
                value = "" if pd.isna(value) else str(value).strip()

            cell = ws.cell(row_idx, col_idx, value)
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=name == "Descripción Técnica",
            )
            if name == "Cant." and value is not None:
                cell.number_format = "0.00"
            elif name == "Precio Unitario (MXN)" and value is not None:
                cell.number_format = '$#,##0.00'

    widths = {
        "Área": 20, "Partida": 30, "Subpartida": 25,
        "Descripción Técnica": 65, "Unidad": 12, "Cant.": 12,
        "Precio Unitario (MXN)": 22,
    }
    for idx, name in enumerate(EDITOR_EXCEL_COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = widths[name]

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:G{max(ws.max_row, 1)}"
    ws.sheet_view.showGridLines = False

    if ws.max_row >= 2:
        from openpyxl.worksheet.table import Table, TableStyleInfo
        table = Table(displayName="TablaPresupuesto", ref=f"A1:G{ws.max_row}")
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False, showLastColumn=False,
            showRowStripes=True, showColumnStripes=False,
        )
        ws.add_table(table)

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


def render_editor_excel_nuevo():
    """Editor independiente para construir un presupuesto Excel desde cero."""
    st.header("Crear presupuesto en Excel")
    st.caption(
        "Construye un archivo nuevo sin cargar un presupuesto existente. "
        "Puedes capturar filas directamente o pegar un bloque copiado desde Excel."
    )

    st.info(
        "El archivo exportado contiene únicamente: Área, Partida, Subpartida, "
        "Descripción Técnica, Unidad, Cant. y Precio Unitario (MXN)."
    )

    def filas_vacias(n=5):
        return pd.DataFrame(
            [
                {
                    "Área": "", "Partida": "", "Subpartida": "",
                    "Descripción Técnica": "", "Unidad": "",
                    "Cant.": 1.0, "Precio Unitario (MXN)": 0.0,
                }
                for _ in range(n)
            ],
            columns=EDITOR_EXCEL_COLUMNS,
        )

    if "new_excel_editor_df" not in st.session_state:
        st.session_state["new_excel_editor_df"] = filas_vacias()

    with st.expander("Pegar actividades desde Excel", expanded=True):
        st.caption(
            "Copia directamente desde Excel. Se detectan encabezados aunque estén "
            "en otro orden. También puedes pegar filas sin encabezado usando el orden estándar."
        )
        paste_text = st.text_area(
            "Pega aquí las filas copiadas",
            height=180,
            placeholder=(
                "Área\tPartida\tSubpartida\tDescripción Técnica\tUnidad\tCant.\tPrecio Unitario (MXN)\n"
                "Cocina\tAcabados\tMuros\tSuministro y aplicación de pintura\tM2\t22.17\t185.00"
            ),
            key="new_excel_paste_text",
        )
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Cargar al editor", type="secondary", use_container_width=True):
                try:
                    pasted = parsear_actividades_pegadas(paste_text)
                    st.session_state["new_excel_editor_df"] = pasted[
                        EDITOR_EXCEL_COLUMNS
                    ].copy()
                    st.success(f"{len(pasted)} actividad(es) cargada(s).")
                except Exception as exc:
                    st.error(f"No fue posible leer el bloque: {exc}")
        with c2:
            if st.button("Limpiar editor", use_container_width=True):
                st.session_state["new_excel_editor_df"] = filas_vacias()
                st.rerun()

    st.subheader("Editor")
    edited = st.data_editor(
        st.session_state["new_excel_editor_df"],
        key="new_excel_editor",
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_order=EDITOR_EXCEL_COLUMNS,
        column_config={
            "Área": st.column_config.TextColumn("Área", width="small"),
            "Partida": st.column_config.TextColumn("Partida", width="medium"),
            "Subpartida": st.column_config.TextColumn("Subpartida", width="medium"),
            "Descripción Técnica": st.column_config.TextColumn(
                "Descripción Técnica", width="large"
            ),
            "Unidad": st.column_config.TextColumn("Unidad", width="small"),
            "Cant.": st.column_config.NumberColumn("Cant.", min_value=0.0, step=0.01),
            "Precio Unitario (MXN)": st.column_config.NumberColumn(
                "Precio Unitario (MXN)", min_value=0.0, step=0.01, format="$ %.2f"
            ),
        },
    )
    st.session_state["new_excel_editor_df"] = edited.copy()

    nonempty = edited.apply(
        lambda row: any(str(v).strip() not in {"", "nan", "None"} for v in row),
        axis=1,
    )
    st.caption(f"{int(nonempty.sum())} fila(s) con contenido.")

    excel_bytes = crear_excel_desde_editor(edited)
    st.download_button(
        "Exportar Excel",
        data=excel_bytes,
        file_name="Presupuesto_nuevo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )


def crear_excel(
    project_code: str,
    project_data: dict,
    result: PresupuestoIA,
    items: list[dict],
    params: dict,
    version: int = 1,
) -> bytes:
    """
    Libro de presupuesto con estructura de Partida/Subpartida como la utilizada
    por la empresa.

    01 Presupuesto:
      Área | Partida | Subpartida | Descripción Técnica | Unidad | Cant. |
      Precio Unitario | Importe interno | Considerar

      El único importe comercial mostrado es el Importe interno, que es la base
      de negociación con subcontrataciones. La marca y el IVA se aplican fuera
      de este módulo, en la etapa posterior de venta al cliente.
      Considerar permite activar/desactivar cada actividad sin borrar la fila.

    """
    wb = Workbook()
    # Forzar recálculo al abrir/guardar para que Excel y hojas compatibles
    # actualicen todos los enlaces entre Presupuesto y Control Interno.
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcOnSave = True
    ws = wb.active
    ws.title = "01 Presupuesto"

    brown = "4A342B"
    brown_light = "EDE7E3"
    gray = "E7E7E7"
    gray_light = "F5F5F5"
    editable_fill = "FFF8E7"
    formula_fill = "F2F2F2"
    dark_gray = "555555"
    white = "FFFFFF"
    internal_blue = "1F4E78"
    trace_orange = "FCE4D6"

    thin_gray = Side(style="thin", color="D4D4D4")

    items = recalcular_areas_items(project_data, items)
    structured_items = estructura_partidas_excel(items)
    ordered_items = [dict(x) for x in structured_items]
    commercial_row_map = {}

    # -----------------------------------------------------
    # 01 PRESUPUESTO - DATOS DEL PROYECTO
    # -----------------------------------------------------
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:I1")
    ws["A1"] = "PRESUPUESTO"
    ws["A1"].font = Font(size=20, bold=True, color=brown)

    ws.merge_cells("A2:I2")
    ws["A2"] = project_data["name"]
    ws["A2"].font = Font(size=12, bold=True, color=brown)

    ws.merge_cells("A3:I3")
    ws["A3"] = (
        f"{project_data['project_type']} · {project_data.get('budget_level', 'Medio-alto')} · "
        f"{project_data['location']} · {project_code} · V{version:02d}"
    )
    ws["A3"].font = Font(size=9, color=dark_gray)

    # -----------------------------------------------------
    # RESUMEN POR PARTIDAS
    # H = Importe interno.
    # -----------------------------------------------------
    ws.merge_cells("A5:I5")
    ws["A5"] = "RESUMEN"
    ws["A5"].font = Font(size=13, bold=True, color=brown)

    summary_row = 6
    summary_map = {}
    sections = []

    for item in structured_items:
        section = normalizar_seccion_comercial(item.get("category"))
        if section not in sections:
            sections.append(section)

    for section_idx, section in enumerate(sections, start=1):
        part_name = nombre_partida_excel(section)
        ws.merge_cells(
            start_row=summary_row,
            start_column=1,
            end_row=summary_row,
            end_column=7,
        )
        ws.cell(summary_row, 1, f"{section_idx}. {part_name}")
        ws.cell(summary_row, 1).fill = PatternFill("solid", fgColor=gray_light)
        ws.cell(summary_row, 8).fill = PatternFill("solid", fgColor=gray_light)
        summary_map[section] = summary_row
        summary_row += 1

    internal_summary_row = summary_row
    ws.merge_cells(
        start_row=summary_row, start_column=1, end_row=summary_row, end_column=7
    )
    ws.cell(summary_row, 1, "Presupuesto interno (MXN)")
    ws.cell(summary_row, 1).font = Font(bold=True)
    ws.cell(summary_row, 1).fill = PatternFill("solid", fgColor=gray)
    ws.cell(summary_row, 8).fill = PatternFill("solid", fgColor=gray)

    # -----------------------------------------------------
    # TABLA DE PARTIDAS Y SUBPARTIDAS
    # -----------------------------------------------------
    table_header_row = summary_row + 2
    headers = [
        "Área",
        "Partida",
        "Subpartida",
        "Descripción Técnica",
        "Unidad",
        "Cant.",
        "Precio Unitario (MXN)",
        "Importe interno (MXN)",
        "Considerar",
    ]

    for col, header in enumerate(headers, 1):
        cell = ws.cell(table_header_row, col, header)
        cell.font = Font(bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=brown)
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

    row = table_header_row + 1
    section_amount_rows = {section: [] for section in sections}

    for item in structured_items:
        commercial_row_map[item["code"]] = row
        section = normalizar_seccion_comercial(item.get("category"))

        ws.cell(row, 1, area_excel_item(item))
        ws.cell(row, 2, item["partida_excel"])
        ws.cell(row, 3, item["subpartida_excel"])
        ws.cell(row, 4, descripcion_excel_item(item))
        ws.cell(row, 5, item["unit"])
        ws.cell(row, 6, float(item["quantity"]))

        # G es el Precio Unitario editable. El Importe interno (H) se calcula
        # siempre como Precio Unitario x Cantidad para facilitar ajustes manuales.
        quantity = float(item["quantity"] or 0.0)
        initial_unit_price = (
            float(item["sale_amount"]) / quantity if quantity else 0.0
        )
        ws.cell(row, 7, initial_unit_price)
        ws.cell(row, 7).comment = Comment(
            f"Fuente: {item.get('price_source') or 'Sin fuente'}\n"
            f"Confianza: {item.get('price_confidence') or 'Sin dato'}\n"
            f"{item.get('price_source_detail') or ''}\n"
            f"{item.get('considerations') or ''}",
            "Presupuesto",
        )
        ws.cell(row, 8, f"=F{row}*G{row}")

        # I controla si la actividad participa o no en el presupuesto.
        ws.cell(row, 9, "Sí" if item_esta_incluido(item) else "No")

        ws.cell(row, 6).number_format = "0.00"
        for col in (7, 8):
            ws.cell(row, col).number_format = '$#,##0.00'

        for col in range(1, 10):
            cell = ws.cell(row, col)
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=col in {1, 2, 3, 4},
                horizontal="center" if col in {5, 6, 9} else "left",
            )
            cell.border = Border(bottom=thin_gray)

        for col in (7, 8):
            ws.cell(row, col).alignment = Alignment(horizontal="right")

        # G es editable; H es fórmula (Precio Unitario x Cantidad); I es editable.
        ws.cell(row, 7).fill = PatternFill("solid", fgColor=editable_fill)
        ws.cell(row, 8).fill = PatternFill("solid", fgColor=formula_fill)
        ws.cell(row, 9).fill = PatternFill("solid", fgColor=editable_fill)

        ws.row_dimensions[row].height = max(
            34,
            min(
                78,
                22 + (len(descripcion_excel_item(item)) // 95) * 14,
            ),
        )

        section_amount_rows[section].append(row)
        row += 1

    first_item_row = table_header_row + 1
    last_item_row = table_header_row + len(structured_items)
    if structured_items:
        include_validation = DataValidation(
            type="list",
            formula1='"Sí,No"',
            allow_blank=False,
        )
        include_validation.error = 'Selecciona únicamente Sí o No.'
        include_validation.errorTitle = 'Valor no válido'
        ws.add_data_validation(include_validation)
        include_validation.add(f"I{first_item_row}:I{last_item_row}")

        # Una actividad desactivada sigue visible, pero se atenúa para que sea
        # evidente que ya no participa en los totales.
        inactive_fill = PatternFill("solid", fgColor="E7E6E6")
        inactive_font = Font(color="7F7F7F")
        ws.conditional_formatting.add(
            f"A{first_item_row}:I{last_item_row}",
            FormulaRule(
                formula=[f'$I{first_item_row}="No"'],
                fill=inactive_fill,
                font=inactive_font,
            ),
        )

    # -----------------------------------------------------
    # TOTAL INTERNO AL FINAL DE LA TABLA
    # -----------------------------------------------------
    row += 1
    internal_detail_row = row
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    ws.cell(row, 1, "Presupuesto interno (MXN)")
    ws.cell(row, 1).font = Font(bold=True)
    ws.cell(
        row,
        8,
        f'=SUMIF(I{table_header_row + 1}:I{row - 2},"Sí",H{table_header_row + 1}:H{row - 2})',
    )
    ws.cell(row, 8).number_format = '$#,##0.00'
    ws.cell(row, 8).font = Font(bold=True)
    ws.cell(row, 8).fill = PatternFill("solid", fgColor=gray)
    ws.cell(row, 1).fill = PatternFill("solid", fgColor=gray)

    # Resumen superior: subtotal interno por partida.
    for section in sections:
        amount_rows = section_amount_rows.get(section) or []
        internal_formula = (
            "+".join(f'IF(I{r}="Sí",H{r},0)' for r in amount_rows)
            if amount_rows else "0"
        )
        sr = summary_map[section]
        ws.cell(sr, 8, f"={internal_formula}")
        ws.cell(sr, 8).number_format = '$#,##0.00'
        ws.cell(sr, 8).alignment = Alignment(horizontal="right")

    ws.cell(internal_summary_row, 8, f"=H{internal_detail_row}")
    ws.cell(internal_summary_row, 8).number_format = '$#,##0.00'


    widths = [18, 23, 25, 68, 11, 11, 21, 22, 13]
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width

    ws.row_dimensions[table_header_row].height = 32
    ws.auto_filter.ref = (
        f"A{table_header_row}:I{table_header_row + len(structured_items)}"
    )

    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.45
    ws.page_margins.bottom = 0.45

    out = BytesIO()
    wb.save(out)
    return out.getvalue()


# =========================================================
# EXCEL — FORMATO CLIENTE (Partidas y totales)
# =========================================================

# Estilo tomado directamente del archivo de ejemplo (AQUI PRO). Se deja como
# constante para conservar el formato cliente y evitar repetir colores.
_CLIENTE_HEADER_FILL = PatternFill(fill_type="solid", fgColor="FF37241B")
_CLIENTE_HEADER_FONT = Font(name="Calibri", size=12, bold=True, color="FFEEEEEE")
_CLIENTE_DATA_FONT = Font(name="Calibri", size=12, bold=False, color="FF37241B")
_CLIENTE_HIGHLIGHT_FILL = PatternFill(fill_type="solid", fgColor="FF7A7776")
_CLIENTE_RIGHT_ALIGN = Alignment(horizontal="right")


def crear_excel_formato_cliente(
    project_code: str,
    project_data: dict,
    items: list[dict],
    params: dict,
    version: int = 1,
    margin_pct: float | None = None,
) -> bytes:
    """
    Libro con el formato "cliente" (una hoja: Partidas con totales), basado
    del ejemplo proporcionado por la empresa.

    Reutiliza estructura_partidas_excel(...) -- la misma función que arma
    01 Presupuesto -- para que los capítulos y su orden sean siempre
    equivalentes entre ambos archivos.

    Coste  = Importe interno de cada partida (item["sale_amount"]).
    Margen = constante de negocio MARGEN_PRESUPUESTO_CLIENTE_PCT (30%).
    Precio = fórmula Coste * (1 + Margen / 100), se recalcula sola en Excel.
    """
    margin_pct = (
        float(margin_pct) if margin_pct is not None else MARGEN_PRESUPUESTO_CLIENTE_PCT
    )
    iva_pct = float(params.get("iva_pct", 16.0))

    incluidos = [it for it in items if item_esta_incluido(it)]
    structured_items = estructura_partidas_excel(incluidos)

    wb = Workbook()
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcOnSave = True

    partidas = wb.active
    partidas.title = "Partidas"
    partidas.sheet_view.showGridLines = True
    partidas.column_dimensions["B"].width = 25
    partidas.column_dimensions["C"].width = 55

    headers = [
        "Código", "Capítulo", "Partida", "Descripción", "Uds.",
        "Tipo Ud.", "Margen", "Coste", "Precio", "% Impuestos",
    ]
    for col, header in enumerate(headers, start=1):
        c = partidas.cell(1, col, header)
        c.font = _CLIENTE_HEADER_FONT
        c.fill = _CLIENTE_HEADER_FILL

    row = 2
    current_chapter = None
    for item in structured_items:
        part_number = item["part_number"]
        subpart_number = item["subpart_number"]

        if part_number != current_chapter:
            current_chapter = part_number
            code_cell = partidas.cell(row, 1, codigo_capitulo_cliente(part_number))
            code_cell.font = _CLIENTE_DATA_FONT
            code_cell.fill = _CLIENTE_HIGHLIGHT_FILL
            cap_cell = partidas.cell(
                row, 2, nombre_partida_excel(item.get("category")).upper()
            )
            cap_cell.font = _CLIENTE_DATA_FONT
            cap_cell.alignment = _CLIENTE_RIGHT_ALIGN
            row += 1

        code_cell = partidas.cell(
            row, 1, codigo_partida_cliente(part_number, subpart_number)
        )
        code_cell.font = _CLIENTE_DATA_FONT
        code_cell.fill = _CLIENTE_HIGHLIGHT_FILL

        values = [
            (3, nombre_subpartida_excel(item)),
            (4, item.get("description", "")),
            (5, float(item.get("quantity", 0.0))),
            (6, str(item.get("unit", "")).lower()),
            (7, margin_pct),
        ]
        for col, val in values:
            c = partidas.cell(row, col, val)
            c.font = _CLIENTE_DATA_FONT
            c.alignment = _CLIENTE_RIGHT_ALIGN

        coste_cell = partidas.cell(row, 8, float(item.get("sale_amount", 0.0)))
        coste_cell.font = _CLIENTE_DATA_FONT
        coste_cell.alignment = _CLIENTE_RIGHT_ALIGN

        precio_cell = partidas.cell(row, 9, f"=H{row}*(1+G{row}/100)")
        precio_cell.font = _CLIENTE_DATA_FONT
        precio_cell.fill = _CLIENTE_HIGHLIGHT_FILL

        row += 1

    last_item_row = row - 1
    subtotal_row = row + 1
    totals = [
        ("Subtotal", f"=SUM(I2:I{last_item_row})" if last_item_row >= 2 else 0),
        ("IVA", f"=I{subtotal_row}*{iva_pct / 100.0:.6f}"),
        ("Total", f"=I{subtotal_row}+I{subtotal_row + 1}"),
    ]
    for offset, (label, formula) in enumerate(totals):
        rr = subtotal_row + offset
        partidas.cell(rr, 8, label).font = _CLIENTE_HEADER_FONT
        partidas.cell(rr, 8).fill = _CLIENTE_HEADER_FILL
        partidas.cell(rr, 9, formula).number_format = '$#,##0.00'
        partidas.cell(rr, 9).font = Font(bold=True)
    for col in ('H', 'I'):
        partidas.column_dimensions[col].width = 20
    partidas.column_dimensions['D'].width = 65
    for rr in range(2, last_item_row + 1):
        partidas.cell(rr, 4).alignment = Alignment(wrap_text=True, vertical='top')
        partidas.cell(rr, 10, iva_pct)
        partidas.cell(rr, 8).number_format = '$#,##0.00'
        partidas.cell(rr, 9).number_format = '$#,##0.00'
    partidas.freeze_panes = 'E2'
    partidas.sheet_properties.pageSetUpPr.fitToPage = True
    partidas.page_setup.orientation = 'landscape'
    partidas.page_setup.fitToWidth = 1
    partidas.page_setup.fitToHeight = 0
    partidas.print_options.horizontalCentered = True
    partidas.print_title_rows = '1:1'
    partidas.oddHeader.center.text = f"{project_data.get('name', '')} | {project_code} | V{version:02d}"

    out = BytesIO()
    wb.save(out)
    return out.getvalue()


def empaquetar_excels_zip(
    excel_interno: bytes, excel_cliente: bytes, project_code: str, version: int
) -> bytes:
    """Empaqueta el Excel interno (negociación) y el Excel cliente (Partidas) en un único .zip para que ambos se descarguen de una sola vez."""
    buf = BytesIO()
    tag = f"{project_code}-V{version:02d}"
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(f"{tag}_Presupuesto_interno.xlsx", excel_interno)
        zf.writestr(f"{tag}_Presupuesto_cliente.xlsx", excel_cliente)
    buf.seek(0)
    return buf.getvalue()


def crear_paquete_excels(
    project_code: str,
    project_data: dict,
    result: PresupuestoIA,
    items: list[dict],
    params: dict,
    version: int = 1,
) -> bytes:
    """
    Genera ambos archivos -- el interno de siempre (crear_excel) y el nuevo
    formato cliente (crear_excel_formato_cliente) -- y los entrega juntos en
    un .zip. Mismo orden de argumentos que crear_excel para poder sustituir
    las llamadas existentes sin tocar el resto de cada flujo.
    """
    excel_interno = crear_excel(
        project_code=project_code,
        project_data=project_data,
        result=result,
        items=items,
        params=params,
        version=version,
    )
    excel_cliente = crear_excel_formato_cliente(
        project_code=project_code,
        project_data=project_data,
        items=items,
        params=params,
        version=version,
    )
    return empaquetar_excels_zip(excel_interno, excel_cliente, project_code, version)


# =========================================================
# DATAFRAMES DE PRESENTACIÓN
# =========================================================


def dataframe_resumen(items: list[dict]) -> pd.DataFrame:
    rows = []
    for x in estructura_partidas_excel(items):
        rows.append(
            {
                "Partida": x["partida_excel"],
                "Subpartida": x["subpartida_excel"],
                "Descripción Técnica": descripcion_excel_item(x),
                "Unidad": x["unit"],
                "Cant.": x["quantity"],
                "Precio Unitario": x["unit_sale"],
                "Importe interno": x["sale_amount"],
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

    (
        tab_concepts,
        tab_prices,
        tab_projects,
        tab_budgets,
        tab_maintenance,
        tab_export,
    ) = st.tabs(
        [
            "Conceptos",
            "Precios",
            "Proyectos",
            "Presupuestos",
            "Mantenimiento",
            "Exportar",
        ]
    )

    source_labels = {
        "COTIZACION_PROVEEDOR": "Cotización de proveedor",
        "COSTO_REAL": "Costo real de obra",
        "IA_ESTIMADO": "Estimación de IA",
        "MANUAL": "Registro manual",
        "BASE_INTERNA": "Base interna",
        "HISTORICO_IA": "Histórico generado por IA",
    }
    status_labels = {
        "VALIDADO": "Validado",
        "COTIZADO_PROVEEDOR": "Cotizado por proveedor",
        "COSTO_REAL": "Costo real",
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
                            "IA_ESTIMADO",
                    "MANUAL",
                ]
                status_options = [
                    "VALIDADO",
                    "COTIZADO_PROVEEDOR",
                    "COSTO_REAL",
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
                            "Importe interno": b.get("sale_before_tax"),
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
                            "Importe interno": st.column_config.NumberColumn(format="$ %.2f"),
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
                    confirm_project = st.checkbox(
                        "Confirmo que deseo eliminar este proyecto completo.",
                        key=f"confirm_delete_project_{project_id}",
                    )
                    if st.button(
                        "Eliminar proyecto completo",
                        key=f"delete_project_button_{project_id}",
                        disabled=not confirm_project,
                    ):
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
                    "Importe interno": b.get("sale_before_tax"),
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
                    "Importe interno": st.column_config.NumberColumn(format="$ %.2f"),
                },
            )

            budget_map = {
                b["id"]: (
                    f"{b.get('project_code') or ''} — "
                    f"{b.get('project_name') or ''} — "
                    f"{formato_moneda(float(b.get('sale_before_tax') or 0))}"
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
                    bm4.metric("Importe interno", formato_moneda(float(budget.get("sale_before_tax") or 0)))

                    st.caption(
                        f"Indirectos: {float(budget.get('indirect_pct') or 0):.2f}% · "
                        f"Utilidad: {float(budget.get('profit_pct') or 0):.2f}% · "
                        f"Estado: {budget.get('status') or ''}"
                    )

                if items:
                    st.markdown("#### Actividades")
                    item_df = pd.DataFrame([
                        {
                            "Sección": i.get("category") or "",
                            "Concepto": i.get("commercial_title") or i.get("subcategory") or "",
                            "Código": i.get("code") or "",
                            "Descripción": i.get("description") or "",
                            "Unidad": i.get("unit") or "",
                            "Cantidad": i.get("quantity"),
                            "Costo unitario": i.get("unit_cost"),
                            "P.U. interno": i.get("unit_sale"),
                            "Importe interno": i.get("sale_amount"),
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
                            "P.U. interno": st.column_config.NumberColumn(format="$ %.2f"),
                            "Importe interno": st.column_config.NumberColumn(format="$ %.2f"),
                        },
                    )
                else:
                    st.info("Este presupuesto no contiene actividades registradas.")

                with st.expander("Opciones avanzadas"):
                    st.warning(
                        "Eliminar un presupuesto elimina sus actividades e historial vinculado. "
                        "Si es el único presupuesto del proyecto, también puede eliminarse el proyecto."
                    )
                    confirm_budget = st.checkbox(
                        "Confirmo que deseo eliminar este presupuesto.",
                        key=f"confirm_delete_budget_{budget_id}",
                    )
                    if st.button(
                        "Eliminar presupuesto",
                        key=f"delete_budget_button_{budget_id}",
                        disabled=not confirm_budget,
                    ):
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
            "afectan datos persistentes. Por ahora no se utiliza clave de eliminación."
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
                db.clear_all_data()
                st.session_state.pop("generated", None)
                st.success("Todos los datos de la aplicación fueron eliminados.")
                st.rerun()

        with st.container(border=True):
            st.markdown("### Carga Masiva de Catálogo Ancla (CSV)")
            st.caption(
                "Sube un archivo CSV con tus precios históricos estandarizados. "
                "Las columnas requeridas son: **Codigo, Partida, Subpartida, Concepto_Generico, Unidad, Costo_Unitario**."
            )

            uploaded_csv = st.file_uploader(
                "Subir archivo de catálogo CSV",
                type=["csv"],
                key="csv_uploader_seed",
            )

            if uploaded_csv:
                if st.button(
                    "Procesar y Cargar Catálogo",
                    type="primary",
                    use_container_width=True,
                    key="process_seed_catalog_csv",
                ):
                    with st.spinner("Procesando carga masiva..."):
                        try:
                            df_csv = pd.read_csv(uploaded_csv)

                            necesarias = [
                                "Codigo",
                                "Partida",
                                "Subpartida",
                                "Concepto_Generico",
                                "Unidad",
                                "Costo_Unitario",
                            ]
                            faltantes = [col for col in necesarias if col not in df_csv.columns]

                            if faltantes:
                                st.error(
                                    "El archivo CSV no tiene el formato correcto. "
                                    f"Faltan las columnas: {', '.join(faltantes)}"
                                )
                            else:
                                count_nuevos = 0
                                count_omitidos = 0
                                for _, row in df_csv.iterrows():
                                    code = str(row["Codigo"]).strip()
                                    category = str(row["Partida"]).strip()
                                    subcategory = str(row["Subpartida"]).strip()
                                    description = str(row["Concepto_Generico"]).strip()
                                    unit = str(row["Unidad"]).strip().upper()

                                    raw_cost = (
                                        str(row["Costo_Unitario"])
                                        .replace("$", "")
                                        .replace(",", "")
                                        .strip()
                                    )
                                    try:
                                        unit_cost = float(raw_cost)
                                    except (TypeError, ValueError):
                                        count_omitidos += 1
                                        continue

                                    if (
                                        not code
                                        or not category
                                        or not subcategory
                                        or not description
                                        or not unit
                                        or unit_cost <= 0
                                    ):
                                        count_omitidos += 1
                                        continue

                                    new_id = db.create_concept(
                                        code, category, subcategory, description, unit
                                    )

                                    db.add_price(
                                        concept_id=new_id,
                                        unit_cost=unit_cost,
                                        source="BASE_INTERNA",
                                        source_detail="Carga masiva histórica CSV",
                                        status="VALIDADO",
                                        confidence="Alta",
                                    )
                                    count_nuevos += 1

                                st.success(
                                    f"✅ Se cargaron exitosamente {count_nuevos} conceptos ancla. "
                                    "La IA ahora los priorizará."
                                )
                                if count_omitidos:
                                    st.warning(
                                        f"Se omitieron {count_omitidos} filas por tener precios o datos inválidos."
                                    )

                        except Exception as e:
                            st.error(f"Error procesando el archivo CSV: {e}")

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
# ESTADO DE APLICACIÓN
# =========================================================


st.title("Generador de presupuestos")

database_url = get_secret("DATABASE_URL")
try:
    db, db_error = get_database(database_url, DATABASE_CACHE_VERSION)

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
    st.error("No fue posible conectar con la base de datos.")
    st.exception(exc)
    st.stop()

with st.sidebar:
    st.header("Navegación")
    section = st.radio(
        "Sección",
        ["Generar presupuesto", "Crear Excel", "Catálogo e historial"],
        key="main_section",
        label_visibility="collapsed",
    )

    if section == "Generar presupuesto":
        st.divider()
        st.header("Parámetros")

        indirect_pct = st.number_input(
            "Indirectos (%)",
            min_value=0.0,
            max_value=100.0,
            value=10.0,
            step=0.5,
            key="indirect_pct",
        )
        profit_pct = st.number_input(
            "Utilidad (%)",
            min_value=0.0,
            max_value=100.0,
            value=18.0,
            step=0.5,
            key="profit_pct",
        )
        iva_pct = st.number_input(
            "IVA (%) · referencia posterior",
            min_value=0.0,
            max_value=100.0,
            value=16.0,
            step=1.0,
            key="iva_pct",
            disabled=True,
        )
        waste_pct = st.number_input(
            "Desperdicio (%) · referencia sin aplicación",
            min_value=0.0,
            max_value=50.0,
            value=4.0,
            step=0.5,
            key="waste_pct",
            disabled=True,
        )

        with st.expander("Configuración"):
            model_name = st.text_input(
                "Modelo Gemini",
                value="gemini-3.8-flash",
                key="model_name",
            )

        st.divider()
        if st.button("Reiniciar página", use_container_width=True):
            st.session_state.clear()
            st.rerun()

# =========================================================
# BASE INTERNA
# =========================================================


if section == "Crear Excel":
    render_editor_excel_nuevo()
    st.stop()

if section == "Catálogo e historial":
    render_admin_database(db)
    st.stop()


# =========================================================
# CHECKPOINTS DE GENERACIÓN
# =========================================================

def firma_generacion(project_data: dict, params: dict, model_name: str) -> str:
    """Firma estable para saber si un checkpoint corresponde a los mismos datos."""
    payload = {
        "engine_version": DATABASE_CACHE_VERSION + "-compact-prices-v1",
        "project_data": project_data,
        "params": params,
        "model_name": model_name or "",
    }
    return json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def guardar_checkpoint_generacion(
    *,
    stage: int,
    status: str,
    input_signature: str,
    result=None,
    items=None,
    mensaje: str = "",
):
    """Guarda el mayor avance alcanzado sin depender de que la siguiente etapa termine."""
    checkpoint = dict(st.session_state.get("generation_checkpoint") or {})
    checkpoint.update({
        "stage": int(stage),
        "status": status,
        "input_signature": input_signature,
        "mensaje": mensaje,
        "updated_at": ahora_iso(),
    })
    if result is not None:
        checkpoint["result"] = result.model_dump() if hasattr(result, "model_dump") else result
    if items is not None:
        checkpoint["items"] = items
    st.session_state["generation_checkpoint"] = checkpoint


# =========================================================
# BORRADOR LOCAL DEL FORMULARIO
# =========================================================

FORM_DRAFT_STORAGE_KEY = "presupuesto_form_draft_v1"


def _get_local_storage():
    """Obtiene el almacenamiento local del navegador cuando está disponible."""
    if LocalStorage is None:
        return None
    try:
        return LocalStorage()
    except Exception:
        return None


def cargar_borrador_local():
    """Recupera el último borrador guardado en el navegador."""
    storage = _get_local_storage()
    if storage is None:
        return None
    try:
        raw = storage.getItem(
            FORM_DRAFT_STORAGE_KEY,
            key="_load_presupuesto_form_draft",
        )
        if not raw:
            return None
        if isinstance(raw, str):
            return json.loads(raw)
        if isinstance(raw, dict):
            return raw
    except Exception:
        pass
    return None


def guardar_borrador_local(data: dict):
    """Guarda los campos principales del formulario en el navegador."""
    storage = _get_local_storage()
    if storage is None:
        return False
    try:
        storage.setItem(
            FORM_DRAFT_STORAGE_KEY,
            json.dumps(data, ensure_ascii=False),
        )
        return True
    except Exception:
        return False


def limpiar_borrador_local():
    """Elimina el borrador guardado localmente sin tocar otras claves."""
    storage = _get_local_storage()
    if storage is None:
        return
    try:
        storage.setItem(FORM_DRAFT_STORAGE_KEY, "")
    except Exception:
        pass


# =========================================================
# FORMULARIO INICIAL
# =========================================================


DESCRIPTION_EXAMPLE = """Ejemplo:

SEGUNDA PLANTA
Recámara principal de aproximadamente 4.20 x 3.80 m.
- Retiro de piso laminado existente.
- Colocación de piso nuevo.
- Reparación y pintura de muros.

AZOTEA
Área aproximada de 9 x 4 m.
- Limpieza y preparación de superficie.
- Impermeabilización completa.
- Revisión de bajadas pluviales.
"""

DEFAULT_GUIDE_TEXT = """- Considerar protección básica de las áreas de trabajo y zonas de tránsito.
- Incluir limpieza durante los trabajos y limpieza final.
- En pintura, considerar preparación básica, resanes menores, sellador cuando sea necesario y dos manos de pintura.
- En instalaciones y elementos nuevos, considerar suministro, colocación, fijaciones y conexiones, cuando correspondan.
- Mantener materiales y acabados coherentes con el nivel de presupuesto seleccionado.
"""

ADJUSTMENT_EXAMPLE = """Ejemplos:
- Falta considerar limpieza fina al terminar la obra.
- El precio de la pintura me parece muy bajo, revísalo.
- La cantidad de impermeabilización debe ser mayor.
- Cambia la cancelería a aluminio línea pesada con cristal templado.
"""


def volver_a_entrada():
    """
    Regresa al formulario conservando exactamente los datos con los que se
    generó el presupuesto.

    Se utiliza un estado intermedio independiente de los widgets porque
    Streamlit puede retirar del session_state los valores de widgets que dejan
    de renderizarse mientras se muestra la pantalla de resultados.
    """
    generated = st.session_state.get("generated")
    if generated:
        project_data = generated.get("project_data") or {}

        st.session_state["_restore_project_form"] = {
            "client_name": project_data.get("name", ""),
            "project_location": project_data.get("location", ""),
            "project_type": project_data.get(
                "project_type",
                "Remodelación interior general",
            ),
            "budget_level": project_data.get("budget_level", "Medio-alto"),
            "project_description": project_data.get("description", ""),
            "guide_text": (
                project_data.get("guide_text")
                if project_data.get("guide_text") is not None
                else DEFAULT_GUIDE_TEXT
            ),
        }

    st.session_state.pop("generated", None)
    st.rerun()


if "generated" not in st.session_state:
    # Restauración explícita al volver desde un presupuesto ya generado.
    # Debe ocurrir ANTES de crear los widgets del formulario.
    restore_data = st.session_state.pop("_restore_project_form", None)
    if restore_data:
        for field_key, field_value in restore_data.items():
            st.session_state[field_key] = field_value

    # En una recarga completa del navegador, session_state se pierde.
    # Recuperamos el último borrador guardado en localStorage.
    if not st.session_state.get("_draft_checked", False):
        draft = cargar_borrador_local()
        if draft:
            for field_key, field_value in draft.items():
                if field_value is not None:
                    st.session_state[field_key] = field_value
        st.session_state["_draft_checked"] = True

    if "guide_text" not in st.session_state:
        st.session_state["guide_text"] = DEFAULT_GUIDE_TEXT

    # -----------------------------------------------------
    # RECARGAR PRESUPUESTO EXISTENTE
    # -----------------------------------------------------
    uploaded_budget = st.file_uploader(
        "Cargar presupuesto Excel",
        type=["xlsx"],
        key="reload_budget_excel",
    )

    if uploaded_budget is not None:
        if st.button(
            "Cargar y continuar editando",
            type="primary",
            use_container_width=True,
            key="load_existing_budget",
        ):
            fallback_params = {
                "indirect_pct": float(indirect_pct),
                "profit_pct": float(profit_pct),
                "iva_pct": float(iva_pct),
                "waste_pct": float(waste_pct),
            }

            with st.spinner("Leyendo presupuesto..."):
                try:
                    imported = importar_presupuesto_excel(
                        uploaded_budget.getvalue(),
                        fallback_params=fallback_params,
                        file_name=uploaded_budget.name,
                    )

                    st.session_state["generated"] = {
                        "project_id": None,
                        "budget_id": None,
                        "saved": False,
                        "pending_revision": False,
                        "project_code": imported["project_code"],
                        "version": imported["version"],
                        "project_data": imported["project_data"],
                        "params": imported["params"],
                        "result": imported["result"].model_dump(),
                        "items": imported["items"],
                        "financials": imported["financials"],
                        "excel_bytes": imported["excel_bytes"],
                        "revision_history": [],
                        "pending_revision_notes": [],
                        "imported_from_excel": True,
                    }
                    st.rerun()
                except Exception as exc:
                    st.exception(exc)

    st.divider()

    # Los widgets están fuera de st.form para que cada cambio pueda
    # guardarse automáticamente en el navegador y sobrevivir a una recarga.
    f1, f2 = st.columns(2)
    with f1:
        client_name = st.text_input(
            "Nombre del cliente",
            placeholder="Ej. Desarrollos de la Vega",
            key="client_name",
        )
    with f2:
        location = st.text_input(
            "Ubicación",
            placeholder="Ej. Coyoacán, CDMX",
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

    budget_level = st.selectbox(
        "Nivel de presupuesto",
        NIVELES_PRESUPUESTO,
        index=2,
        key="budget_level",
    )

    description = st.text_area(
        "Descripción general de trabajos",
        placeholder=DESCRIPTION_EXAMPLE,
        height=430,
        key="project_description",
    )

    guide_text = st.text_area(
        "Texto guía",
        height=300,
        key="guide_text",
    )

    # Autoguardado local del borrador. No guarda API keys ni resultados de Gemini.
    current_draft = {
        "client_name": client_name,
        "project_location": location,
        "project_type": project_type,
        "budget_level": budget_level,
        "project_description": description,
        "guide_text": guide_text,
    }
    draft_signature = json.dumps(
        current_draft, ensure_ascii=False, sort_keys=True, separators=(",", ":")
    )
    if draft_signature != st.session_state.get("_last_saved_draft_signature"):
        if guardar_borrador_local(current_draft):
            st.session_state["_last_saved_draft_signature"] = draft_signature

    c1, c2 = st.columns([4, 1])
    with c1:
        generate = st.button(
            "Generar presupuesto",
            type="primary",
            use_container_width=True,
        )
    with c2:
        clear_draft = st.button(
            "Borrar borrador",
            use_container_width=True,
        )

    if clear_draft:
        limpiar_borrador_local()
        for field_key in (
            "client_name",
            "project_location",
            "project_description",
            "guide_text",
        ):
            st.session_state.pop(field_key, None)
        st.session_state.pop("_last_saved_draft_signature", None)
        st.rerun()

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
            "name": client_name.strip(),
            "project_type": project_type,
            "budget_level": budget_level,
            "location": location.strip(),
            "dimension_mode": "Integradas en descripción",
            "dimensions_text": "",
            "description": description.strip(),
            "guide_text": guide_text.strip(),
        }

        if st.session_state.get("generation_in_progress", False):
            st.warning("Ya hay una generación en curso. Espera a que termine.")
            st.stop()

        st.session_state["generation_in_progress"] = True
        st.subheader("Generación del presupuesto")
        st.caption("El registro se actualiza en cada etapa y cada 5 segundos mientras Gemini responde.")
        progress_bar = st.progress(0)
        progress_text = st.empty()
        log_placeholder = st.empty()
        st.session_state["generation_log"] = []
        st.session_state.pop("generation_error_trace", None)
        progress_state = {"pct": 0}

        def ui_progress(pct: int, message: str):
            # El porcentaje representa etapas completadas; no retrocede al esperar.
            progress_state["pct"] = max(progress_state["pct"], max(0, min(int(pct), 100)))
            message = limpiar_log_generacion(message)
            entries = st.session_state["generation_log"]
            entries.append(f"[{datetime.now():%H:%M:%S}] {message}")
            progress_bar.progress(progress_state["pct"])
            progress_text.text(f"{progress_state['pct']}% · {message}")
            mostrar_log_generacion(log_placeholder, entries)

        try:
            input_signature = firma_generacion(project_data, params, model_name)
            old_checkpoint = st.session_state.get("generation_checkpoint") or {}
            same_input = old_checkpoint.get("input_signature") == input_signature
            if not same_input:
                guardar_checkpoint_generacion(
                    stage=0,
                    status="iniciando",
                    input_signature=input_signature,
                    mensaje="Preparando una nueva corrida.",
                )
                old_checkpoint = st.session_state["generation_checkpoint"]

            stage = int(old_checkpoint.get("stage") or 0) if same_input else 0
            checkpoint_result = old_checkpoint.get("result")
            checkpoint_items = old_checkpoint.get("items")

            # ETAPA 1 -------------------------------------------------------
            if stage >= 1 and checkpoint_result:
                result = PresupuestoIA.model_validate(checkpoint_result)
                ui_progress(25, "1/3 · Recuperando estructura ya generada")
            else:
                ui_progress(3, "Validando datos y preparando el proyecto")
                ui_progress(4, "1/3 · Interpretando alcance y generando partidas")
                result = generar_presupuesto_ia(
                    api_key=api_key,
                    model_name=model_name,
                    project_data=project_data,
                    params=params,
                    progress_callback=lambda _pct, msg: ui_progress(12, msg),
                )
                guardar_checkpoint_generacion(
                    stage=1,
                    status="completada",
                    input_signature=input_signature,
                    result=result,
                    mensaje="Partidas y alcance generados en una sola lectura.",
                )
                stage = 1

            # ETAPA 2 -------------------------------------------------------
            checkpoint = st.session_state.get("generation_checkpoint") or {}
            if stage >= 2 and checkpoint.get("result"):
                result = PresupuestoIA.model_validate(checkpoint["result"])
                ui_progress(45, "1/3 · Recuperando partidas preparadas")
            else:
                ui_progress(38, "1/3 · Preparando clasificación de partidas")
                # La clasificación y la secuencia se normalizan en Python.
                # No se realiza otra llamada a Gemini para auditar las partidas.
                guardar_checkpoint_generacion(
                    stage=2,
                    status="completada",
                    input_signature=input_signature,
                    result=result,
                    mensaje="Partidas, subpartidas y secuencia auditadas.",
                )
                stage = 2

            # ETAPA 3 -------------------------------------------------------
            checkpoint = st.session_state.get("generation_checkpoint") or {}
            checkpoint_items = checkpoint.get("items")
            if stage >= 3 and checkpoint_items:
                items = checkpoint_items
                ui_progress(78, "2/3 · Recuperando precios ya revisados")
            else:
                ui_progress(52, "2/3 · Buscando precios históricos internos")
                # Dejamos explícito que estamos trabajando en esta etapa antes de
                # entrar a Gemini. Si la etapa 3 falla, las etapas 1 y 2 siguen
                # guardadas y la siguiente corrida comenzará aquí.
                guardar_checkpoint_generacion(
                    stage=2,
                    status="etapa_3_en_curso",
                    input_signature=input_signature,
                    result=result,
                    mensaje="Consultando referencias y valuando precios.",
                )
                items = resolver_items(
                    db, result, project_data, params,
                    api_key=api_key,
                    model_name=model_name,
                    progress_callback=lambda pct, msg: ui_progress(max(52, min(pct, 92)), msg),
                )
                if not items:
                    raise RuntimeError("La IA no generó actividades utilizables.")
                guardar_checkpoint_generacion(
                    stage=3,
                    status="completada",
                    input_signature=input_signature,
                    result=result,
                    items=items,
                    mensaje="Precios valuados y partidas convertidas en items.",
                )
                stage = 3

            # ETAPA 4 -------------------------------------------------------
            ui_progress(93, "3/3 · Calculando importes y preparando los dos Excel")
            items = asignar_codigos_jerarquicos(items)
            financials = calcular_financieros(items, params)
            provisional_code = db.next_project_code(
                project_data["name"], project_data["location"]
            )
            excel_bytes = crear_paquete_excels(
                project_code=provisional_code, project_data=project_data, result=result,
                items=items, params=params, version=1,
            )

            guardar_checkpoint_generacion(
                stage=4,
                status="completada",
                input_signature=input_signature,
                result=result,
                items=items,
                mensaje="Excel preparado.",
            )
            ui_progress(100, "Presupuesto terminado")
            st.session_state["generated"] = {
                "project_id": None, "budget_id": None, "saved": False,
                "pending_revision": False, "project_code": provisional_code, "version": 1,
                "project_data": project_data, "params": params,
                "result": result.model_dump(), "items": items, "financials": financials,
                "excel_bytes": excel_bytes, "revision_history": [], "pending_revision_notes": [],
            }
            st.session_state["generation_in_progress"] = False
            st.session_state.pop("generation_last_error", None)
            st.rerun()
        except Exception as exc:
            st.session_state["generation_in_progress"] = False
            checkpoint = st.session_state.get("generation_checkpoint") or {}
            st.session_state["generation_last_error"] = limpiar_log_generacion(exc)
            st.session_state["generation_error_trace"] = limpiar_log_generacion(traceback.format_exc())
            ui_progress(progress_state["pct"], f"ERROR del programa · {type(exc).__name__}: {exc}")
            st.text(st.session_state["generation_error_trace"])
            stage = int(checkpoint.get("stage") or 0)
            if error_gemini_transitorio(exc):
                st.error(
                    "Gemini sigue rechazando temporalmente la solicitud después de todos los reintentos. "
                    f"Se conservó el avance hasta la etapa {stage}/4. "
                    "Al volver a pulsar Generar presupuesto se reanudará desde el último checkpoint compatible."
                )
            else:
                st.error(
                    "No fue posible completar la generación. "
                    f"Se conservó el avance hasta la etapa {stage}/4. Detalle: {limpiar_log_generacion(exc)}"
                )
        finally:
            st.session_state["generation_in_progress"] = False


# =========================================================
# RESULTADO
# =========================================================


else:
    g = st.session_state["generated"]
    result = PresupuestoIA.model_validate(g["result"])
    items = g["items"]
    financials = g["financials"]
    version = int(g.get("version") or 1)
    saved = bool(g.get("saved"))

    if st.session_state.get("generation_log"):
        with st.expander("Registro de generación", expanded=False):
            mostrar_log_generacion(st.empty(), st.session_state["generation_log"])
            st.download_button(
                "Descargar registro", data="\n".join(st.session_state["generation_log"]),
                file_name="registro_generacion.txt", mime="text/plain",
                key="download_generation_log",
            )

    st.subheader(g["project_code"])
    if saved:
        st.caption(f"Guardado · V{version:02d}")
    elif g.get("project_id"):
        st.caption(f"Cambios sin guardar · próxima versión V{version:02d}")
    elif g.get("imported_from_excel"):
        st.caption("Presupuesto recargado desde Excel · cambios sin guardar")
    else:
        st.caption("Borrador sin guardar")

    p1, p2, p3, p4 = st.columns(4)
    with p1:
        st.text_input(
            "Cliente",
            value=g["project_data"]["name"],
            disabled=True,
            key=f"locked_client_{version}_{saved}",
        )
    with p2:
        st.text_input(
            "Ubicación",
            value=g["project_data"]["location"],
            disabled=True,
            key=f"locked_location_{version}_{saved}",
        )
    with p3:
        st.text_input(
            "Tipo de obra",
            value=g["project_data"]["project_type"],
            disabled=True,
            key=f"locked_type_{version}_{saved}",
        )
    with p4:
        st.text_input(
            "Nivel",
            value=g["project_data"].get("budget_level", "Medio-alto"),
            disabled=True,
            key=f"locked_level_{version}_{saved}",
        )

    st.text_area(
        "Descripción general de trabajos",
        value=g["project_data"]["description"],
        height=220,
        disabled=True,
        key=f"locked_description_{version}_{saved}",
    )
    st.text_area(
        "Texto guía",
        value=g["project_data"]["guide_text"] or "",
        height=160,
        disabled=True,
        key=f"locked_guide_{version}_{saved}",
    )

    # Este módulo termina en el Importe interno. La marca y el IVA se agregan
    # posteriormente, fuera de este presupuesto de negociación con subcontratistas.
    presupuesto_interno = float(financials["sale_before_tax"])

    st.metric(
        "Presupuesto interno",
        formato_moneda(presupuesto_interno),
    )

    with st.expander("Detalle interno"):
        i1, i2, i3 = st.columns(3)
        i1.metric("Costo directo", formato_moneda(financials["direct_cost"]))
        i2.metric("Indirectos", formato_moneda(financials["indirect_cost"]))
        i3.metric("Utilidad", formato_moneda(financials["profit"]))

    df = dataframe_resumen(items)
    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Cant.": st.column_config.NumberColumn(format="%.2f"),
            "Precio Unitario": st.column_config.NumberColumn(format="$ %.2f"),
            "Importe interno": st.column_config.NumberColumn(format="$ %.2f"),
        },
    )

    if result.consideraciones_generales or result.datos_faltantes:
        with st.expander("Consideraciones"):
            for item in result.consideraciones_generales:
                st.write(f"- {item}")
            if result.datos_faltantes:
                st.markdown("**Datos por confirmar**")
                for item in result.datos_faltantes:
                    st.write(f"- {item}")

    file_status = f"V{version:02d}" if g.get("project_id") else "BORRADOR"
    st.download_button(
        "Descargar Excel (interno + cliente)",
        data=g["excel_bytes"],
        file_name=f"{g['project_code']}-{file_status}_Presupuesto.zip",
        mime="application/zip",
        use_container_width=True,
    )
    st.caption(
        "El .zip incluye el Excel interno (01 Presupuesto) y el Excel cliente "
        "(Partidas con subtotal, IVA y total). Cada archivo contiene una sola hoja."
    )

    # -----------------------------------------------------
    # EDITOR DIRECTO + PEGAR DESDE EXCEL
    # -----------------------------------------------------
    st.divider()
    st.subheader("Editor del presupuesto")
    st.caption(
        "Puedes editar cualquier fila directamente. El identificador interno queda oculto, "
        "así que agregar o mover filas ya no desordena las actividades existentes."
    )

    editor_columns = [
        "Área", "Partida", "Subpartida", "Descripción Técnica",
        "Unidad", "Cant.", "Precio Unitario (MXN)",
    ]
    editor_rows = [
        {
            "__code": str(item.get("code") or ""),
            "Área": area_excel_item(item),
            "Partida": item.get("category") or "",
            "Subpartida": item.get("subcategory") or "",
            "Descripción Técnica": item.get("description") or "",
            "Unidad": item.get("unit") or "",
            "Cant.": float(item.get("quantity") or 0.0),
            "Precio Unitario (MXN)": float(item.get("unit_sale") or 0.0),
        }
        for item in items
    ]
    editor_df = pd.DataFrame(editor_rows, columns=["__code"] + editor_columns)
    edited_df = st.data_editor(
        editor_df,
        key=f"excel_editor_{g['project_code']}_{version}",
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_order=editor_columns,
        column_config={
            "__code": None,
            "Área": st.column_config.TextColumn("Área", width="small"),
            "Partida": st.column_config.TextColumn("Partida", width="medium"),
            "Subpartida": st.column_config.TextColumn("Subpartida", width="medium"),
            "Descripción Técnica": st.column_config.TextColumn("Descripción Técnica", width="large"),
            "Unidad": st.column_config.TextColumn("Unidad", width="small"),
            "Cant.": st.column_config.NumberColumn("Cant.", min_value=0.0, step=0.01, format="%.2f"),
            "Precio Unitario (MXN)": st.column_config.NumberColumn(
                "Precio Unitario (MXN)", min_value=0.0, step=0.01, format="$ %.2f"
            ),
        },
    )

    if st.button(
        "Aplicar cambios del editor",
        type="primary",
        use_container_width=True,
        key=f"apply_excel_editor_{g['project_code']}_{version}",
    ):
        try:
            revised_items = preparar_items_desde_editor_excel(
                edited_df, items, g["params"], g["project_data"]
            )
            revised_items = asignar_codigos_jerarquicos(revised_items)
            revised_result = PresupuestoIA(
                nombre_proyecto=g["project_data"]["name"],
                actividad_principal=g["project_data"]["project_type"],
                alcance_resumido="Presupuesto editado manualmente desde el editor.",
                consideraciones_generales=["Cambios aplicados mediante editor manual."],
                datos_faltantes=[],
                actividades=[item_a_actividad(item) for item in revised_items],
            )
            revised_financials = calcular_financieros(revised_items, g["params"])
            revised_excel = crear_paquete_excels(
                project_code=g["project_code"],
                project_data=g["project_data"],
                result=revised_result,
                items=revised_items,
                params=g["params"],
                version=version,
            )
            g.update({
                "saved": False,
                "result": revised_result.model_dump(),
                "items": revised_items,
                "financials": revised_financials,
                "excel_bytes": revised_excel,
            })
            st.session_state["generated"] = g
            st.rerun()
        except Exception as exc:
            st.error(f"No fue posible aplicar los cambios: {exc}")

    with st.expander("Pegar actividades desde Excel", expanded=False):
        st.caption(
            "Copia directamente de Excel y pega aquí. Se detectan encabezados aunque estén en otro orden. "
            "Con encabezados usa: Área · Partida · Subpartida · Descripción Técnica · Unidad · Cant. · Precio Unitario (MXN)."
        )
        paste_text = st.text_area(
            "Pega aquí las filas copiadas",
            height=180,
            placeholder=(
                "Área\tPartida\tSubpartida\tDescripción Técnica\tUnidad\tCant.\tPrecio Unitario (MXN)\n"
                "Cocina\tACABADOS Y RECUBRIMIENTOS\tMuros\tSuministro y aplicación de pintura...\tM2\t22.17\t185.00"
            ),
            key=f"paste_activities_{g['project_code']}_{version}",
        )
        paste_mode = st.selectbox(
            "Dónde colocarlas",
            [
                "Automático (misma Partida/Subpartida)",
                "Al final",
                "Antes de una actividad",
                "Después de una actividad",
            ],
            key=f"paste_mode_{g['project_code']}_{version}",
        )
        if paste_mode in {"Antes de una actividad", "Después de una actividad"} and items:
            labels = [f"{i+1}. {item.get('subcategory') or item.get('description','')[:70]} [{item.get('code')}]" for i,item in enumerate(items)]
            selected = st.selectbox(
                "Actividad de referencia",
                options=range(len(items)),
                format_func=lambda i: labels[i],
                key=f"paste_anchor_{g['project_code']}_{version}",
            )
            anchor_code = str(items[selected].get("code") or "")
        else:
            anchor_code = ""

        if st.button(
            "Pegar e insertar actividades",
            type="secondary",
            use_container_width=True,
            key=f"insert_pasted_{g['project_code']}_{version}",
        ):
            try:
                pasted_df = parsear_actividades_pegadas(paste_text)
                mode = paste_mode
                revised_items = insertar_actividades_pegadas(
                    items,
                    pasted_df,
                    g["params"],
                    g["project_data"],
                    mode=mode,
                    anchor_code=anchor_code,
                )
                revised_items = asignar_codigos_jerarquicos(revised_items)
                revised_result = PresupuestoIA(
                    nombre_proyecto=g["project_data"]["name"],
                    actividad_principal=g["project_data"]["project_type"],
                    alcance_resumido="Presupuesto actualizado con actividades pegadas desde Excel.",
                    consideraciones_generales=["Actividades insertadas desde un bloque copiado desde Excel."],
                    datos_faltantes=[],
                    actividades=[item_a_actividad(item) for item in revised_items],
                )
                revised_financials = calcular_financieros(revised_items, g["params"])
                revised_excel = crear_paquete_excels(
                    project_code=g["project_code"],
                    project_data=g["project_data"],
                    result=revised_result,
                    items=revised_items,
                    params=g["params"],
                    version=version,
                )
                g.update({
                    "saved": False,
                    "result": revised_result.model_dump(),
                    "items": revised_items,
                    "financials": revised_financials,
                    "excel_bytes": revised_excel,
                })
                st.session_state["generated"] = g
                st.rerun()
            except Exception as exc:
                st.error(f"No fue posible insertar el bloque pegado: {exc}")

    # -----------------------------------------------------
    # Ajuste sencillo con IA
    # -----------------------------------------------------
    st.divider()
    st.subheader("Ajustar con IA")

    adjustment_request = st.text_area(
        "¿Qué quieres revisar, agregar o cambiar?",
        placeholder=ADJUSTMENT_EXAMPLE,
        height=180,
        key=f"adjustment_request_{version}_{saved}",
    )

    if st.button(
        "Aplicar ajuste",
        type="primary",
        use_container_width=True,
        key=f"apply_adjustment_{version}_{saved}",
    ):
        if not adjustment_request.strip():
            st.error("Escriba el cambio que desea realizar.")
        else:
            api_key = get_api_key_runtime()
            if not api_key:
                st.error("Falta GEMINI_API_KEY en Streamlit Secrets.")
            else:
                with st.spinner("Aplicando ajuste..."):
                    try:
                        revision = revisar_presupuesto_ia(
                            api_key=api_key,
                            model_name=model_name,
                            project_data=g["project_data"],
                            params=g["params"],
                            current_result=result,
                            current_items=items,
                            revision_request=adjustment_request.strip(),
                        )

                        revised_result, revised_items, change_log = aplicar_revision_estructural(
                            db=db,
                            current_result=result,
                            current_items=items,
                            revision=revision,
                            project_data=g["project_data"],
                            params=g["params"],
                            api_key=api_key,
                            model_name=model_name,
                        )
                        revised_items = recalcular_areas_items(
                            g["project_data"],
                            revised_items,
                        )
                        revised_items = asignar_codigos_jerarquicos(revised_items)
                        revised_financials = calcular_financieros(
                            revised_items,
                            g["params"],
                        )

                        # Si el proyecto ya existe en la base, el ajuste queda como
                        # borrador de la siguiente versión hasta que el usuario lo guarde.
                        if g.get("project_id"):
                            target_version = version if g.get("pending_revision") else version + 1
                            pending_revision = True
                        else:
                            target_version = 1
                            pending_revision = False

                        excel_bytes = crear_paquete_excels(
                            project_code=g["project_code"],
                            project_data=g["project_data"],
                            result=revised_result,
                            items=revised_items,
                            params=g["params"],
                            version=target_version,
                        )

                        history = list(g.get("revision_history") or [])
                        history.append(
                            {
                                "request": adjustment_request.strip(),
                                "summary": revision.resumen_revision,
                                "changes": change_log,
                            }
                        )
                        pending_notes = list(g.get("pending_revision_notes") or [])
                        if g.get("project_id"):
                            pending_notes.append(adjustment_request.strip())

                        g.update(
                            {
                                "saved": False,
                                "pending_revision": pending_revision,
                                "version": target_version,
                                "result": revised_result.model_dump(),
                                "items": revised_items,
                                "financials": revised_financials,
                                "excel_bytes": excel_bytes,
                                "revision_history": history,
                                "pending_revision_notes": pending_notes,
                            }
                        )
                        st.session_state["generated"] = g
                        st.rerun()
                    except Exception as exc:
                        st.exception(exc)

    history = g.get("revision_history") or []
    if history:
        with st.expander("Ajustes realizados"):
            for i, rev in enumerate(history, start=1):
                st.markdown(f"**Ajuste {i}: {rev['summary']}**")
                st.caption(rev["request"])
                for change in rev.get("changes") or []:
                    st.write(f"- {change}")

    # -----------------------------------------------------
    # Acciones del proyecto
    # -----------------------------------------------------
    st.divider()
    a1, a2, a3 = st.columns(3)

    with a1:
        save_label = "Guardado en base" if saved else (
            "Guardar nueva versión" if g.get("project_id") else "Guardar en base"
        )
        if st.button(
            save_label,
            type="primary" if not saved else "secondary",
            use_container_width=True,
            disabled=saved,
            key=f"save_project_{version}_{saved}",
        ):
            try:
                if not g.get("project_id"):
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
                    real_version = 1
                else:
                    budget_id, real_version = db.save_revision(
                        project_id=g["project_id"],
                        parent_budget_id=g["budget_id"],
                        result=result,
                        items=items,
                        params=g["params"],
                        financials=financials,
                        revision_instruction=(
                            "\n\n".join(g.get("pending_revision_notes") or [])
                            or (history[-1]["request"] if history else "Ajuste del presupuesto")
                        ),
                    )
                    project_id = g["project_id"]
                    real_code = g["project_code"]

                excel_bytes = crear_paquete_excels(
                    project_code=real_code,
                    project_data=g["project_data"],
                    result=result,
                    items=items,
                    params=g["params"],
                    version=real_version,
                )

                clean_items = []
                for saved_item in items:
                    cleaned = dict(saved_item)
                    cleaned.pop("record_new_price", None)
                    clean_items.append(cleaned)

                g.update(
                    {
                        "project_id": project_id,
                        "budget_id": budget_id,
                        "project_code": real_code,
                        "version": real_version,
                        "saved": True,
                        "pending_revision": False,
                        "items": clean_items,
                        "excel_bytes": excel_bytes,
                        "pending_revision_notes": [],
                    }
                )
                st.session_state["generated"] = g
                st.rerun()
            except Exception as exc:
                st.exception(exc)

    with a2:
        if st.button(
            "Editar entrada inicial",
            use_container_width=True,
            key=f"edit_initial_{version}_{saved}",
        ):
            volver_a_entrada()

    with a3:
        if st.button(
            "Eliminar proyecto",
            use_container_width=True,
            key=f"delete_project_result_{version}_{saved}",
        ):
            try:
                if g.get("project_id"):
                    db.delete_project(g["project_id"])
                st.session_state.pop("generated", None)
                st.rerun()
            except Exception as exc:
                st.exception(exc)
