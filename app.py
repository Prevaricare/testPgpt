import os
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st
from google import genai
from google.genai import types
from pydantic import BaseModel, Field
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================

st.set_page_config(
    page_title="Presupuestador IA",
    page_icon=None,
    layout="wide",
)

st.title("Sistema de Presupuestación Asistida")
st.caption("Presupuestos preliminares para remodelación e interiorismo.")



# =========================================================
# MODELOS DE RESPUESTA ESTRUCTURADA DE GEMINI
# =========================================================

class Partida(BaseModel):
    clave: str = Field(description="Clave corta única, por ejemplo DEM-01")
    categoria: str = Field(description="Categoría o capítulo del presupuesto")
    concepto: str = Field(description="Descripción clara y breve del trabajo")
    unidad: str = Field(description="Unidad: m2, m3, ml, pza, lote, etc.")
    cantidad: float = Field(
        ge=0,
        description="Cantidad preliminar calculada o estimada"
    )
    precio_unitario_estimado: float = Field(
        ge=0,
        description="Precio unitario preliminar estimado en MXN"
    )
    criterio_cuantificacion: str = Field(
        description="Método o fórmula utilizada para obtener la cantidad"
    )
    fundamento_inclusion: str = Field(
        description="Razón técnica breve por la que esta partida forma parte del alcance"
    )
    datos_utilizados: str = Field(
        description="Datos del usuario o supuestos específicos empleados para cuantificar"
    )
    nivel_confianza: str = Field(
        description="Alta, Media o Baja según la calidad de la información disponible"
    )
    requiere_cotizacion: bool = Field(
        description="True si el precio depende fuertemente de marca/modelo/proveedor"
    )
    observaciones: str = Field(
        description="Supuestos, restricciones o advertencias importantes"
    )


class PresupuestoIA(BaseModel):
    nombre_proyecto: str
    alcance_resumido: str
    supuestos_generales: list[str]
    datos_faltantes: list[str]
    partidas: list[Partida]


# =========================================================
# FUNCIONES
# =========================================================

def get_api_key():
    """Busca la clave primero en Streamlit Secrets y luego en variables de entorno."""
    try:
        if "GEMINI_API_KEY" in st.secrets:
            return st.secrets["GEMINI_API_KEY"]
    except Exception:
        pass

    return os.getenv("GEMINI_API_KEY")


def analizar_con_gemini(
    api_key: str,
    model_name: str,
    tipo_obra: str,
    nombre_proyecto: str,
    ubicacion: str,
    largo: float,
    ancho: float,
    altura: float,
    descripcion: str,
    desperdicio: float,
):
    client = genai.Client(api_key=api_key)

    prompt = f"""
Eres un asistente técnico de presupuestación de obras de remodelación e
interiorismo en México.

Tu tarea es generar un PRESUPUESTO PRELIMINAR para ejercicio académico.

DATOS DEL PROYECTO
Nombre: {nombre_proyecto}
Tipo de obra: {tipo_obra}
Ubicación: {ubicacion if ubicacion else "No indicada"}
Dimensiones generales proporcionadas:
- Largo: {largo:.3f} m
- Ancho: {ancho:.3f} m
- Altura: {altura:.3f} m
Desperdicio de referencia: {desperdicio:.1f} %

DESCRIPCIÓN DEL USUARIO
{descripcion}

REGLAS IMPORTANTES
1. Desglosa el alcance en partidas razonables para una remodelación pequeña.
2. Usa unidades habituales: m2, m3, ml, pza, lote.
3. Calcula cantidades solo cuando los datos lo permitan.
4. Si tienes que asumir algo, explícalo en observaciones.
5. Para pisos puedes usar largo x ancho cuando corresponda.
6. Para muros puedes usar perímetro x altura o altura de recubrimiento cuando
   sea razonable, pero debes explicar el criterio.
7. No inventes puertas, ventanas o cantidades de mobiliario que no hayan sido
   mencionadas. Si hacen falta, repórtalas como datos faltantes.
8. Los precios unitarios son únicamente una ESTIMACIÓN académica en MXN.
   No los presentes como cotizaciones reales ni como precios actuales
   garantizados.
9. Marca requiere_cotizacion=True para muebles, luminarias, cancelería,
   electrodomésticos, accesorios de marca o conceptos muy variables.
10. Evita duplicar conceptos.
11. Incluye, cuando aplique, preliminares, demoliciones/retiros, albañilería,
    acabados, instalaciones y limpieza.
12. No agregues IVA, indirectos ni utilidad como partidas: eso lo calculará
    Python después.
13. La clave debe ser corta y ordenada, como PRE-01, DEM-01, ACA-01.
14. Si la descripción es insuficiente, genera solo lo justificable y enumera
    claramente los datos faltantes.
15. Para cada partida incluye un fundamento técnico breve de por qué debe existir.
16. Indica explícitamente qué datos utilizaste para cuantificar cada concepto.
17. Clasifica el nivel de confianza de cada cantidad como Alta, Media o Baja.
18. Explica los criterios de cuantificación de forma verificable, usando fórmulas,
    relaciones geométricas o supuestos claros cuando corresponda.
19. No muestres razonamiento interno ni cadenas de pensamiento. Entrega únicamente
    una justificación técnica resumida, suficiente para que otra persona pueda revisar
    el criterio aplicado.
20. Si un precio es especialmente incierto, indícalo en observaciones y marca que
    requiere cotización.
"""

    # Intentamos primero el modelo seleccionado. Si Google devuelve un error
    # de modelo no disponible (404), probamos alternativas estables.
    modelos_a_probar = []
    for modelo in [
        model_name,
        "gemini-3.6-flash",
        "gemini-3.5-flash",
        "gemini-3.5-flash-lite",
    ]:
        if modelo and modelo not in modelos_a_probar:
            modelos_a_probar.append(modelo)

    ultimo_error = None

    for modelo in modelos_a_probar:
        try:
            response = client.models.generate_content(
                model=modelo,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=PresupuestoIA,
                ),
            )

            if not response.text:
                raise RuntimeError(
                    f"Gemini ({modelo}) devolvió una respuesta vacía."
                )

            resultado = PresupuestoIA.model_validate_json(response.text)
            return resultado

        except Exception as e:
            ultimo_error = e
            mensaje = str(e).lower()

            # Solo hacemos fallback automático cuando parece un problema
            # de disponibilidad/nombre del modelo. Otros errores (API key,
            # cuota, permisos, esquema, etc.) se muestran directamente.
            es_error_modelo = (
                "404" in mensaje
                or "not_found" in mensaje
                or "not found" in mensaje
                or "model" in mensaje and (
                    "no longer available" in mensaje
                    or "not available" in mensaje
                    or "not supported" in mensaje
                )
            )

            if not es_error_modelo:
                raise

    raise RuntimeError(
        "No se pudo usar ninguno de los modelos Gemini disponibles. "
        f"Último error: {ultimo_error}"
    )


def preparar_dataframe(resultado: PresupuestoIA) -> pd.DataFrame:
    filas = []

    for p in resultado.partidas:
        filas.append(
            {
                "Clave": p.clave,
                "Categoría": p.categoria,
                "Concepto": p.concepto,
                "Unidad": p.unidad,
                "Cantidad": round(float(p.cantidad), 3),
                "P.U. estimado": round(float(p.precio_unitario_estimado), 2),
                "Importe": round(
                    float(p.cantidad) * float(p.precio_unitario_estimado), 2
                ),
                "Requiere cotización": "Sí" if p.requiere_cotizacion else "No",
                "Criterio de cuantificación": p.criterio_cuantificacion,
                "Fundamento de inclusión": p.fundamento_inclusion,
                "Datos utilizados": p.datos_utilizados,
                "Nivel de confianza": p.nivel_confianza,
                "Observaciones": p.observaciones,
            }
        )

    return pd.DataFrame(filas)


def recalcular_importes(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    for col in ["Cantidad", "P.U. estimado"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["Importe"] = (df["Cantidad"] * df["P.U. estimado"]).round(2)
    return df


def crear_excel(
    df: pd.DataFrame,
    resultado: PresupuestoIA,
    datos_proyecto: dict,
    indirectos_pct: float,
    utilidad_pct: float,
    iva_pct: float,
) -> bytes:
    wb = Workbook()

    # Colores
    fill_header = PatternFill("solid", fgColor="1F4E78")
    fill_subheader = PatternFill("solid", fgColor="D9EAF7")
    fill_total = PatternFill("solid", fgColor="E2F0D9")
    white_font = Font(color="FFFFFF", bold=True)
    bold_font = Font(bold=True)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # -----------------------------------------------------
    # Hoja 1: Presupuesto
    # -----------------------------------------------------
    ws = wb.active
    ws.title = "Presupuesto"

    ws["A1"] = resultado.nombre_proyecto
    ws["A1"].font = Font(size=16, bold=True)
    ws.merge_cells("A1:M1")

    ws["A2"] = "Presupuesto preliminar generado con asistencia de IA"
    ws.merge_cells("A2:M2")

    encabezados = list(df.columns)
    fila_header = 4

    for c, valor in enumerate(encabezados, start=1):
        celda = ws.cell(row=fila_header, column=c, value=valor)
        celda.fill = fill_header
        celda.font = white_font
        celda.alignment = Alignment(horizontal="center", vertical="center")
        celda.border = border

    fila_inicio = fila_header + 1

    for r, (_, row) in enumerate(df.iterrows(), start=fila_inicio):
        for c, col in enumerate(encabezados, start=1):
            valor = row[col]
            celda = ws.cell(row=r, column=c, value=valor)
            celda.border = border
            celda.alignment = Alignment(vertical="top", wrap_text=True)

        ws.cell(r, 5).number_format = '0.000'
        ws.cell(r, 6).number_format = '$#,##0.00'
        ws.cell(r, 7).number_format = '$#,##0.00'

    fila_fin = fila_inicio + len(df) - 1
    fila_resumen = fila_fin + 3

    subtotal = float(df["Importe"].sum())
    indirectos = subtotal * indirectos_pct / 100
    utilidad = (subtotal + indirectos) * utilidad_pct / 100
    antes_iva = subtotal + indirectos + utilidad
    iva = antes_iva * iva_pct / 100
    total = antes_iva + iva

    resumen = [
        ("Costo directo", subtotal),
        (f"Indirectos ({indirectos_pct:.2f}%)", indirectos),
        (f"Utilidad ({utilidad_pct:.2f}%)", utilidad),
        ("Subtotal antes de IVA", antes_iva),
        (f"IVA ({iva_pct:.2f}%)", iva),
        ("TOTAL", total),
    ]

    for i, (texto, valor) in enumerate(resumen, start=fila_resumen):
        ws.cell(i, 6, texto).font = bold_font
        ws.cell(i, 7, valor).number_format = '$#,##0.00'
        if texto == "TOTAL":
            ws.cell(i, 6).fill = fill_total
            ws.cell(i, 7).fill = fill_total
            ws.cell(i, 7).font = bold_font

    widths = {
        "A": 12,
        "B": 20,
        "C": 42,
        "D": 10,
        "E": 12,
        "F": 16,
        "G": 16,
        "H": 18,
        "I": 34,
        "J": 38,
        "K": 38,
        "L": 18,
        "M": 42,
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:M{fila_fin}"

    # -----------------------------------------------------
    # Hoja 2: Resumen
    # -----------------------------------------------------
    wr = wb.create_sheet("Resumen")

    wr["A1"] = "Resumen del presupuesto"
    wr["A1"].font = Font(size=16, bold=True)

    resumen_categoria = (
        df.groupby("Categoría", as_index=False)["Importe"]
        .sum()
        .sort_values("Importe", ascending=False)
    )

    wr.append([])
    wr.append(["Categoría", "Importe"])

    for cell in wr[3]:
        cell.fill = fill_header
        cell.font = white_font

    for _, row in resumen_categoria.iterrows():
        wr.append([row["Categoría"], float(row["Importe"])])

    for row in wr.iter_rows(min_row=4, min_col=2, max_col=2):
        row[0].number_format = '$#,##0.00'

    wr.column_dimensions["A"].width = 35
    wr.column_dimensions["B"].width = 18

    # -----------------------------------------------------
    # Hoja 3: Datos_Proyecto
    # -----------------------------------------------------
    wd = wb.create_sheet("Datos_Proyecto")

    wd["A1"] = "Datos del proyecto"
    wd["A1"].font = Font(size=16, bold=True)

    fila = 3
    for clave, valor in datos_proyecto.items():
        wd.cell(fila, 1, clave).font = bold_font
        wd.cell(fila, 2, valor)
        fila += 1

    fila += 1
    wd.cell(fila, 1, "Alcance resumido").font = bold_font
    wd.cell(fila, 2, resultado.alcance_resumido)

    wd.column_dimensions["A"].width = 28
    wd.column_dimensions["B"].width = 80

    # -----------------------------------------------------
    # Hoja 4: Revision_IA
    # -----------------------------------------------------
    wn = wb.create_sheet("Revision_IA")

    wn["A1"] = "Revisión y advertencias"
    wn["A1"].font = Font(size=16, bold=True)

    wn["A3"] = "Supuestos generales"
    wn["A3"].fill = fill_subheader
    wn["A3"].font = bold_font

    fila = 4
    for item in resultado.supuestos_generales:
        wn.cell(fila, 1, "• " + item)
        fila += 1

    fila += 1
    wn.cell(fila, 1, "Datos faltantes")
    wn.cell(fila, 1).fill = fill_subheader
    wn.cell(fila, 1).font = bold_font
    fila += 1

    if resultado.datos_faltantes:
        for item in resultado.datos_faltantes:
            wn.cell(fila, 1, "• " + item)
            fila += 1
    else:
        wn.cell(fila, 1, "No se identificaron datos faltantes importantes.")

    fila += 2
    wn.cell(
        fila, 1,
        "NOTA: Los precios y cantidades son preliminares y deben validarse "
        "con planos, levantamiento, catálogo de conceptos y cotizaciones."
    )
    wn.cell(fila, 1).font = Font(italic=True)

    wn.column_dimensions["A"].width = 110

    # Metadato de generación
    wn.cell(fila + 2, 1, f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:
    st.header("Parámetros")

    model_name = st.text_input(
        "Modelo",
        value="gemini-3.6-flash",
        help="Modelo principal. La app utiliza alternativas si no está disponible."
    )

    indirectos_pct = st.number_input(
        "Indirectos (%)",
        min_value=0.0,
        max_value=100.0,
        value=10.0,
        step=1.0,
    )

    utilidad_pct = st.number_input(
        "Utilidad (%)",
        min_value=0.0,
        max_value=100.0,
        value=10.0,
        step=1.0,
    )

    iva_pct = st.number_input(
        "IVA (%)",
        min_value=0.0,
        max_value=100.0,
        value=16.0,
        step=1.0,
    )

    desperdicio = st.number_input(
        "Desperdicio referencia (%)",
        min_value=0.0,
        max_value=50.0,
        value=10.0,
        step=1.0,
    )


# =========================================================
# FORMULARIO
# =========================================================

with st.form("form_proyecto"):
    col1, col2 = st.columns(2)

    with col1:
        nombre_proyecto = st.text_input(
            "Nombre del proyecto",
            value="Remodelación de baño"
        )

        tipo_obra = st.selectbox(
            "Tipo de obra",
            [
                "Baño",
                "Cocina",
                "Recámara",
                "Sala / comedor",
                "Local comercial",
                "Oficina",
                "Remodelación interior general",
                "Otro",
            ],
        )

        ubicacion = st.text_input(
            "Ubicación",
            placeholder="Ciudad o zona de referencia"
        )

    with col2:
        c1, c2, c3 = st.columns(3)

        largo = c1.number_input(
            "Largo (m)",
            min_value=0.0,
            value=2.40,
            step=0.10
        )

        ancho = c2.number_input(
            "Ancho (m)",
            min_value=0.0,
            value=1.80,
            step=0.10
        )

        altura = c3.number_input(
            "Altura (m)",
            min_value=0.0,
            value=2.50,
            step=0.10
        )

    descripcion = st.text_area(
        "Descripción de trabajos",
        height=180,
        placeholder="Trabajos, materiales, elementos a retirar o instalar y restricciones conocidas.",
    )

    generar = st.form_submit_button(
        "Generar presupuesto",
        type="primary",
        use_container_width=True,
    )


# =========================================================
# GENERACIÓN
# =========================================================

if generar:
    api_key = get_api_key()

    if not api_key:
        st.error(
            "No encontré GEMINI_API_KEY. Agrégala en los Secrets de Streamlit "
            "o como variable de entorno."
        )
        st.stop()

    if not descripcion.strip():
        st.error("Escribe una descripción de los trabajos.")
        st.stop()

    if largo <= 0 or ancho <= 0 or altura <= 0:
        st.error("Largo, ancho y altura deben ser mayores que cero.")
        st.stop()

    with st.spinner("Procesando información..."):
        try:
            resultado = analizar_con_gemini(
                api_key=api_key,
                model_name=model_name,
                tipo_obra=tipo_obra,
                nombre_proyecto=nombre_proyecto,
                ubicacion=ubicacion,
                largo=largo,
                ancho=ancho,
                altura=altura,
                descripcion=descripcion,
                desperdicio=desperdicio,
            )

            st.session_state["resultado"] = resultado
            st.session_state["df_presupuesto"] = preparar_dataframe(resultado)
            st.session_state["datos_proyecto"] = {
                "Nombre": nombre_proyecto,
                "Tipo de obra": tipo_obra,
                "Ubicación": ubicacion,
                "Largo (m)": largo,
                "Ancho (m)": ancho,
                "Altura (m)": altura,
                "Descripción": descripcion,
                "Desperdicio referencia (%)": desperdicio,
            }

        except Exception as e:
            st.exception(e)
            st.stop()


# =========================================================
# RESULTADOS
# =========================================================

if "resultado" in st.session_state:
    resultado = st.session_state["resultado"]
    df_original = st.session_state["df_presupuesto"]

    st.divider()
    st.subheader("1. Alcance interpretado")
    st.write(resultado.alcance_resumido)

    c1, c2 = st.columns(2)

    with c1:
        st.markdown("**Supuestos**")
        if resultado.supuestos_generales:
            for item in resultado.supuestos_generales:
                st.write("•", item)
        else:
            st.write("Sin supuestos relevantes.")

    with c2:
        st.markdown("**Datos pendientes de confirmar**")
        if resultado.datos_faltantes:
            for item in resultado.datos_faltantes:
                st.write("•", item)
        else:
            st.write("No se detectaron datos faltantes importantes.")

    st.subheader("2. Presupuesto preliminar")

    st.caption("Cantidades y precios unitarios son editables. Los importes se recalculan automáticamente.")

    columnas_editables = [
        "Clave",
        "Categoría",
        "Concepto",
        "Unidad",
        "Cantidad",
        "P.U. estimado",
        "Requiere cotización",
        "Criterio de cuantificación",
        "Fundamento de inclusión",
        "Datos utilizados",
        "Nivel de confianza",
        "Observaciones",
    ]

    df_editor = st.data_editor(
        df_original[columnas_editables],
        use_container_width=True,
        num_rows="dynamic",
        hide_index=True,
        column_config={
            "Cantidad": st.column_config.NumberColumn(
                "Cantidad",
                min_value=0.0,
                format="%.3f",
            ),
            "P.U. estimado": st.column_config.NumberColumn(
                "P.U. estimado",
                min_value=0.0,
                format="$ %.2f",
            ),
        },
        key="editor_presupuesto",
    )

    df_final = recalcular_importes(df_editor)

    st.subheader("3. Criterios técnicos")

    criterios_df = df_final[
        [
            "Clave",
            "Concepto",
            "Criterio de cuantificación",
            "Fundamento de inclusión",
            "Datos utilizados",
            "Nivel de confianza",
            "Observaciones",
        ]
    ].copy()

    st.dataframe(
        criterios_df,
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("4. Resumen económico")

    costo_directo = float(df_final["Importe"].sum())
    indirectos = costo_directo * indirectos_pct / 100
    utilidad = (costo_directo + indirectos) * utilidad_pct / 100
    subtotal_antes_iva = costo_directo + indirectos + utilidad
    iva = subtotal_antes_iva * iva_pct / 100
    total = subtotal_antes_iva + iva

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Costo directo", f"${costo_directo:,.2f}")
    m2.metric("Indirectos", f"${indirectos:,.2f}")
    m3.metric("IVA", f"${iva:,.2f}")
    m4.metric("TOTAL", f"${total:,.2f}")

    with st.expander("Tabla consolidada"):
        st.dataframe(
            df_final[
                [
                    "Clave",
                    "Categoría",
                    "Concepto",
                    "Unidad",
                    "Cantidad",
                    "P.U. estimado",
                    "Importe",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )

    excel_bytes = crear_excel(
        df=df_final,
        resultado=resultado,
        datos_proyecto=st.session_state["datos_proyecto"],
        indirectos_pct=indirectos_pct,
        utilidad_pct=utilidad_pct,
        iva_pct=iva_pct,
    )

    st.subheader("5. Exportación")

    csv_presupuesto = df_final.to_csv(index=False).encode("utf-8-sig")
    csv_criterios = criterios_df.to_csv(index=False).encode("utf-8-sig")

    resumen_categoria_export = (
        df_final.groupby("Categoría", as_index=False)["Importe"]
        .sum()
        .sort_values("Importe", ascending=False)
    )
    csv_resumen = resumen_categoria_export.to_csv(index=False).encode("utf-8-sig")

    e1, e2, e3 = st.columns(3)

    with e1:
        st.download_button(
            "CSV",
            data=csv_presupuesto,
            file_name="presupuesto_preliminar.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with e2:
        st.download_button(
            "Descargar criterios CSV",
            data=csv_criterios,
            file_name="criterios_tecnicos.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with e3:
        st.download_button(
            "Descargar resumen CSV",
            data=csv_resumen,
            file_name="resumen_categorias.csv",
            mime="text/csv",
            use_container_width=True,
        )

    st.download_button(
        "Descargar Excel",
        data=excel_bytes,
        file_name="presupuesto_preliminar.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True,
        type="primary",
    )
