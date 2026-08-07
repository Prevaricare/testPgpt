# Presupuestador IA - Streamlit

## Archivos
- `app.py`
- `requirements.txt`

## Secret requerido

En Streamlit Community Cloud:

1. Abre tu aplicación.
2. Ve a Settings / Secrets.
3. Agrega:

```toml
GEMINI_API_KEY = "TU_CLAVE_AQUI"
```

No subas la API key directamente a GitHub.

## Ejecución local

Instala dependencias:

```bash
pip install -r requirements.txt
```

Configura la variable `GEMINI_API_KEY` o crea `.streamlit/secrets.toml`.

Después:

```bash
streamlit run app.py
```

## Qué hace esta versión

1. Recibe dimensiones y descripción.
2. Gemini genera un presupuesto preliminar estructurado.
3. Muestra supuestos y datos faltantes.
4. Permite editar cantidades y P.U.
5. Calcula costo directo, indirectos, utilidad e IVA.
6. Exporta un archivo Excel con:
   - Presupuesto
   - Resumen
   - Datos_Proyecto
   - Revision_IA

Los precios producidos por IA son únicamente preliminares y deben sustituirse o
validarse con un catálogo de precios/cotizaciones para uso profesional.
