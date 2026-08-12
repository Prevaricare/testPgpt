# Sistema de Presupuestación Asistida
## Aplicación interna para presupuestos preliminares de remodelación e interiorismo

Esta aplicación fue diseñada para apoyar la generación inicial de presupuestos de una empresa de interiorismo y remodelación que trabaja principalmente mediante subcontratación.

El objetivo no es sustituir la revisión profesional ni convertirse en un sistema tradicional de análisis de precios unitarios. Su función es transformar una descripción general de obra en una propuesta inicial organizada, consistente y reutilizable, consultar antecedentes de costos de la empresa, generar una estimación económica y producir archivos que puedan continuar revisándose en Excel.

La aplicación utiliza:

- Streamlit para la interfaz.
- Google Gemini para interpretar el alcance y proponer actividades.
- Python para todos los cálculos económicos.
- PostgreSQL/Supabase para la base de datos persistente en producción.
- SQLite únicamente como respaldo para desarrollo local.
- `streamlit-authenticator` para controlar el acceso.
- OpenPyXL para generar el archivo Excel.
- Pandas para manipular y exportar tablas.

---

# 1. Objetivo general

El flujo de trabajo está pensado para una empresa que:

- recibe descripciones generales de trabajos de remodelación e interiorismo;
- necesita generar presupuestos con rapidez;
- subcontrata la mayor parte de las actividades;
- no dispone todavía de una base histórica extensa;
- requiere construir gradualmente su propio catálogo de conceptos y costos;
- necesita conservar información de proyectos y presupuestos anteriores;
- realiza la revisión y corrección final directamente en Excel.

La aplicación genera una primera propuesta. El Excel descargado continúa siendo el documento de trabajo final para revisión.

---

# 2. Flujo general

El flujo principal es:

1. El usuario inicia sesión.
2. Selecciona **Generar presupuesto**.
3. Define parámetros comerciales:
   - indirectos;
   - utilidad;
   - IVA;
   - desperdicio de referencia.
4. Introduce datos del proyecto:
   - nombre;
   - tipo de obra;
   - ubicación;
   - dimensiones definidas o variables;
   - descripción detallada;
   - texto guía opcional.
5. Gemini interpreta el alcance.
6. Gemini propone partidas, subpartidas y actividades generales subcontratables.
7. La aplicación intenta encontrar antecedentes de precio en la base interna.
8. Si no encuentra una referencia interna suficientemente similar, utiliza la estimación propuesta por Gemini.
9. Python calcula los importes.
10. La aplicación muestra un resumen de consulta.
11. Se generan:
    - Excel;
    - TXT auxiliar;
    - ZIP con ambos archivos.
12. Si no se activó el modo simulación, el proyecto, presupuesto, actividades, conceptos y precios correspondientes se guardan en la base.
13. La revisión detallada y las correcciones finales se realizan en Excel.

---

# 3. Filosofía de los costos

La empresa subcontrata actividades completas.

Por esa razón, el sistema NO intenta obligatoriamente descomponer cada concepto en:

- material;
- mano de obra;
- herramienta;
- cuadrilla;
- rendimiento.

En su lugar, una actividad puede representar un servicio integral del subcontratista.

Ejemplo:

```text
Partida:
HERRERÍA Y CANCELERÍA

Subpartida:
Cancelería

Actividad:
Suministro e instalación de cancel fijo en cristal templado,
incluyendo fabricación, herrajes, sellos, transporte e instalación.

Unidad:
M2

Cantidad:
4.50

Costo unitario:
Costo integral estimado o histórico del subcontratista.
```

La aplicación puede utilizar `LOTE`, `PZA`, `JGO`, `M2`, `ML`, `M3` u otra unidad apropiada.

---

# 4. Cálculo económico

Gemini no calcula los totales económicos.

Los cálculos se realizan en Python para mantener consistencia.

La lógica es:

```text
Costo directo
= suma de los costos directos de las actividades

Indirectos
= Costo directo × porcentaje de indirectos

Base de utilidad
= Costo directo + Indirectos

Utilidad
= Base de utilidad × porcentaje de utilidad

Venta antes de IVA
= Costo directo + Indirectos + Utilidad

IVA
= Venta antes de IVA × porcentaje de IVA

Total
= Venta antes de IVA + IVA
```

El desperdicio se utiliza únicamente como referencia cuando técnicamente corresponda.

No se aplica automáticamente como porcentaje general a todas las actividades, porque muchos costos ya representan servicios integrales de subcontratistas.

---

# 5. Modo simulación

El formulario incluye un interruptor llamado **Simulación**.

Cuando está activo:

- Gemini genera normalmente el presupuesto;
- se consulta la base interna existente;
- se realizan todos los cálculos;
- se genera Excel;
- se genera TXT;
- se genera ZIP;
- se muestran resultados.

Pero NO se guardan nuevos datos en:

- proyectos;
- presupuestos;
- conceptos;
- historial de precios;
- actividades del presupuesto.

Esto permite:

- probar cambios de prompt;
- capacitar usuarios;
- realizar ejemplos;
- probar proyectos hipotéticos;
- evitar contaminar el historial empresarial.

Una simulación puede LEER precios de la base existente, pero no escribir en ella.

---

# 6. Archivos generados

La aplicación produce un paquete asociado al proyecto.

El código se genera a partir de la ubicación, el tipo de proyecto y un consecutivo.

Ejemplo:

```text
COY-BAN-0004
```

Una salida típica contiene:

```text
COY-BAN-0004/
├── COY-BAN-0004_Presupuesto.xlsx
└── COY-BAN-0004_Captura_Plataforma.txt
```

El ZIP permite descargar ambos archivos juntos.

## 6.1 Excel

El Excel es el documento principal de revisión.

Está diseñado para que la empresa pueda:

- revisar descripciones;
- ajustar cantidades;
- corregir costos;
- modificar alcances;
- revisar parámetros;
- continuar trabajando sin depender de Streamlit.

Las fórmulas económicas importantes permanecen en el archivo.

## 6.2 TXT auxiliar

El TXT NO se carga automáticamente en la plataforma externa.

Su objetivo es facilitar la captura manual.

Contiene la información organizada para copiar:

- partida;
- subpartida;
- actividad;
- cantidad;
- unidad;
- margen;
- beneficio;
- costo;
- venta.

---

# 7. Base de datos

La aplicación utiliza cinco tablas principales.

## 7.1 `projects`

Guarda información general del proyecto.

Campos principales:

- `id`
- `code`
- `name`
- `project_type`
- `location`
- `dimension_mode`
- `dimensions_text`
- `description`
- `guide_text`
- `main_activity`
- `created_at`

## 7.2 `budgets`

Guarda información económica de cada presupuesto.

Campos principales:

- proyecto asociado;
- versión;
- estado;
- indirectos;
- utilidad;
- IVA;
- desperdicio;
- costo directo;
- indirectos calculados;
- utilidad calculada;
- venta antes de IVA;
- IVA;
- total;
- resumen de alcance;
- fecha.

## 7.3 `concepts`

Es el catálogo interno de actividades.

Cada concepto puede incluir:

- código;
- partida;
- subpartida;
- descripción;
- unidad;
- descripción normalizada;
- presupuesto donde fue creado;
- fecha.

Un concepto no representa necesariamente un APU tradicional.

Representa una actividad reutilizable para futuros presupuestos.

## 7.4 `price_history`

Guarda múltiples precios para cada concepto.

Nunca se pretende que un concepto tenga únicamente "un precio definitivo".

Ejemplo:

```text
Cancel de cristal templado
├── 2026-05-15  $4,400/m2  Estimación IA
├── 2026-06-08  $4,650/m2  Referencia externa
├── 2026-07-20  $4,850/m2  Cotización proveedor
└── 2026-08-05  $5,050/m2  Costo real
```

Esto permite que el catálogo gane valor con el tiempo.

## 7.5 `budget_items`

Guarda la fotografía exacta de las actividades utilizadas en cada presupuesto.

Esto es importante porque un concepto del catálogo puede cambiar posteriormente, pero el presupuesto histórico debe conservar:

- descripción utilizada;
- cantidad;
- unidad;
- costo usado;
- venta;
- fuente;
- criterios;
- consideraciones.

---

# 8. Fuentes y estados de precios

Los precios pueden tener distintos orígenes.

Ejemplos:

- `COTIZACION_PROVEEDOR`
- `COSTO_REAL`
- `REFERENCIA_EXTERNA`
- `IA_ESTIMADO`
- `MANUAL`
- `BASE_INTERNA`

La interfaz administrativa traduce estos códigos a nombres más legibles.

También existe un estado, por ejemplo:

- Validado.
- Cotizado por proveedor.
- Costo real.
- Referencia externa.
- Estimado por IA.

Y un nivel de confianza:

- Alta.
- Media.
- Baja.

La intención es evitar que una estimación inicial de IA se confunda con un costo real confirmado.

---

# 9. Catálogo e historial

Los administradores tienen una sección llamada:

**Catálogo e historial**

Esta reemplaza la antigua interfaz que se parecía demasiado a un editor de base de datos.

La interfaz está dividida en:

1. Conceptos.
2. Precios.
3. Proyectos.
4. Presupuestos.
5. Exportar.

---

# 10. Conceptos

La pestaña **Conceptos** funciona como un catálogo.

Permite:

- buscar por descripción;
- buscar por código;
- buscar por partida;
- buscar por subpartida;
- filtrar por partida;
- abrir la ficha de un concepto;
- consultar su último costo;
- consultar cuántas veces se utilizó;
- ver precios recientes;
- editar sus datos descriptivos;
- crear conceptos manualmente;
- eliminar un concepto con confirmación.

La edición no se realiza directamente sobre una tabla.

Primero se abre una ficha y después se presiona:

```text
Editar concepto
```

Modificar un concepto NO modifica automáticamente:

- presupuestos históricos;
- precios históricos.

Esto permite corregir nombres o clasificaciones sin destruir la trazabilidad.

---

# 11. Precios

La pestaña **Precios** está separada de Conceptos deliberadamente.

El procedimiento es:

1. Buscar un concepto.
2. Seleccionarlo.
3. Revisar su historial.
4. Agregar una nueva referencia.
5. Indicar:
   - costo unitario;
   - origen;
   - estado;
   - confianza;
   - proveedor o nota.
6. Guardar.

Agregar un precio nuevo NO borra los anteriores.

Los precios antiguos deben conservarse normalmente.

La opción de eliminar existe únicamente para corregir registros erróneos y exige una confirmación explícita.

---

# 12. Proyectos

La pestaña **Proyectos** funciona como archivo histórico.

Permite:

- buscar por código;
- nombre;
- ubicación;
- tipo de proyecto.

Al abrir un proyecto se muestra:

- código;
- nombre;
- tipo;
- ubicación;
- actividad principal;
- dimensiones;
- descripción inicial;
- presupuestos asociados.

Un administrador puede corregir datos descriptivos.

Editar estos datos NO recalcula presupuestos existentes.

La eliminación de un proyecto es una acción avanzada porque elimina también sus presupuestos asociados.

---

# 13. Presupuestos

La pestaña **Presupuestos** permite consultar lo que ya se guardó.

Muestra:

- proyecto;
- nombre;
- costo directo;
- indirectos;
- utilidad;
- venta antes de IVA;
- total;
- fecha;
- estado.

Al abrir uno se muestran también sus actividades.

Esta sección es principalmente de consulta.

La corrección detallada de un presupuesto continúa realizándose en Excel.

---

# 14. Exportación administrativa

La pestaña **Exportar** permite descargar CSV de:

- conceptos;
- historial de precios;
- proyectos;
- presupuestos;
- actividades de presupuestos.

Estas descargas son útiles para:

- respaldo;
- auditoría;
- revisión;
- análisis externo;
- migración;
- depuración.

Descargar un CSV no modifica la base.

---

# 15. Autenticación

Toda la aplicación está protegida con `streamlit-authenticator`.

Las credenciales no están en GitHub.

Se cargan desde Streamlit Secrets.

Existen dos roles principales.

## Usuario

Puede:

- generar presupuestos;
- usar simulación;
- consultar resultados de la generación actual;
- descargar Excel;
- descargar TXT;
- descargar ZIP.

## Administrador

Puede hacer todo lo anterior y además:

- abrir Catálogo e historial;
- editar conceptos;
- agregar precios;
- eliminar registros;
- consultar proyectos;
- consultar presupuestos;
- exportar la base.

---

# 16. Cookie de sesión

`streamlit-authenticator` utiliza una cookie para mantener la sesión.

La duración se controla con:

```toml
[auth_cookie]
expiry_days = 30
```

Esto evita pedir usuario y contraseña en cada visita durante el periodo de validez de la sesión.

La cookie no sustituye a las credenciales del servidor.

---

# 17. Configuración de Secrets

En Streamlit Community Cloud:

1. Abra la aplicación.
2. Entre a **Settings**.
3. Abra **Secrets**.
4. Pegue una configuración TOML.

Ejemplo:

```toml
GEMINI_API_KEY = "TU_API_KEY"

DATABASE_URL = "postgresql://..."

[auth_cookie]
name = "presupuestador_empresa_auth"
key = "UNA_CLAVE_PRIVADA_LARGA_Y_ALEATORIA"
expiry_days = 30

[credenciales_app.usernames.admin]
email = "admin@empresa.com"
first_name = "Administrador"
last_name = "Empresa"
password = "CONTRASENA_ADMIN"
roles = ["admin"]

[credenciales_app.usernames.usuario1]
email = "usuario@empresa.com"
first_name = "Usuario"
last_name = "Empresa"
password = "CONTRASENA_USUARIO"
roles = ["usuario"]
```

El archivo `secrets_ejemplo.toml` incluido en este proyecto contiene una plantilla.

NO suba un archivo real de secrets a GitHub.

---

# 18. Gemini API

La API key de Gemini se lee desde:

```toml
GEMINI_API_KEY = "..."
```

en Streamlit Secrets.

La clave no se solicita al usuario final.

Esto permite que la empresa controle una única credencial desde el servidor.

La aplicación utiliza respuestas estructuradas mediante Pydantic.

Gemini devuelve datos como:

- partida;
- subpartida;
- código sugerido;
- descripción;
- unidad;
- cantidad;
- estimación de costo;
- criterio de cantidad;
- fundamento de inclusión;
- confianza;
- consideraciones.

Posteriormente Python procesa esos datos.

---

# 19. Búsqueda de precios internos

Después de recibir actividades de Gemini, la aplicación compara cada actividad con el catálogo existente.

La búsqueda considera:

- descripción normalizada;
- coincidencia de palabras;
- similitud textual;
- unidad.

Cuando encuentra un concepto interno suficientemente parecido puede utilizar su precio histórico más reciente como referencia.

La idea es que, a medida que la empresa acumule proyectos reales, la dependencia de las estimaciones de IA disminuya.

---

# 20. ¿Qué ocurre cuando no existe un concepto?

Cuando no hay coincidencia interna suficiente:

1. se utiliza inicialmente la actividad propuesta por Gemini;
2. se utiliza el costo estimado disponible;
3. se identifica su procedencia;
4. si el presupuesto es real, el concepto puede incorporarse al catálogo;
5. el precio se conserva con su fuente y nivel de confianza.

Esto permite comenzar con una base vacía.

---

# 21. Supabase / PostgreSQL

Para producción se recomienda PostgreSQL.

Supabase se utiliza como proveedor sencillo de PostgreSQL administrado.

Procedimiento general:

1. Crear una cuenta en Supabase.
2. Crear un proyecto.
3. Abrir las opciones de conexión.
4. Copiar la cadena PostgreSQL correspondiente.
5. Guardarla en Streamlit Secrets:

```toml
DATABASE_URL = "postgresql://..."
```

La aplicación crea automáticamente las tablas necesarias si no existen.

No es necesario crear manualmente:

- `projects`;
- `budgets`;
- `concepts`;
- `price_history`;
- `budget_items`.

---

# 22. SQLite

Si `DATABASE_URL` no existe, la aplicación utiliza:

```text
presupuestador_empresa.db
```

junto al archivo `app.py`.

Esto es útil para:

- pruebas locales;
- desarrollo;
- demostraciones.

NO debe tratarse como almacenamiento empresarial persistente dentro de Streamlit Community Cloud.

Cuando existe `DATABASE_URL` y PostgreSQL falla, la aplicación se detiene.

No cambia silenciosamente a SQLite.

Esto evita que un usuario crea que un presupuesto fue guardado en Supabase cuando realmente quedó en un archivo temporal.

---

# 23. Instalación local

Se recomienda Python moderno compatible con las dependencias.

Crear un entorno virtual:

```bash
python -m venv .venv
```

Activarlo.

Windows:

```bash
.venv\Scripts\activate
```

Linux/macOS:

```bash
source .venv/bin/activate
```

Instalar:

```bash
pip install -r requirements.txt
```

Ejecutar:

```bash
streamlit run app.py
```

---

# 24. Secrets locales

Para ejecutar localmente puede crear:

```text
.streamlit/
└── secrets.toml
```

Utilice el mismo formato de `secrets_ejemplo.toml`.

Nunca suba `.streamlit/secrets.toml` con información real a un repositorio público.

Conviene agregarlo a `.gitignore`.

Ejemplo:

```gitignore
.streamlit/secrets.toml
presupuestador_empresa.db
__pycache__/
.venv/
```

---

# 25. Despliegue en Streamlit Community Cloud

Estructura mínima del repositorio:

```text
repositorio/
├── app.py
├── requirements.txt
└── README.md
```

Después:

1. subir el repositorio a GitHub;
2. crear una app en Streamlit Community Cloud;
3. seleccionar `app.py`;
4. configurar Secrets;
5. desplegar.

El repositorio puede ser público siempre que NO contenga:

- API keys;
- DATABASE_URL;
- contraseñas;
- claves de cookies;
- secrets reales.

---

# 26. Seguridad

Principios aplicados:

## 26.1 Secrets del lado servidor

Se mantienen en Streamlit Secrets:

- Gemini API key;
- conexión PostgreSQL;
- usuarios;
- contraseñas;
- clave de cookie.

## 26.2 Roles

La administración de datos solo aparece para usuarios con rol:

```text
admin
```

## 26.3 Confirmaciones de eliminación

Las operaciones destructivas requieren escribir frases como:

```text
ELIMINAR CONCEPTO
ELIMINAR PRECIO
ELIMINAR PROYECTO
ELIMINAR PRESUPUESTO
```

## 26.4 Historial en lugar de sobrescritura

Los precios se agregan al historial en vez de reemplazar automáticamente el anterior.

## 26.5 PostgreSQL persistente

La información empresarial no depende del almacenamiento temporal de Streamlit.

---

# 27. Modificar datos iniciales

Después de generar un presupuesto existe una opción para volver al formulario inicial.

Su finalidad es corregir errores importantes del alcance.

El comportamiento esperado es:

1. descartar la generación actual;
2. volver al formulario;
3. conservar los datos iniciales para facilitar la corrección;
4. modificar la descripción o parámetros;
5. generar nuevamente desde cero.

La aplicación no intenta "parchar" parcialmente una generación anterior.

Esto reduce inconsistencias.

---

# 28. Diferencia entre editar en Streamlit y editar en Excel

Streamlit sirve para:

- capturar datos iniciales;
- generar;
- consultar;
- administrar catálogo e históricos.

Excel sirve para:

- revisión final;
- correcciones específicas;
- negociación;
- ajustes comerciales;
- entrega interna o al cliente.

No se pretende duplicar en Streamlit todas las funciones de edición de Excel.

---

# 29. Recomendaciones de operación

## Al comenzar

La base puede estar prácticamente vacía.

Es normal que la IA tenga mayor participación.

## Conforme se reciban cotizaciones

Agregar al historial:

- costo;
- proveedor;
- fecha;
- confianza;
- estado.

## Cuando termine una obra

Cuando sea posible, registrar costos reales.

Los costos reales son más valiosos que las estimaciones iniciales.

## No borrar precios antiguos por estar desactualizados

Es preferible conservarlos y agregar un registro nuevo.

El histórico tiene valor para comparar evolución de costos.

## Usar simulación para pruebas

Evita llenar la base con conceptos ficticios.

---

# 30. Respaldo

Aunque Supabase almacene la base de manera persistente, es recomendable descargar periódicamente los CSV desde:

**Catálogo e historial → Exportar**

Se recomienda conservar respaldos fechados.

Ejemplo:

```text
backup_2026-08-31/
├── conceptos.csv
├── historial_precios.csv
├── proyectos.csv
├── presupuestos.csv
└── actividades_presupuestos.csv
```

---

# 31. Limitaciones actuales

La aplicación sigue siendo un sistema de apoyo.

Actualmente:

- Gemini puede proponer actividades incorrectas;
- una estimación de IA no equivale a una cotización;
- la similitud textual puede asociar conceptos parecidos que no sean equivalentes;
- los precios deben validarse cuando el riesgo económico sea importante;
- la revisión final sigue siendo responsabilidad de la empresa;
- no existe todavía una integración automática completa con catálogos externos;
- el TXT es auxiliar y la captura en la plataforma externa continúa siendo manual.

---

# 32. Evolución recomendada

Posibles mejoras futuras:

1. integración de catálogos externos;
2. importación masiva de conceptos desde CSV/XLSX;
3. proveedores vinculados a conceptos;
4. panel de evolución histórica de precios;
5. comparación entre costo presupuestado y costo real;
6. versiones formales de un mismo presupuesto;
7. bitácora de cambios;
8. permisos más detallados;
9. auditoría por usuario;
10. integración con almacenamiento documental;
11. búsqueda semántica más avanzada;
12. generación de solicitudes de cotización a proveedores.

---

# 33. Archivos del proyecto

## `app.py`

Contiene:

- autenticación;
- interfaz;
- Gemini;
- cálculos;
- base de datos;
- catálogo;
- historial;
- generación de archivos.

## `requirements.txt`

Lista las dependencias de Python necesarias.

## `secrets_ejemplo.toml`

Plantilla de configuración.

No contiene secretos reales.

## `README.md`

Este documento.

---

# 34. Dependencias principales

La versión actual utiliza:

```text
streamlit
streamlit-authenticator
pandas
openpyxl
google-genai
pydantic
psycopg
```

Las versiones y restricciones concretas se encuentran en `requirements.txt`.

---

# 35. Resumen operativo

Para un usuario normal:

```text
Iniciar sesión
→ Generar presupuesto
→ Ingresar alcance
→ Generar
→ Revisar resumen
→ Descargar Excel/TXT/ZIP
→ Corregir en Excel
```

Para un administrador:

```text
Iniciar sesión
→ Catálogo e historial
→ Conceptos / Precios / Proyectos / Presupuestos
→ Consultar o administrar
→ Exportar respaldos cuando sea necesario
```

Para una prueba:

```text
Activar Simulación
→ Generar normalmente
→ Descargar archivos
→ No se escribe nada nuevo en la base
```

---

# 36. Principio central del sistema

La intención es que el sistema evolucione de:

```text
IA estima casi todo
```

hacia:

```text
IA interpreta el proyecto
+
la base histórica aporta experiencia real
+
Python calcula
+
la empresa revisa
```

Cuantos más costos reales, cotizaciones y proyectos se registren correctamente, mayor será el valor de la base interna y menor la dependencia de estimaciones puramente generadas por IA.
