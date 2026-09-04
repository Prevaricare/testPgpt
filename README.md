# Presupuestador con IA — V15

Aplicación en Streamlit para generar presupuestos preliminares de remodelación e interiorismo con apoyo de Gemini, base de datos interna, referencias externas y exportación a Excel.

La V15 trabaja con **actividades separadas por área física**. Un mismo oficio realizado en espacios distintos debe mantenerse como conceptos distintos para poder asignarlo a proveedores diferentes, revisar costos por zona y editar el presupuesto con mayor facilidad.

---

## 1. Cómo darle información al programa

La calidad del presupuesto depende mucho más de **cómo se organiza la información** que de escribir un texto muy largo.

El formato recomendado es:

1. Identificar claramente cada área.
2. Dar dimensiones o cantidades conocidas dentro de esa misma área.
3. Escribir una actividad por renglón.
4. Indicar qué existe actualmente y qué se desea hacer.
5. Especificar materiales, acabados o alcances solamente cuando ya estén definidos.
6. Separar trabajos generales de los trabajos de cada espacio.

No es necesario calcular precios, indirectos, utilidad, IVA o el 30 % de marca. La aplicación se encarga de esa parte.

### Formato recomendado para la descripción

```text
ÁREA 1
Dimensiones o datos generales del área.
- Trabajo 1.
- Trabajo 2.
- Trabajo 3.

ÁREA 2
Dimensiones o datos generales del área.
- Trabajo 1.
- Trabajo 2.

GENERAL
- Trabajos que afectan a toda la obra.
```

### Evitar

```text
Remodelar cocina y baños, cambiar pisos, arreglar muros, hacer instalaciones,
pintar y poner muebles.
```

Aunque Gemini puede interpretar una descripción así, tendrá que inferir áreas, cantidades, relaciones entre trabajos y alcances.

### Preferir

```text
COCINA
Área aproximada de 3.80 x 3.20 m.
- Retiro de piso existente y colocación de piso porcelánico nuevo.
- Resane y pintura de muros.
- Sustitución de frentes de cocina existentes.

BAÑO 1
Área aproximada de 2.20 x 1.80 m.
- Demolición de recubrimientos existentes.
- Colocación de nuevo recubrimiento en piso y muros.
- Sustitución de WC y lavabo.
```

---

## 2. Ejemplo integral de captura

El siguiente ejemplo está pensado para darle a la aplicación información suficiente para trabajar con pocas suposiciones y conservar correctamente las áreas.

### Nombre del cliente

```text
Desarrollos de la Vega
```

Usar el nombre real o una identificación clara del cliente. El sistema lo utiliza también para generar el código del proyecto.

### Ubicación

```text
Coyoacán, CDMX
```

Conviene incluir alcaldía o municipio y ciudad. La ubicación ayuda a contextualizar costos y características generales del proyecto.

### Tipo de obra

```text
Remodelación interior general
```

Usar la categoría que describa el proyecto completo. Si intervienen cocina, baños, recámaras y otras zonas, normalmente conviene utilizar **Remodelación interior general** en lugar de escoger solamente una de las áreas.

### Nivel de presupuesto

```text
Medio-alto
```

Debe representar la calidad comercial esperada de materiales, acabados y proveedores. No es necesario detallar precios dentro de la descripción.

### Descripción general de trabajos

```text
COCINA
Área aproximada de 4.00 x 3.20 m. Altura libre aproximada de 2.60 m.
Mobiliario de cocina existente que se conservará parcialmente.
- Retiro de piso cerámico existente.
- Suministro y colocación de piso porcelánico nuevo en toda el área.
- Retiro de salpicadero existente y colocación de recubrimiento nuevo en muro de trabajo.
- Resanes menores y pintura de muros y plafón.
- Sustitución de frentes visibles de muebles de cocina existentes, aproximadamente 7.20 m2.
- Fabricación de un módulo nuevo para refrigerador.
- Adecuación de contactos eléctricos para pequeños electrodomésticos.
- Revisión y adecuación de conexiones hidráulicas y sanitarias del fregadero.

BAÑO 1
Área aproximada de 2.40 x 1.80 m. Altura aproximada de 2.40 m.
- Demolición de recubrimiento existente en piso y muros.
- Retiro de WC, lavabo y accesorios existentes.
- Revisión y adecuación de instalaciones hidráulicas y sanitarias.
- Impermeabilización en zona húmeda.
- Suministro y colocación de piso y recubrimiento cerámico nuevo.
- Suministro e instalación de WC y lavabo nuevos.
- Pintura de plafón.
- Colocación de accesorios básicos.

BAÑO 2
Área aproximada de 2.10 x 1.70 m. Altura aproximada de 2.40 m.
- Demolición de recubrimiento existente en piso y muros.
- Retiro de muebles sanitarios existentes.
- Revisión y adecuación de instalaciones hidráulicas y sanitarias.
- Impermeabilización en zona húmeda.
- Suministro y colocación de piso y recubrimiento nuevo.
- Suministro e instalación de WC, lavabo y mezcladora nuevos.
- Pintura de plafón.

BAÑO 3
Área aproximada de 1.80 x 1.60 m. Altura aproximada de 2.40 m.
- Retiro de acabados existentes.
- Adecuación menor de instalaciones hidráulicas y sanitarias.
- Colocación de piso y recubrimiento nuevo.
- Sustitución de WC y lavabo.
- Pintura de plafón.

RECÁMARA PRINCIPAL
Área aproximada de 4.20 x 3.80 m.
- Retiro de piso laminado existente.
- Preparación de firme para recibir acabado nuevo.
- Suministro y colocación de piso laminado nuevo con zoclo.
- Resanes menores y pintura de muros y plafón.
- Fabricación e instalación de clóset de aproximadamente 3.20 m de frente.

SALA / COMEDOR
Área conjunta aproximada de 7.00 x 4.50 m.
- Protección de elementos que se conservarán.
- Resanes menores en muros.
- Pintura de muros y plafón.
- Fabricación de lambrín decorativo en muro de sala, aproximadamente 22.17 m2.
- Revisión de contactos y apagadores existentes; sustituir piezas dañadas.

FACHADA
Frente aproximado de 5.16 m y altura aproximada de 8.00 m.
- Preparación de superficie.
- Resanes menores.
- Sellador cuando sea necesario.
- Aplicación de pintura para exterior en dos manos.

GENERAL
- Protección básica de circulaciones y zonas que se conservarán.
- Acarreos internos y retiro ordinario de residuos derivados de los trabajos.
- Limpieza durante la ejecución.
- Limpieza fina al terminar la obra.
```

### Texto guía

```text
- Considerar protección básica de las áreas de trabajo y zonas de tránsito.
- Incluir limpieza durante los trabajos y limpieza final.
- En pintura, considerar preparación básica, resanes menores, sellador cuando sea necesario y dos manos de pintura.
- En instalaciones y elementos nuevos, considerar suministro, colocación, fijaciones y conexiones, cuando correspondan.
- Mantener materiales y acabados coherentes con el nivel de presupuesto seleccionado.
- Mantener separados los trabajos de cada área, aunque pertenezcan al mismo oficio.
- No combinar Cocina, Baño 1, Baño 2, Baño 3 u otras áreas en un solo concepto.
- Utilizar General únicamente para trabajos que realmente correspondan al conjunto de la obra.
- Si faltan dimensiones suficientes para obtener un metraje confiable, utilizar una unidad comercial adecuada como PZA, LOTE, JGO o PTO antes que inventar cantidades.
- En carpintería, mantener separados los muebles que sean distintos por función, diseño, dimensiones o especificación; únicamente agrupar unidades realmente equivalentes.
```

---

## 3. Por qué este formato funciona mejor

La aplicación ya recibe por separado el nombre, ubicación, tipo y nivel del proyecto. Por eso la **Descripción general de trabajos** debe concentrarse principalmente en el alcance técnico.

Separar el texto por áreas permite que Gemini pueda generar directamente conceptos como:

```text
Cocina  | Albañilería y estructura | Muros | ...
Baño 1  | Albañilería y estructura | Muros | ...
Baño 2  | Albañilería y estructura | Muros | ...
Baño 3  | Albañilería y estructura | Muros | ...
```

en lugar de crear una sola actividad de albañilería para todos los espacios.

Esto facilita:

- asignar proveedores diferentes por área;
- revisar costos individualmente;
- corregir cantidades sin afectar otras zonas;
- ejecutar la obra por etapas;
- comparar presupuestos de distintos proveedores;
- identificar rápidamente dónde está concentrado el costo.

---

## 4. Reglas de captura recomendadas

### Áreas

Usar nombres estables y específicos:

```text
Cocina
Baño 1
Baño 2
Baño 3
Recámara principal
Recámara 2
Sala / Comedor
Fachada
Azotea
General
```

Evitar cambiar de nombre al mismo espacio dentro de una misma descripción. Por ejemplo, no usar primero `Baño visitas` y después `Medio baño` si se trata del mismo sitio.

### Dimensiones

Cuando se conozcan, escribirlas dentro del área correspondiente:

```text
Área aproximada de 3.20 x 2.40 m.
Muro de 4.10 m de largo x 2.60 m de alto.
Barra de 1.70 x 0.51 m y 0.91 m de altura.
Frente de clóset de 3.20 m.
22.17 m2 de lambrín.
```

No es necesario convertir manualmente todas las dimensiones a m2. Cuando el cálculo sea justificable, la aplicación puede realizarlo.

### Actividades

Cada viñeta debe describir una entrega o trabajo identificable:

```text
- Retiro de piso existente.
- Colocación de piso porcelánico nuevo.
- Resane y pintura de muros.
```

No es necesario escribir cada material, herramienta o cuadrilla. La aplicación sigue generando conceptos comerciales, no análisis de precios unitarios completos.

### Especificaciones

Agregar solamente lo que realmente se conoce o sea importante para el costo:

```text
Piso porcelánico formato 60 x 120 cm.
Carpintería en MDF con acabado laminado.
Cancel templado de 9 mm.
Pintura para exterior.
```

Si todavía no se ha elegido marca, modelo o acabado exacto, es preferible indicar el nivel esperado y permitir que el presupuesto siga siendo preliminar.

---

## 5. Qué debe ir en “Texto guía”

El Texto guía sirve para establecer **reglas generales del proyecto**, no para repetir todas las actividades.

Aquí deben colocarse criterios como:

- qué debe considerarse incluido normalmente;
- criterios generales de preparación y limpieza;
- nivel de acabados;
- reglas especiales del presupuesto;
- exclusiones o condiciones que se aplican de manera general.

Si una instrucción corresponde solamente a un baño o una cocina, debe colocarse en la **Descripción general de trabajos**, dentro de esa área, no en el Texto guía.

---

## 6. Funcionamiento general de la aplicación

Flujo principal:

```text
Datos del proyecto
        ↓
Descripción organizada por áreas
        ↓
Gemini genera actividades estructuradas
        ↓
Auditoría de partida, subpartida y secuencia
        ↓
Búsqueda de precios internos / referencias disponibles
        ↓
Cálculos financieros en Python
        ↓
Revisión en Streamlit
        ↓
Ajustar con IA, si es necesario
        ↓
Excel final
        ↓
Guardar en base de datos
```

Gemini se utiliza principalmente para interpretar el alcance, estructurar las actividades, estimar cantidades cuando sea justificable y apoyar en precios cuando no existe una referencia mejor.

Los cálculos de indirectos, utilidad, marca e IVA se realizan posteriormente en la aplicación.

---

## 7. Separación por área en V15

Cada actividad generada contiene ahora un campo `area`.

Regla principal:

> Una actividad pertenece a una sola área física. El mismo oficio ejecutado en áreas diferentes debe mantenerse como actividades independientes.

Ejemplo correcto:

```text
Cocina  → Albañilería → Resanes y preparación
Baño 1  → Albañilería → Resanes y preparación
Baño 2  → Albañilería → Resanes y preparación
Baño 3  → Albañilería → Resanes y preparación
```

Ejemplo que debe evitarse:

```text
General → Albañilería en cocina y tres baños
```

La categoría `General` se reserva para trabajos verdaderamente globales.

---

## 8. Excel generado

### 01 Presupuesto

Columnas actuales:

```text
Área
Partida
Subpartida
Descripción Técnica
Unidad
Cant.
Precio Unitario (MXN)
Importe interno (MXN)
Importe Final (MXN)
```

**Importe interno** corresponde al importe comercial interno antes de aplicar el 30 % de marca de franquicia y antes del IVA.

**Importe Final** aplica automáticamente:

```text
Importe interno × 1.30 × (1 + IVA)
```

La primera hoja mantiene además el resumen por partidas y los totales generales.

### 02 Control Interno

Incluye el área correspondiente a cada concepto y conserva los datos utilizados para revisar costos, indirectos, utilidad, precio comercial y diferencias.

### 03 Trazabilidad

Conserva información de origen de precio, criterios, confianza y consideraciones de cada actividad.

### 04 Costos por Área

Mantiene una revisión interna de los costos agrupados por espacio.

---

## 9. Recargar un Excel existente

La aplicación permite cargar nuevamente un presupuesto `.xlsx` para continuar editándolo sin necesitar la sesión original.

Flujo:

```text
Cargar presupuesto Excel
        ↓
Cargar y continuar editando
        ↓
Reconstrucción del presupuesto
        ↓
Revisión / Ajustar con IA
        ↓
Descargar nuevamente en el formato actual
```

Los archivos generados con versiones recientes contienen mayor información interna y permiten una reconstrucción más completa. Para archivos antiguos o externos, la aplicación intenta reconocer estructuralmente las columnas disponibles.

---

## 10. Consejos para obtener mejores resultados

La entrada ideal no es la más larga: es la que deja menos decisiones ambiguas.

**Mejor:**

```text
BAÑO 2
2.10 x 1.70 m, altura 2.40 m.
- Retiro de recubrimientos existentes.
- Impermeabilización de zona de regadera.
- Colocación de piso y recubrimiento cerámico nuevo.
- Sustitución de WC y lavabo.
```

**Peor:**

```text
Arreglar el segundo baño completo.
```

Siempre que sea posible indicar:

```text
DÓNDE → QUÉ EXISTE → QUÉ SE RETIRA → QUÉ SE MODIFICA → QUÉ SE COLOCA → CUÁNTO
```

No es necesario proporcionar información que todavía no existe. Cuando un dato sea desconocido, es preferible omitirlo o indicarlo como pendiente antes que inventarlo.

---

## 11. Tecnologías principales

- Python
- Streamlit
- Google Gemini API
- Pydantic
- OpenPyXL
- SQLite para desarrollo local
- PostgreSQL / Supabase para persistencia en producción

---

## 12. Objetivo del sistema

El programa no busca sustituir una cotización definitiva de proveedor ni convertirse en un sistema de APU completo.

Su objetivo es transformar rápidamente información inicial de un proyecto de remodelación en un presupuesto comercial estructurado, editable, trazable y reutilizable, manteniendo suficiente separación por área para trabajar con diferentes proveedores y tomar decisiones durante la ejecución.
