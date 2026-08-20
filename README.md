# Guía de Uso: Generación de Presupuestos

Esta sección explica cómo utilizar el módulo principal de la aplicación para generar presupuestos de obra asistidos por Inteligencia Artificial (Gemini).

## Parámetros Comerciales

Antes de iniciar, dirígete a la barra lateral izquierda y ajusta los porcentajes financieros que aplicarán a tu presupuesto: Indirectos (%), Utilidad (%), IVA (%) y Desperdicio de referencia (%).

## Nombre del cliente

Escriba el nombre del cliente o de la empresa que solicita el proyecto.

Ejemplo:

```text
Desarrollos de la Vega
```

## Ubicación

Escriba la ubicación principal del proyecto. Esto ayuda a contextualizar los costos de referencia.

Ejemplo:

```text
Farallón, Álvaro Obregón, CDMX
```

## Tipo de obra

Seleccione del menú desplegable la categoría que mejor describa el proyecto.

Ejemplo:

```text
Remodelación interior general
```

## Descripción general de trabajos

Describa el alcance exacto del proyecto. Sea lo más específico posible indicando las zonas, las medidas conocidas y los trabajos particulares a realizar. Las dimensiones exactas permiten que la IA genere mejores cálculos automáticos.

Ejemplo:

```text
Casa habitación.

SEGUNDA PLANTA
Recámara principal de aproximadamente 4.20 x 3.80 m.
- Retiro de piso laminado existente.
- Colocación de piso nuevo.
- Reparación y pintura de muros.
- Sustitución de luminarias.
```

## Texto guía

Escriba las consideraciones generales, reglas operativas o exclusiones que aplicarán a todo el proyecto. 

**¿Para qué se usa en la IA?** 
Este texto se inyecta directamente en las instrucciones de la Inteligencia Artificial (prompt) y funciona como una "regla global". Le indica a la IA qué criterios asumir cuando no hay detalles específicos en la descripción. Por ejemplo, sirve para indicarle que siempre considere acarreo de escombros, que los materiales son de calidad "premium", o que excluya ciertos trámites de los costos.

Ejemplo:

```text
- Considerar protección básica de las áreas de trabajo.
- Evitar asumir trabajos que no estén explícitamente descritos.
- En trabajos de pintura, considerar preparación básica, resanes y dos manos.
- Considerar materiales y acabados de lujo para los pisos.
```

## Modo Simulación

Si desea generar un presupuesto rápido o de prueba sin que quede guardado en el historial de proyectos y presupuestos de la base de datos interna, active el interruptor de **Simulación**.

## Botón "Generar presupuesto"

Haga clic en este botón para procesar la información. La IA estructurará las actividades y el sistema buscará automáticamente los precios unitarios en la base histórica de la empresa o en el tabulador de referencia. Si falta algún precio, la IA realizará una estimación inicial.
