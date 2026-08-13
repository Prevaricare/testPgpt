# Presupuestador de Remodelaciones

Esta aplicación ayuda a generar un primer presupuesto para proyectos de remodelación e interiorismo.

La app genera una propuesta inicial.  
La revisión final de precios, cantidades y conceptos se realiza en el archivo Excel.

---

# 1. Iniciar sesión

Abra la aplicación e ingrese su usuario y contraseña.

Los usuarios normales pueden generar presupuestos.

Los administradores también pueden entrar a **Catálogo e historial** para revisar conceptos, precios y proyectos guardados.

---

# 2. Crear un presupuesto

Entre a:

**Generar presupuesto**

Primero defina los parámetros generales:

- Indirectos
- Utilidad
- IVA
- Desperdicio de referencia

Después ingrese los datos del proyecto.

## Nombre del cliente

Escriba el nombre del cliente.

Ejemplo:

```text
Desarrollos de la Vega
```

## Ubicación

Escriba la ubicación principal del proyecto.

Ejemplo:

```text
Farallón, Álvaro Obregón, CDMX
```

La aplicación utiliza el cliente y la ubicación para generar un código.

Ejemplo:

```text
DDV-FAR-0001
```

## Tipo de obra

Seleccione la opción que más se acerque al proyecto.

Por ejemplo:

- Remodelación interior general
- Baño
- Cocina
- Oficina
- Local comercial
- Caseta / acceso

---

# 3. Descripción general de trabajos

Este es el campo más importante.

Aquí debe escribir:

- zonas del proyecto;
- medidas conocidas;
- trabajos solicitados;
- elementos que se van a retirar;
- elementos nuevos;
- materiales o acabados importantes;
- cualquier información necesaria para entender la obra.

No es necesario escribirlo de forma muy técnica.

Ejemplo:

```text
Casa habitación.

SEGUNDA PLANTA

Recámara principal de aproximadamente 4.20 x 3.80 m.

- Retiro de piso laminado existente.
- Colocación de piso nuevo.
- Reparación y pintura de muros.
- Cambio de luminarias.

AZOTEA

Área aproximada de 9 x 4 m.

- Limpieza de superficie.
- Reparación de fisuras.
- Impermeabilización completa.
- Revisión de bajadas pluviales.
```

Mientras más clara sea la descripción, mejor será el resultado.

---

# 4. Texto guía

Este campo sirve para dar instrucciones generales que deben aplicarse a todo el presupuesto.

No es para repetir los trabajos.

Puede utilizarlo para indicar:

- nivel de calidad;
- tipo de acabados;
- restricciones;
- horarios;
- protecciones;
- trabajos que no deben incluirse;
- criterios generales.

Ejemplo:

```text
- Considerar acabados de gama media.
- Incluir protección con plástico y cartón engomado en zonas de tránsito.
- Considerar retiro de desperdicios y limpieza final.
- Los trabajos solo pueden realizarse de lunes a viernes.
- No considerar jardinería.
- Si falta una especificación, utilizar una opción comercial de gama media.
```

Este campo es opcional.

---

# 5. Simulación

Active **Simulación** cuando solo quiera hacer una prueba.

En simulación la aplicación:

- genera el presupuesto;
- genera Excel;
- genera TXT;
- permite hacer revisiones;

pero **no guarda el proyecto en la base de datos**.

Úselo para:

- pruebas;
- capacitación;
- ejemplos;
- verificar cómo interpreta Gemini un proyecto.

Si el resultado de una simulación le gusta, al final puede utilizar:

**Guardar versión actual en la base de datos**

No es necesario volver a generar el presupuesto.

---

# 6. Generar presupuesto

Cuando termine de ingresar la información presione:

**Generar presupuesto**

La aplicación analizará el proyecto y generará:

- partidas;
- subpartidas;
- actividades;
- cantidades aproximadas;
- costos;
- indirectos;
- utilidad;
- IVA;
- total.

También intenta utilizar precios existentes del catálogo interno cuando encuentra conceptos similares.

---

# 7. Revisar el resultado

Después de generar, se muestran nuevamente los datos originales del proyecto.

Estos datos aparecen bloqueados para evitar modificarlos accidentalmente.

Después se muestra:

- costo directo;
- indirectos;
- utilidad;
- total;
- tabla de actividades;
- origen de los precios;
- consideraciones importantes.

Esta pantalla sirve principalmente para revisar que el presupuesto tenga sentido general.

Los cambios pequeños deben hacerse directamente en Excel.

---

# 8. Descargar archivos

La app genera un paquete ZIP.

Ejemplo:

```text
DDV-FAR-0001-V01.zip
```

Dentro se encuentran:

```text
DDV-FAR-0001-V01_Presupuesto.xlsx
DDV-FAR-0001-V01_Captura_Plataforma.txt
```

## Excel

Es el archivo principal de trabajo.

Aquí puede:

- cambiar precios;
- cambiar cantidades;
- corregir descripciones;
- agregar notas;
- hacer ajustes finales.

## TXT

Es un archivo auxiliar.

Sirve para facilitar la captura manual de información en la plataforma de presupuestos de la empresa.

---

# 9. Revisiones importantes del presupuesto

Si después de generar descubre que falta una parte importante del proyecto, use:

**Revisión estructural del presupuesto**

Esta función es para cambios grandes.

Ejemplos:

```text
Se olvidó contemplar toda la impermeabilización de la azotea.
Agregar preparación de superficie, reparación de fisuras e
impermeabilización para el área de 9 x 4 m.
```

Otro ejemplo:

```text
La cancelería debe cambiar completamente.

Ya no será aluminio ligero. Se requiere aluminio línea pesada
con cristal templado de 10 mm.
```

La aplicación genera una nueva versión:

```text
V01
V02
V03
```

No elimina automáticamente las versiones anteriores.

---

# 10. Cuándo NO usar una revisión

No use la revisión estructural para cambios pequeños.

Por ejemplo:

```text
Cambiar $4,500 por $4,800
```

o:

```text
Cambiar cantidad de 3 a 4 piezas
```

o:

```text
Corregir una palabra
```

Para esos cambios use directamente el Excel.

La revisión con Gemini debe reservarse para cambios importantes de alcance.

---

# 11. Catálogo e historial

Los administradores pueden entrar a:

**Catálogo e historial**

Ahí pueden consultar:

- Conceptos
- Precios
- Proyectos
- Presupuestos
- Mantenimiento
- Exportación

## Conceptos

Permite revisar las actividades que la empresa ha ido acumulando.

## Precios

Permite agregar nuevas referencias.

Ejemplo:

```text
Concepto:
Suministro e instalación de cancelería

Costo:
$5,200 / m2

Fuente:
Cotización de proveedor

Estado:
Validado
```

No es necesario borrar un precio viejo cuando llega uno nuevo.

Es mejor agregar el nuevo y conservar el historial.

## Proyectos

Permite consultar proyectos anteriores.

## Presupuestos

Permite revisar las versiones guardadas de los presupuestos.

---

# 12. Borrar una prueba guardada por error

En la sección de generación existe:

**Corrección de última carga**

Sirve para eliminar el último proyecto guardado si se subió por error.

Por ejemplo:

- una prueba que debía ser simulación;
- un proyecto con datos incorrectos;
- una generación que no debía guardarse.

La función pide una clave de eliminación.

---

# 13. Recomendaciones

Para obtener mejores resultados:

1. Describa claramente cada zona.
2. Incluya las medidas que conozca.
3. Separe las actividades por planta, espacio o área.
4. Indique materiales importantes cuando ya estén definidos.
5. Utilice Texto guía para criterios generales.
6. Use Simulación cuando esté haciendo pruebas.
7. Corrija precios y detalles menores directamente en Excel.
8. Use Revisiones solo cuando cambie una parte importante del proyecto.
9. Cuando exista una cotización real de proveedor, agréguela al historial de precios.

---

# 14. Ejemplo completo de captura

## Cliente

```text
Corporativo Cocoteros
```

## Ubicación

```text
Coyoacán, CDMX
```

## Tipo

```text
Caseta / acceso
```

## Descripción general

```text
Caseta de vigilancia.

Medidas aproximadas 1.29 x 1.48 m.

- Sellar puertas louver para evitar entrada de frío.
- Colar una losa para cerrar la parte superior.
- Hacer apertura en muro de piedra volcánica.
- Colocar cancel de seguridad.
- Abrir ranura para recepción de correo.
- Fabricar escritorio en escuadra con cajones.
- Colocar silla para personal.
- Colocar frigobar.
- Agregar contactos eléctricos.
- Instalar iluminación.
- Considerar sistema de CCTV.
```

## Texto guía

```text
- Considerar materiales de gama media.
- Incluir instalación completa.
- Considerar protección y limpieza.
- Cuando un elemento requiera proveedor especializado,
  marcarlo como referencia aproximada.
```

Después presione:

**Generar presupuesto**

Revise el resultado general y descargue el Excel para realizar los ajustes finales.
