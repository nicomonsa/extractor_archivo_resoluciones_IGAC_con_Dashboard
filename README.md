# 🏛️ Parser de Archivos Catastrales y Tablero de Seguimiento

**Descripción Corta:** Conjunto de herramientas en Python para la ingeniería inversa, extracción y consolidación de archivos planos de ancho fijo del catastro multipropósito colombiano (IGAC), con generación automática de tableros HTML de seguimiento.

## Sobre el Proyecto

Desarrollado para la supervisión municipal de un convenio de actualización catastral con enfoque multipropósito. Los archivos de resoluciones de conservación y de la base catastral se entregan como texto posicional sin delimitador ni diccionario de datos: la estructura se descifró por análisis de frecuencias y validación cruzada contra los valores declarados en los propios registros, y posteriormente se confirmó contra el layout oficial.

El resultado convierte entregas indescifrables en tablas relacionales listas para PostgreSQL, Power BI o ArcGIS Pro.

## Funcionalidades Principales

🔍 **Extracción Posicional Documentada:** Diccionario de posiciones explícito (`read_fwf` con `colspecs`) para dos formatos distintos: resoluciones de conservación (410/428 caracteres) y cortes de la base catastral R1/R2 (312/330 caracteres). Los campos sin confirmar se extraen en crudo y quedan marcados como pendientes, no interpretados.

🔑 **Reconstrucción del Código Predial Nacional:** Ensambla el identificador de 30 posiciones (Resolución IGAC 1149 de 2021) a partir de sus bloques —zona, sector, comuna, barrio, manzana, terreno, condición, torre, piso y unidad— y valida su longitud registro a registro.

🧩 **Resolución de Granularidad Múltiple:** Separa los tres tipos de registro encadenados por predio, cada uno con cardinalidad propia (por propietario, por matrícula, por vigencia), y los reduce a un consolidado de un renglón por predio mediante pivoteo de cancelación e inscripción a columnas.

🔎 **Lectura del Movimiento Efectivo:** Compara campo por campo los estados anterior y posterior de cada predio para determinar qué cambió realmente —titular, área, avalúo, destino económico o solo la escritura del nombre—, de forma separada de la clase de mutación declarada en el archivo.

⚠️ **Validación de Duplicidad en Tres Niveles:** Detecta archivos idénticos por huella de contenido (MD5), la misma resolución presente en varias entregas y registros repetidos por llave compuesta. Depura antes de consolidar para no inflar cifras, conservando el inventario completo de lo recibido y registrando en alertas qué copia se descartó.

📊 **Tablero HTML Autocontenido:** Genera un dashboard por corrida más un índice navegable. Embebe la totalidad de los registros como arreglo comprimido con diccionarios de valores repetidos (~70 bytes por predio frente a ~390 en HTML plano), con búsqueda por número predial sobre el universo completo y no sobre la página visible.

📂 **Salida Estructurada:** CSV con punto y coma y coma decimal para configuración regional en español, Excel multi-hoja con identificadores forzados a texto —evita la corrupción del código predial por notación científica—, y subcarpeta con marca temporal por ejecución para trazabilidad entre entregas.

## Escala Verificada

19 archivos procesados: ~22.000 registros de resoluciones de conservación y 186.615 registros de cortes de la base catastral, consolidados en 77.373 predios.

## Tecnologías

- Python 3.x
- Pandas (`read_fwf`, pivot_table)
- OpenPyXL (Excel I/O)
- hashlib (detección de duplicados)
- Bootstrap 5 · DataTables · Chart.js (capa de visualización)

## Nota

El repositorio contiene únicamente la lógica de extracción y los diccionarios de estructura. No incluye datos catastrales ni información de titulares.
