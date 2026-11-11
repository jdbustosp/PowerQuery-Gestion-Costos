# Consultas compartidas para controles Oracle

Cada libro conserva `ProyectoActual`, las tablas de captura y un registro local `ORA_CONFIG` con `Sitio`, `CarpetaReportes` y `CarpetaDescargas`. Los módulos de esta carpeta aplican la misma lógica a cualquier libro con ese esquema.

Cada consulta de esta carpeta se vincula en Excel mediante `FxGitHub("Oracle/<nombre>")`. `BD` apunta a `ORA_BD`, `DESCARGAS` a `ORA_DESCARGAS` y `DISPONIBLE` a `ORA_DISPONIBLE`. Cada libro obtiene el código publicado al actualizar Power Query.

| Consulta | Función |
| --- | --- |
| `BD_BASE` | Construye la base de costos con presupuesto, asegurado, consumido, comparativos, provisiones, adicionales, almacén, facturas y teóricos. Calcula la proyección con el asegurado y el saldo por adjudicar. |
| `ORA_FUENTES` | Localiza los informes de Oracle Unifier y normaliza encabezados y números. |
| `ORA_ARBOL` | Reconstruye la jerarquía CBS y los paquetes de trabajo. Genera claves de actividad y nombres de insumos. |
| `ORA_PPTO_BD` | Convierte el presupuesto inicial V1 en filas por insumo y actividad. |
| `ORA_DESCARGAS_LIBRO` | Localiza y abre el libro de descargas del proyecto. |
| `ORA_DESCARGAS` | Lee las asignaciones a comparativos de la tabla DESCARGA. |
| `ORA_PPTO_V2_BD` | Lee y valida el presupuesto vigente V2 por insumo en Actualizacion_ppto. |
| `ORA_DISPONIBLE` | Cruza las descargas con V2 para obtener lo adjudicado y el saldo por adjudicar. |
| `ORA_PPTO_V2` | Prepara las filas que muestran V2 como medida independiente en BD. |
| `ORA_ITEMSINSUMOS` | Extrae los importes proyectados y consumidos del Cost Sheet por actividad CBS. |
| `ORA_MOV` | Normaliza órdenes y contratos, con valor de orden y valor de recepción. |
| `ORA_ASEGURADO` | Normaliza el reporte de asegurado por registro, proceso e insumo. |
| `ORA_COMPRAS` | Adapta compras, otros costos asegurados y almacén. Conserva los signos de devoluciones y reintegros. |
| `ORA_CONTRATOS` | Adapta contratos y actas a las columnas de control. |
| `ORA_DESCUENTOS` | Identifica descuentos y los relaciona con su orden o contrato. |
| `ORA_FACTURAS` | Normaliza facturas y recepciones por orden o contrato, con exclusión de cancelaciones, rechazos y duplicados. |
| `ORA_BD` | Añade a `BD_BASE` las filas de V2 y los niveles CBS para entregar BD. |
