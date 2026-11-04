# Formato Turpial: cambios y hallazgos

Esta carpeta guarda el M que usa solo el libro de Turpial mientras se termina su formato. Los demás libros leen
`Consultas/<nombre>.m`, así que nada de aquí les llega. Turpial lo lee con `FxGitHub("Turpial/<nombre>")`.

Repositorio público: aquí van reglas y hallazgos técnicos, sin contratistas, valores ni números de OC.

## Cuándo pasar el formato a los demás libros
Cuando Juan dé por terminado el formato de Turpial:
1. Aplicar `para-todos-los-libros.patch` en la raíz del repositorio (`git apply`). Cubre `BD.m`, `COMPRAS.m` y `TD3_OC.m`.
2. Medir antes en una copia de cada libro. Revisar lo pendiente de abajo.

## 2026-11-03

### COMPRAS confundía insumos con paréntesis
- `FnClaveLimpia` corta el nombre en el primer `(`. Por eso cruzar `Ins` por nombre unía insumos distintos, por ejemplo
  concretos que solo se diferencian por lo que va entre paréntesis.
- Arreglo: `COMPRAS` cruza por `Codigo ins` + actividad, luego por `Codigo ins` y al final por nombre.
- No cambiar `FnClaveLimpia`: otras consultas dependen de su comportamiento.
- Comprobado con el reporte de almacén: salidas, entradas y OC quedan sin diferencias por insumo.

### El mismo comparativo salía en dos renglones
- El `# CC - Comparativo` llegaba con distinta capitalización según la fuente.
- Regla: todo en mayúsculas, aplicado en `LlavesLimpias` de `BD`, antes de cualquier cruce por comparativo.
- Los textos fijos también quedan en mayúsculas: `SIN CUADRO COMPARATIVO` y `POR ADJUDICAR`.
- `F_Globales` no se tocó: todo lo que entra a `BD` pasa por `LlavesLimpias`. Las fuentes siguen con su capitalización propia.

### Clasificador del asegurado de SINCO
Regla de Juan: contratos y órdenes de compra toman el clasificador solo de los comparativos; lo demás va a `zOtros`.
- Las filas de `CONTRATO` y `COMPRAS` no traen comparativo propio: se enlazan con él por la OC, en las filas de Det_CC.
- Una línea asegurada sin clasificador toma el más frecuente de las filas aseguradas de su misma OC. Si no hay, `zOtros`.
- Las provisiones ya no clasifican a nadie, ni por OC ni por insumo. La fila de provisión conserva su propia etiqueta.
- `TD3_OC` calcula el clasificador de la OC sin contar provisiones.
- El `MaestroInsumos` sigue mandando sobre el insumo, pero no clasifica el resto de la OC.
- Entradas y salidas no siguen esta regla, para no alterar el cálculo de desperdicio. Por eso una OC en `zOtros` puede
  tener entradas en otro clasificador.

## Pendiente para el paso a los demás libros
- Sin medir en ningún otro libro. La prueba de Pamplona se detuvo antes de empezar porque ese libro no tiene la consulta
  `TD3_OC`: quitarla del lote de prueba de los libros que no la tengan.
- Los libros sin `zOtros` hoy (Reserva de Guayacán y Verde Esperanza) tendrán ese renglón nuevo en sus dinámicas.
- `ComparativosContratados` compara el comparativo distinguiendo mayúsculas. Al unificarlas, más filas de `ADJUDICADO`
  coinciden y proyectan 0, lo que puede mover `VT Proyectado Colsubsidio`.
- `CONTRATOS` usa la misma llave por nombre que tenía `COMPRAS`; falta verificar si tiene el mismo problema.
- Corrección 2026-11-04: Versalles sí recibe `BD.m` (`ORA_BD` parte de `BD_BASE = FxGitHub("BD")`). Karakali tiene `BD`
  incrustada, pero `COMPRAS` y `TD3_OC` leen de GitHub, así que esos dos sí le llegan.
- Pedido de Juan, en pausa: una sola versión de las consultas para todos los libros de SINCO, sin carpeta `Turpial/`.
  Pasar el código a `Consultas/` llega a esos libros en su próximo refresco; se hace solo cuando Juan lo autorice.

## 2026-11-04

### El libro volvía a la versión anterior
- Excel abre la ruta sincronizada como documento de SharePoint: `FullName` es la URL y usa su propia caché. No aparece `~$`
  aunque el libro esté abierto, así que esa comprobación no sirve en esta carpeta.
- Reemplazar el archivo por fuera con una copia se perdió 2 veces. Causa probable; falta confirmar que no se repite.
- Método vigente: abrir el libro en Excel, editar ahí y guardar desde Excel (`Save`).
- Las dinámicas con segmentadores de clasificador guardados cambian de total si sus OC cambian de clasificador. En
  Turpial, la de depuración del asegurado quedó guardada solo con un clasificador marcado y por eso fue la única que
  cambió. Revisar en cada libro qué clasificadores quedan marcados y que `zOtros` se vea donde haga falta.

### Todo desde BD, sin modelo de datos (Turpial)
- `TablaDinámica3` y `TD_DESPERDICIOS` ya no usan Power Pivot: leen la tabla BD. Una dinámica del modelo no cambia de
  caché: se borra y se recrea en la misma celda, con su nombre, títulos y segmentadores.
- BD trae ahora las facturas por OC (filas Tipo `FACTURA`) y filas de apoyo. Las demás dinámicas, `TablaDinámica4` y la
  consulta `SINCO` las excluyen.
- Filtrar una provisión sin combinar botones: una OC tiene varias provisiones, así que cada par OC-provisión lleva los
  valores de la OC en columnas aparte (`Aseg prov`, `Comp prov`, `Inv prov`, `Desc prov`). Un campo calculado las usa
  solo cuando el filtro deja fuera las filas reales (`Fila real`). El teórico hace lo mismo con `Teorico fijo`.
- Límite: con varias provisiones marcadas, una OC que tenga ambas se cuenta doble en asegurado y entradas; con varios
  contratistas de almacén marcados, el teórico se multiplica. Con una sola selección da igual que el modelo de datos.
- `SINCO` está incrustada en el libro y lee la tabla BD (`SINCO.m` en esta carpeta es copia de registro). Así el
  refresco baja de unos 150 s a unos 65 s porque BD se calcula una vez.
- Al unificar las consultas: estas columnas y filas llegarían a los otros libros, y sus dinámicas tendrían que excluir
  `FACTURA` igual que Turpial.
- Las filas de apoyo ya no llevan un centro de costos inventado: usan el de su OC y, el teorico, el de su fila; asi BD solo tiene centros reales y varios conviven bien. Sin el valor `(FACTURA)` ya no hace falta ocultar nada en la dinamica por centro de costos.
- Bajo filtro de provision, asegurado, entradas y comparativos viven en la fila `(sin factura)`: si se quita esa fila del filtro `Factura`, esas columnas quedan en 0 aunque el bruto siga.
