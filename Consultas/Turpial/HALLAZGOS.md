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
- Versalles usa `ORA_BD` y Karakali está congelado: el cambio de `BD.m` no les llega.
