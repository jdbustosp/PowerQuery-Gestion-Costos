# Consultas/SINCO — formato único de Control interno SINCO

Fuente única para todos los libros de Control interno SINCO con el formato vigente (MEMORIAS por bloques,
ASIGNACION, catálogo por CC, Approvals clasificado). Cada consulta del libro es una línea:

```
FxGitHub("SINCO/<nombre>")
```

con `FxGitHub` apuntando a `refs/heads/main/Consultas/`. Un cambio aquí llega a TODOS esos libros en su
próximo refresco: probar en copia de al menos dos libros antes de subir.

| Consulta del libro | Archivo | Origen (15/11/2026) |
|---|---|---|
| BD | `BD.m` | clasificador común (66a8c94) |
| BD_Origen | `BD_Origen.m` | `Pamplona/AjustesTexto/BD.m` |
| CACHE_SEGUIMIENTO, CATALOGO_INSUMOS_CC, COMPARATIVOS, DESCARGAS, DISPONIBLE, F_Globales | mismo nombre | `Pamplona/AjustesTexto/` |
| CATALOGO_INS_CONCRETO, FX_INSUMO_CC, SINCO, SP_Seguimiento_Parsed | mismo nombre | `Pamplona/` |
| COMPRAS, TEORICO_TODO, TEORICO_MAMPOSTERIA | mismo nombre | `Turpial/` |

Las demás consultas (APROBACIONES_SP, CONTRATOS, DESCUENTOS, ITEMSINSUMOS, PPTO_BD, PROVISIONES_SP,
SINCO_FACTURAS, SP_*) siguen en la raíz de `Consultas/`, comunes a todos los libros SINCO.

`Pamplona/`, `Pamplona/AjustesTexto/` y `Turpial/` quedan congeladas hasta que ningún libro las lea; no
editarlas. Lo propio de cada proyecto vive en el libro (ProyectoActual, tablas de captura), nunca aquí.
