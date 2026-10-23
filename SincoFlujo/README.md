# SincoFlujo

Copia independiente del código M que usan los libros `SINCO FLUJO.xlsx` (flujo SINCO de Power Automate Desktop).
En esos libros, `FxGitHub` lee `SincoFlujo/<consulta>.m`; los libros "Control interno" siguen leyendo `Consultas/`.

- Un cambio aquí solo afecta a los SINCO FLUJO; un cambio en `Consultas/` ya no les llega.
- Separada de `Consultas/` el 2026-10-21 a partir del commit 5f91a00.
- Si se quiere llevar una mejora de `Consultas/` a los SINCO FLUJO, copiar el `.m` aquí a propósito.

## Consultas de las hojas del flujo (2026-10-22)
Leen la carpeta Actual del centro de costos (`SP_Archivos_Flujo`, parametro `CentroCostoActual` = CONFIG FLUJO!K5)
y conservan entre refrescos lo que escribe el flujo (`FxConservarEstado`).

| Consulta | Hoja | Fuente |
|---|---|---|
| PLANTILLA_PPTO_SINCO / _ORACLE | PLANTILLA SINCO | SEGUIMIENTO + APU |
| ACTIVIDADES_NUEVAS | ACTIVIDADES NUEVAS | SEGUIMIENTO vs APU |
| INSUMOS_NUEVOS | INSUMOS NUEVOS | SEGUIMIENTO vs APU |
| PROYECCION_INSUMOS | PROYECCION INSUMOS | SEGUIMIENTO (proyeccion distinta del presupuesto) |
| CONTRATOS_SINCO | CONTRATOS | ESTADO DE CONTRATOS + SEGUIMIENTO + APU |
| ORDENES_COMPRA | ORDENES COMPRA | INFORMEORDEN + ESTADO DE ORDENES |
| ASIGNACION_OC | ASIGNACION | INFORMEORDEN + ESTADO DE ORDENES |
