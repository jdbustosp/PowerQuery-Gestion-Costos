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
| PLANTILLA_PPTO_ACTUALIZADO | PLANTILLA SINCO | presupuesto actualizado de "0. Descargas pptos" (en lugar del ERP del constructor) |
| ACTIVIDADES_NUEVAS | ACTIVIDADES NUEVAS | SEGUIMIENTO vs APU |
| INSUMOS_NUEVOS | INSUMOS NUEVOS | SEGUIMIENTO vs APU |
| PROYECCION_INSUMOS | PROYECCION INSUMOS | SEGUIMIENTO (proyeccion distinta del presupuesto) |
| CONTRATOS_SINCO | CONTRATOS | ESTADO DE CONTRATOS + SEGUIMIENTO + APU |
| ORDENES_COMPRA | ORDENES COMPRA | INFORMEORDEN + ESTADO DE ORDENES |
| ASIGNACION_OC | ASIGNACION | INFORMEORDEN + ESTADO DE ORDENES |
| ENTRADAS (tabla ENTRADAS_SINCO) | ENTRADAS | Informe masivo de entradas + Facturacion de entradas y actas |
| ACTAS (tabla ACTAS_SINCO) | ACTAS | Detallado actas + Informe relacion de pagos (factura); cantidad por valor cuando el grupo junta varias lineas |

## Limitación conocida (2026-11-05): códigos de actividad repetidos

Códigos de actividad repetidos en dos subcapítulos (2026-11-05): `PLANTILLA_PPTO_SINCO` trabaja por OCURRENCIA. El
subcapítulo se cruza por código + descripción normalizada + orden; la suma de insumos y la exclusión de actividades
en cero van por la fila de la actividad (antes iban por código: valores inflados, subcapítulo de la primera y
ocurrencias borradas). Sin códigos repetidos el resultado es idéntico al anterior. Control: la suma de `Vr Total`
debe igualar el TOTAL PROYECTO de SEGUIMIENTO. Pendiente fuera del M: el importador de SINCO enlaza los insumos por
código y deja los de todas las ocurrencias en una sola. Por eso (paso 5b) la primera ocurrencia con cantidad
conserva su código y las siguientes reciben `<capítulo>.901`, `.902`… (orden del APU, saltando códigos usados);
sus insumos heredan el código nuevo. Paso 7b: actividades con igual descripción y distinto valor reciben
` (<código>)` al final, porque SINCO parece compartir el APU por descripción. Sin repetidos, sin cambios.
Pendiente: las consultas de proyección, contratos, asignación y salidas siguen usando el código de origen.
