// ORA_ASEGURADO: ASEGURADO.xls de Oracle normalizado (una fila por registro y articulo).
// Proceso: CONTRATO, ORDEN DE COMPRA, PRESUPUESTO, GESTION RECURSOS HUMANOS, CAJA MENOR,
// SERVICIO PUBLICO, COSTOS DISTRIBUIBLES, DESCUENTOS, TRANSFERENCIA.
// "V.r Unitrio" (asi viene) no se usa: el informe la trae rota; se usa V.r Total.
let
    Fuentes = ORA_FUENTES,
    Arbol = ORA_ARBOL,
    FnNum = Fuentes[FnNum],
    Limpiar = Arbol[Limpiar],
    T = Fuentes[ASEGURADO],
    FnCol = (nombre as text) as list =>
        if Table.HasColumns(T, nombre) then Table.Column(T, nombre)
        else error Error.Record("Columna no encontrada", "ASEGURADO.xls: falta la columna '" & nombre & "'", "Columnas: " & Text.Combine(Table.ColumnNames(T), " | ")),
    // la segunda "Descripción" (la del articulo) llega renombrada por Power Query
    ColDescArt = List.First(List.Select(Table.ColumnNames(T), each Text.StartsWith(_, "Descripci") and _ <> "Descripción" and _ <> "Descripcion"), null),
    Base = Table.FromColumns({
        List.Transform(FnCol("Paquete de Trabajo"), Limpiar),
        List.Transform(FnCol("Cod CBS"), Limpiar),
        List.Transform(FnCol("Articulo"), Limpiar),
        if ColDescArt = null then List.Repeat({""}, Table.RowCount(T)) else List.Transform(Table.Column(T, ColDescArt), Limpiar),
        List.Transform(FnCol("U Medida"), Limpiar),
        List.Transform(FnCol("Proceso"), (v) => Text.Upper(Limpiar(v))),
        List.Transform(FnCol("Registro"), Limpiar),
        List.Transform(FnCol("Cantidad"), FnNum),
        List.Transform(FnCol("V.r Total"), FnNum)
    }, {"Paquete", "CBS", "Articulo", "Desc Articulo", "UM", "Proceso", "Registro", "Cantidad", "VT"}),
    Resultado = Table.Buffer(Table.SelectRows(Base, each [Proceso] <> "" and [Proceso] <> "PROCESO"))
in
    Resultado
