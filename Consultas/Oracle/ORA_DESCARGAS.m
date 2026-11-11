// Lee y normaliza la tabla DESCARGA del libro del proyecto. La clave del insumo puede ser alfanumérica.
let
    Columnas = {
        "Proyecto", "Centro de Costos", "Subcapitulo", "Capitulo", "Actividad", "Codigo ins", "Ins",
        "Cantidad ppto (CC)", "V/U ppto (CC)", "Valor Total ppto (CC)",
        "# CC - Comparativo", "# CC", "Comparativo"
    },
    Libro = ORA_DESCARGAS_LIBRO,
    Origen = Libro{[Item = "DESCARGA", Kind = "Table"]}[Data],
    Seleccion = Table.SelectColumns(Origen, Columnas),
    Texto = (v as any) as nullable text => if v = null then null else Text.Trim(Text.From(v)),
    ConTextos = Table.TransformColumns(Seleccion,
        List.Transform(List.Difference(Columnas, {"Cantidad ppto (CC)", "V/U ppto (CC)", "Valor Total ppto (CC)"}),
            (c) => {c, Texto, type nullable text})),
    Proyecto = Text.Upper(Text.Trim(Text.From(ProyectoActual))),
    FilasProyecto = Table.SelectRows(ConTextos, each
        [Proyecto] <> null and Text.Upper([Proyecto]) = Proyecto and
        [Centro de Costos] <> null and [Centro de Costos] <> ""),
    FnNum = ORA_FUENTES[FnNum],
    ConNumeros = Table.TransformColumns(FilasProyecto, {
        {"Cantidad ppto (CC)", FnNum, type nullable number}, {"V/U ppto (CC)", FnNum, type nullable number},
        {"Valor Total ppto (CC)", FnNum, type nullable number}}),
    // Construir el comparativo desde las dos entradas manuales evita depender del cache de una formula.
    SinComp = Table.RemoveColumns(ConNumeros, {"# CC - Comparativo"}),
    ConComp = Table.AddColumn(SinComp, "# CC - Comparativo", each
        if [#"# CC"] = null or [#"# CC"] = "" or [Comparativo] = null or [Comparativo] = ""
        then null else [#"# CC"] & " - " & [Comparativo], type nullable text),
    Resultado = Table.Buffer(Table.ReorderColumns(ConComp, Columnas))
in
    Resultado
