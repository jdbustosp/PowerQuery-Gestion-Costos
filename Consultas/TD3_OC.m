let
    // TD3_OC (2026-10-05): dimension del modelo de datos de la dinamica de contratos/OC.
    // Una fila por contrato u orden de compra de SINCO (filas CONTRATO/COMPRAS de BD).
    // Tipo, centro de costos, clasificador y contratista son UNO por OC y se aplican a todas
    // sus filas (comparativos, provisiones, descuentos): asi un segmentador trae la OC completa.
    FnTexto = (v as any) as text => let t = try Text.Trim(Text.From(v)) otherwise null in if t = null then "" else t,
    FnOC = (v as any) as text => Text.Select(FnTexto(v), {"0".."9"}),
    FnModa = (l as list) as nullable text =>
        let x = List.Select(List.Transform(l, FnTexto), each _ <> "") in if List.IsEmpty(x) then null else List.Mode(x),
    Base = Table.Buffer(Table.SelectRows(
        // Lee la TABLA BD ya cargada en la hoja: recalcular la consulta BD por cada tabla
        // del modelo sumaba ~9 min a cada actualizacion.
        Table.AddColumn(Table.SelectColumns(Excel.CurrentWorkbook(){[Name = "BD"]}[Content], {"Centro de Costos", "Tipo", "# OC / Contrato", "Nombre Contratista", "Clasificador"}), "OC", each FnOC([#"# OC / Contrato"]), type text),
        each [OC] <> "")),
    Sinco = Table.SelectRows(Base, each [Tipo] = "CONTRATO" or [Tipo] = "COMPRAS"),
    PorOC = Table.Group(Sinco, {"OC"}, {
        {"Tipo", each if List.Contains([Tipo], "CONTRATO") then "Contrato" else "Orden de compra", type text},
        {"Centro de Costos", each FnModa([Centro de Costos]), type text},
        {"Contratista", each FnModa([Nombre Contratista]), type text}}),
    // Clasificador: el mas frecuente entre TODAS las filas de la OC (BD ya lo trae desde aprobaciones)
    ClasifOC = Table.Group(Base, {"OC"}, {{"Clasificador", each FnModa([Clasificador]), type text}}),
    Unido = Table.ExpandTableColumn(Table.NestedJoin(PorOC, {"OC"}, ClasifOC, {"OC"}, "C", JoinKind.LeftOuter), "C", {"Clasificador"}),
    Final = Table.TransformColumns(Unido, {
        {"Clasificador", each if _ = null then "(sin clasificador)" else _, type text},
        {"Contratista", each if _ = null then "(sin contratista)" else _, type text}})
in
    Final
