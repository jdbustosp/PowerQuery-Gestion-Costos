
let
    Norm = (v as any, cc as text) as text =>
        let t0 = F_Globales[FnNormalizeComparativo](v),
            t = if t0 <> null and cc <> "" and Text.StartsWith(t0, cc & "-") then F_Globales[FnNormalizeComparativo](Text.AfterDelimiter(t0, "-")) else t0
        in if t = null then "" else Text.Upper(F_Globales[FnRemoveAccentsSymbols](t)),
    Txt = (v as any) as text => if v = null then "" else F_Globales[FnTrimText](v),
    Num = (v as any) as number => if v = null then 0 else Number.From(v),
    Centro = (v as any) as text => let t = Txt(v), p = Text.BeforeDelimiter(t & "-", "-") in Text.Trim(p),
    Aprob = Table.AddColumn(Table.AddColumn(APROBACIONES_SP, "__CC", each Centro([Centro de Costos]), type text),
        "__Comp", each Norm([#"# CC - Comparativo"], [__CC]), type text),
    Grupos = Table.Group(Table.SelectRows(Aprob, each [__Comp] <> ""), {"__CC", "__Comp"}, {
        {"Centro de Costos", each List.First([Centro de Costos]), type nullable text},
        {"# CC - Comparativo", each List.First([#"# CC - Comparativo"]), type text},
        {"Clasificador Approvals", each Text.Combine(List.Sort(List.Distinct(List.Select(List.Transform([Clasificador], Txt), each _ <> ""))), " | "), type text},
        {"Valor aprobado", each List.Sum(List.Transform([VT CC cons], Num)), type number}}),
    Desc = DESCARGA,
    ConClave = Table.AddColumn(Table.AddColumn(Desc, "__CC", each Centro([Centro de Costos]), type text),
        "__Comp", each Norm([#"# CC - Comparativo"], [__CC]), type text),
    Registrados = Table.Distinct(Table.SelectColumns(Table.SelectRows(ConClave, each [__Comp] <> ""), {"__CC", "__Comp"})),
    Pendientes = Table.NestedJoin(Table.SelectRows(Grupos, each [Valor aprobado] > 0), {"__CC", "__Comp"}, Registrados, {"__CC", "__Comp"}, "__Desc", JoinKind.LeftAnti),
    ConProyecto = Table.AddColumn(Pendientes, "Proyecto", each Text.Trim(ProyectoActual), type text),
    ConEstado = Table.AddColumn(ConProyecto, "Estado", each if [__CC] = "" then "REVISAR CENTRO DE COSTOS" else "APROBADO SIN REGISTRO EN DESCARGAS", type text),
    Final = Table.Sort(Table.SelectColumns(ConEstado, {"Proyecto", "Centro de Costos", "# CC - Comparativo", "Clasificador Approvals", "Valor aprobado", "Estado"}),
        {{"Centro de Costos", Order.Ascending}, {"# CC - Comparativo", Order.Ascending}})
in
    Final
