
let
    Norm = (v as any) as text => let t = try Text.Upper(Text.Trim(F_Globales[FnRemoveAccentMarks](Text.From(v)))) otherwise null
        in if t = null then "" else Text.Combine(List.Select(Text.Split(t, " "), each _ <> ""), " "),
    BDGuardada = Excel.CurrentWorkbook(){[Name="BD"]}[Content],
    Fuente = Table.SelectColumns(Table.SelectRows(BDGuardada, each [Tipo] = "ITEMS"), {"Centro de Costos", "Codigo ins", "Ins", "Clasificador"}),
    Reales = Table.SelectRows(Fuente, each Norm([Centro de Costos]) <> "" and Norm([Ins]) <> ""),
    Catalogo = Table.Group(Reales, {"Centro de Costos", "Codigo ins", "Ins"}, {
        {"__ClasesBD", each List.Sort(List.Distinct(List.Select(List.Transform([Clasificador], (v) => if v = null then "" else Text.Trim(Text.From(v))), (v) => v <> "" and v <> "zOtros"))), type list}}),
    Manual = Excel.CurrentWorkbook(){[Name="MaestroInsumos"]}[Content],
    Reglas = Table.Group(Table.AddColumn(Table.SelectRows(Manual, each Norm([Ins]) <> "" and Norm([Clasificador]) <> ""), "__Ins", each Norm([Ins]), type text), {"__Ins"},
        {{"__Clases", each List.Distinct(List.Transform([Clasificador], Text.Trim)), type list}}),
    Validas = Table.SelectRows(Reglas, each List.Count([__Clases]) = 1),
    Mapa = Record.FromList(List.Transform(Validas[__Clases], each _{0}), Validas[__Ins]),
    Clasificado = Table.AddColumn(Catalogo, "Clasificador manual", each Record.FieldOrDefault(Mapa, Norm([Ins]), null), type nullable text),
    ClaseBD = Table.AddColumn(Clasificado, "Clasificador", each Text.Combine([__ClasesBD], " | "), type text),
    Estado = Table.AddColumn(ClaseBD, "Estado", each if List.Count([__ClasesBD]) > 1 then "REVISAR CLASIFICADORES"
        else if List.IsEmpty([__ClasesBD]) then "PENDIENTE DE CLASIFICAR" else "CLASIFICADO", type text),
    Final = Table.Sort(Table.SelectColumns(Estado, {"Centro de Costos", "Codigo ins", "Ins", "Clasificador", "Clasificador manual", "Estado"}),
        {{"Centro de Costos", Order.Ascending}, {"Ins", Order.Ascending}, {"Codigo ins", Order.Ascending}})
in
    Final
