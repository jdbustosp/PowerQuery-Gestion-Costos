
let
    Norm = (v as any) as text => if v = null then "" else Text.Upper(Text.Trim(Text.From(v))),
    Catalogo = CATALOGO_INSUMOS_CC,
    Reales = Table.SelectRows(Catalogo, each Norm([Clasificador manual]) = "CONCRETO" or Norm([Clasificador]) = "CONCRETO" or
        ((Text.StartsWith(Norm([Ins]), "CONCRETO ") or Text.StartsWith(Norm([Ins]), "M_CONCRETO ")) and Text.EndsWith(Norm([Ins]), "(M3)"))),
    Manual = Excel.CurrentWorkbook(){[Name="MaestroInsumos"]}[Content],
    Referencias = Table.SelectRows(Manual, each Norm([Clasificador]) = "CONCRETO" and Norm([Ins]) <> ""),
    Nombres = Table.Combine({Table.SelectColumns(Reales, {"Ins"}), Table.SelectColumns(Referencias, {"Ins"})}),
    Final = Table.Sort(Table.Distinct(Nombres, {"Ins"}), {{"Ins", Order.Ascending}})
in
    Final
