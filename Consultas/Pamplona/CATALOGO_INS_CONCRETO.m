let
    Norm = (v as any) as text => if v = null then "" else Text.Upper(Text.Trim(Text.From(v))),
    Catalogo = CATALOGO_INSUMOS_CC,
    Reales = Table.SelectRows(Catalogo, each
        (Text.StartsWith(Norm([Ins]), "CONCRETO ") or Text.StartsWith(Norm([Ins]), "M_CONCRETO ")) and
        Text.EndsWith(Norm([Ins]), "(M3)")),
    Final = Table.Sort(Table.Distinct(Table.SelectColumns(Reales, {"Ins"}), {"Ins"}), {{"Ins", Order.Ascending}})
in
    Final
