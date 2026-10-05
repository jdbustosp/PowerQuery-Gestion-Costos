let
    // TD3_FACTURAS (2026-10-05): invertido por factura (actas y entradas facturadas de SINCO)
    // para el nivel 4 de la dinamica de contratos/OC. Por cada OC se agrega una fila
    // "(sin factura)" con lo invertido que aun no tiene factura (o la diferencia contra
    // VT Cortes), para que las facturas siempre sumen el invertido de la OC.
    Fact = Table.SelectRows(SINCO_FACTURAS, each [OC] <> ""),
    PorFactura = Table.Group(Fact, {"OC", "FacturaClave"}, {
        {"Factura", each List.First(List.Select([Factura], each _ <> "" and _ <> null), "(sin factura)"), type text},
        {"Invertido", each List.Sum([Valor]), type number}}),
    ConLlave = Table.AddColumn(PorFactura, "FactKey", each [OC] & "|" & (if [FacturaClave] = "" then "SIN" else [FacturaClave]), type text),
    SoloOC = Table.Join(ConLlave, {"OC"}, Table.SelectColumns(TD3_OC, {"OC"}), {"OC"}, JoinKind.Inner),
    // Cuadre contra el invertido de la OC (VT Cortes de BD)
    SumaFact = Table.Group(SoloOC, {"OC"}, {{"Facturado", each List.Sum([Invertido]), type number}}),
    Cuadre = Table.ExpandTableColumn(
        Table.NestedJoin(Table.SelectColumns(TD3_FACT_OC, {"OC", "Invertido"}), {"OC"}, SumaFact, {"OC"}, "F", JoinKind.LeftOuter),
        "F", {"Facturado"}),
    Saldo = Table.SelectRows(
        Table.AddColumn(Cuadre, "Dif", each [Invertido] - (if [Facturado] = null then 0 else [Facturado]), type number),
        each Number.Abs([Dif]) >= 1),
    FilasSaldo = Table.FromRecords(List.Transform(Table.ToRecords(Saldo), (r) =>
        [OC = r[OC], FacturaClave = "", Factura = "(sin factura)", Invertido = r[Dif], FactKey = r[OC] & "|SIN"]),
        type table [OC = text, FacturaClave = text, Factura = text, Invertido = number, FactKey = text]),
    Todo = Table.Combine({Table.SelectColumns(SoloOC, {"OC", "FacturaClave", "Factura", "Invertido", "FactKey"}), FilasSaldo}),
    Final = Table.Group(Todo, {"FactKey", "OC"}, {
        {"Factura", each List.First([Factura]), type text},
        {"Invertido", each List.Sum([Invertido]), type number}})
in
    Final
