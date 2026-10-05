let
    // TD3_DIMFACT (2026-10-05): una fila por OC + factura (nivel 4 de la dinamica de
    // contratos/OC). Une las facturas de SINCO y las de provisiones; el texto que se muestra
    // es el de SINCO cuando existe. Cada OC tiene ademas su fila "(sin factura)".
    DeSinco = Table.AddColumn(Table.SelectColumns(TD3_FACTURAS, {"FactKey", "OC", "Factura"}), "Orden", each 1, Int64.Type),
    DeProv = Table.AddColumn(Table.SelectColumns(TD3_PROV, {"FactKey", "OC", "Factura"}), "Orden", each 2, Int64.Type),
    SinFactura = Table.FromRecords(List.Transform(TD3_OC[OC], (oc) => [FactKey = oc & "|SIN", OC = oc, Factura = "(sin factura)", Orden = 3]),
        type table [FactKey = text, OC = text, Factura = text, Orden = Int64.Type]),
    Todo = Table.Combine({DeSinco, DeProv, SinFactura}),
    Final = Table.Group(Todo, {"FactKey", "OC"}, {{"Factura", each Table.Min(_, "Orden")[Factura], type text}})
in
    Final
