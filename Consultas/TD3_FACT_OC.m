let
    // TD3_FACT_OC (2026-10-05): valores por contrato/OC para la dinamica de contratos/OC.
    // Cada columna de BD solo tiene valor en su tipo de fila: VT Asegurada y VT Cortes en
    // CONTRATO/COMPRAS, VR total aprobacion en CC (Det_CC), Valor descuento en DESCUENTO.
    FnOC = (v as any) as text => let t = try Text.Trim(Text.From(v)) otherwise null in if t = null then "" else Text.Select(t, {"0".."9"}),
    FnNum = (v as any) as number => let n = try Number.From(v) otherwise null in if n = null then 0 else n,
    Base = Table.SelectRows(
        // Lee la TABLA BD ya cargada en la hoja (ver TD3_OC)
        Table.AddColumn(Table.SelectColumns(Excel.CurrentWorkbook(){[Name = "BD"]}[Content], {"# OC / Contrato", "VT Asegurada", "VR total aprobacion", "VT Cortes", "Valor descuento"}), "OC", each FnOC([#"# OC / Contrato"]), type text),
        each [OC] <> ""),
    PorOC = Table.Group(Base, {"OC"}, {
        {"Asegurado", each List.Sum(List.Transform([VT Asegurada], FnNum)), type number},
        {"Comparativos", each List.Sum(List.Transform([VR total aprobacion], FnNum)), type number},
        {"Invertido", each List.Sum(List.Transform([VT Cortes], FnNum)), type number},
        {"Descuentos", each List.Sum(List.Transform([Valor descuento], FnNum)), type number}}),
    SoloOC = Table.Join(PorOC, {"OC"}, Table.SelectColumns(TD3_OC, {"OC"}), {"OC"}, JoinKind.Inner)
in
    SoloOC
