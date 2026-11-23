let
    // ============================================================
    // POR_CONTRATAR (2026-11-23): comparativos aprobados con saldo sin contratar en SINCO
    // (aprobado CC CONSOLIDADO menos asegurado de las OC ligadas a sus CC). Lee la tabla BD
    // ya cargada (refrescar BD antes). Estado: Abierto (defecto) / Cerrado, lo pone el
    // usuario en esta misma tabla y se conserva al refrescar (lo lee FnPorContratar).
    // Cerrado = no se contrata nada mas: el ADJUDICADO de ese comparativo proyecta 0.
    // ============================================================
    BD0 = try Excel.CurrentWorkbook(){[Name = "BD"]}[Content] otherwise #table({"Tipo"}, {}),
    PC = F_Globales[FnPorContratar](BD0),
    Visibles = Table.SelectRows(PC[Grupos], each [Mostrar]),
    Final = Table.Sort(Table.SelectColumns(Visibles, {"Centro de Costos", "Comparativo", "Valor", "Estado"}),
        {{"Centro de Costos", Order.Ascending}, {"Comparativo", Order.Ascending}}),
    Tipos = Table.TransformColumnTypes(Final, {{"Centro de Costos", type text}, {"Comparativo", type text},
        {"Valor", type number}, {"Estado", type text}})
in
    Tipos
