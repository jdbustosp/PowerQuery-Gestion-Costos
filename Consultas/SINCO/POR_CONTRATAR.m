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
    // 2026-11-24: tambien los aprobados sin filas ADJUDICADO (antes COMPARATIVOS_POR_SUBIR), mismo umbral.
    // Valor = saldo (aprobado - asegurado - suelto). Contratacion: Parcial si ya hay algo contratado.
    N = (v as any) as number => if v = null then 0 else try Number.From(v) otherwise 0,
    Visibles = Table.SelectRows(PC[Grupos], each [Mostrar] or
        (N(Record.FieldOrDefault(_, "PptoAdj", 0)) = 0 and [Valor] >= 1000000 and [Valor] >= 0.02 * N([Aprobado]))),
    ConTipo = Table.AddColumn(Visibles, "Contratacion", each
        if N([Asegurado]) + N(Record.FieldOrDefault(_, "Suelto", 0)) <> 0 then "Parcial" else "Todo por subir", type text),
    Final = Table.Sort(Table.SelectColumns(ConTipo, {"Centro de Costos", "Comparativo", "Valor", "Estado", "Contratacion"}),
        {{"Centro de Costos", Order.Ascending}, {"Comparativo", Order.Ascending}}),
    Tipos = Table.TransformColumnTypes(Final, {{"Centro de Costos", type text}, {"Comparativo", type text},
        {"Valor", type number}, {"Estado", type text}, {"Contratacion", type text}})
in
    Tipos
