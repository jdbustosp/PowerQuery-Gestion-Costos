let
    // =========================================================
    // Parseo unificado Seguimiento + APU. La logica vive en
    // F_Globales[FxProcesarCentroCosto] para no duplicarse con
    // PPTO_TODOS_PROYECTOS.
    // =========================================================
    FxProcesarCentroCosto = F_Globales[FxProcesarCentroCosto],

    ConCentroCosto = SP_Archivos_Proyecto,
    Agrupado = Table.Group(ConCentroCosto, {"Centro de Costos"}, {{"Binarios", each let
        FilaPres = Table.SelectRows(_, each Text.Contains([Name], "ANALISIS DE PRECIOS UNITARIOS", Comparer.OrdinalIgnoreCase)),
        FilaSeg = Table.SelectRows(_, each Text.Contains([Name], "SEGUIMIENTO POR ITEMS", Comparer.OrdinalIgnoreCase))
        in if Table.RowCount(FilaPres) > 0 and Table.RowCount(FilaSeg) > 0
        then [Bin_P = Binary.Buffer(FilaPres{0}[Content]), Bin_S = Binary.Buffer(FilaSeg{0}[Content])]
        else null
    }}),
    CentrosCompletos = Table.SelectRows(Agrupado, each [Binarios] <> null),
    TablaConDatos = Table.AddColumn(CentrosCompletos, "Datos", each FxProcesarCentroCosto([Binarios][Bin_S], [Binarios][Bin_P])),
    Expandido = Table.ExpandTableColumn(TablaConDatos, "Datos", {"Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo", "Cantidad Presupuesto", "VT Presupuesto", "Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido"}),
    ColumnasUtiles = Table.SelectColumns(Expandido, {"Centro de Costos", "Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo", "Cantidad Presupuesto", "VT Presupuesto", "Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido"}),
    TiposFinales = Table.TransformColumnTypes(ColumnasUtiles,{{"Centro de Costos", type text}, {"Codigo ins", Int64.Type}, {"Cantidad Presupuesto", type number}, {"VT Presupuesto", Currency.Type}, {"Cantidad Proyectado", type number}, {"VT Proyectado", Currency.Type}, {"Cantidad Consumido", type number}, {"VT Consumido", Currency.Type}}),
    // BUFFER MAESTRO: ITEMSINSUMOS y PPTO_BD comparten este resultado
    Resultado = Table.Buffer(TiposFinales)
in
    Resultado
