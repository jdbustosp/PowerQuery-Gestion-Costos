let
    // =========================================================
    // Parseo unificado Seguimiento + APU. La logica vive en
    // F_Globales[FxProcesarCentroCosto] para no duplicarse con
    // PPTO_TODOS_PROYECTOS.
    // =========================================================
    FxProcesarCentroCosto = F_Globales[FxProcesarCentroCosto],
    FnReadSPBinary = F_Globales[FnReadSPBinary],
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",

    ConCentroCosto = SP_Archivos_Proyecto,

    PickLatestBinary = (t as table, containsText as text) as nullable binary =>
        let
            candidatos = Table.Sort(
                Table.SelectRows(t, each Text.Contains([Name], containsText, Comparer.OrdinalIgnoreCase)),
                {{"TimeLastModified", Order.Descending}, {"Name", Order.Ascending}}
            ),
            path = if Table.RowCount(candidatos) = 0 then null else candidatos{0}[ServerRelativeUrl]
        in
            if path = null then null else FnReadSPBinary(SiteUrl, path),

    Agrupado = Table.Group(ConCentroCosto, {"Centro de Costos"}, {{"Binarios", each
        let
            binPres = PickLatestBinary(_, "ANALISIS DE PRECIOS UNITARIOS"),
            binSeg = PickLatestBinary(_, "SEGUIMIENTO POR ITEMS")
        in
            if binPres <> null and binSeg <> null then [Bin_P = binPres, Bin_S = binSeg] else null
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
