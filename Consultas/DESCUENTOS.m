let
    // ============================================================
    // FUNCIONES AUXILIARES GLOBALES
    // ============================================================
    FnFormatCodigoAct = F_Globales[FnFormatCodigoAct],
    FxToNumberFlex = F_Globales[FxToNumberFlex],

    // ============================================================
    // FUNCIÓN MÁGICA: PROCESAR DESCUENTOS
    // ============================================================
    FxProcesarDescuentos = (Binario as binary) =>
        let
            HtmlTexto = Text.FromBinary(Binario, 1252),
            Columnas_HTML = List.Transform({1..15}, each {"Columna" & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}),
            RawTable = Html.Table(HtmlTexto, Columnas_HTML, [RowSelector="tr"]),
            
            FilasLimpias = Table.SelectRows(RawTable, each [Columna1] <> "GRAN TOTAL" and [Columna1] <> "DESCUENTOS SALIDAS" and [Columna1] <> null),
            FillContrato = Table.FillDown(Table.AddColumn(FilasLimpias, "ContratoInfo", each let txt = Text.Trim(Text.From([Columna1] ?? "")) in if Text.StartsWith(Text.Upper(txt), "CONTRATO") then txt else null), {"ContratoInfo"}),
            AddOCContrato = Table.AddColumn(FillContrato, "# OC / Contrato", each let raw = Text.Select(Text.From([ContratoInfo] ?? ""), {"0".."9"}) in if raw = "" then null else raw, type text),
            AddCodigoAct = Table.AddColumn(AddOCContrato, "Codigo act", each let txt = Text.Trim(Text.From([Columna3] ?? "")), baseCod = if Text.Contains(txt, " ") then Text.BeforeDelimiter(txt, " ") else txt in if baseCod = "" then null else FnFormatCodigoAct(baseCod), type text),
            
            AddValorDescuento = Table.AddColumn(AddCodigoAct, "Valor descuento", each let rawTxt = Text.Remove(Text.Trim(Text.From([Columna7] ?? "")), {"$", " "}) in FxToNumberFlex(rawTxt), Currency.Type),
            
            BaseFinal = Table.SelectColumns(Table.SelectRows(AddValorDescuento, each [#"# OC / Contrato"] <> null and [Codigo act] <> null and [Valor descuento] <> null and [Valor descuento] <> 0), {"# OC / Contrato", "Codigo act", "Valor descuento"})
        in BaseFinal,

    // ============================================================
    // CONEXIÓN A SHAREPOINT (LECTURA DESDE CONSULTA COMPARTIDA)
    // ============================================================
    ArchivosProyecto = Table.SelectRows(SP_Archivos_Proyecto, each 
        Text.Contains([Name], "DESCUENTOS", Comparer.OrdinalIgnoreCase)
    ),
    ConCentroCosto = ArchivosProyecto,
    
    // 🔥 EL SALVAVIDAS
    Agrupado = Table.Group(ConCentroCosto, {"Centro de Costos"}, {{"Binario", each Binary.Buffer(_{0}[Content])}}),
    TablaConDatos = Table.AddColumn(Agrupado, "Datos", each FxProcesarDescuentos([Binario])),
    
    SinBinario = Table.RemoveColumns(TablaConDatos, {"Binario"}),
    Expandido = Table.ExpandTableColumn(SinBinario, "Datos", {"# OC / Contrato", "Codigo act", "Valor descuento"}),
    
    Descuentos_Clean = Table.TransformColumns(Expandido, {
        {"Centro de Costos", each if _ = null then null else Text.Upper(Text.Trim(Text.From(_))), type text},
        {"Codigo act", each FnFormatCodigoAct(_), type text},
        {"# OC / Contrato", each if _ = null then null else Text.Trim(Text.From(_)), type text}
    }, null, MissingField.Ignore),
    
    BaseDescuentos_EnMemoria = Descuentos_Clean,

    // ============================================================
    // LECTURA DIRECTA DE CONSULTAS (Memoria)
    // ============================================================
    SourceContratos = CONTRATOS,
    CONTRATOS_Clean = Table.TransformColumns(SourceContratos, {
        {"# OC / Contrato", each if _ = null then null else Text.Trim(Text.From(_)), type text}
    }, null, MissingField.Ignore),
    ContratosPorOC = Table.Group(CONTRATOS_Clean, {"# OC / Contrato"}, {{"Nombre Contratista", each List.First([Nombre Contratista]), type text}, {"Descripcion contrato", each List.First([Descripcion contrato]), type text}}),

    SourceItems = ITEMSINSUMOS,
    ITEMS_Clean = Table.TransformColumns(SourceItems, {
        {"Codigo act", each FnFormatCodigoAct(_), type text}
    }, null, MissingField.Ignore),
    ItemsPorCodigo = Table.Buffer(Table.Group(ITEMS_Clean, {"Codigo act"}, {{"Actividad", each List.First([Actividad]), type text}, {"Capitulo", each List.First([Capitulo]), type text}, {"Subcapitulo", each List.First([Subcapitulo]), type text}})),

    // ============================================================
    // CRUCES FINALES Y SELECCIÓN ESTRICTA
    // ============================================================
    MergeContratos = Table.NestedJoin(BaseDescuentos_EnMemoria, {"# OC / Contrato"}, ContratosPorOC, {"# OC / Contrato"}, "C", JoinKind.LeftOuter),
    ExpandContratos = Table.ExpandTableColumn(MergeContratos, "C", {"Nombre Contratista", "Descripcion contrato"}, {"Nombre Contratista", "Descripcion contrato"}),

    MergeItems = Table.NestedJoin(ExpandContratos, {"Codigo act"}, ItemsPorCodigo, {"Codigo act"}, "I", JoinKind.LeftOuter),
    ExpandItems = Table.ExpandTableColumn(MergeItems, "I", {"Actividad", "Capitulo", "Subcapitulo"}, {"Actividad", "Capitulo", "Subcapitulo"}),

    AgregadoTipo = Table.AddColumn(ExpandItems, "Tipo", each "Descuento", type text),
    
    SelectedFinal = Table.SelectColumns(AgregadoTipo, {"Centro de Costos", "Codigo act", "Actividad", "Capitulo", "Subcapitulo", "# OC / Contrato", "Nombre Contratista", "Descripcion contrato", "Valor descuento", "Tipo"}),
    TypedFinal = Table.TransformColumnTypes(SelectedFinal, {{"Centro de Costos", type text}, {"Codigo act", type text}, {"Actividad", type text}, {"Capitulo", type text}, {"Subcapitulo", type text}, {"# OC / Contrato", type text}, {"Nombre Contratista", type text}, {"Descripcion contrato", type text}, {"Valor descuento", Currency.Type}, {"Tipo", type text}}),
    
    FilteredZeros = Table.SelectRows(TypedFinal, each [Valor descuento] <> 0 and [Valor descuento] <> null),

    TablaFinal = FilteredZeros
in
    TablaFinal
