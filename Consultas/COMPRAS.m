let
    // ============================================================
    // FUNCIONES AUXILIARES GLOBALES
    // ============================================================
    FnFormatCodigoAct = (raw as any) as nullable text => let txtRaw = if raw = null then null else Text.Trim(Text.From(raw)), result = if txtRaw = null or txtRaw = "" then null else let txtNorm = Text.Replace(Text.Replace(txtRaw, ",", "."), " ", ""), hasDot = Text.Contains(txtNorm, ".") in if hasDot then txtNorm else let digits = Text.Select(txtNorm, {"0".."9"}), len = Text.Length(digits) in if len <= 3 then null else Text.Range(digits, 0, len - 3) & "." & Text.Range(digits, len - 3, 3) in result,
    FnPrepareTableWithHeader = (tbl as table) as table => let firstColName = Table.ColumnNames(tbl){0}, firstColValues = Table.Column(tbl, firstColName), headerFlags = List.Transform(firstColValues, (x) => let txt = Text.Upper(if x = null then "" else Text.From(x)), txtNorm = Text.Replace(txt, "Ó", "O") in Text.Contains(txtNorm, "COD")), hasHeader = List.Contains(headerFlags, true), promoted = if hasHeader then let headerIndex = List.PositionOf(headerFlags, true), skipped = Table.Skip(tbl, headerIndex) in Table.PromoteHeaders(skipped, [PromoteAllScalars = true]) else tbl in promoted,
    FxToNumberFlex = (value as any) as nullable number => let v = value, numeroDirecto = if Value.Is(v, type number) then Number.From(v) else null in if numeroDirecto <> null then numeroDirecto else let t0 = if v=null then "" else Text.From(v), t = Text.Trim(Text.Replace(Text.Replace(t0, "#(00A0)", ""), " ", "")) in if t="" then null else let tryUS = try Number.FromText(t, "en-US"), valUS = if tryUS[HasError] then null else tryUS[Value] in if valUS<>null then valUS else let tryES = try Number.FromText(t, "es-ES"), valES = if tryES[HasError] then null else tryES[Value] in valES,
    replacements = {{"á","a"},{"Á","A"},{"é","e"},{"É","E"},{"í","i"},{"Í","I"},{"ó","o"},{"Ó","O"},{"ú","u"},{"Ú","U"},{"º",""},{"°",""},{"¨",""}},
    Normalize = (t as nullable text) => let txt = if t=null then "" else Text.From(t), result = List.Accumulate(replacements, txt, (state,current)=> Text.Replace(state, current{0}, current{1})) in Text.Trim(result),
    FxClaveTexto = (t as nullable text) => let clean = Normalize(t), sinPar = if Text.Contains(clean,"(") then Text.BeforeDelimiter(clean,"(") else clean, palabras = List.Select(Text.Split(sinPar," "), each _<> ""), base = if List.Count(palabras)=0 then null else Text.Lower(Text.Combine(palabras," ")) in base,
    FnMapColumn = (rec as record, cols as list, keywords as list) => let match = List.First(List.Select(cols, (c) => List.AnyTrue(List.Transform(keywords, (k) => Text.Contains(Text.Upper(c), k))))) in if match = null then null else Record.Field(rec, match),
    Columnas_OC = List.Transform({1..10}, each {"Columna" & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}),

    // ============================================================
    // FUNCIÓN MÁGICA: PROCESAR COMPRAS
    // ============================================================
    FxProcesarCompras = (BinDetalles as binary, BinOC as binary) => let
        HtmlOC = Text.FromBinary(Binary.Buffer(BinOC), 65001),
        RawOC = Html.Table(HtmlOC, Columnas_OC, [RowSelector="tr"]),
        AddOCKey = Table.AddColumn(RawOC, "OC_Key_Temp", each let v = Text.From([Columna1] ?? "") in if Text.StartsWith(v, "Orden de Compra No.") then Text.Trim(Text.Replace(v, "Orden de Compra No.", "")) else null, type text),
        Ordenes_Agrupadas = Table.RenameColumns(Table.Group(Table.SelectRows(Table.FillDown(AddOCKey, {"OC_Key_Temp"}), each [OC_Key_Temp] <> null), {"OC_Key_Temp"}, {{"Proveedor_Raw", each let l = List.RemoveNulls([Columna2]), l2 = List.Select(l, (x) => let t = Text.Trim(Text.From(x ?? "")) in t <> "Proveedor" and t <> "Insumo") in if List.IsEmpty(l2) then null else List.First(l2), type text}}), {{"OC_Key_Temp", "OC_Key"}}),

        LibroExcel = Excel.Workbook(Binary.Buffer(BinDetalles), null, true),
        DetallesCrudos = FnPrepareTableWithHeader(LibroExcel{0}[Data]),
        Cols = Table.ColumnNames(DetallesCrudos),
        MapStd = Table.AddColumn(DetallesCrudos, "Std", each [ Codigo_ins = FnMapColumn(_, Cols, {"CÓDIGO", "CODIGO", "COD."}), Ins = FnMapColumn(_, Cols, {"INSUMO", "DESCRIPCIÓN", "DESCRIPCION"}), Act = FnMapColumn(_, Cols, {"ACTIVIDAD", "DESTINO", "FRENTE", "ITEM", "ÍTEM"}), Cant = FnMapColumn(_, Cols, {"CANTIDAD", "CANT."}), VU_Crudo = try Record.FieldValues(_){10} otherwise FnMapColumn(_, Cols, {"VALOR UNITARIO", "VLR UNIT", "UNITARIO"}), IVA_Crudo = try Record.FieldValues(_){11} otherwise FnMapColumn(_, Cols, {"IVA %", "IVA", "% IVA"}), VT = try Record.FieldValues(_){12} otherwise FnMapColumn(_, Cols, {"VALOR TOTAL", "VLR TOTAL", "TOTAL"}), OC = FnMapColumn(_, Cols, {"ORDEN", "PEDIDO", "O.C"}) ]),
        DetallesStd = Table.ExpandRecordColumn(MapStd, "Std", {"Codigo_ins", "Ins", "Act", "Cant", "VT", "VU_Crudo", "IVA_Crudo", "OC"}, {"Codigo ins", "Ins", "Actividad", "Cantidad Comprado", "VT Comprado", "VU_Crudo", "IVA_Crudo", "# OC / Contrato"}),
        DetConKeyOC = Table.AddColumn(DetallesStd, "OC_Key", each Text.Trim(Text.From([#"# OC / Contrato"] ?? "")), type text),
        DetConCodAct = Table.AddColumn(DetConKeyOC, "Codigo act", each let c = Text.Trim(Text.BeforeDelimiter(Text.Trim(Text.From([Actividad] ?? "")), "-", 0)) in if c = "" then null else c, type text),
        DetConClave = Table.AddColumn(DetConCodAct, "InsClave", each FxClaveTexto([Ins]), type text),
        MergedOC = Table.NestedJoin(DetConClave, {"OC_Key"}, Ordenes_Agrupadas, {"OC_Key"}, "ORD", JoinKind.LeftOuter),
        ExpandedOC = Table.ExpandTableColumn(MergedOC, "ORD", {"Proveedor_Raw"}, {"Proveedor_Raw"}),
        AddedNombreContratista = Table.AddColumn(ExpandedOC, "Nombre Contratista", each let p = try Text.From([Proveedor_Raw]) otherwise null, t = if p = null then null else let pos = Text.PositionOf(p, "-") in if pos < 0 then Text.Trim(p) else Text.Trim(Text.Range(p, pos + 1)) in t, type text)
    in AddedNombreContratista,

    // ============================================================
    // CONEXIÓN A SHAREPOINT
    // ============================================================
    RutaBase = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv", ArchivosSharePoint = SharePoint.Files(RutaBase, [ApiVersion = 15]),
    ParamProyecto = Text.Trim(ProyectoActual),
    ArchivosProyecto = Table.SelectRows(ArchivosSharePoint, each Text.Contains(Text.Upper([Folder Path]), "/" & Text.Upper(ParamProyecto) & "/") and Text.EndsWith([Folder Path], "/Actual/") and (Text.Contains(Text.Upper([Name]), "INFORMEORDEN") or Text.Contains(Text.Upper([Name]), "ESTADO DE ORDENES")) and not Text.StartsWith([Name], "~$")),
    ConCentroCosto = Table.AddColumn(ArchivosProyecto, "Centro de Costos", each Text.Trim(Text.Replace(Text.AfterDelimiter([Folder Path], "/" & ParamProyecto & "/"), "/Actual/", ""))),
    
    Agrupado = Table.Group(ConCentroCosto, {"Centro de Costos"}, {{"Binarios", each let FilaDet = Table.SelectRows(_, each Text.Contains(Text.Upper([Name]), "INFORMEORDEN")), FilaOC = Table.SelectRows(_, each Text.Contains(Text.Upper([Name]), "ESTADO DE ORDENES")) in if Table.RowCount(FilaDet) > 0 and Table.RowCount(FilaOC) > 0 then [Bin_Det = Binary.Buffer(FilaDet{0}[Content]), Bin_OC = Binary.Buffer(FilaOC{0}[Content])] else null}}),
    CentrosCompletos = Table.SelectRows(Agrupado, each [Binarios] <> null),
    TablaConDatos = Table.AddColumn(CentrosCompletos, "Datos", each FxProcesarCompras([Binarios][Bin_Det], [Binarios][Bin_OC])),
    Expandido = Table.ExpandTableColumn(TablaConDatos, "Datos", {"Codigo ins", "Ins", "Actividad", "Codigo act", "InsClave", "# OC / Contrato", "Cantidad Comprado", "VT Comprado", "VU_Crudo", "IVA_Crudo", "Nombre Contratista"}),

    Expandido_Clean = Table.TransformColumns(Expandido, {
        {"Centro de Costos", each if _ = null then null else Text.Upper(Text.Trim(Text.From(_))), type text},
        {"Codigo act", each FnFormatCodigoAct(_), type text}
    }, null, MissingField.Ignore),
    
    Compras_Unicas = Table.Buffer(Table.Distinct(Expandido_Clean)),

    // ============================================================
    // LECTURA DIRECTA A EXCEL (MAESTROS LOCALES - 🔥 MODO ESTRICTO)
    // ============================================================
    ITEMS_TablaLocal = Table.Buffer(Excel.CurrentWorkbook(){[Name="TbItems"]}[Content]),
    ITEMS_Clean = Table.TransformColumns(ITEMS_TablaLocal, {
        {"Centro de Costos", each if _ = null then null else Text.Upper(Text.Trim(Text.From(_))), type text},
        {"Codigo act", each FnFormatCodigoAct(_), type text}
    }, null, MissingField.Ignore),
    ITEMS_Base = Table.SelectColumns(ITEMS_Clean, {"Centro de Costos", "Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo"}),

    ITEMS_Insumos_Dist = Table.Buffer(Table.Distinct(Table.AddColumn(ITEMS_Base, "InsClave", each FxClaveTexto([Ins]), type text), {"Centro de Costos", "Codigo act", "InsClave"})),
    ITEMS_Respaldo = Table.Buffer(Table.Distinct(ITEMS_Insumos_Dist, {"Centro de Costos", "InsClave"})),

    ItemsPorCodigo_Estricto = Table.Buffer(Table.Group(ITEMS_Base, {"Centro de Costos", "Codigo act"}, {
        {"Act_Estricto", each List.First(List.RemoveNulls([Actividad])), type text},
        {"Cap_Estricto", each List.First(List.RemoveNulls([Capitulo])), type text}, 
        {"Sub_Estricto", each List.First(List.RemoveNulls([Subcapitulo])), type text}
    })),
    
    ItemsPorCodigo_Generico = Table.Buffer(Table.Group(ITEMS_Base, {"Codigo act"}, {
        {"Act_Gen", each List.First(List.RemoveNulls([Actividad])), type text},
        {"Cap_Gen", each List.First(List.RemoveNulls([Capitulo])), type text}, 
        {"Sub_Gen", each List.First(List.RemoveNulls([Subcapitulo])), type text}
    })),

    // ============================================================
    // CRUCES FINALES 
    // ============================================================
    MergedExacto = Table.NestedJoin(Compras_Unicas, {"Centro de Costos", "Codigo act", "InsClave"}, ITEMS_Insumos_Dist, {"Centro de Costos", "Codigo act", "InsClave"}, "EXACTO", JoinKind.LeftOuter),
    ExpandedExacto = Table.ExpandTableColumn(MergedExacto, "EXACTO", {"Ins"}, {"Ex.Ins"}),
    
    MergedRescate = Table.NestedJoin(ExpandedExacto, {"Centro de Costos", "InsClave"}, ITEMS_Respaldo, {"Centro de Costos", "InsClave"}, "RESCATE", JoinKind.LeftOuter),
    ExpandedRescate = Table.ExpandTableColumn(MergedRescate, "RESCATE", {"Codigo act", "Actividad", "Ins"}, {"Rs.Codigo act", "Rs.Actividad", "Rs.Ins"}),

    MergedEstricto = Table.NestedJoin(ExpandedRescate, {"Centro de Costos", "Codigo act"}, ItemsPorCodigo_Estricto, {"Centro de Costos", "Codigo act"}, "EST", JoinKind.LeftOuter),
    ExpandedEstricto = Table.ExpandTableColumn(MergedEstricto, "EST", {"Act_Estricto", "Cap_Estricto", "Sub_Estricto"}, {"Act_Estricto", "Cap_Estricto", "Sub_Estricto"}),

    MergedGenerico = Table.NestedJoin(ExpandedEstricto, {"Codigo act"}, ItemsPorCodigo_Generico, {"Codigo act"}, "GEN", JoinKind.LeftOuter),
    ExpandedGenerico = Table.ExpandTableColumn(MergedGenerico, "GEN", {"Act_Gen", "Cap_Gen", "Sub_Gen"}, {"Act_Gen", "Cap_Gen", "Sub_Gen"}),

    AddedCoalesced = Table.AddColumn(ExpandedGenerico, "FinalCols", each 
        let 
            e = [Ex.Ins] <> null, 
            ca = if e then [Codigo act] else ([Rs.Codigo act] ?? [Codigo act]), 
            a0 = Text.Trim(Text.From([Actividad] ?? "")), 
            aOrig = if a0 = "" then null else if ca <> null and not Text.StartsWith(a0, Text.From(ca)) then Text.From(ca) & " - " & a0 else a0,
            
            ActOficial = [Act_Estricto] ?? [Act_Gen] ?? aOrig,
            CapFinal = [Cap_Estricto] ?? [Cap_Gen],
            SubCapFinal = [Sub_Estricto] ?? [Sub_Gen]
        in [ 
            InsFinal = if e then [Ex.Ins] else ([Rs.Ins] ?? [Ins]), 
            CodActFinal = ca, 
            ActFinal = ActOficial, 
            CapFinal = CapFinal, 
            SubCapFinal = SubCapFinal 
        ]),
    
    ExpandedFinalCols = Table.ExpandRecordColumn(Table.RemoveColumns(AddedCoalesced, {"Ins", "Actividad", "Codigo act", "Ex.Ins", "Rs.Codigo act", "Rs.Actividad", "Rs.Ins", "Act_Estricto", "Cap_Estricto", "Sub_Estricto", "Act_Gen", "Cap_Gen", "Sub_Gen"}), "FinalCols", {"InsFinal", "CodActFinal", "ActFinal", "CapFinal", "SubCapFinal"}, {"Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo"}),

    NumericColumns = Table.TransformColumns(ExpandedFinalCols, {{"Cantidad Comprado", each FxToNumberFlex(_), type number}, {"VT Comprado", each FxToNumberFlex(_), type number}, {"VU_Crudo", each FxToNumberFlex(_), type number}, {"IVA_Crudo", each FxToNumberFlex(_), type number}}),
    Added_VU = Table.AddColumn(NumericColumns, "V/U Comprado", each let vb = [VU_Crudo], iva = [IVA_Crudo], p = if iva = null then 0 else if iva >= 1 then iva / 100 else iva, vc = if vb = null then null else vb * (1 + p) in if vc = null then null else Number.Round(vc, 0), type number),
    FilteredZeros = Table.SelectRows(Added_VU, each try [VT Comprado] <> null and [VT Comprado] <> 0 otherwise false),

    SelectedFinal = Table.SelectColumns(Table.AddColumn(Table.AddColumn(FilteredZeros, "Tipo", each "COMPRAS", type text), "Descripcion contrato", each "pedido obra", type text), {"Centro de Costos", "Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo", "# OC / Contrato", "Nombre Contratista", "Descripcion contrato", "Cantidad Comprado", "VT Comprado", "V/U Comprado", "Tipo"}),
    TypedFinal = Table.TransformColumnTypes(SelectedFinal, {{"Centro de Costos", type text}, {"Codigo ins", Int64.Type}, {"Cantidad Comprado", type number}, {"VT Comprado", type number}, {"V/U Comprado", type number}}),
    TablaFinal = Table.Buffer(TypedFinal)
in
    TablaFinal
