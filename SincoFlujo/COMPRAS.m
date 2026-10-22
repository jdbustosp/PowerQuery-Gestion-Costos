let
    // ============================================================
    // FUNCIONES AUXILIARES GLOBALES
    // ============================================================
    FnFormatCodigoAct = F_Globales[FnFormatCodigoAct],
    FnDecodeHtml = F_Globales[FnDecodeHtml],
    FnPrepareTableWithHeader = F_Globales[FnPrepareTableWithHeader],
    FxToNumberFlex = F_Globales[FxToNumberFlex],
    FnClaveLimpia = F_Globales[FnClaveLimpia],
    FnMapColumn = F_Globales[FnMapColumn],
    FnReadSPBinary = F_Globales[FnReadSPBinary],
    FnRemoveAccentMarks = F_Globales[FnRemoveAccentMarks],
    FnRemoveAccentsSymbols = F_Globales[FnRemoveAccentsSymbols],
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    Columnas_OC = F_Globales[FnBuildColumnas](10),
    Columnas_Entradas = F_Globales[FnBuildColumnas](10),
    Columnas_Salidas = F_Globales[FnBuildColumnas](12),

    ColumnasBase = {
        "Codigo ins", "Ins", "Actividad", "Codigo act", "InsClave", "# OC / Contrato",
        "Cantidad Comprado", "VT Comprado", "VU_Crudo", "IVA_Crudo", "Nombre Contratista",
        "#ENTRADA", "Cantidad Cortes", "VT Cortes", "#SALIDA", "Cantidad Cons Cols", "VT Cons Cols"
    },

    EmptyCompras = #table(ColumnasBase, {}),

    FnText = (v as any) as text =>
        try Text.Trim(Text.From(if v = null then "" else v)) otherwise "",

    FnCleanDisplay = (v as any) as nullable text =>
        let
            t = FnText(v),
            clean = if t = "" then null else Text.Upper(Text.Trim(FnRemoveAccentMarks(t)))
        in
            clean,

    FnBuildInsUM = (desc as any, um as any) as nullable text =>
        let
            d = FnCleanDisplay(desc),
            u = FnCleanDisplay(um)
        in
            if d = null then null else if u = null or u = "" then d else d & " (" & u & ")",

    FnCleanContratistaFromDash = (v as any) as nullable text =>
        let
            t = FnText(v),
            afterDash = if Text.Contains(t, "-") then Text.Trim(Text.AfterDelimiter(t, "-")) else t,
            clean = FnCleanDisplay(afterDash)
        in
            clean,

    FnRenameSequential = (tbl as table) as table =>
        let
            cols = Table.ColumnNames(tbl),
            renamed = Table.RenameColumns(tbl, List.Zip({cols, List.Transform({1..List.Count(cols)}, each "Columna" & Text.From(_))}))
        in
            renamed,

    // ============================================================
    // PROCESAR INFORMEORDEN + ESTADO DE ORDENES
    // ============================================================
    FxProcesarCompras = (BinDetalles as binary, BinOC as binary) as table => let
        RawOC_Raw = try Excel.Workbook(BinOC, null, true){0}[Data]
                otherwise Html.Table(FnDecodeHtml(BinOC), Columnas_OC, [RowSelector="tr"]),
        RawOC = FnRenameSequential(RawOC_Raw),

        AddOCKey = Table.AddColumn(RawOC, "OC_Key_Temp", each let v = FnText([Columna1]) in if Text.StartsWith(v, "Orden de Compra No.") then Text.Trim(Text.Replace(v, "Orden de Compra No.", "")) else null, type text),
        Ordenes_Agrupadas = Table.RenameColumns(Table.Group(Table.SelectRows(Table.FillDown(AddOCKey, {"OC_Key_Temp"}), each [OC_Key_Temp] <> null), {"OC_Key_Temp"}, {{"Proveedor_Raw", each let l = List.RemoveNulls([Columna2]), l2 = List.Select(l, (x) => let t = FnText(x) in t <> "Proveedor" and t <> "Insumo") in if List.IsEmpty(l2) then null else List.First(l2), type text}}), {{"OC_Key_Temp", "OC_Key"}}),

        LibroExcel = Excel.Workbook(Binary.Buffer(BinDetalles), null, true),
        DetallesCrudos = FnPrepareTableWithHeader(LibroExcel{0}[Data]),
        Cols = Table.ColumnNames(DetallesCrudos),
        // FnMapColumn no depende de la fila, solo de Cols+keywords (siempre fijos aqui) —
        // resolver el nombre de columna UNA sola vez en vez de re-buscarlo en cada fila
        // (antes: 8 busquedas completas de texto por fila x N filas, todas con el mismo resultado).
        // Misma logica de match que FnMapColumn (F_Globales), pero devolviendo el NOMBRE de
        // columna en vez del valor, para poder resolverla una sola vez fuera del loop por fila.
        FnResolverCol = (keywords as list) as nullable text =>
            let
                norm = (x as any) as text =>
                    let
                        txt = try Text.From(x) otherwise "",
                        clean = FnRemoveAccentsSymbols(txt)
                    in Text.Upper(if clean = null then "" else clean),
                match = List.First(List.Select(Cols, (c) => List.AnyTrue(List.Transform(keywords, (k) => Text.Contains(norm(c), norm(k))))), null)
            in match,
        ColCodigoIns = FnResolverCol({"CÓDIGO", "CODIGO", "COD."}),
        ColIns = FnResolverCol({"INSUMO", "DESCRIPCIÓN", "DESCRIPCION"}),
        ColAct = FnResolverCol({"ACTIVIDAD", "DESTINO", "FRENTE", "ITEM", "ÍTEM"}),
        ColCant = FnResolverCol({"CANTIDAD", "CANT."}),
        ColVU = FnResolverCol({"VALOR UNITARIO", "VLR UNIT", "UNITARIO"}),
        ColIVA = FnResolverCol({"IVA %", "IVA", "% IVA"}),
        ColVT = FnResolverCol({"VALOR TOTAL", "VLR TOTAL", "TOTAL"}),
        ColOC = FnResolverCol({"ORDEN", "PEDIDO", "O.C"}),
        MapStd = Table.AddColumn(DetallesCrudos, "Std", each [
            Codigo_ins = if ColCodigoIns = null then null else Record.Field(_, ColCodigoIns),
            Ins = if ColIns = null then null else Record.Field(_, ColIns),
            Act = if ColAct = null then null else Record.Field(_, ColAct),
            Cant = if ColCant = null then null else Record.Field(_, ColCant),
            VU_Crudo = try Record.FieldValues(_){10} otherwise (if ColVU = null then null else Record.Field(_, ColVU)),
            IVA_Crudo = try Record.FieldValues(_){11} otherwise (if ColIVA = null then null else Record.Field(_, ColIVA)),
            VT = try Record.FieldValues(_){12} otherwise (if ColVT = null then null else Record.Field(_, ColVT)),
            OC = if ColOC = null then null else Record.Field(_, ColOC)
        ]),
        DetallesStd = Table.ExpandRecordColumn(MapStd, "Std", {"Codigo_ins", "Ins", "Act", "Cant", "VT", "VU_Crudo", "IVA_Crudo", "OC"}, {"Codigo ins", "Ins", "Actividad", "Cantidad Comprado", "VT Comprado", "VU_Crudo", "IVA_Crudo", "# OC / Contrato"}),
        DetConKeyOC = Table.AddColumn(DetallesStd, "OC_Key", each FnText([#"# OC / Contrato"]), type text),
        DetConCodAct = Table.AddColumn(DetConKeyOC, "Codigo act", each let c = Text.Trim(Text.BeforeDelimiter(FnText([Actividad]), "-", 0)) in if c = "" then null else c, type text),
        DetConClave = Table.AddColumn(DetConCodAct, "InsClave", each FnClaveLimpia([Ins]), type text),
        MergedOC = Table.NestedJoin(DetConClave, {"OC_Key"}, Ordenes_Agrupadas, {"OC_Key"}, "ORD", JoinKind.LeftOuter),
        ExpandedOC = Table.ExpandTableColumn(MergedOC, "ORD", {"Proveedor_Raw"}, {"Proveedor_Raw"}),
        AddedNombreContratista = Table.AddColumn(ExpandedOC, "Nombre Contratista", each FnCleanContratistaFromDash([Proveedor_Raw]), type text),
        Selected = Table.SelectColumns(AddedNombreContratista, ColumnasBase, MissingField.UseNull)
    in Selected,

    // ============================================================
    // PROCESAR INFORME ENTRADAS DE ALMACEN DETALLADAS
    // ============================================================
    FxProcesarEntradas = (BinEntradas as binary) as table => let
        Raw_Raw = try Excel.Workbook(Binary.Buffer(BinEntradas), null, true){0}[Data]
                otherwise Html.Table(FnDecodeHtml(BinEntradas), Columnas_Entradas, [RowSelector="tr"]),
        // FillDown O(N): se marca cada fila de encabezado una sola vez y se arrastra
        // hacia abajo, en lugar de re-escanear toda la tabla por cada fila (O(N^2)).
        Raw = Table.Buffer(FnRenameSequential(Raw_Raw)),
        ConFlagMeta = Table.AddColumn(Raw, "__esMeta", each
            FnText([Columna1]) <> "" and FnText([Columna2]) <> "" and FnText([Columna4]) <> "" and
            (try Date.From([Columna1]) otherwise null) <> null and
            (try Number.FromText(FnText([Columna2])) otherwise null) <> null and
            (try Number.FromText(FnText([Columna4])) otherwise null) <> null, type logical),
        ConMetaRec = Table.AddColumn(ConFlagMeta, "__meta", each if [__esMeta] then _ else null, type nullable record),
        WithMeta = Table.FillDown(ConMetaRec, {"__meta"}),
        Items = Table.SelectRows(WithMeta, each
            let
                cod = FnText([Columna1]),
                codNum = try Number.FromText(cod) otherwise null,
                ins = FnText([Columna2]),
                cant = FxToNumberFlex([Columna4]),
                vt = FxToNumberFlex([Columna7])
            in
                codNum <> null and ins <> "" and (cant <> null or vt <> null)
        ),
        AddStd = Table.AddColumn(Items, "Std", each
            let
                m = [__meta],
                entrada = if m = null then null else FnText(Record.Field(m, "Columna2")),
                oc = if m = null then null else FnText(Record.Field(m, "Columna4")),
                proveedor = if m = null then null else Record.Field(m, "Columna3"),
                insFinal = FnBuildInsUM([Columna2], [Columna3])
            in [
                #"Codigo ins" = FnText([Columna1]),
                Ins = insFinal,
                Actividad = null,
                #"Codigo act" = null,
                InsClave = FnClaveLimpia(insFinal),
                #"# OC / Contrato" = oc,
                #"Cantidad Comprado" = null,
                #"VT Comprado" = null,
                VU_Crudo = null,
                IVA_Crudo = null,
                #"Nombre Contratista" = FnCleanContratistaFromDash(proveedor),
                #"#ENTRADA" = entrada,
                #"Cantidad Cortes" = FxToNumberFlex([Columna4]),
                #"VT Cortes" = FxToNumberFlex([Columna7]),
                #"#SALIDA" = null,
                #"Cantidad Cons Cols" = null,
                #"VT Cons Cols" = null
            ]),
        Expanded = Table.ExpandRecordColumn(AddStd, "Std", ColumnasBase, ColumnasBase),
        Selected = Table.SelectColumns(Expanded, ColumnasBase, MissingField.UseNull)
    in Selected,

    // ============================================================
    // PROCESAR MASIVO SALIDAS DETALLADO
    // ============================================================
    FxProcesarSalidas = (BinSalidas as binary) as table => let
        Raw_Raw = Html.Table(FnDecodeHtml(BinSalidas), Columnas_Salidas, [RowSelector="tr"]),
        // FillDown O(N): mismo patron que en Entradas para evitar el re-escaneo O(N^2).
        Raw = Table.Buffer(FnRenameSequential(Raw_Raw)),
        ConFlagMeta = Table.AddColumn(Raw, "__esMeta", each
            FnText([Columna1]) <> "" and FnText([Columna2]) <> "" and FnText([Columna3]) <> "" and
            (try Date.From([Columna1]) otherwise null) <> null and
            (try Number.FromText(FnText([Columna2])) otherwise null) <> null, type logical),
        ConMetaRec = Table.AddColumn(ConFlagMeta, "__meta", each if [__esMeta] then _ else null, type nullable record),
        WithMeta = Table.FillDown(ConMetaRec, {"__meta"}),
        Items = Table.SelectRows(WithMeta, each
            let
                cod = FnText([Columna1]),
                codNum = try Number.FromText(cod) otherwise null,
                ins = FnText([Columna2]),
                cant = FxToNumberFlex([Columna5]),
                vt = FxToNumberFlex([Columna8])
            in
                codNum <> null and ins <> "" and (cant <> null or vt <> null)
        ),
        AddStd = Table.AddColumn(Items, "Std", each
            let
                m = [__meta],
                salida = if m = null then null else FnText(Record.Field(m, "Columna2")),
                contratista = if m = null then null else Record.Field(m, "Columna3"),
                insFinal = FnBuildInsUM([Columna2], [Columna4]),
                codAct = FnFormatCodigoAct([Columna3])
            in [
                #"Codigo ins" = FnText([Columna1]),
                Ins = insFinal,
                Actividad = null,
                #"Codigo act" = codAct,
                InsClave = FnClaveLimpia(insFinal),
                #"# OC / Contrato" = null,
                #"Cantidad Comprado" = null,
                #"VT Comprado" = null,
                VU_Crudo = null,
                IVA_Crudo = null,
                #"Nombre Contratista" = FnCleanContratistaFromDash(contratista),
                #"#ENTRADA" = null,
                #"Cantidad Cortes" = null,
                #"VT Cortes" = null,
                #"#SALIDA" = salida,
                #"Cantidad Cons Cols" = FxToNumberFlex([Columna5]),
                #"VT Cons Cols" = FxToNumberFlex([Columna8])
            ]),
        Expanded = Table.ExpandRecordColumn(AddStd, "Std", ColumnasBase, ColumnasBase),
        Selected = Table.SelectColumns(Expanded, ColumnasBase, MissingField.UseNull)
    in Selected,

    // ============================================================
    // CONEXIÓN A SHAREPOINT (LECTURA DESDE CONSULTA COMPARTIDA)
    // ============================================================
    ArchivosProyecto = Table.SelectRows(SP_Archivos_Proyecto, each
        Text.Contains([Name], "INFORMEORDEN", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "ESTADO DE ORDENES", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "INFORME ENTRADAS DE ALMACEN", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "INFORME ENTRADAS DE ALMACÉN", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "ENTRADAS POR INSUMO", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "MASIVO SALIDAS", Comparer.OrdinalIgnoreCase)
    ),
    ConCentroCosto = ArchivosProyecto,

    PickLatestBinary = (t as table, containsText as text) as nullable binary =>
        let
            candidatos = Table.Sort(
                Table.SelectRows(t, each Text.Contains([Name], containsText, Comparer.OrdinalIgnoreCase)),
                {{"TimeLastModified", Order.Descending}, {"Name", Order.Ascending}}
            ),
            path = if Table.RowCount(candidatos) = 0 then null else candidatos{0}[ServerRelativeUrl]
        in
            if path = null then null else FnReadSPBinary(SiteUrl, path),

    PickLatestBinaryAny = (t as table, containsTexts as list) as nullable binary =>
        let
            candidatos = Table.Sort(
                Table.SelectRows(t, each List.AnyTrue(List.Transform(containsTexts, (needle) => Text.Contains([Name], needle, Comparer.OrdinalIgnoreCase)))),
                {{"TimeLastModified", Order.Descending}, {"Name", Order.Ascending}}
            ),
            path = if Table.RowCount(candidatos) = 0 then null else candidatos{0}[ServerRelativeUrl]
        in
            if path = null then null else FnReadSPBinary(SiteUrl, path),

    // === Preferir formato liviano (plano, sin fill-down) si existe en SharePoint,
    // con respaldo automatico al formato "detallado" viejo si el proyecto aun no
    // tiene el reporte nuevo subido. Nunca rompe: si no hay ninguno, retorna null.
    PickPreferido = (t as table, containsLiviano as text, containsViejo as text) as record =>
        let
            candLiviano = Table.Sort(Table.SelectRows(t, each Text.Contains([Name], containsLiviano, Comparer.OrdinalIgnoreCase)), {{"TimeLastModified", Order.Descending}, {"Name", Order.Ascending}}),
            hayLiviano = Table.RowCount(candLiviano) > 0,
            candViejo = Table.Sort(Table.SelectRows(t, each Text.Contains([Name], containsViejo, Comparer.OrdinalIgnoreCase) and not Text.Contains([Name], containsLiviano, Comparer.OrdinalIgnoreCase)), {{"TimeLastModified", Order.Descending}, {"Name", Order.Ascending}}),
            path = if hayLiviano then candLiviano{0}[ServerRelativeUrl] else if Table.RowCount(candViejo) > 0 then candViejo{0}[ServerRelativeUrl] else null,
            bin = if path = null then null else FnReadSPBinary(SiteUrl, path)
        in [Binario = bin, EsLiviano = hayLiviano],

    // === List.Transform en vez de Table.Group + AddColumn anidado ===
    // Confirmado por prueba: el mismo trabajo via Table.Group tardaba 402s vs
    // 133s via List.Transform (3x mas rapido), para el mismo Centro de Costo.
    CCsConArchivos = List.Distinct(ConCentroCosto[Centro de Costos]),
    ResultadosPorCC = List.Transform(CCsConArchivos, (cc) =>
        let
            filtrado = Table.SelectRows(ConCentroCosto, each [Centro de Costos] = cc),
            binDet = PickLatestBinary(filtrado, "INFORMEORDEN"),
            binOC = PickLatestBinary(filtrado, "ESTADO DE ORDENES"),
            resEntradas = PickPreferido(filtrado, "ENTRADAS POR INSUMO", "INFORME ENTRADAS DE ALMACEN"),
            resSalidas = PickPreferido(filtrado, "DESCRIPTIVAS", "MASIVO SALIDAS"),
            tCompras = if binDet <> null and binOC <> null then FxProcesarCompras(binDet, binOC) else EmptyCompras,
            tEntradas = if resEntradas[Binario] = null then EmptyCompras
                        else if resEntradas[EsLiviano] then F_Globales[FxProcesarEntradasPorInsumo](resEntradas[Binario], ColumnasBase)
                        else FxProcesarEntradas(resEntradas[Binario]),
            tSalidas = if resSalidas[Binario] = null then EmptyCompras
                       else if resSalidas[EsLiviano] then F_Globales[FxProcesarSalidasDescriptivas](resSalidas[Binario], ColumnasBase)
                       else FxProcesarSalidas(resSalidas[Binario]),
            combinado = Table.Combine({tCompras, tEntradas, tSalidas}),
            conCC = Table.AddColumn(combinado, "Centro de Costos", each cc, type text)
        in
            conCC
    ),
    Expandido = Table.Combine(ResultadosPorCC),

    Expandido_Clean = Table.TransformColumns(Expandido, {
        {"Centro de Costos", each if _ = null then null else Text.Upper(Text.Trim(Text.From(_))), type text},
        {"Codigo act", each FnFormatCodigoAct(_), type text}
    }, null, MissingField.Ignore),

    Compras_Unicas = Table.Buffer(Table.Distinct(Expandido_Clean)),

    // ============================================================
    // LECTURA DIRECTA DE CONSULTAS (Memoria)
    // ============================================================
    ITEMS_TablaLocal = ITEMSINSUMOS,
    ITEMS_Clean = Table.TransformColumns(ITEMS_TablaLocal, {
        {"Centro de Costos", each if _ = null then null else Text.Upper(Text.Trim(Text.From(_))), type text},
        {"Codigo act", each FnFormatCodigoAct(_), type text}
    }, null, MissingField.Ignore),
    // Buffer: ITEMS_Base se usa 3 veces (ITEMS_Insumos_Dist, ItemsPorCodigo_Estricto,
    // ItemsPorCodigo_Generico) — sin buffer, cada uno recalcula TransformColumns+SelectColumns
    // desde cero en vez de reusar el resultado ya calculado.
    ITEMS_Base = Table.Buffer(Table.SelectColumns(ITEMS_Clean, {"Centro de Costos", "Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo"})),

    ITEMS_Insumos_Dist = Table.Buffer(Table.Distinct(Table.AddColumn(ITEMS_Base, "InsClave", each FnClaveLimpia([Ins]), type text), {"Centro de Costos", "Codigo act", "InsClave"})),
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
            ca = if [#"#ENTRADA"] <> null then null else if e then [Codigo act] else (if [Rs.Codigo act] <> null then [Rs.Codigo act] else [Codigo act]),
            a0 = FnText([Actividad]),
            aOrig = if a0 = "" then null else if ca <> null and not Text.StartsWith(a0, Text.From(ca)) then Text.From(ca) & " - " & a0 else a0,

            ActOficial = if [#"#ENTRADA"] <> null then null else if [Act_Estricto] <> null then [Act_Estricto] else if [Act_Gen] <> null then [Act_Gen] else aOrig,
            CapFinal = if [#"#ENTRADA"] <> null then null else if [Cap_Estricto] <> null then [Cap_Estricto] else [Cap_Gen],
            SubCapFinal = if [#"#ENTRADA"] <> null then null else if [Sub_Estricto] <> null then [Sub_Estricto] else [Sub_Gen]
        in [
            InsFinal = if e then [Ex.Ins] else (if [Rs.Ins] <> null then [Rs.Ins] else [Ins]),
            CodActFinal = ca,
            ActFinal = ActOficial,
            CapFinal = CapFinal,
            SubCapFinal = SubCapFinal
        ]),

    ExpandedFinalCols = Table.ExpandRecordColumn(Table.RemoveColumns(AddedCoalesced, {"Ins", "Actividad", "Codigo act", "Ex.Ins", "Rs.Codigo act", "Rs.Actividad", "Rs.Ins", "Act_Estricto", "Cap_Estricto", "Sub_Estricto", "Act_Gen", "Cap_Gen", "Sub_Gen"}), "FinalCols", {"InsFinal", "CodActFinal", "ActFinal", "CapFinal", "SubCapFinal"}, {"Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo"}),

    NumericColumns = Table.TransformColumns(ExpandedFinalCols, {
        {"#ENTRADA", each FxToNumberFlex(_), Int64.Type},
        {"#SALIDA", each FxToNumberFlex(_), Int64.Type},
        {"Cantidad Comprado", each FxToNumberFlex(_), type number},
        {"VT Comprado", each FxToNumberFlex(_), type number},
        {"VU_Crudo", each FxToNumberFlex(_), type number},
        {"IVA_Crudo", each FxToNumberFlex(_), type number},
        {"Cantidad Cortes", each FxToNumberFlex(_), type number},
        {"VT Cortes", each FxToNumberFlex(_), type number},
        {"Cantidad Cons Cols", each FxToNumberFlex(_), type number},
        {"VT Cons Cols", each FxToNumberFlex(_), type number}
    }),
    Added_VU = Table.AddColumn(NumericColumns, "V/U Comprado", each let vb = [VU_Crudo], iva = [IVA_Crudo], p = if iva = null then 0 else if iva >= 1 then iva / 100 else iva, vc = if vb = null then null else vb * (1 + p) in if vc = null then null else Number.Round(vc, 0), type number),
    FilteredZeros = Table.SelectRows(Added_VU, each try
        ([VT Comprado] <> null and [VT Comprado] <> 0) or
        ([VT Cortes] <> null and [VT Cortes] <> 0) or
        ([VT Cons Cols] <> null and [VT Cons Cols] <> 0)
    otherwise false),

    SelectedFinal = Table.SelectColumns(Table.AddColumn(Table.AddColumn(FilteredZeros, "Tipo", each "COMPRAS", type text), "Descripcion contrato", each "pedido obra", type text), {
        "Centro de Costos", "Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo",
        "# OC / Contrato", "#ENTRADA", "#SALIDA", "Nombre Contratista", "Descripcion contrato",
        "Cantidad Comprado", "VT Comprado", "V/U Comprado", "Cantidad Cortes", "VT Cortes", "Cantidad Cons Cols", "VT Cons Cols", "Tipo"
    }, MissingField.Ignore),
    TypedFinal = Table.TransformColumnTypes(SelectedFinal, {
        {"Centro de Costos", type text}, {"Codigo ins", Int64.Type}, {"Cantidad Comprado", type number},
        {"#ENTRADA", Int64.Type}, {"#SALIDA", Int64.Type}, {"VT Comprado", type number}, {"V/U Comprado", type number}, {"Cantidad Cortes", type number},
        {"VT Cortes", type number}, {"Cantidad Cons Cols", type number}, {"VT Cons Cols", type number}
    }),
    TablaFinal = TypedFinal
in
    TablaFinal
