let
    // =========================================================
    // ASEGURAMOS EL PARÁMETRO GLOBAL
    // =========================================================
    ParamProyecto = Text.Trim(ProyectoActual),

    // =========================================================
    // FUNCIONES AUXILIARES GLOBALES
    // =========================================================
    FnFormatCodigoAct = F_Globales[FnFormatCodigoAct],
    FnParseNumber = F_Globales[FxToNumberFlex],
    FnRemoveAccentsSymbols = F_Globales[FnRemoveAccentsSymbols],
    FnPrepareTableWithHeader = F_Globales[FnPrepareTableWithHeader],
    Columnas_HTML = List.Transform({1..40}, each {"Columna " & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}),

    // =========================================================
    // FUNCIÓN MÁGICA
    // =========================================================
    FxProcesarCentroCosto = (BinarioSeguimiento as binary, BinarioPresupuesto as binary) =>
        let
            // 🔥 LECTURA FORZADA EN UTF-8 (65001) PARA LAS "Ñ" y TILDES
            HtmlSeguimiento  = Text.FromBinary(Binary.Buffer(BinarioSeguimiento), 65001),
            OrigenItems      = Table.Buffer(Html.Table(HtmlSeguimiento, Columnas_HTML, [RowSelector="tr"])),
            ItemsPrepared    = FnPrepareTableWithHeader(OrigenItems),

            ItemsColNames    = Table.ColumnNames(ItemsPrepared),
            ItemsCodColName  = ItemsColNames{0}, ItemsDescColName = ItemsColNames{1}, ItemsTipoColName = ItemsColNames{2}, ItemsUMColName   = ItemsColNames{3},

            ItemsWithTipoFila = Table.AddColumn(ItemsPrepared, "TipoFila", (r as record) => let codValue = Record.Field(r, ItemsCodColName), descValue = Record.Field(r, ItemsDescColName), tipoValue = Record.Field(r, ItemsTipoColName), umValue = Record.Field(r, ItemsUMColName), codText = if codValue = null then "" else Text.Trim(Text.From(codValue)), descText = if descValue = null then "" else Text.Trim(Text.From(descValue)), tipoText = if tipoValue = null then "" else Text.Trim(Text.From(tipoValue)), umText = if umValue = null then "" else Text.Trim(Text.From(umValue)), codUpper = Text.Upper(codText), descUpper = Text.Upper(descText), tryNum = try Number.FromText(codText), isNumeric = not tryNum[HasError], numValue = if isNumeric then tryNum[Value] else 0, tipoFila = if codText = "" then "Otro" else if Text.StartsWith(codUpper, "SUBCAP") or Text.StartsWith(descUpper, "SUBCAP") then "SubCapitulo" else if Text.Contains(codUpper, "CAPITULO") or Text.Contains(descUpper, "CAPITULO") then "Capitulo" else if isNumeric and tipoText = "" and umText = "" and (Text.Length(codText) <= 2 or (numValue >= 1000 and Number.Mod(numValue, 1000) = 0)) then "Capitulo" else if isNumeric and tipoText = "" and umText = "" then "Actividad" else if isNumeric then "Insumo" else "Otro" in tipoFila, type text),
            ItemsWithCapitulo = Table.AddColumn(ItemsWithTipoFila, "Capitulo", (r as record) => let tipo = Record.Field(r, "TipoFila"), codRaw = Record.Field(r, ItemsCodColName), descRaw = Record.Field(r, ItemsDescColName), codTxt = if codRaw = null then "" else Text.Trim(Text.From(codRaw)), descTxt = if descRaw = null then "" else Text.Trim(Text.From(descRaw)), codCap = if codTxt = "00" then codTxt else let num = try Number.FromText(codTxt) in if not num[HasError] and num[Value] >= 1000 and Number.Mod(num[Value], 1000) = 0 then Text.From(num[Value] / 1000) else codTxt, capTxt = if descTxt = "" then codCap else codCap & "-" & descTxt in if tipo = "Capitulo" then FnRemoveAccentsSymbols(capTxt) else null, type text),
            ItemsCapituloFillDown = Table.FillDown(ItemsWithCapitulo, {"Capitulo"}),
            ItemsWithSubcapitulo = Table.AddColumn(ItemsCapituloFillDown, "Subcapitulo", (r as record) => let tipo = Record.Field(r, "TipoFila"), codRaw = Record.Field(r, ItemsCodColName), descRaw = Record.Field(r, ItemsDescColName), codTxt = if codRaw = null then "" else Text.From(codRaw), descTxt = if descRaw = null then "" else Text.From(descRaw), fuenteRaw = if Text.Contains(Text.Upper(codTxt), "SUBCAP") then codTxt else if Text.Contains(Text.Upper(descTxt), "SUBCAP") then descTxt else "", subTxt = if tipo <> "SubCapitulo" or fuenteRaw = "" then null else let baseTxt = if Text.Contains(fuenteRaw, ":") then Text.AfterDelimiter(fuenteRaw, ":") else fuenteRaw in FnRemoveAccentsSymbols(Text.Trim(baseTxt)) in subTxt, type text),
            ItemsSubcapituloFillDown = Table.FillDown(ItemsWithSubcapitulo, {"Subcapitulo"}),
            ItemsWithCodActRaw = Table.AddColumn(ItemsSubcapituloFillDown, "CodigoActRaw", (r as record) => let tipo = Record.Field(r, "TipoFila") in if tipo = "Actividad" then Text.From(Record.Field(r, ItemsCodColName)) else null, type text),
            ItemsCodActRawFillDown = Table.FillDown(ItemsWithCodActRaw, {"CodigoActRaw"}),
            ItemsWithCodigoAct = Table.AddColumn(ItemsCodActRawFillDown, "Codigo act", each FnFormatCodigoAct([CodigoActRaw]), type text),
            ItemsSoloInsumos = Table.SelectRows(ItemsWithCodigoAct, each [TipoFila] = "Insumo"),

            ItemsColsInsumos = Table.ColumnNames(ItemsSoloInsumos),
            CantProySourceCol = if List.Count(ItemsColsInsumos) > 7 then ItemsColsInsumos{7} else null, VTProySourceCol = if List.Count(ItemsColsInsumos) > 9 then ItemsColsInsumos{9} else null, CantConsSourceCol = if List.Count(ItemsColsInsumos) > 19 then ItemsColsInsumos{19} else null, VTConsSourceCol = if List.Count(ItemsColsInsumos) > 21 then ItemsColsInsumos{21} else null,

            ItemsAddCantProy = Table.AddColumn(ItemsSoloInsumos, "Cantidad Proyectado", (r as record) => if CantProySourceCol = null then null else Record.Field(r, CantProySourceCol), type any),
            ItemsAddVTProy = Table.AddColumn(ItemsAddCantProy, "VT Proyectado", (r as record) => if VTProySourceCol = null then null else Record.Field(r, VTProySourceCol), type any),
            ItemsAddCantCons = Table.AddColumn(ItemsAddVTProy, "Cantidad Consumido", (r as record) => if CantConsSourceCol = null then null else Record.Field(r, CantConsSourceCol), type any),
            ItemsAddVTCons = Table.AddColumn(ItemsAddCantCons, "VT Consumido", (r as record) => if VTConsSourceCol = null then null else Record.Field(r, VTConsSourceCol), type any),

            ItemsWithCodigoIns = Table.AddColumn(ItemsAddVTCons, "Codigo ins", each Text.From(Record.Field(_, ItemsCodColName)), type text),
            ItemsWithIns = Table.AddColumn(ItemsWithCodigoIns, "Ins", (r as record) => let descIns = Record.Field(r, ItemsDescColName), umIns = Record.Field(r, ItemsUMColName), dTxt0 = if descIns = null then "" else Text.Trim(Text.From(descIns)), umTxt = if umIns  = null then "" else Text.Trim(Text.From(umIns)), baseTxt = if umTxt = "" then dTxt0 else dTxt0 & " (" & umTxt & ")", clean = FnRemoveAccentsSymbols(baseTxt) in clean, type text),

            // 🔥 LECTURA FORZADA EN UTF-8 (65001)
            HtmlPresupuesto  = Text.FromBinary(Binary.Buffer(BinarioPresupuesto), 65001),
            OrigenSeg        = Table.Buffer(Html.Table(HtmlPresupuesto, Columnas_HTML, [RowSelector="tr"])),
            SegPrepared      = FnPrepareTableWithHeader(OrigenSeg),
            SegCols          = Table.ColumnNames(SegPrepared),
            SegCodCol = if List.Count(List.Select(SegCols, (c) => Text.Contains(Text.Replace(Text.Upper(c), "Ó", "O"), "COD"))) > 0 then List.Select(SegCols, (c) => Text.Contains(Text.Replace(Text.Upper(c), "Ó", "O"), "COD")){0} else SegCols{0},
            SegItemCol = if List.Count(List.Select(SegCols, (c) => Text.Contains(Text.Upper(c), "ITEM"))) > 0 then List.Select(SegCols, (c) => Text.Contains(Text.Upper(c), "ITEM")){0} else if List.Count(SegCols) > 2 then SegCols{2} else SegCols{1},
            SegUMCol = if List.Count(List.Select(SegCols, (c) => Text.Contains(Text.Upper(c), "UM"))) > 0 then List.Select(SegCols, (c) => Text.Contains(Text.Upper(c), "UM")){0} else if List.Count(SegCols) > 3 then SegCols{3} else SegCols{1},

            SegWithTipoFila = Table.AddColumn(SegPrepared, "TipoFilaSeg", (r as record) => let codText = Text.Trim(Text.From(Record.Field(r, SegCodCol) ?? "")), tryNum = try Number.FromText(codText) in if not tryNum[HasError] then "Actividad" else "Otro", type text),
            SegSoloActividades = Table.SelectRows(SegWithTipoFila, each [TipoFilaSeg] = "Actividad"),
            SegWithCodigoAct = Table.AddColumn(SegSoloActividades, "Codigo act", each FnFormatCodigoAct(Record.Field(_, SegCodCol)), type text),
            SegWithActividad = Table.AddColumn(SegWithCodigoAct, "Actividad", (r as record) => let itemTxt = Text.Trim(Text.From(Record.Field(r, SegItemCol) ?? "")), umTxt = Text.Trim(Text.From(Record.Field(r, SegUMCol) ?? "")), codTxt = Text.Trim(Text.From(Record.Field(r, "Codigo act") ?? "")), baseTxt = if umTxt = "" then itemTxt else itemTxt & " (" & umTxt & ")", actTxt = if codTxt = "" then baseTxt else codTxt & "-" & baseTxt in FnRemoveAccentsSymbols(actTxt), type text),

            SegForJoinRaw = Table.SelectColumns(SegWithActividad, {"Codigo act", "Actividad"}),
            SegForJoin    = Table.Distinct(SegForJoinRaw, {"Codigo act"}),

            ItemsJoinSeg = Table.NestedJoin(ItemsWithIns, {"Codigo act"}, SegForJoin, {"Codigo act"}, "Seg", JoinKind.LeftOuter),
            ItemsExpandedSeg = Table.ExpandTableColumn(ItemsJoinSeg, "Seg", {"Actividad"}, {"Actividad"}),

            ItemsNumsTyped = Table.TransformColumns(ItemsExpandedSeg, {{"Cantidad Proyectado", each FnParseNumber(_), type number}, {"VT Proyectado", each FnParseNumber(_), Currency.Type}, {"Cantidad Consumido", each FnParseNumber(_), type number}, {"VT Consumido", each FnParseNumber(_), Currency.Type}}),
            ITEMSINSUMOS_Final = Table.SelectColumns(ItemsNumsTyped, {"Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo", "Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido"})
        in ITEMSINSUMOS_Final,

    // =========================================================
    // EXTRACCIÓN MAESTRA
    // =========================================================
    RutaBase = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv", ArchivosSharePoint = SharePoint.Files(RutaBase, [ApiVersion = 15]),
    ArchivosProyecto = Table.Buffer(Table.SelectRows(ArchivosSharePoint, each Text.Contains(Text.Upper([Folder Path]), "/" & Text.Upper(ParamProyecto) & "/") and Text.EndsWith([Folder Path], "/Actual/"))),
    ConCentroCosto = Table.AddColumn(ArchivosProyecto, "Centro de Costos", each Text.Trim(Text.Replace(Text.AfterDelimiter([Folder Path], "/" & ParamProyecto & "/"), "/Actual/", ""))),
    Agrupado = Table.Group(ConCentroCosto, {"Centro de Costos"}, {{"Binarios", each let FilaPres = Table.SelectRows(_, each Text.Contains(Text.Upper([Name]), "PRESUPUESTO ITEMS")), FilaSeg = Table.SelectRows(_, each Text.Contains(Text.Upper([Name]), "SEGUIMIENTO POR ITEMS")) in if Table.RowCount(FilaPres) > 0 and Table.RowCount(FilaSeg) > 0 then [Bin_P = FilaPres{0}[Content], Bin_S = FilaSeg{0}[Content]] else null}}),
    CentrosCompletos = Table.SelectRows(Agrupado, each [Binarios] <> null),
    TablaConDatos = Table.AddColumn(CentrosCompletos, "Datos", each FxProcesarCentroCosto([Binarios][Bin_S], [Binarios][Bin_P])),
    Expandido = Table.ExpandTableColumn(TablaConDatos, "Datos", {"Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo", "Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido"}),
    ColumnasUtiles = Table.SelectColumns(Expandido, {"Centro de Costos", "Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo", "Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido"}),
    TiposFinales = Table.TransformColumnTypes(ColumnasUtiles,{{"Centro de Costos", type text}, {"Codigo ins", Int64.Type}, {"Cantidad Proyectado", type number}, {"VT Proyectado", Currency.Type}, {"Cantidad Consumido", type number}, {"VT Consumido", Currency.Type}}),
    
    TablaEnMemoria = TiposFinales
in 
    TablaEnMemoria
