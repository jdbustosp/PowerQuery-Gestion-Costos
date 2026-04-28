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
            HtmlSeguimiento  = Text.FromBinary(BinarioSeguimiento, 65001),
            OrigenItems      = Html.Table(HtmlSeguimiento, Columnas_HTML, [RowSelector="tr"]),
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

            // 🔥 LECTURA FORZADA EN UTF-8 (65001) - NUEVO DICCIONARIO APU
            HtmlAPU  = Text.FromBinary(BinarioPresupuesto, 65001),
            OrigenAPU = Html.Table(HtmlAPU, Columnas_HTML, [RowSelector="tr"]),
            
            APU_Paso1 = Table.AddColumn(OrigenAPU, "Cod_Temp", each 
                let 
                    c1 = Text.Trim(Text.From([#"Columna 1"] ?? "")),
                    hasDash = Text.Contains(c1, "-"),
                    preDash = if hasDash then Text.Trim(Text.BeforeDelimiter(c1, "-")) else "",
                    esNum = try Number.FromText(preDash) otherwise null
                in if hasDash and esNum <> null then FnFormatCodigoAct(preDash) else null
            ),
            
            APU_Paso2 = Table.SelectRows(APU_Paso1, each [Cod_Temp] <> null),
            
            APU_Diccionario = Table.AddColumn(APU_Paso2, "NombreActAPU", each 
                let 
                    rawName = Text.AfterDelimiter(Text.From([#"Columna 1"] ?? ""), "-"),
                    cleanName = Text.Trim(Text.Replace(Text.Replace(Text.Replace(rawName, "#(lf)", " "), "#(cr)", " "), "#(00A0)", " "))
                in cleanName, type text
            ),
            
            APU_DiccionarioLimpio = Table.SelectColumns(APU_Diccionario, {"Cod_Temp", "NombreActAPU", "Columna 3"}),
            APU_DiccionarioRenombrado = Table.RenameColumns(APU_DiccionarioLimpio, {{"Cod_Temp", "CodigoActAPU"}, {"Columna 3", "UM_Actividad"}}),
            DiccionarioAPU_Unico = Table.Distinct(APU_DiccionarioRenombrado, {"CodigoActAPU"}),

            // CRUCE CONTRA SEGUIMIENTO
            ItemsJoinAPU = Table.NestedJoin(ItemsWithIns, {"Codigo act"}, DiccionarioAPU_Unico, {"CodigoActAPU"}, "APU", JoinKind.LeftOuter),
            ItemsExpandedAPU = Table.ExpandTableColumn(ItemsJoinAPU, "APU", {"NombreActAPU", "UM_Actividad"}, {"NombreActAPU", "UM_Actividad"}),

            // 🔥 Creamos el nombre oficial de la Actividad (Código - Nombre (UM))
            ItemsExpandedSeg = Table.AddColumn(ItemsExpandedAPU, "Actividad", each 
                let 
                    codTxt = [Codigo act] ?? "",
                    nombreExtraido = Text.Trim(Text.From([NombreActAPU] ?? "")),
                    nombreReal = if nombreExtraido = "" then "Actividad " & codTxt else nombreExtraido,
                    
                    subcapTxt = Text.Trim(Text.From([Subcapitulo] ?? "")),
                    nombreSinSubcap = if subcapTxt <> "" then Text.Replace(nombreReal, subcapTxt, "") else nombreReal,
                    
                    umTxt = Text.Trim(Text.From([UM_Actividad] ?? "")),
                    
                    nombreLimpio = Text.Combine(List.Select(Text.Split(nombreSinSubcap, " "), each _ <> ""), " "),
                    actTxt = if umTxt = "" then codTxt & "-" & nombreLimpio else codTxt & "-" & nombreLimpio & " (" & umTxt & ")"
                in FnRemoveAccentsSymbols(actTxt), type text
            ),

            ItemsNumsTyped = Table.TransformColumns(ItemsExpandedSeg, {{"Cantidad Proyectado", each FnParseNumber(_), type number}, {"VT Proyectado", each FnParseNumber(_), Currency.Type}, {"Cantidad Consumido", each FnParseNumber(_), type number}, {"VT Consumido", each FnParseNumber(_), Currency.Type}}),
            ITEMSINSUMOS_Final = Table.SelectColumns(ItemsNumsTyped, {"Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo", "Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido"})
        in ITEMSINSUMOS_Final,

    // =========================================================
    // EXTRACCIÓN MAESTRA (LECTURA DESDE CONSULTA COMPARTIDA)
    // =========================================================
    ConCentroCosto = SP_Archivos_Proyecto,
    Agrupado = Table.Group(ConCentroCosto, {"Centro de Costos"}, {{"Binarios", each let FilaPres = Table.SelectRows(_, each Text.Contains(Text.Upper([Name]), "ANALISIS DE PRECIOS UNITARIOS")), FilaSeg = Table.SelectRows(_, each Text.Contains(Text.Upper([Name]), "SEGUIMIENTO POR ITEMS")) in if Table.RowCount(FilaPres) > 0 and Table.RowCount(FilaSeg) > 0 then [Bin_P = FilaPres{0}[Content], Bin_S = FilaSeg{0}[Content]] else null}}),
    CentrosCompletos = Table.SelectRows(Agrupado, each [Binarios] <> null),
    TablaConDatos = Table.AddColumn(CentrosCompletos, "Datos", each FxProcesarCentroCosto([Binarios][Bin_S], [Binarios][Bin_P])),
    Expandido = Table.ExpandTableColumn(TablaConDatos, "Datos", {"Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo", "Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido"}),
    ColumnasUtiles = Table.SelectColumns(Expandido, {"Centro de Costos", "Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo", "Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido"}),
    TiposFinales = Table.TransformColumnTypes(ColumnasUtiles,{{"Centro de Costos", type text}, {"Codigo ins", Int64.Type}, {"Cantidad Proyectado", type number}, {"VT Proyectado", Currency.Type}, {"Cantidad Consumido", type number}, {"VT Consumido", Currency.Type}}),
    
    TablaEnMemoria = TiposFinales
in 
    TablaEnMemoria
