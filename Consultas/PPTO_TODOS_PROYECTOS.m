let
    // =========================================================
    // FUNCIONES AUXILIARES GLOBALES
    // =========================================================
    FnFormatCodigoAct = F_Globales[FnFormatCodigoAct],
    FnParseNumber = F_Globales[FxToNumberFlex],
    FnRemoveAccentsSymbols = F_Globales[FnRemoveAccentsSymbols],
    FnPrepareTableWithHeader = F_Globales[FnPrepareTableWithHeader],
    Columnas_HTML = List.Transform({1..25}, each {"Columna " & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}),

    // =========================================================
    // API REST DE SHAREPOINT (TODOS LOS PROYECTOS)
    // =========================================================
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    BasePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/Reportes EDT",
    Headers = [Accept="application/json;odata=nometadata"],

    FnEncode = (path as text) as text => 
        Text.Combine(List.Transform(Text.Split(path, "/"), each Uri.EscapeDataString(_)), "/"),

    // =========================================================
    // FUNCIÓN: PROCESAR CENTRO DE COSTO
    // (Copia exacta de FxProcesarCentroCosto de SP_Seguimiento_Parsed)
    // =========================================================
    FxProcesarCentroCosto = (BinarioSeguimiento as binary, BinarioPresupuesto as binary) =>
        let
            // 🚀 Excel.Workbook es 3-5x más rápido que Html.Table
            OrigenItems = try Excel.Workbook(BinarioSeguimiento, null, true){0}[Data]
                          otherwise Html.Table(Text.FromBinary(BinarioSeguimiento, 65001), Columnas_HTML, [RowSelector="tr"]),
            ItemsPrepared    = Table.Buffer(FnPrepareTableWithHeader(OrigenItems)),

            ItemsColNames    = Table.ColumnNames(ItemsPrepared),
            ItemsCodColName  = ItemsColNames{0}, ItemsDescColName = ItemsColNames{1}, ItemsTipoColName = ItemsColNames{2}, ItemsUMColName = ItemsColNames{3},

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

            // Presupuesto
            CantPresCol = if List.Count(ItemsColsInsumos) > 4 then ItemsColsInsumos{4} else null,
            VTPresCol   = if List.Count(ItemsColsInsumos) > 6 then ItemsColsInsumos{6} else null,
            // Proyectado
            CantProyCol = if List.Count(ItemsColsInsumos) > 7 then ItemsColsInsumos{7} else null,
            VTProyCol   = if List.Count(ItemsColsInsumos) > 9 then ItemsColsInsumos{9} else null,
            // Consumido
            CantConsCol = if List.Count(ItemsColsInsumos) > 19 then ItemsColsInsumos{19} else null,
            VTConsCol   = if List.Count(ItemsColsInsumos) > 21 then ItemsColsInsumos{21} else null,

            A1 = Table.AddColumn(ItemsSoloInsumos, "Cantidad Presupuesto", (r) => if CantPresCol = null then null else Record.Field(r, CantPresCol)),
            A2 = Table.AddColumn(A1, "VT Presupuesto", (r) => if VTPresCol = null then null else Record.Field(r, VTPresCol)),
            A3 = Table.AddColumn(A2, "Cantidad Proyectado", (r) => if CantProyCol = null then null else Record.Field(r, CantProyCol)),
            A4 = Table.AddColumn(A3, "VT Proyectado", (r) => if VTProyCol = null then null else Record.Field(r, VTProyCol)),
            A5 = Table.AddColumn(A4, "Cantidad Consumido", (r) => if CantConsCol = null then null else Record.Field(r, CantConsCol)),
            A6 = Table.AddColumn(A5, "VT Consumido", (r) => if VTConsCol = null then null else Record.Field(r, VTConsCol)),

            ItemsWithCodigoIns = Table.AddColumn(A6, "Codigo ins", each Text.From(Record.Field(_, ItemsCodColName)), type text),
            ItemsWithIns = Table.AddColumn(ItemsWithCodigoIns, "Ins", (r as record) => let descIns = Record.Field(r, ItemsDescColName), umIns = Record.Field(r, ItemsUMColName), dTxt0 = if descIns = null then "" else Text.Trim(Text.From(descIns)), umTxt = if umIns = null then "" else Text.Trim(Text.From(umIns)), baseTxt = if umTxt = "" then dTxt0 else dTxt0 & " (" & umTxt & ")", clean = FnRemoveAccentsSymbols(baseTxt) in clean, type text),

            // 🚀 PARSEO APU - Excel.Workbook (más rápido)
            OrigenAPU_Raw = try Excel.Workbook(BinarioPresupuesto, null, true){0}[Data]
                            otherwise Html.Table(Text.FromBinary(BinarioPresupuesto, 65001), 
                                List.Transform({1..3}, each {"Columna " & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}), [RowSelector="tr"]),
            OrigenAPU_Cols = Table.SelectColumns(OrigenAPU_Raw, List.FirstN(Table.ColumnNames(OrigenAPU_Raw), 3)),
            OrigenAPU = Table.RenameColumns(OrigenAPU_Cols, List.Zip({Table.ColumnNames(OrigenAPU_Cols), {"Columna 1", "Columna 2", "Columna 3"}})),
            
            APU_Paso1 = Table.AddColumn(OrigenAPU, "Cod_Temp", each 
                let 
                    c1Value = if [#"Columna 1"] = null then "" else [#"Columna 1"],
                    c1 = Text.Trim(Text.From(c1Value)),
                    hasDash = Text.Contains(c1, "-"),
                    preDash = if hasDash then Text.Trim(Text.BeforeDelimiter(c1, "-")) else "",
                    esNum = try Number.FromText(preDash) otherwise null
                in if hasDash and esNum <> null then FnFormatCodigoAct(preDash) else null
            ),
            APU_Paso2 = Table.SelectRows(APU_Paso1, each [Cod_Temp] <> null),
            APU_Diccionario = Table.AddColumn(APU_Paso2, "NombreActAPU", each 
                let 
                    c1Value = if [#"Columna 1"] = null then "" else [#"Columna 1"],
                    rawName = Text.AfterDelimiter(Text.From(c1Value), "-"),
                    cleanName = Text.Trim(Text.Replace(Text.Replace(Text.Replace(rawName, "#(lf)", " "), "#(cr)", " "), "#(00A0)", " "))
                in cleanName, type text
            ),
            APU_DiccionarioLimpio = Table.SelectColumns(APU_Diccionario, {"Cod_Temp", "NombreActAPU", "Columna 3"}),
            APU_DiccionarioRenombrado = Table.RenameColumns(APU_DiccionarioLimpio, {{"Cod_Temp", "CodigoActAPU"}, {"Columna 3", "UM_Actividad"}}),
            DiccionarioAPU_Unico = Table.Buffer(Table.Distinct(APU_DiccionarioRenombrado, {"CodigoActAPU"})),

            // CRUCE CONTRA SEGUIMIENTO
            ItemsJoinAPU = Table.NestedJoin(ItemsWithIns, {"Codigo act"}, DiccionarioAPU_Unico, {"CodigoActAPU"}, "APU", JoinKind.LeftOuter),
            ItemsExpandedAPU = Table.ExpandTableColumn(ItemsJoinAPU, "APU", {"NombreActAPU", "UM_Actividad"}, {"NombreActAPU", "UM_Actividad"}),

            // Nombre oficial de Actividad
            ItemsWithActividad = Table.AddColumn(ItemsExpandedAPU, "Actividad", each 
                let 
                    codTxt = if [Codigo act] = null then "" else [Codigo act],
                    nombreExtraido = Text.Trim(Text.From(if [NombreActAPU] = null then "" else [NombreActAPU])),
                    nombreReal = if nombreExtraido = "" then "Actividad " & codTxt else nombreExtraido,
                    subcapTxt = Text.Trim(Text.From(if [Subcapitulo] = null then "" else [Subcapitulo])),
                    nombreSinSubcap = if subcapTxt <> "" then Text.Replace(nombreReal, subcapTxt, "") else nombreReal,
                    umTxt = Text.Trim(Text.From(if [UM_Actividad] = null then "" else [UM_Actividad])),
                    nombreLimpio = Text.Combine(List.Select(Text.Split(nombreSinSubcap, " "), each _ <> ""), " "),
                    actTxt = if umTxt = "" then codTxt & "-" & nombreLimpio else codTxt & "-" & nombreLimpio & " (" & umTxt & ")"
                in FnRemoveAccentsSymbols(actTxt), type text
            ),

            // Parsear números
            NumsTyped = Table.TransformColumns(ItemsWithActividad, {
                {"Cantidad Presupuesto", each FnParseNumber(_), type number},
                {"VT Presupuesto", each FnParseNumber(_), Currency.Type},
                {"Cantidad Proyectado", each FnParseNumber(_), type number},
                {"VT Proyectado", each FnParseNumber(_), Currency.Type},
                {"Cantidad Consumido", each FnParseNumber(_), type number},
                {"VT Consumido", each FnParseNumber(_), Currency.Type}
            }),

            Final = Table.SelectColumns(NumsTyped, {"Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo", "Cantidad Presupuesto", "VT Presupuesto", "Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido"})
        in Final,

    // =========================================================
    // NIVEL 1: LISTAR TODOS LOS PROYECTOS
    // =========================================================
    ProyectosUrl = SiteUrl & "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(BasePath) & "')/Folders?$select=Name",
    ProyectosFolders = Table.FromRecords(Json.Document(Web.Contents(ProyectosUrl, [Headers=Headers]))[value]),

    // =========================================================
    // NIVEL 2: POR CADA PROYECTO, LISTAR CENTROS DE COSTOS
    // =========================================================
    WithCCs = Table.AddColumn(ProyectosFolders, "CCs", each
        let
            projPath = BasePath & "/" & [Name],
            ccsUrl = SiteUrl & "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(projPath) & "')/Folders?$select=Name",
            result = try Json.Document(Web.Contents(ccsUrl, [Headers=Headers])) otherwise null
        in
            if result <> null and (try List.Count(result[value]) otherwise 0) > 0
            then Table.FromRecords(result[value]) else null
    ),
    ValidProjects = Table.SelectRows(WithCCs, each [CCs] <> null),
    ExpandedCCs = Table.ExpandTableColumn(ValidProjects, "CCs", {"Name"}, {"Centro de Costos"}),
    WithProyecto = Table.RenameColumns(ExpandedCCs, {{"Name", "Proyecto"}}),

    // =========================================================
    // NIVEL 3: POR CADA CC, LISTAR ARCHIVOS EN /Actual/
    // =========================================================
    WithFiles = Table.AddColumn(WithProyecto, "Archivos", each
        let
            ccActualPath = BasePath & "/" & [Proyecto] & "/" & [Centro de Costos] & "/Actual",
            filesUrl = SiteUrl & "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(ccActualPath) & "')/Files?$select=Name,ServerRelativeUrl",
            result = try Json.Document(Web.Contents(filesUrl, [Headers=Headers])) otherwise null
        in
            if result <> null and (try List.Count(result[value]) otherwise 0) > 0
            then Table.FromRecords(result[value]) else null
    ),
    ValidCCs = Table.SelectRows(WithFiles, each [Archivos] <> null),
    ExpandedFiles = Table.ExpandTableColumn(ValidCCs, "Archivos", {"Name", "ServerRelativeUrl"}, {"FileName", "ServerRelativeUrl"}),

    // Solo SEGUIMIENTO POR ITEMS y ANALISIS DE PRECIOS UNITARIOS
    Relevant = Table.SelectRows(ExpandedFiles, each
        not Text.StartsWith([FileName], "~$") and (
        Text.Contains([FileName], "SEGUIMIENTO POR ITEMS", Comparer.OrdinalIgnoreCase) or
        Text.Contains([FileName], "ANALISIS DE PRECIOS UNITARIOS", Comparer.OrdinalIgnoreCase))
    ),

    // Descargar contenido con Binary.Buffer (evita re-descargas)
    WithContent = Table.AddColumn(Relevant, "Content", each
        Binary.Buffer(Web.Contents(SiteUrl & "/_api/web/GetFileByServerRelativeUrl('" & FnEncode([ServerRelativeUrl]) & "')/$value"))
    ),

    // =========================================================
    // AGRUPAR POR PROYECTO + CC Y PROCESAR
    // =========================================================
    Agrupado = Table.Group(WithContent, {"Proyecto", "Centro de Costos"}, {{"Binarios", each
        let
            FilaSeg = Table.SelectRows(_, each Text.Contains([FileName], "SEGUIMIENTO POR ITEMS", Comparer.OrdinalIgnoreCase)),
            FilaPres = Table.SelectRows(_, each Text.Contains([FileName], "ANALISIS DE PRECIOS UNITARIOS", Comparer.OrdinalIgnoreCase))
        in if Table.RowCount(FilaSeg) > 0 and Table.RowCount(FilaPres) > 0
           then [Bin_S = FilaSeg{0}[Content], Bin_P = FilaPres{0}[Content]]
           else null
    }}),
    CentrosCompletos = Table.SelectRows(Agrupado, each [Binarios] <> null),

    // Aplicar FxProcesarCentroCosto a cada par de binarios
    TablaConDatos = Table.AddColumn(CentrosCompletos, "Datos", each 
        FxProcesarCentroCosto([Binarios][Bin_S], [Binarios][Bin_P])),

    // =========================================================
    // EXPANDIR Y APLICAR LÓGICA PPTO_BD
    // =========================================================
    Expandido = Table.ExpandTableColumn(TablaConDatos, "Datos", 
        {"Codigo ins", "Ins", "Actividad", "Capitulo", "Subcapitulo", "Cantidad Presupuesto", "VT Presupuesto"}),

    // V/U Presupuesto (VT / Cantidad)
    AddVU = Table.AddColumn(Expandido, "V/U Presupuesto", each 
        if [Cantidad Presupuesto] = null or [Cantidad Presupuesto] = 0 or [VT Presupuesto] = null then null 
        else [VT Presupuesto] / [Cantidad Presupuesto], Currency.Type),

    // Filtrar donde VT Presupuesto <> 0
    Filtered = Table.SelectRows(AddVU, each [VT Presupuesto] <> null and [VT Presupuesto] <> 0),

    // Selección y orden final de columnas (sin Codigo act, sin Tipo)
    Selected = Table.SelectColumns(Filtered, 
        {"Proyecto", "Centro de Costos", "Subcapitulo", "Capitulo", "Actividad", "Codigo ins", "Ins", 
         "Cantidad Presupuesto", "VT Presupuesto", "V/U Presupuesto"}),

    // Tipos de datos finales
    Typed = Table.TransformColumnTypes(Selected, {
        {"Proyecto", type text}, {"Centro de Costos", type text}, 
        {"Subcapitulo", type text}, {"Capitulo", type text}, {"Actividad", type text},
        {"Codigo ins", Int64.Type}, {"Ins", type text},
        {"Cantidad Presupuesto", type number}, {"VT Presupuesto", Currency.Type}, 
        {"V/U Presupuesto", Currency.Type}
    }),

    TablaFinal = Typed
in
    TablaFinal

