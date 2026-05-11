let
    Funciones = [
        FnFormatCodigoAct = (raw as any) as nullable text =>
            let
                txtRaw  = if raw = null then null else Text.Trim(Text.From(raw)),
                result =
                    if txtRaw = null or txtRaw = "" then null
                    else
                        let
                            txtNorm = Text.Replace(Text.Replace(txtRaw, ",", "."), " ", ""),
                            hasDot  = Text.Contains(txtNorm, ".")
                        in
                            if hasDot then txtNorm
                            else
                                let
                                    digits = Text.Select(txtNorm, {"0".."9"}),
                                    len    = Text.Length(digits)
                                in
                                    if len <= 3 then null
                                    else Text.Range(digits, 0, len - 3) & "." & Text.Range(digits, len - 3, 3)
            in result,

        FxToNumberFlex = (value as any) as nullable number =>
            let
                v = value,
                numeroDirecto = if Value.Is(v, type number) then Number.From(v) else null
            in
                if numeroDirecto <> null then numeroDirecto else
                let t0 = if v=null then "" else Text.From(v), t = Text.Trim(Text.Replace(Text.Replace(t0, "#(00A0)", ""), " ", "")) in
                if t="" then null else
                let tryUS = try Number.FromText(t, "en-US"), valUS = if tryUS[HasError] then null else tryUS[Value] in
                if valUS<>null then valUS else
                let tryES = try Number.FromText(t, "es-ES"), valES = if tryES[HasError] then null else tryES[Value] in valES,

        // Incluye Ñ/ñ y mantiene las mojibakes UTF-8 mal interpretadas
        // El try inicial protege contra entradas Error/Record/List que romperian Text.From
        FnRemoveAccentsSymbols = (t as any) as nullable text =>
            let
                initial = try (if t = null then null else Text.From(t)) otherwise null,
                replacements = {
                    {“á”,”a”},{“Á”,”A”},{“é”,”e”},{“É”,”E”},{“í”,”i”},{“Í”,”I”},{“ó”,”o”},{“Ó”,”O”},{“ú”,”u”},{“Ú”,”U”},
                    {“ñ”,”n”},{“Ñ”,”N”},
                    {“º”,””},{“°”,””},{“¨”,””},
                    {“Ã””,”O”}, {“Ã’”,”N”}, {“Ã¡”,”A”}, {“Ã©”,”E”}, {“Ã³”,”O”}, {“Ãº”,”U”}, {“Ã”,”A”},
                    {“#(lf)”, “ “}, {“#(cr)”, “ “}
                },
                result = if initial = null then null else List.Accumulate(replacements, initial, (state, current) => Text.Replace(state, current{0}, current{1}))
            in result,

        FnClaveLimpia = (t as nullable text) as nullable text =>
            if t = null then null
            else let
                sinUnidad = if Text.Contains(t, "(") then Text.BeforeDelimiter(t, "(") else t,
                t1 = Text.Upper(Text.Trim(sinUnidad)),
                t2 = List.Accumulate({{"Á","A"},{"É","E"},{"Í","I"},{"Ó","O"},{"Ú","U"},{"Ñ","N"},{"Ü","U"}}, t1, (state, current) => Text.Replace(state, current{0}, current{1})),
                t3 = Text.Select(t2, {"A".."Z", "0".."9"})
            in if t3 = "" then null else t3,

        // try defensivo: si t es Error/Record/List, Text.From lanza excepcion → devolvemos null
        FnCleanText = (t as any) as nullable text =>
            try (if t = null then null else let txt = Text.Trim(Text.From(t)) in if txt = "" then null else Text.Upper(txt)) otherwise null,

        FnTrimText = (t as any) as nullable text =>
            try (if t = null then null else Text.Trim(Text.From(t))) otherwise null,

        FnPrepareTableWithHeader = (tbl as table) as table =>
            let
                firstColName = Table.ColumnNames(tbl){0},
                firstColValues = Table.Column(tbl, firstColName),
                headerFlags = List.Transform(firstColValues, (x) => let txt = Text.Upper(if x = null then "" else Text.From(x)), txtNorm = Text.Replace(txt, "Ó", "O") in Text.Contains(txtNorm, "COD")),
                hasHeader = List.Contains(headerFlags, true),
                promoted = if hasHeader then let headerIndex = List.PositionOf(headerFlags, true), skipped  = Table.Skip(tbl, headerIndex) in Table.PromoteHeaders(skipped, [PromoteAllScalars = true]) else tbl
            in promoted,

        FnEncode = (path as nullable text) as nullable text =>
            if path = null then null
            else Text.Combine(List.Transform(Text.Split(path, "/"), each Uri.EscapeDataString(_)), "/"),

        FnBuildColumnas = (n as number) as list => List.Transform({1..n}, each {"Columna " & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}),

        FnCleanContratista = (t as any) as nullable text =>
            let
                safe = try (if t = null then null else Text.From(t)) otherwise null,
                t2 = if safe = null then null else Text.Replace(safe, Character.FromNumber(65533), Character.FromNumber(78)),
                t3 = if t2 = null then null else Text.Trim(Text.Upper(t2)),
                replacements = {
                    {Character.FromNumber(193),"A"},{Character.FromNumber(201),"E"},
                    {Character.FromNumber(205),"I"},{Character.FromNumber(211),"O"},
                    {Character.FromNumber(218),"U"},{Character.FromNumber(209),"N"}
                },
                t3_clean = if t3 = null then null else List.Accumulate(replacements, t3, (state, current) => Text.Replace(state, current{0}, current{1})),
                suffixes = {" S.A.S.", " S.A.S", " SAS.", " SAS", " S.A.", " S.A", " SA.", " SA", " LTDA.", " LTDA", " S EN C", " S. EN C."},
                t4 = if t3_clean = null then null else List.Accumulate(suffixes, t3_clean, (state, suffix) =>
                    if Text.EndsWith(state, suffix) then Text.Trim(Text.Range(state, 0, Text.Length(state) - Text.Length(suffix))) else state
                ),
                result = if t4 = null or t4 = "" then null else t4
            in result,

        // Mapea una columna por palabras clave dentro de una lista de nombres de columna (movido desde COMPRAS.m)
        FnMapColumn = (rec as record, cols as list, keywords as list) =>
            let match = List.First(List.Select(cols, (c) => List.AnyTrue(List.Transform(keywords, (k) => Text.Contains(Text.Upper(c), k)))))
            in if match = null then null else Record.Field(rec, match),

        // Construye un Record {prefijoCC -> nombreCarpetaCompleto} para lookup O(1) en mapeo Pamplona 1
        FnBuildFolderPrefixMap = (carpetas as list) as record =>
            let
                pares = List.Transform(carpetas, (x) =>
                    let prefix = if Text.Contains(x, "-") then Text.Trim(Text.BeforeDelimiter(x, "-")) else x
                    in {prefix, x}),
                claves = List.Transform(pares, each _{0}),
                valores = List.Transform(pares, each _{1})
            in
                Record.FromList(valores, claves),

        // Heuristica unificada APROBACIONES_SP / PROVISIONES_SP: ultima palabra coincide
        // o nombre sin espacios coincide. Usa FnRemoveAccentsSymbols para normalizar.
        FnMatchFolder = (proyectoExcel as text, listaCarpetas as list) as text =>
            let
                count = List.Count(listaCarpetas)
            in
                if count = 0 then proyectoExcel
                else if count = 1 then listaCarpetas{0}
                else
                    let
                        proyClean = FnRemoveAccentsSymbols(Text.Upper(proyectoExcel)),
                        matches = List.Select(listaCarpetas, each
                            let
                                baseName = if Text.Contains(_, "-") then Text.Trim(Text.AfterDelimiter(_, "-")) else Text.Trim(_),
                                baseClean = FnRemoveAccentsSymbols(Text.Upper(baseName)),
                                lastWordFolder = List.Last(Text.Split(baseClean, " ")),
                                lastWordProy = List.Last(Text.Split(proyClean, " "))
                            in
                                Text.Contains(proyClean, lastWordFolder) or
                                Text.Contains(baseClean, lastWordProy) or
                                Text.Replace(baseClean, " ", "") = Text.Replace(proyClean, " ", "")
                        )
                    in
                        if List.Count(matches) = 1 then matches{0} else proyectoExcel,

        // Descarga un Excel desde SharePoint y devuelve la primera hoja con headers promocionados.
        // Web.Contents con RelativePath: DataSourcePath estable en siteUrl, evita cache de URLs especificas.
        // Devuelve null si la lectura falla, para que el caller decida la forma de la tabla vacia.
        FnReadSPExcel = (siteUrl as text, filePath as text) as nullable table =>
            let
                binario = Web.Contents(siteUrl, [
                    RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(filePath) & "')/$value"
                ]),
                libro = try Excel.Workbook(Binary.Buffer(binario), null, true) otherwise null
            in
                if libro = null then null else Table.PromoteHeaders(libro{0}[Data], [PromoteAllScalars=true]),

        // Funcion compartida por SP_Seguimiento_Parsed y PPTO_TODOS_PROYECTOS:
        // toma binarios de Seguimiento e APU y devuelve la tabla de insumos limpia.
        FxProcesarCentroCosto = (BinarioSeguimiento as binary, BinarioPresupuesto as binary) as table =>
            let
                Columnas_HTML = FnBuildColumnas(25),
                Columnas_APU  = FnBuildColumnas(3),

                OrigenItems = try Excel.Workbook(BinarioSeguimiento, null, true){0}[Data]
                              otherwise Html.Table(Text.FromBinary(BinarioSeguimiento, 65001), Columnas_HTML, [RowSelector="tr"]),
                ItemsPrepared = Table.Buffer(FnPrepareTableWithHeader(OrigenItems)),

                ItemsColNames = Table.ColumnNames(ItemsPrepared),
                ItemsCodColName = ItemsColNames{0}, ItemsDescColName = ItemsColNames{1}, ItemsTipoColName = ItemsColNames{2}, ItemsUMColName = ItemsColNames{3},

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

                CantPresCol = if List.Count(ItemsColsInsumos) > 4 then ItemsColsInsumos{4} else null,
                VTPresCol   = if List.Count(ItemsColsInsumos) > 6 then ItemsColsInsumos{6} else null,
                CantProyCol = if List.Count(ItemsColsInsumos) > 7 then ItemsColsInsumos{7} else null,
                VTProyCol   = if List.Count(ItemsColsInsumos) > 9 then ItemsColsInsumos{9} else null,
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

                OrigenAPU_Raw = try Excel.Workbook(BinarioPresupuesto, null, true){0}[Data]
                                otherwise Html.Table(Text.FromBinary(BinarioPresupuesto, 65001), Columnas_APU, [RowSelector="tr"]),
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
                APU_DiccionarioLimpio = Table.SelectColumns(APU_Diccionario, {"Cod_Temp", "NombreActAPU", "Columna 3"}, MissingField.Ignore),
                APU_DiccionarioRenombrado = Table.RenameColumns(APU_DiccionarioLimpio,
                    List.Select({{"Cod_Temp", "CodigoActAPU"}, {"Columna 3", "UM_Actividad"}}, each Table.HasColumns(APU_DiccionarioLimpio, _{0}))),
                DiccionarioAPU_Unico = Table.Buffer(Table.Distinct(APU_DiccionarioRenombrado, {"CodigoActAPU"})),

                ItemsJoinAPU = Table.NestedJoin(ItemsWithIns, {"Codigo act"}, DiccionarioAPU_Unico, {"CodigoActAPU"}, "APU", JoinKind.LeftOuter),
                ItemsExpandedAPU = Table.ExpandTableColumn(ItemsJoinAPU, "APU", {"NombreActAPU", "UM_Actividad"}, {"NombreActAPU", "UM_Actividad"}),

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

                NumsTyped = Table.TransformColumns(ItemsWithActividad, {
                    {"Cantidad Presupuesto", each FxToNumberFlex(_), type number},
                    {"VT Presupuesto", each FxToNumberFlex(_), Currency.Type},
                    {"Cantidad Proyectado", each FxToNumberFlex(_), type number},
                    {"VT Proyectado", each FxToNumberFlex(_), Currency.Type},
                    {"Cantidad Consumido", each FxToNumberFlex(_), type number},
                    {"VT Consumido", each FxToNumberFlex(_), Currency.Type}
                }),

                Final = Table.SelectColumns(NumsTyped, {"Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo", "Cantidad Presupuesto", "VT Presupuesto", "Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido"})
            in Final
    ]
in
    Funciones
