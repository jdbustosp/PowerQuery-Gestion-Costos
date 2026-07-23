let
    Funciones = [
        FnFormatCodigoAct = (raw as any) as nullable text =>
            let
                // El #(00A0) (espacio no separable) llega invisible desde los reportes de
                // SharePoint; si no se limpia aqui, "2.041" y "2.041 " (con NBSP) quedan
                // como codigos distintos y rompen el fill-down/join por codigo de actividad.
                txtRaw  = if raw = null then null else Text.Trim(Text.Replace(Text.From(raw), "#(00A0)", " ")),
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
                v              = value,
                isNum          = Value.Is(v, type number),
                numeroDirecto  = if isNum then Number.From(v) else null,
                t0             = if v = null then "" else Text.From(v),
                t              = Text.Trim(Text.Replace(Text.Replace(t0, "#(00A0)", ""), " ", "")),
                tryUS          = try Number.FromText(t, "en-US"),
                valUS          = if tryUS[HasError] then null else tryUS[Value],
                tryES          = try Number.FromText(t, "es-ES"),
                valES          = if tryES[HasError] then null else tryES[Value],
                result         = if numeroDirecto <> null then numeroDirecto
                                 else if t = "" then null
                                 else if valUS <> null then valUS
                                 else valES
            in result,

        FnRemoveAccentsSymbols = (t as any) as nullable text =>
            let
                initial = try (if t = null then null else Text.From(t)) otherwise null,
                replacements = {
                    {"#(00E1)","a"},{"#(00C1)","A"},
                    {"#(00E9)","e"},{"#(00C9)","E"},
                    {"#(00ED)","i"},{"#(00CD)","I"},
                    {"#(00F3)","o"},{"#(00D3)","O"},
                    {"#(00FA)","u"},{"#(00DA)","U"},
                    {"#(00F1)","n"},{"#(00D1)","N"},
                    {"#(00BA)",""},{"#(00B0)",""},{"#(00A8)",""},
                    {"#(lf)", " "}, {"#(cr)", " "}
                },
                result = if initial = null then null
                         else List.Accumulate(replacements, initial, (state, current) => Text.Replace(state, current{0}, current{1}))
            in result,

        FnRemoveAccentMarks = (t as any) as nullable text =>
            let
                initial = try (if t = null then null else Text.From(t)) otherwise null,
                replacements = {
                    {"#(00E1)","a"},{"#(00C1)","A"},
                    {"#(00E9)","e"},{"#(00C9)","E"},
                    {"#(00ED)","i"},{"#(00CD)","I"},
                    {"#(00F3)","o"},{"#(00D3)","O"},
                    {"#(00FA)","u"},{"#(00DA)","U"},
                    {"#(00DC)","U"},{"#(00FC)","u"},
                    {"#(00BA)",""},{"#(00B0)",""},{"#(00A8)",""},
                    {"#(lf)", " "}, {"#(cr)", " "}
                },
                result = if initial = null then null
                         else List.Accumulate(replacements, initial, (state, current) => Text.Replace(state, current{0}, current{1}))
            in result,
        FnClaveLimpia = (t as nullable text) as nullable text =>
            let
                sinUnidad = if t = null then null
                            else if Text.Contains(t, "(") then Text.BeforeDelimiter(t, "(")
                            else t,
                t1 = if sinUnidad = null then null else Text.Upper(Text.Trim(sinUnidad)),
                repl = {
                    {"#(00C1)","A"},{"#(00C9)","E"},{"#(00CD)","I"},
                    {"#(00D3)","O"},{"#(00DA)","U"},{"#(00D1)","N"},{"#(00DC)","U"}
                },
                t2 = if t1 = null then null
                     else List.Accumulate(repl, t1, (state, current) => Text.Replace(state, current{0}, current{1})),
                t3 = if t2 = null then null else Text.Select(t2, {"A".."Z", "0".."9"}),
                result = if t3 = null or t3 = "" then null else t3
            in result,

        FnCleanText = (t as any) as nullable text =>
            try (if t = null then null else let txt = Text.Trim(Text.From(t)) in if txt = "" then null else Text.Upper(txt)) otherwise null,

        FnTrimText = (t as any) as nullable text =>
            try (if t = null then null else Text.Trim(Text.From(t))) otherwise null,

        FnPrepareTableWithHeader = (tbl as table) as table =>
            let
                firstColName   = Table.ColumnNames(tbl){0},
                firstColValues = Table.Column(tbl, firstColName),
                headerFlags    = List.Transform(firstColValues, (x) =>
                    let
                        txt     = Text.Upper(if x = null then "" else Text.From(x)),
                        txtNorm = Text.Replace(txt, "#(00D3)", "O")
                    in Text.Contains(txtNorm, "COD")),
                hasHeader = List.Contains(headerFlags, true),
                promoted  = if hasHeader then
                    let
                        headerIndex = List.PositionOf(headerFlags, true),
                        skipped     = Table.Skip(tbl, headerIndex)
                    in Table.PromoteHeaders(skipped, [PromoteAllScalars = true])
                    else tbl
            in promoted,

        FnEncode = (path as nullable text) as nullable text =>
            if path = null then null
            else Text.Combine(List.Transform(Text.Split(path, "/"), each Uri.EscapeDataString(_)), "/"),

        FnBuildColumnas = (n as number) as list =>
            List.Transform({1..n}, each {"Columna " & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}),

        FnCleanContratista = (t as any) as nullable text =>
            let
                safe       = try (if t = null then null else Text.From(t)) otherwise null,
                t2         = if safe = null then null else Text.Replace(safe, Character.FromNumber(65533), Character.FromNumber(78)),
                t3         = if t2 = null then null else Text.Trim(Text.Upper(t2)),
                repl       = {
                    {Character.FromNumber(193),"A"},{Character.FromNumber(201),"E"},
                    {Character.FromNumber(205),"I"},{Character.FromNumber(211),"O"},
                    {Character.FromNumber(218),"U"},{Character.FromNumber(209),"N"}
                },
                t3_clean   = if t3 = null then null
                             else List.Accumulate(repl, t3, (state, current) => Text.Replace(state, current{0}, current{1})),
                suffixes   = {" S.A.S.", " S.A.S", " SAS.", " SAS", " S.A.", " S.A", " SA.", " SA", " LTDA.", " LTDA", " S EN C", " S. EN C."},
                t4         = if t3_clean = null then null
                             else List.Accumulate(suffixes, t3_clean, (state, suffix) =>
                                 if Text.EndsWith(state, suffix)
                                 then Text.Trim(Text.Range(state, 0, Text.Length(state) - Text.Length(suffix)))
                                 else state),
                result     = if t4 = null or t4 = "" then null else t4
            in result,

        FnMapColumn = (rec as record, cols as list, keywords as list) =>
            let
                norm = (x as any) as text =>
                    let
                        txt = try Text.From(x) otherwise "",
                        clean = FnRemoveAccentsSymbols(txt)
                    in Text.Upper(if clean = null then "" else clean),
                match = List.First(
                    List.Select(cols, (c) =>
                        List.AnyTrue(List.Transform(keywords, (k) => Text.Contains(norm(c), norm(k))))
                    ),
                    null
                )
            in if match = null then null else Record.Field(rec, match),

        FnBuildFolderPrefixMap = (carpetas as list) as record =>
            let
                pares = List.Transform(carpetas, (x) =>
                    let
                        nombre = try Text.From(x) otherwise "",
                        prefix = if Text.Contains(nombre, "-") then Text.Trim(Text.BeforeDelimiter(nombre, "-")) else Text.Trim(nombre)
                    in {prefix, nombre}),
                validos = List.Select(pares, each _{0} <> null and _{0} <> ""),
                tabla = Table.Distinct(Table.FromRows(validos, {"Clave", "Valor"}), {"Clave"})
            in Record.FromList(tabla[Valor], tabla[Clave]),

        FnMatchFolder = (proyectoExcel as text, listaCarpetas as list) as text =>
            let
                count = List.Count(listaCarpetas)
            in
                if count = 0 then proyectoExcel
                else if count = 1 then listaCarpetas{0}
                else
                    let
                        proyClean = FnRemoveAccentsSymbols(Text.Upper(proyectoExcel)),
                        matches   = List.Select(listaCarpetas, each
                            let
                                baseName       = if Text.Contains(_, "-") then Text.Trim(Text.AfterDelimiter(_, "-")) else Text.Trim(_),
                                baseClean      = FnRemoveAccentsSymbols(Text.Upper(baseName)),
                                lastWordFolder = List.Last(Text.Split(baseClean, " ")),
                                lastWordProy   = List.Last(Text.Split(proyClean, " "))
                            in
                                Text.Contains(proyClean, lastWordFolder) or
                                Text.Contains(baseClean, lastWordProy) or
                                Text.Replace(baseClean, " ", "") = Text.Replace(proyClean, " ", "")
                        )
                    in if List.Count(matches) = 1 then matches{0} else proyectoExcel,

        FnReadSPBinary = (siteUrl as text, filePath as text) as nullable binary =>
            let
                raw = try Web.Contents(siteUrl, [
                    RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(filePath) & "')/$value",
                    Headers = [Accept = "*/*"],
                    Timeout = #duration(0, 0, 10, 0),
                    ManualStatusHandling = {404, 429, 500, 502, 503, 504}
                ]) otherwise null,
                status = if raw = null then null else try Value.Metadata(raw)[Response.Status] otherwise 200,
                result = if raw = null or status >= 400 then null else Binary.Buffer(raw)
            in
                result,

        FnReadSPExcel = (siteUrl as text, filePath as text) as nullable table =>
            let
                binario = FnReadSPBinary(siteUrl, filePath),
                libro = if binario = null then null else try Excel.Workbook(binario, null, true) otherwise null,
                data = if libro = null or Table.RowCount(libro) = 0 then null else try libro{0}[Data] otherwise null,
                result = if data = null then null else try Table.PromoteHeaders(data, [PromoteAllScalars=true]) otherwise null
            in
                result,

        FxProcesarCentroCosto = (BinarioSeguimiento as binary, BinarioPresupuesto as binary) as table =>
            let
                Columnas_HTML = FnBuildColumnas(25),
                Columnas_APU  = FnBuildColumnas(3),

                OrigenItems   = try Excel.Workbook(BinarioSeguimiento, null, true){0}[Data]
                                otherwise Html.Table(Text.FromBinary(BinarioSeguimiento, 65001), Columnas_HTML, [RowSelector="tr"]),
                ItemsPrepared = Table.Buffer(FnPrepareTableWithHeader(OrigenItems)),

                ItemsColNames     = Table.ColumnNames(ItemsPrepared),
                ItemsCodColName   = ItemsColNames{0},
                ItemsDescColName  = ItemsColNames{1},
                ItemsTipoColName  = ItemsColNames{2},
                ItemsUMColName    = ItemsColNames{3},

                ItemsWithTipoFila = Table.AddColumn(ItemsPrepared, "TipoFila", (r as record) =>
                    let
                        codValue  = Record.Field(r, ItemsCodColName),
                        descValue = Record.Field(r, ItemsDescColName),
                        tipoValue = Record.Field(r, ItemsTipoColName),
                        umValue   = Record.Field(r, ItemsUMColName),
                        codText   = if codValue  = null then "" else Text.Trim(Text.Replace(Text.From(codValue), "#(00A0)", " ")),
                        descText  = if descValue = null then "" else Text.Trim(Text.From(descValue)),
                        tipoText  = if tipoValue = null then "" else Text.Trim(Text.From(tipoValue)),
                        umText    = if umValue   = null then "" else Text.Trim(Text.From(umValue)),
                        codUpper  = Text.Upper(codText),
                        descUpper = Text.Upper(descText),
                        // Sin este replace, un codigo de Actividad con NBSP pegado hace fallar
                        // Number.FromText -> la fila cae en "Otro" -> el fill-down de abajo
                        // arrastra el codigo de la actividad ANTERIOR sobre un bloque de insumos
                        // que en realidad pertenece a otra actividad (insumos "ajenos").
                        codTextNum = Text.Replace(codText, " ", ""),
                        tryNum    = try Number.FromText(codTextNum),
                        isNumeric = not tryNum[HasError],
                        numValue  = if isNumeric then tryNum[Value] else 0,
                        tipoFila  =
                            if codText = "" then "Otro"
                            else if Text.StartsWith(codUpper, "SUBCAP") or Text.StartsWith(descUpper, "SUBCAP") then "SubCapitulo"
                            else if Text.Contains(codUpper, "CAPITULO") or Text.Contains(descUpper, "CAPITULO") then "Capitulo"
                            else if isNumeric and tipoText = "" and umText = "" and (Text.Length(codText) <= 2 or (numValue >= 1000 and Number.Mod(numValue, 1000) = 0)) then "Capitulo"
                            else if isNumeric and tipoText = "" and umText = "" then "Actividad"
                            else if isNumeric then "Insumo"
                            else "Otro"
                    in tipoFila, type text),

                ItemsWithCapitulo = Table.AddColumn(ItemsWithTipoFila, "Capitulo", (r as record) =>
                    let
                        tipo   = Record.Field(r, "TipoFila"),
                        codRaw = Record.Field(r, ItemsCodColName),
                        descRaw= Record.Field(r, ItemsDescColName),
                        codTxt = if codRaw  = null then "" else Text.Trim(Text.From(codRaw)),
                        descTxt= if descRaw = null then "" else Text.Trim(Text.From(descRaw)),
                        tryN   = try Number.FromText(codTxt),
                        codCap = if codTxt = "00" then codTxt
                                 else if not tryN[HasError] and tryN[Value] >= 1000 and Number.Mod(tryN[Value], 1000) = 0
                                 then Text.From(tryN[Value] / 1000)
                                 else codTxt,
                        capTxt = if descTxt = "" then codCap else codCap & "-" & descTxt
                    in if tipo = "Capitulo" then capTxt else null, type text),

                ItemsCapituloFillDown   = Table.FillDown(ItemsWithCapitulo, {"Capitulo"}),

                ItemsWithSubcapitulo = Table.AddColumn(ItemsCapituloFillDown, "Subcapitulo", (r as record) =>
                    let
                        tipo      = Record.Field(r, "TipoFila"),
                        codRaw    = Record.Field(r, ItemsCodColName),
                        descRaw   = Record.Field(r, ItemsDescColName),
                        codTxt    = if codRaw  = null then "" else Text.From(codRaw),
                        descTxt   = if descRaw = null then "" else Text.From(descRaw),
                        fuenteRaw = if Text.Contains(Text.Upper(codTxt), "SUBCAP") then codTxt
                                    else if Text.Contains(Text.Upper(descTxt), "SUBCAP") then descTxt
                                    else "",
                        subTxt    = if tipo <> "SubCapitulo" or fuenteRaw = "" then null
                                    else let baseTxt = if Text.Contains(fuenteRaw, ":") then Text.AfterDelimiter(fuenteRaw, ":") else fuenteRaw
                                         in Text.Trim(baseTxt)
                    in subTxt, type text),

                ItemsSubcapituloFillDown  = Table.FillDown(ItemsWithSubcapitulo, {"Subcapitulo"}),
                ItemsWithCodActRaw        = Table.AddColumn(ItemsSubcapituloFillDown, "CodigoActRaw", (r as record) =>
                    let tipo = Record.Field(r, "TipoFila") in if tipo = "Actividad" then Text.From(Record.Field(r, ItemsCodColName)) else null, type text),
                // Descripcion real de la fila-Actividad en SEGUIMIENTO POR ITEMS, capturada y
                // arrastrada junto con el codigo. El codigo de APU es una numeracion
                // independiente que puede coincidir con el de SEGUIMIENTO por pura casualidad
                // sin ser la misma actividad; esta descripcion (que SI viene del mismo
                // reporte que trae los insumos) es la fuente confiable del nombre.
                ItemsWithDescActRaw       = Table.AddColumn(ItemsWithCodActRaw, "DescActRaw", (r as record) =>
                    let tipo = Record.Field(r, "TipoFila") in if tipo = "Actividad" then Record.Field(r, ItemsDescColName) else null, type text),
                ItemsCodActRawFillDown    = Table.FillDown(ItemsWithDescActRaw, {"CodigoActRaw", "DescActRaw"}),
                ItemsWithCodigoAct        = Table.AddColumn(ItemsCodActRawFillDown, "Codigo act", each FnFormatCodigoAct([CodigoActRaw]), type text),
                ItemsSoloInsumos          = Table.SelectRows(ItemsWithCodigoAct, each [TipoFila] = "Insumo"),
                ItemsColsInsumos          = Table.ColumnNames(ItemsSoloInsumos),

                CantPresCol = if List.Count(ItemsColsInsumos) > 4  then ItemsColsInsumos{4}  else null,
                VTPresCol   = if List.Count(ItemsColsInsumos) > 6  then ItemsColsInsumos{6}  else null,
                CantProyCol = if List.Count(ItemsColsInsumos) > 7  then ItemsColsInsumos{7}  else null,
                VTProyCol   = if List.Count(ItemsColsInsumos) > 9  then ItemsColsInsumos{9}  else null,
                CantConsCol = if List.Count(ItemsColsInsumos) > 19 then ItemsColsInsumos{19} else null,
                VTConsCol   = if List.Count(ItemsColsInsumos) > 21 then ItemsColsInsumos{21} else null,

                A1 = Table.AddColumn(ItemsSoloInsumos, "Cantidad Presupuesto", (r) => if CantPresCol = null then null else Record.Field(r, CantPresCol)),
                A2 = Table.AddColumn(A1, "VT Presupuesto",    (r) => if VTPresCol   = null then null else Record.Field(r, VTPresCol)),
                A3 = Table.AddColumn(A2, "Cantidad Proyectado",(r) => if CantProyCol = null then null else Record.Field(r, CantProyCol)),
                A4 = Table.AddColumn(A3, "VT Proyectado",     (r) => if VTProyCol   = null then null else Record.Field(r, VTProyCol)),
                A5 = Table.AddColumn(A4, "Cantidad Consumido", (r) => if CantConsCol = null then null else Record.Field(r, CantConsCol)),
                A6 = Table.AddColumn(A5, "VT Consumido",      (r) => if VTConsCol   = null then null else Record.Field(r, VTConsCol)),

                ItemsWithCodigoIns = Table.AddColumn(A6, "Codigo ins", each Text.From(Record.Field(_, ItemsCodColName)), type text),
                ItemsWithIns = Table.AddColumn(ItemsWithCodigoIns, "Ins", (r as record) =>
                    let
                        descIns = Record.Field(r, ItemsDescColName),
                        umIns   = Record.Field(r, ItemsUMColName),
                        dTxt0   = if descIns = null then "" else Text.Trim(Text.From(descIns)),
                        umTxt   = if umIns   = null then "" else Text.Trim(Text.From(umIns)),
                        baseTxt = if umTxt = "" then dTxt0 else dTxt0 & " (" & umTxt & ")"
                    in baseTxt, type text),

                OrigenAPU_Raw = try Excel.Workbook(BinarioPresupuesto, null, true){0}[Data]
                                otherwise Html.Table(Text.FromBinary(BinarioPresupuesto, 65001), Columnas_APU, [RowSelector="tr"]),
                OrigenAPU_Cols = Table.SelectColumns(OrigenAPU_Raw, List.FirstN(Table.ColumnNames(OrigenAPU_Raw), 3)),
                OrigenAPU = Table.RenameColumns(OrigenAPU_Cols, List.Zip({Table.ColumnNames(OrigenAPU_Cols), {"Columna 1", "Columna 2", "Columna 3"}})),

                APU_Paso1 = Table.AddColumn(OrigenAPU, "Cod_Temp", each
                    let
                        c1Value = if [#"Columna 1"] = null then "" else [#"Columna 1"],
                        c1      = Text.Trim(Text.From(c1Value)),
                        hasDash = Text.Contains(c1, "-"),
                        preDash = if hasDash then Text.Trim(Text.BeforeDelimiter(c1, "-")) else "",
                        esNum   = try Number.FromText(preDash) otherwise null
                    in if hasDash and esNum <> null then FnFormatCodigoAct(preDash) else null),

                APU_Paso2 = Table.SelectRows(APU_Paso1, each [Cod_Temp] <> null),

                APU_Diccionario = Table.AddColumn(APU_Paso2, "NombreActAPU", each
                    let
                        c1Value  = if [#"Columna 1"] = null then "" else [#"Columna 1"],
                        rawName  = Text.AfterDelimiter(Text.From(c1Value), "-"),
                        cleanName= Text.Trim(Text.Replace(Text.Replace(Text.Replace(rawName, "#(lf)", " "), "#(cr)", " "), "#(00A0)", " "))
                    in cleanName, type text),

                APU_DiccionarioLimpio = Table.SelectColumns(APU_Diccionario, {"Cod_Temp", "NombreActAPU", "Columna 3"}, MissingField.Ignore),
                APU_DiccionarioRenombrado = Table.RenameColumns(APU_DiccionarioLimpio,
                    List.Select({{"Cod_Temp", "CodigoActAPU"}, {"Columna 3", "UM_Actividad"}}, each Table.HasColumns(APU_DiccionarioLimpio, _{0}))),
                // Ordenar antes de Distinct: si el mismo codigo aparece con 2 nombres reales
                // en el APU (no por NBSP), Distinct siempre se queda con el mismo (el
                // alfabeticamente primero) en vez de uno arbitrario que cambia entre refrescos.
                APU_DiccionarioOrdenado = Table.Sort(APU_DiccionarioRenombrado, {{"NombreActAPU", Order.Ascending}}),
                DiccionarioAPU_Unico = Table.Buffer(Table.Distinct(APU_DiccionarioOrdenado, {"CodigoActAPU"})),

                ItemsJoinAPU     = Table.NestedJoin(ItemsWithIns, {"Codigo act"}, DiccionarioAPU_Unico, {"CodigoActAPU"}, "APU", JoinKind.LeftOuter),
                ItemsExpandedAPU = Table.ExpandTableColumn(ItemsJoinAPU, "APU", {"NombreActAPU", "UM_Actividad"}, {"NombreActAPU", "UM_Actividad"}),

                ItemsWithActividad = Table.AddColumn(ItemsExpandedAPU, "Actividad", each
                    let
                        codTxt        = if [Codigo act]  = null then "" else [Codigo act],
                        // Prioridad: primero la descripcion real de SEGUIMIENTO (misma fuente
                        // que el insumo, siempre correcta para ese codigo); si viene vacia,
                        // el nombre de APU; si tampoco hay, un texto generico.
                        // Se limpia el NBSP y se colapsan espacios aqui mismo (no solo en el
                        // codigo) para que el patron " - Subcapitulo" se detecte de forma
                        // fiable mas abajo, sin importar espacios/caracteres invisibles sueltos
                        // pegados en el texto libre del reporte.
                        descSegRaw    = Text.Trim(Text.Replace(Text.From(if [DescActRaw] = null then "" else [DescActRaw]), "#(00A0)", " ")),
                        descSegTxt    = Text.Combine(List.Select(Text.Split(descSegRaw, " "), each _ <> ""), " "),
                        nombreExtraido= Text.Trim(Text.From(if [NombreActAPU] = null then "" else [NombreActAPU])),
                        nombreReal    = if descSegTxt <> "" then descSegTxt
                                        else if nombreExtraido <> "" then nombreExtraido
                                        else "Actividad " & codTxt,
                        subcapTxt     = Text.Trim(Text.From(if [Subcapitulo] = null then "" else [Subcapitulo])),
                        // Si el Subcapitulo viene pegado con un guion separador ("Texto - SUBCAP"),
                        // se quita el bloque completo (guion incluido) para no dejar un guion
                        // huerfano colgando antes de la unidad. Si no aparece con ese patron
                        // exacto, se cae al comportamiento anterior (quitar solo el texto).
                        conGuion      = if subcapTxt = "" then "" else " - " & subcapTxt,
                        nombreSinSub  = if subcapTxt = "" then nombreReal
                                        else if Text.Contains(nombreReal, conGuion) then Text.Replace(nombreReal, conGuion, "")
                                        else Text.Replace(nombreReal, subcapTxt, ""),
                        umTxt         = Text.Trim(Text.From(if [UM_Actividad] = null then "" else [UM_Actividad])),
                        nombreLimpio  = Text.Combine(List.Select(Text.Split(nombreSinSub, " "), each _ <> ""), " "),
                        actTxt        = if umTxt = "" then codTxt & "-" & nombreLimpio
                                        else codTxt & "-" & nombreLimpio & " (" & umTxt & ")"
                    in actTxt, type text),

                NumsTyped = Table.TransformColumns(ItemsWithActividad, {
                    {"Cantidad Presupuesto", each FxToNumberFlex(_), type number},
                    {"VT Presupuesto",       each FxToNumberFlex(_), Currency.Type},
                    {"Cantidad Proyectado",  each FxToNumberFlex(_), type number},
                    {"VT Proyectado",        each FxToNumberFlex(_), Currency.Type},
                    {"Cantidad Consumido",   each FxToNumberFlex(_), type number},
                    {"VT Consumido",         each FxToNumberFlex(_), Currency.Type}
                }),

                Final = Table.SelectColumns(NumsTyped, {
                    "Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo",
                    "Cantidad Presupuesto", "VT Presupuesto",
                    "Cantidad Proyectado",  "VT Proyectado",
                    "Cantidad Consumido",   "VT Consumido"
                })
            in Final
    ]
in
    Funciones
