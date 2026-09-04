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

        // Normaliza claves de texto para cruces entre fuentes (ej. "# CC - Comparativo"):
        // 1) espacios duros (nbsp) -> normales, 2) colapsa espacios repetidos internos,
        // 3) quita espacios alrededor de guiones ("002 - NOMBRE" -> "002-NOMBRE", forma
        //    compacta que usa la tabla manual Det_CC), 4) recorta extremos.
        // Debe aplicarse EN AMBOS LADOS de un join con la misma funcion.
        FnNormalizeSpaces = (t as any) as nullable text =>
            try (
                if t = null then null
                else
                    let
                        txt = Text.Replace(Text.From(t), "#(00A0)", " "),
                        partes = List.Select(Text.Split(txt, " "), each _ <> ""),
                        unido = Text.Combine(partes, " "),
                        sinEspGuion = Text.Replace(Text.Replace(unido, " -", "-"), "- ", "-")
                    in if sinEspGuion = "" then null else sinEspGuion
            ) otherwise null,

        // Decodifica un binario HTML/texto de los reportes SINCO detectando la codificacion:
        // intenta UTF-8 y, si el resultado trae el caracter de reemplazo U+FFFD (tipico de
        // decodificar Latin-1/Windows-1252 como UTF-8: la enie y tildes se vuelven "?"),
        // re-decodifica como ISO-8859-1. Los reportes de SINCO/Oracle vienen mezclados en
        // ambas codificaciones segun el modulo que los exporta — NUNCA usar un codepage
        // fijo en Text.FromBinary para estos archivos, usar siempre esta funcion.
        FnDecodeHtml = (bin as binary) as text =>
            let
                buf = Binary.Buffer(bin),
                utf8 = try Text.FromBinary(buf, TextEncoding.Utf8) otherwise null,
                usarLatin1 = utf8 = null or Text.Contains(utf8, "#(FFFD)"),
                result = if usarLatin1 then Text.FromBinary(buf, 28591) else utf8
            in
                result,

        // ============================================================
        // Subcapitulo embebido en el nombre (proyectos tipo TURPIAL)
        // ============================================================
        // Sufijos de "frente"/especialidad que el reporte agrega DESPUES del
        // subcapitulo real ("... - CUARTO DE BASURAS - ELECTRICO"): NO son
        // subcapitulos. Se reconocen tambien sus truncaduras (ELE, ELEC...).
        SubcapSufijosIgnorados = {"ELECTRICO"},
        // Truncaduras irrecuperables (texto cortado en TODOS los reportes),
        // confirmadas por el usuario: derivado -> subcapitulo real.
        SubcapOverrides = [#"APTOS (U" = "TORRES"],

        FnEsSufijoSubcapIgnorado = (t as text) as logical =>
            let norm = Text.Upper(FnRemoveAccentsSymbols(t))
            in List.AnyTrue(List.Transform(SubcapSufijosIgnorados, (s) =>
                norm = s or (Text.Length(norm) >= 3 and Text.StartsWith(s, norm)))),

        FnAplicarOverrideSubcap = (v as nullable text) as nullable text =>
            if v = null then null
            else let o = try Record.Field(SubcapOverrides, v) otherwise null
                 in if o <> null then o else v,

        // Quita guiones sueltos colgando al final de un texto (recursivo).
        FnQuitarGuionFinal = (t as text) as text =>
            let r = Text.Trim(t)
            in if Text.EndsWith(r, "-") then @FnQuitarGuionFinal(Text.Range(r, 0, Text.Length(r) - 1)) else r,

        // Extrae la cola tras el ultimo " - " (con limpieza de guion colgante),
        // saltando sufijos ignorados de forma recursiva. null si no hay cola valida.
        FnExtraerSubcapDeTexto = (txt as text) as nullable text =>
            let
                pos    = Text.PositionOf(txt, " - ", Occurrence.Last),
                cola   = if pos < 0 then "" else FnQuitarGuionFinal(Text.Trim(Text.Range(txt, pos + 3))),
                cabeza = if pos < 0 then "" else Text.Trim(Text.Range(txt, 0, pos)),
                valida = cola <> "" and cabeza <> "" and Text.Length(cola) <= 60
            in
                if not valida then null
                else if FnEsSufijoSubcapIgnorado(cola) then @FnExtraerSubcapDeTexto(cabeza)
                else cola,

        // Quita del final de un nombre los sufijos ignorados (" - ELECTRICO").
        FnQuitarSufijosSubcapIgnorados = (txt as text) as text =>
            let
                pos    = Text.PositionOf(txt, " - ", Occurrence.Last),
                cola   = if pos < 0 then "" else FnQuitarGuionFinal(Text.Trim(Text.Range(txt, pos + 3))),
                cabeza = if pos < 0 then "" else Text.Trim(Text.Range(txt, 0, pos))
            in if pos >= 0 and cola <> "" and FnEsSufijoSubcapIgnorado(cola) then @FnQuitarSufijosSubcapIgnorados(cabeza) else txt,

        // Separa "NOMBRE - SUBCAP (UM)" en [Nombre, Subcap]: desprende la unidad final
        // "(UM)" si existe, quita sufijos ignorados, extrae el subcapitulo (con overrides)
        // y devuelve el nombre sin el subcapitulo, re-anexando la unidad si no quedo ya.
        // Para nombres cuyo subcapitulo viene DESPUES de la unidad ("X (M3) - TANQUE")
        // o antes ("X - TANQUE (M3)") funciona en ambos ordenes.
        FnSepararSubcapDeNombre = (nombreRaw as nullable text) as record =>
            let
                txt0   = if nombreRaw = null then "" else Text.Trim(Text.Replace(Text.From(nombreRaw), "#(00A0)", " ")),
                txt    = Text.Combine(List.Select(Text.Split(txt0, " "), each _ <> ""), " "),
                // desprender "(UM)" final si existe
                tieneUM = Text.EndsWith(txt, ")") and Text.PositionOf(txt, "(", Occurrence.Last) >= 0,
                posPar  = if tieneUM then Text.PositionOf(txt, "(", Occurrence.Last) else -1,
                um      = if tieneUM then Text.Range(txt, posPar) else "",
                cuerpo0 = if tieneUM then Text.Trim(Text.Range(txt, 0, posPar)) else txt,
                cuerpo  = FnQuitarGuionFinal(FnQuitarSufijosSubcapIgnorados(cuerpo0)),
                subcapX = FnExtraerSubcapDeTexto(cuerpo),
                subcap  = FnAplicarOverrideSubcap(subcapX),
                posTail = if subcapX = null then -1 else Text.PositionOf(cuerpo, " - ", Occurrence.Last),
                cabeza  = if posTail < 0 then cuerpo else FnQuitarGuionFinal(Text.Trim(Text.Range(cuerpo, 0, posTail))),
                nombreF0 = if subcapX = null then cuerpo else cabeza,
                nombreF  = if um = "" or Text.EndsWith(nombreF0, um) then nombreF0 else nombreF0 & " " & um
            in
                [Nombre = nombreF, Subcap = subcap],

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
                    Timeout = #duration(0, 0, 2, 0),
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
                                otherwise Html.Table(FnDecodeHtml(BinarioSeguimiento), Columnas_HTML, [RowSelector="tr"]),
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

                // true si el SEGUIMIENTO trae al menos una fila explicita "SUBCAPITULO:".
                // Proyectos como TURPIAL no las traen: alli el subcapitulo viene pegado
                // como sufijo del nombre de la actividad ("... - GENERALES") y se deriva
                // mas abajo (SubcapDerivado). El gate evita aplicar esa heuristica en
                // proyectos que si declaran subcapitulos (alli un " - X" final puede ser
                // parte legitima del nombre y no un subcapitulo).
                TieneSubcapExplicito = List.Contains(List.Buffer(Table.Column(ItemsWithTipoFila, "TipoFila")), "SubCapitulo"),
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
                                otherwise Html.Table(FnDecodeHtml(BinarioPresupuesto), Columnas_APU, [RowSelector="tr"]),
                OrigenAPU_Cols = Table.SelectColumns(OrigenAPU_Raw, List.FirstN(Table.ColumnNames(OrigenAPU_Raw), 3)),
                OrigenAPU = Table.RenameColumns(OrigenAPU_Cols, List.Zip({Table.ColumnNames(OrigenAPU_Cols), {"Columna 1", "Columna 2", "Columna 3"}})),

                APU_Paso1 = Table.AddColumn(OrigenAPU, "Cod_Temp", each
                    let
                        c1Value = if [#"Columna 1"] = null then "" else [#"Columna 1"],
                        c1      = Text.Trim(Text.From(c1Value)),
                        hasDash = Text.Contains(c1, "-"),
                        preDash = if hasDash then Text.Trim(Text.BeforeDelimiter(c1, "-")) else "",
                        // El archivo APU mezcla, en la misma columna, filas de Actividad
                        // (codigo YA con punto, ej. "10.002") con filas de detalle de
                        // material/insumo del catalogo (codigo entero SIN punto, ej. "11016",
                        // "6400"). FnFormatCodigoAct le inserta un punto a los enteros sueltos
                        // ("11016" -> "11.016"), lo que hace que un codigo de MATERIAL choque
                        // por pura casualidad numerica con un codigo de ACTIVIDAD distinto.
                        // Por eso solo se acepta como codigo de actividad si YA trae el punto
                        // en el texto original: descarta los codigos de material del catalogo.
                        tienePunto = Text.Contains(preDash, "."),
                        esNum   = try Number.FromText(preDash) otherwise null
                    in if hasDash and esNum <> null and tienePunto then FnFormatCodigoAct(preDash) else null),

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
                // Un mismo codigo puede tener MAS DE UNA actividad real distinta en el APU
                // (ej. "2.041" = "Esmaltado Parqueadero" (m2) en un bloque y "Cinta PVC" (ml)
                // en otro). No se reduce a un solo candidato por codigo (Table.Distinct no
                // garantiza cual "gana" de forma confiable ni estable entre refrescos); se
                // agrupan TODOS los candidatos por codigo y, para cada insumo, se elige el
                // candidato cuyo nombre realmente coincide con la descripcion propia de
                // SEGUIMIENTO en esa fila (misma fuente que el insumo, asi que es la senal
                // mas confiable para saber cual de las actividades es).
                DiccionarioAPU_Candidatos = Table.Buffer(Table.Group(APU_DiccionarioRenombrado, {"CodigoActAPU"}, {
                    {"CandidatosAPU", each Table.SelectColumns(_, {"NombreActAPU", "UM_Actividad"}), type table}
                })),

                ItemsJoinAPU       = Table.NestedJoin(ItemsWithIns, {"Codigo act"}, DiccionarioAPU_Candidatos, {"CodigoActAPU"}, "APU", JoinKind.LeftOuter),
                ItemsExpandedAPU0  = Table.ExpandTableColumn(ItemsJoinAPU, "APU", {"CandidatosAPU"}, {"CandidatosAPU"}),
                // Normaliza para comparar: SEGUIMIENTO trae el separador con guion normal
                // ("Cinta PVC - PARQUEADERO BLOQUE A") mientras que APU lo trae con NBSP y
                // sin guion ("Cinta PVC[NBSP]PARQUEADERO BLOQUE A"). Sin normalizar esa
                // diferencia, ninguno de los 2 textos "contiene" al otro y la coincidencia
                // nunca se detecta.
                FnNormalizarParaComparar = (t as any) as text =>
                    let
                        base      = Text.Upper(Text.Trim(Text.From(if t = null then "" else t))),
                        sinNBSP   = Text.Replace(base, "#(00A0)", " "),
                        sinGuion  = Text.Replace(sinNBSP, " - ", " "),
                        colapsado = Text.Combine(List.Select(Text.Split(sinGuion, " "), each _ <> ""), " ")
                    in colapsado,

                // Algunos nombres de actividad vienen del reporte con un guion final
                // colgando y nada detras (p.ej. "Remate muros - Aptos -"), sin que exista
                // Subcapitulo alguno que explique ese guion (viene asi de crudo en
                // DescActRaw/NombreActAPU). Se quita cualquier guion suelto al final del
                // nombre, de forma recursiva por si quedara mas de uno.
                FnQuitarGuionColgante = (t as text) as text =>
                    let
                        recortado = Text.Trim(t),
                        limpio    = if Text.EndsWith(recortado, "-")
                                    then @FnQuitarGuionColgante(Text.Range(recortado, 0, Text.Length(recortado) - 1))
                                    else recortado
                    in limpio,

                // Mismo problema que el guion final, pero al INICIO del nombre
                // (p.ej. "-SC - TOPELLANTAS (Un) - URBANISMO INTERIOR -" trae un
                // guion colgando antes de "SC" sin nada delante que lo explique).
                FnQuitarGuionInicial = (t as text) as text =>
                    let
                        recortado = Text.Trim(t),
                        limpio    = if Text.StartsWith(recortado, "-")
                                    then @FnQuitarGuionInicial(Text.Range(recortado, 1))
                                    else recortado
                    in limpio,

                // Algunos nombres traen la unidad ya incrustada en el texto libre
                // (ej. "TOPELLANTAS (Un) - URBANISMO INTERIOR"), redundante con la
                // unidad que esta misma consulta agrega al final entre parentesis.
                // Si el parentesis embebido coincide (sin distinguir mayusculas) con
                // la unidad final, se quita para no duplicarla; si no coincide (ej.
                // "(bloque fachada)" cuando la unidad final es "M2") se deja intacto
                // porque es contenido real del nombre, no una unidad repetida.
                FnQuitarUnidadEmbebida = (t as text, um as text) as text =>
                    let
                        patron    = "(" & um & ")",
                        tUpper    = Text.Upper(t),
                        pos       = if um = "" then -1 else Text.PositionOf(tUpper, Text.Upper(patron)),
                        sinUnidad = if pos < 0 then t else Text.RemoveRange(t, pos, Text.Length(patron)),
                        colapsado = Text.Combine(List.Select(Text.Split(sinUnidad, " "), each _ <> ""), " ")
                    in colapsado,

                ItemsConAPUElegido = Table.AddColumn(ItemsExpandedAPU0, "APUElegido", each
                    let
                        candidatos      = [CandidatosAPU],
                        hayCandidatos   = candidatos <> null and Table.RowCount(candidatos) > 0,
                        descPropia      = FnNormalizarParaComparar(if [DescActRaw] = null then "" else [DescActRaw]),
                        conCoincidencia = if not hayCandidatos or descPropia = "" then null
                            else Table.SelectRows(candidatos, each
                                let nombreNorm = FnNormalizarParaComparar([NombreActAPU])
                                in Text.Contains(descPropia, nombreNorm) or Text.Contains(nombreNorm, descPropia)),
                        tieneCoincidencia = conCoincidencia <> null and Table.RowCount(conCoincidencia) > 0
                    in
                        if not hayCandidatos then [NombreActAPU = null, UM_Actividad = null]
                        else if tieneCoincidencia then conCoincidencia{0}
                        else candidatos{0}),
                ItemsExpandedAPU   = Table.ExpandRecordColumn(ItemsConAPUElegido, "APUElegido", {"NombreActAPU", "UM_Actividad"}, {"NombreActAPU", "UM_Actividad"}),

                // Subcapitulo DERIVADO del nombre de la actividad, solo cuando el proyecto
                // no declara subcapitulos en el SEGUIMIENTO (ej. TURPIAL): el reporte de
                // "Analisis De Precios Unitarios Items Presupuesto Detallado" agrega el
                // subcapitulo como sufijo tras el ultimo " - " ("1.01 - COMISION ... - GENERALES").
                // Se toma la cola despues del ULTIMO " - " como subcapitulo; el recorte del
                // nombre lo hace la logica ya existente en ItemsWithActividad (patron conGuion).
                // Helpers de subcapitulo embebido: definidos a nivel de F_Globales
                // (los usa tambien DISPONIBLE para los POR ADJUDICAR); aqui solo un alias.
                FnQuitarSufijosIgnorados = FnQuitarSufijosSubcapIgnorados,

                ItemsConSubcapDerivado = Table.AddColumn(ItemsExpandedAPU, "SubcapDerivado", each
                    let
                        subcapSeg = if [Subcapitulo] = null then "" else Text.Trim(Text.From([Subcapitulo])),
                        aplica    = (TieneSubcapExplicito = false) and subcapSeg = "",
                        descRaw   = Text.Trim(Text.Replace(Text.From(if [DescActRaw] = null then "" else [DescActRaw]), "#(00A0)", " ")),
                        descTxt   = Text.Combine(List.Select(Text.Split(descRaw, " "), each _ <> ""), " "),
                        apuRaw    = Text.Trim(Text.Replace(Text.From(if [NombreActAPU] = null then "" else [NombreActAPU]), "#(00A0)", " ")),
                        apuTxt    = Text.Combine(List.Select(Text.Split(apuRaw, " "), each _ <> ""), " "),
                        desdeDesc = if (not aplica) or descTxt = "" then null else FnExtraerSubcapDeTexto(descTxt),
                        // Un "(" sin ")" en la cola delata texto TRUNCADO por el reporte
                        // ("APTOS (U" en vez de "...APTOS (Un) - TORRES"): en ese caso se
                        // intenta con el nombre del APU (otro reporte, normalmente completo).
                        truncado  = desdeDesc <> null and Text.Contains(desdeDesc, "(") and not Text.Contains(desdeDesc, ")"),
                        desdeApu  = if (not aplica) or apuTxt = "" then null else FnExtraerSubcapDeTexto(apuTxt),
                        elegido   = if not aplica then null
                                    else if desdeDesc = null or truncado then (if desdeApu <> null then desdeApu else desdeDesc)
                                    else desdeDesc
                    in FnAplicarOverrideSubcap(elegido), type text),

                // Canonicaliza subcapitulos derivados TRUNCADOS por el reporte (SINCO corta
                // el texto en algunas filas: "ELE", "ELEC", "ELECTRIC" en vez de "ELECTRICO";
                // "SALON SOCIA" en vez de "SALON SOCIAL"). Regla: si un valor derivado es
                // PREFIJO (sin tildes, sin mayusculas) de otro valor derivado MAS LARGO y
                // MAS FRECUENTE del mismo archivo, se reemplaza por el completo. Un subcapitulo
                // real corto no se toca salvo que exista uno mas largo que empiece igual y
                // tenga mas filas (las truncaduras son artefactos de pocas filas).
                SubcapDerivadosLista = List.Buffer(
                    let
                        vals = List.RemoveNulls(Table.Column(ItemsConSubcapDerivado, "SubcapDerivado")),
                        dist = List.Distinct(vals)
                    in List.Transform(dist, (d) => [V = d, N = List.Count(List.Select(vals, (x) => x = d)), Norm = Text.Upper(FnRemoveAccentsSymbols(d))])),

                FnCanonSubcap = (v as nullable text) as nullable text =>
                    if v = null then null else
                    let
                        normV  = Text.Upper(FnRemoveAccentsSymbols(v)),
                        propio = List.First(List.Select(SubcapDerivadosLista, each [Norm] = normV), [N = 0]),
                        cands  = List.Select(SubcapDerivadosLista, each [Norm] <> normV and Text.StartsWith([Norm], normV) and [N] > propio[N]),
                        mejor  = if List.Count(cands) = 0 then null
                                 else List.Accumulate(cands, null, (s, c) => if s = null or Text.Length(c[Norm]) > Text.Length(s[Norm]) then c else s)
                    in if mejor = null then v else mejor[V],

                ItemsWithActividad = Table.AddColumn(ItemsConSubcapDerivado, "Actividad", each
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
                        nombreReal0   = if descSegTxt <> "" then descSegTxt
                                        else if nombreExtraido <> "" then nombreExtraido
                                        else "Actividad " & codTxt,
                        // Los sufijos ignorados ("- ELECTRICO") tampoco van en el nombre.
                        nombreReal    = FnQuitarSufijosIgnorados(nombreReal0),
                        // Subcapitulo normalizado igual que la descripcion (NBSP->espacio,
                        // espacios colapsados) para que la busqueda de mas abajo no falle por
                        // una diferencia de espacios/caracteres invisibles entre los 2 campos
                        // (vienen de columnas distintas del mismo reporte, no siempre coinciden
                        // caracter por caracter aunque se vean iguales).
                        subcapFuenteSeg = if [Subcapitulo] = null then "" else Text.Trim(Text.From([Subcapitulo])),
                        subcapFuente  = if subcapFuenteSeg <> "" then subcapFuenteSeg
                                        else Text.From(if [SubcapDerivado] = null then "" else [SubcapDerivado]),
                        subcapRaw     = Text.Trim(Text.Replace(subcapFuente, "#(00A0)", " ")),
                        subcapTxt     = Text.Combine(List.Select(Text.Split(subcapRaw, " "), each _ <> ""), " "),
                        // Si el Subcapitulo viene pegado con un guion separador ("Texto - SUBCAP"),
                        // se quita el bloque completo (guion incluido) para no dejar un guion
                        // huerfano colgando antes de la unidad. Si no aparece con ese patron
                        // exacto, se cae al comportamiento anterior (quitar solo el texto).
                        // La busqueda es insensible a mayusculas (Subcapitulo y la descripcion
                        // no siempre coinciden en mayusculas/minusculas), pero el recorte se
                        // hace sobre el texto ORIGINAL para no alterar su capitalizacion real.
                        conGuion      = if subcapTxt = "" then "" else " - " & subcapTxt,
                        nombreRealUpper = Text.Upper(nombreReal),
                        posConGuion   = if conGuion = "" then -1 else Text.PositionOf(nombreRealUpper, Text.Upper(conGuion)),
                        posSubcap     = if subcapTxt = "" then -1 else Text.PositionOf(nombreRealUpper, Text.Upper(subcapTxt)),
                        nombreSinSub  = if subcapTxt = "" then nombreReal
                                        else if posConGuion >= 0 then Text.RemoveRange(nombreReal, posConGuion, Text.Length(conGuion))
                                        else if posSubcap >= 0 then Text.RemoveRange(nombreReal, posSubcap, Text.Length(subcapTxt))
                                        else nombreReal,
                        umTxt         = Text.Trim(Text.From(if [UM_Actividad] = null then "" else [UM_Actividad])),
                        nombreColapsado = Text.Combine(List.Select(Text.Split(nombreSinSub, " "), each _ <> ""), " "),
                        // Limpieza en 2 pasadas del texto crudo de la actividad (nada de
                        // esto tiene relacion con Subcapitulo, ya se quito arriba si aplicaba):
                        // 1) quitar la unidad si ya viene incrustada en el nombre, redundante
                        //    con la que se agrega mas abajo entre parentesis;
                        // 2) quitar guiones sueltos al inicio y/o al final que a veces
                        //    vienen asi de crudo en el reporte (ver funciones mas arriba).
                        nombreSinUnidad  = FnQuitarUnidadEmbebida(nombreColapsado, umTxt),
                        nombreLimpio  = FnQuitarGuionInicial(FnQuitarGuionColgante(nombreSinUnidad)),
                        actTxt        = if umTxt = "" then codTxt & "-" & nombreLimpio
                                        else codTxt & "-" & nombreLimpio & " (" & umTxt & ")"
                    in actTxt, type text),

                // Unifica el Subcapitulo: el explicito del SEGUIMIENTO gana; si no hay,
                // usa el derivado del sufijo del nombre (proyectos tipo TURPIAL).
                SubcapUnificado0 = Table.AddColumn(ItemsWithActividad, "SubcapituloFinal", each
                    let s = if [Subcapitulo] = null then "" else Text.Trim(Text.From([Subcapitulo]))
                    in if s <> "" then s else FnCanonSubcap([SubcapDerivado]), type text),
                SubcapUnificado = Table.RenameColumns(
                    Table.RemoveColumns(SubcapUnificado0, {"Subcapitulo", "SubcapDerivado"}),
                    {{"SubcapituloFinal", "Subcapitulo"}}),

                NumsTyped = Table.TransformColumns(SubcapUnificado, {
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
            in Final,

        // Procesa "Masivo salidas DESCRIPTIVAS" (formato plano nuevo: una fila completa
        // por insumo, sin bloques repetidos meta/subheader/total, no necesita fill-down).
        // ColumnasBase debe venir del llamador (mismo shape que usa FxProcesarSalidas).
        FxProcesarSalidasDescriptivas = (BinSalidas as binary, ColumnasBase as list) as table =>
            let
                FnText = (v as any) as text => try Text.Trim(Text.From(if v = null then "" else v)) otherwise "",
                FnCleanDisplay = (v as any) as nullable text =>
                    let t = FnText(v), clean = if t = "" then null else Text.Upper(Text.Trim(FnRemoveAccentMarks(t)))
                    in clean,
                FnBuildInsUM = (desc as any, um as any) as nullable text =>
                    let d = FnCleanDisplay(desc), u = FnCleanDisplay(um)
                    in if d = null then null else if u = null or u = "" then d else d & " (" & u & ")",
                FnCleanContratistaFromDash = (v as any) as nullable text =>
                    let t = FnText(v), afterDash = if Text.Contains(t, "-") then Text.Trim(Text.AfterDelimiter(t, "-")) else t, clean = FnCleanDisplay(afterDash)
                    in clean,
                Columnas = FnBuildColumnas(13),
                Raw = Table.Buffer(Html.Table(FnDecodeHtml(BinSalidas), Columnas, [RowSelector="tr"])),
                AddStd = Table.AddColumn(Raw, "Std", each
                    let
                        salidaNo = FnText(Record.Field(_, "Columna 2")),
                        contratista = Record.Field(_, "Columna 4"),
                        codigoIns = FnText(Record.Field(_, "Columna 7")),
                        descripcion = Record.Field(_, "Columna 8"),
                        item = Record.Field(_, "Columna 9"),
                        um = Record.Field(_, "Columna 10"),
                        cant = Record.Field(_, "Columna 11"),
                        vrTotal = Record.Field(_, "Columna 13"),
                        insFinal = FnBuildInsUM(descripcion, um),
                        codAct = FnFormatCodigoAct(item)
                    in [
                        #"Codigo ins" = codigoIns,
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
                        #"#SALIDA" = salidaNo,
                        #"Cantidad Cons Cols" = FxToNumberFlex(cant),
                        #"VT Cons Cols" = FxToNumberFlex(vrTotal)
                    ]),
                Expanded = Table.ExpandRecordColumn(AddStd, "Std", ColumnasBase, ColumnasBase),
                Filtrado = Table.SelectRows(Expanded, each try (Number.FromText([#"Codigo ins"]) <> null) otherwise false),
                Selected = Table.SelectColumns(Filtrado, ColumnasBase, MissingField.UseNull)
            in Selected,

        // Procesa "Informe entradas por insumo" (formato plano nuevo: una fila completa
        // por entrada, sin bloques repetidos meta/item que hay que reconstruir con
        // FillDown como en FxProcesarEntradas). Mismo patron que FxProcesarSalidasDescriptivas.
        // Columnas reales del reporte: 1 Sucursal, 2 Cod, 3 Descripcion, 4 Agrupacion, 5 UM,
        // 6 No. OC, 7 No. EA, 8 Fecha, 9 Cantidad, 10 Vr. Unitario, 11 IVA, 12 Valor Total,
        // 13 Proveedor, 14 Obs.
        FxProcesarEntradasPorInsumo = (BinEntradas as binary, ColumnasBase as list) as table =>
            let
                FnText = (v as any) as text => try Text.Trim(Text.From(if v = null then "" else v)) otherwise "",
                FnCleanDisplay = (v as any) as nullable text =>
                    let t = FnText(v), clean = if t = "" then null else Text.Upper(Text.Trim(FnRemoveAccentMarks(t)))
                    in clean,
                FnBuildInsUM = (desc as any, um as any) as nullable text =>
                    let d = FnCleanDisplay(desc), u = FnCleanDisplay(um)
                    in if d = null then null else if u = null or u = "" then d else d & " (" & u & ")",
                FnCleanContratistaFromDash = (v as any) as nullable text =>
                    let t = FnText(v), afterDash = if Text.Contains(t, "-") then Text.Trim(Text.AfterDelimiter(t, "-")) else t, clean = FnCleanDisplay(afterDash)
                    in clean,
                Columnas = FnBuildColumnas(14),
                Raw = Table.Buffer(Html.Table(FnDecodeHtml(BinEntradas), Columnas, [RowSelector="tr"])),
                AddStd = Table.AddColumn(Raw, "Std", each
                    let
                        codigoIns = FnText(Record.Field(_, "Columna 2")),
                        descripcion = Record.Field(_, "Columna 3"),
                        um = Record.Field(_, "Columna 5"),
                        ocNo = FnText(Record.Field(_, "Columna 6")),
                        eaNo = FnText(Record.Field(_, "Columna 7")),
                        cantidad = Record.Field(_, "Columna 9"),
                        valorTotal = Record.Field(_, "Columna 12"),
                        proveedor = Record.Field(_, "Columna 13"),
                        insFinal = FnBuildInsUM(descripcion, um)
                    in [
                        #"Codigo ins" = codigoIns,
                        Ins = insFinal,
                        Actividad = null,
                        #"Codigo act" = null,
                        InsClave = FnClaveLimpia(insFinal),
                        #"# OC / Contrato" = ocNo,
                        #"Cantidad Comprado" = null,
                        #"VT Comprado" = null,
                        VU_Crudo = null,
                        IVA_Crudo = null,
                        #"Nombre Contratista" = FnCleanContratistaFromDash(proveedor),
                        #"#ENTRADA" = eaNo,
                        #"Cantidad Cortes" = FxToNumberFlex(cantidad),
                        #"VT Cortes" = FxToNumberFlex(valorTotal),
                        #"#SALIDA" = null,
                        #"Cantidad Cons Cols" = null,
                        #"VT Cons Cols" = null
                    ]),
                Expanded = Table.ExpandRecordColumn(AddStd, "Std", ColumnasBase, ColumnasBase),
                Filtrado = Table.SelectRows(Expanded, each try (Number.FromText([#"Codigo ins"]) <> null) otherwise false),
                Selected = Table.SelectColumns(Filtrado, ColumnasBase, MissingField.UseNull)
            in Selected
    ]
in
    Funciones
