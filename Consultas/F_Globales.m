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

        FnRemoveAccentsSymbols = (t as nullable text) as nullable text => 
            if t = null then null 
            else let 
                initial = Text.From(t), 
                replacements = {
                    {"á","a"},{"Á","A"},{"é","e"},{"É","E"},{"í","i"},{"Í","I"},{"ó","o"},{"Ó","O"},{"ú","u"},{"Ú","U"},
                    {"º",""},{"°",""},{"¨",""},
                    {"Ã“","O"}, {"Ã‘","N"}, {"Ã¡","A"}, {"Ã©","E"}, {"Ã³","O"}, {"Ãº","U"}, {"Ã","A"},
                    {"#(lf)", " "}, {"#(cr)", " "}
                }, 
                result = List.Accumulate(replacements, initial, (state, current) => Text.Replace(state, current{0}, current{1})) 
            in result,

        FnClaveLimpia = (t as nullable text) as nullable text => 
            if t = null then null 
            else let 
                sinUnidad = if Text.Contains(t, "(") then Text.BeforeDelimiter(t, "(") else t,
                t1 = Text.Upper(Text.Trim(sinUnidad)),
                t2 = List.Accumulate({{"Á","A"},{"É","E"},{"Í","I"},{"Ó","O"},{"Ú","U"},{"Ñ","N"},{"Ü","U"}}, t1, (state, current) => Text.Replace(state, current{0}, current{1})),
                t3 = Text.Select(t2, {"A".."Z", "0".."9"})
            in if t3 = "" then null else t3,

        FnPrepareTableWithHeader = (tbl as table) as table => 
            let 
                firstColName = Table.ColumnNames(tbl){0}, 
                firstColValues = Table.Column(tbl, firstColName), 
                headerFlags = List.Transform(firstColValues, (x) => let txt = Text.Upper(if x = null then "" else Text.From(x)), txtNorm = Text.Replace(txt, "Ó", "O") in Text.Contains(txtNorm, "COD")), 
                hasHeader = List.Contains(headerFlags, true), 
                promoted = if hasHeader then let headerIndex = List.PositionOf(headerFlags, true), skipped  = Table.Skip(tbl, headerIndex) in Table.PromoteHeaders(skipped, [PromoteAllScalars = true]) else tbl 
            in promoted
    ]
in
    Funciones
