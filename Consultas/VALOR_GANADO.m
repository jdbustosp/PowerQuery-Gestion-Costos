let
    // ============================================================
    // VALOR_GANADO: indicadores de valor ganado del proyecto actual,
    // leidos del archivo de query de valor ganado (tablas SPI y CPI).
    // Toma el archivo MAS RECIENTE de la carpeta cuyo nombre contenga
    // "valor ganado" (el nombre cambia con la fecha: 260831_query valor ganado.xlsx).
    // Montos en millones, tal como vienen en el archivo.
    // ============================================================
    ParamProyecto = Text.Upper(Text.Trim(ProyectoActual)),
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    Carpeta = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/Valor ganado/0. Query",
    FnEncode = F_Globales[FnEncode],
    Headers = [Accept = "application/json;odata=nometadata"],

    ArchivosJson = Json.Document(Web.Contents(SiteUrl, [
        RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(Carpeta) & "')/Files",
        Query = [#"$select" = "Name,ServerRelativeUrl,TimeLastModified"],
        Headers = Headers
    ])),
    Archivos = Table.FromRecords(ArchivosJson[value], {"Name", "ServerRelativeUrl", "TimeLastModified"}, MissingField.UseNull),
    Candidatos = Table.Sort(
        Table.SelectRows(Archivos, each
            Text.Contains([Name], "valor ganado", Comparer.OrdinalIgnoreCase)
            and not Text.StartsWith([Name], "~$")
            and Text.EndsWith(Text.Lower([Name]), ".xlsx")),
        {{"TimeLastModified", Order.Descending}}),
    Archivo =
        if Table.IsEmpty(Candidatos)
        then error Error.Record("Sin archivo de valor ganado", "No hay un .xlsx con 'valor ganado' en el nombre en " & Carpeta)
        else Candidatos{0},

    Binario = Binary.Buffer(Web.Contents(SiteUrl, [
        RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(Archivo[ServerRelativeUrl]) & "')/$value",
        Headers = [Accept = "*/*"]
    ])),
    Libro = Excel.Workbook(Binario, true),
    TablaSPI = Libro{[Item = "SPI", Kind = "Table"]}[Data],
    TablaCPI = Libro{[Item = "CPI", Kind = "Table"]}[Data],

    FnCol = (t as table, nombre as text) as text =>
        // Los encabezados traen saltos de linea ("% #(lf)ejecutado"): se busca sin ellos.
        let
            limpio = (s as text) => Text.Upper(Text.Combine(List.Select(Text.SplitAny(s, " #(lf)#(cr)"), each _ <> ""), " ")),
            hit = List.First(List.Select(Table.ColumnNames(t), each limpio(_) = limpio(nombre)), null)
        in if hit = null then error "Columna '" & nombre & "' no encontrada en la tabla de valor ganado" else hit,
    FnValor = (t as table, nombre as text) => let c = FnCol(t, nombre) in if Table.IsEmpty(t) then null else Table.Column(t, c){0},

    FilaSPI = Table.SelectRows(TablaSPI, each Text.Upper(Text.Trim(Text.From([PROYECTO]))) = ParamProyecto),
    FilaCPI = Table.SelectRows(TablaCPI, each Text.Upper(Text.Trim(Text.From([PROYECTO]))) = ParamProyecto),
    FnNum = (v as any) as nullable number => try Number.From(v) otherwise null,
    FnFecha = (v as any) as nullable date =>
        if v = null then null
        else try Date.From(v) otherwise (try Date.From(DateTime.From(Number.From(v))) otherwise null),

    Resultado = #table(
        type table [
            Proyecto = text, SPI = number, #"Pct programado" = number, #"Pct ejecutado" = number,
            #"Fecha corte" = date, #"Fecha inicio obra" = date, #"Fecha fin obra" = date, #"Duracion dias" = number,
            #"CPI archivo" = number, #"Ppto inicial" = number, Vivienda = number, #"Urb interior" = number,
            Reembolsables = number, Inflacion = number, Imprevistos = number, #"Urb exterior" = number,
            #"Fecha ppto inicial" = date, #"Ppto proyectado archivo" = number, #"Fecha proyectado" = date,
            #"Archivo origen" = text
        ],
        {{
            ParamProyecto,
            FnNum(FnValor(FilaSPI, "SPI")), FnNum(FnValor(FilaSPI, "% programado")), FnNum(FnValor(FilaSPI, "% ejecutado")),
            FnFecha(FnValor(FilaSPI, "Fecha de corte")), FnFecha(FnValor(FilaSPI, "Fecha Inicio obra")), FnFecha(FnValor(FilaSPI, "Fecha Fin obra")),
            FnNum(FnValor(FilaSPI, "Duración (dias)")),
            FnNum(FnValor(FilaCPI, "CPI")), FnNum(FnValor(FilaCPI, "TOTAL PPTO IC")), FnNum(FnValor(FilaCPI, "Vivienda:")),
            FnNum(FnValor(FilaCPI, "Urb. Interior:")), FnNum(FnValor(FilaCPI, "Reembolsables:")),
            FnNum(FnValor(FilaCPI, "Incrementos por inflación:")), FnNum(FnValor(FilaCPI, "Imprevistos:")),
            FnNum(FnValor(FilaCPI, "Urb. exterior:")), FnFecha(FnValor(FilaCPI, "Fecha Fact IC")),
            FnNum(FnValor(FilaCPI, "TOTAL PROYECTADO")), FnFecha(FnValor(FilaCPI, "Fecha proy")),
            Archivo[Name]
        }})
in
    Resultado
