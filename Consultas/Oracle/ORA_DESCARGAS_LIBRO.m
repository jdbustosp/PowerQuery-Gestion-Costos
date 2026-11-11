// ORA_DESCARGAS_LIBRO: archivo de descargas del proyecto LEIDO DE SHAREPOINT (no de una ruta local):
// "ORA_CONFIG[CarpetaDescargas]/<ProyectoActual>.xlsx" (nombre exacto primero, luego
// "contiene"; misma busqueda que DESCARGAS de SINCO). Devuelve la tabla de navegacion de Excel.Workbook
// (tablas y hojas: DESCARGA, Actualizacion_ppto...). Sin archivo o ilegible = error explicito.
let
    SiteUrl = ORA_CONFIG[Sitio],
    Carpeta = ORA_CONFIG[CarpetaDescargas],
    Proyecto = Text.Upper(Text.Trim(Text.From(ProyectoActual))),
    FnEncode = F_Globales[FnEncode],
    Listado = Json.Document(Web.Contents(SiteUrl, [
        RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(Carpeta) & "')/Files",
        Query = [#"$select" = "Name,ServerRelativeUrl,TimeLastModified"],
        Headers = [Accept = "application/json;odata=nometadata"],
        Timeout = #duration(0, 0, 2, 0)])),
    Archivos = Table.SelectRows(Table.FromRecords(Listado[value]), each
        not Text.StartsWith([Name], "~$") and Text.EndsWith(Text.Upper([Name]), ".XLSX")),
    Base = (n as text) as text => Text.Upper(Text.Trim(Text.Start(n, Text.Length(n) - 5))),
    Exactos = Table.SelectRows(Archivos, each Base([Name]) = Proyecto),
    Contienen = Table.SelectRows(Archivos, each Text.Contains(Text.Upper([Name]), Proyecto)),
    Candidatos = if Table.RowCount(Exactos) > 0 then Exactos else Contienen,
    Ruta = if Table.RowCount(Candidatos) = 0
        then error Error.Record("Falta archivo de descargas", "No hay " & Proyecto & ".xlsx en " & Carpeta)
        else Candidatos{0}[ServerRelativeUrl],
    Binario = F_Globales[FnReadSPBinary](SiteUrl, Ruta),
    Libro = if Binario = null
        then error Error.Record("Archivo de descargas ilegible", "No se pudo descargar " & Ruta & " de SharePoint")
        else Excel.Workbook(Binary.Buffer(Binario), null, true)
in
    Libro
