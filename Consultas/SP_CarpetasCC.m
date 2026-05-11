let
    // Lista TODAS las carpetas (Centro de Costos) del proyecto actual,
    // sin filtrar por presencia de archivos. Usado por APROBACIONES_SP
    // y PROVISIONES_SP para mapear nombres de proyecto -> CC sin perder
    // CCs que no tengan archivos SEGUIMIENTO/INFORMEORDEN/etc.
    ParamProyecto = Text.Trim(ProyectoActual),
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    BasePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. Reportes EDT - Control costos interno/" & ParamProyecto,
    Headers = [Accept="application/json;odata=nometadata"],
    FnEncode = F_Globales[FnEncode],

    // Web.Contents con RelativePath/Query: DataSourcePath estable, sin problemas de cache
    Respuesta = Json.Document(Web.Contents(SiteUrl, [
        RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(BasePath) & "')/Folders",
        Query = [#"$select" = "Name"],
        Headers = Headers
    ])),
    AsTable = Table.FromRecords(Respuesta[value]),
    Renamed = Table.RenameColumns(AsTable, {{"Name", "Centro de Costos"}}),
    Final = Table.Buffer(Renamed)
in
    Final
