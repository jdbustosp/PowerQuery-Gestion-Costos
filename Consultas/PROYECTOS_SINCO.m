let
    // Lista las carpetas de proyectos disponibles en SharePoint (una por proyecto).
    // Se carga en la hoja CONFIG para alimentar el desplegable de ProyectoActual.
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    BasePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. Reportes EDT - Control costos interno",
    FnEncode = F_Globales[FnEncode],

    Resp = try Json.Document(Web.Contents(SiteUrl, [
        RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(BasePath) & "')/Folders",
        Query = [#"$select" = "Name"],
        Headers = [Accept = "application/json;odata=nometadata"],
        Timeout = #duration(0, 0, 2, 0)
    ])) otherwise null,

    Carpetas =
        if Resp = null or not Record.HasFields(Resp, "value")
        then #table({"Name"}, {})
        else Table.FromRecords(Resp[value]),

    SinSistema = Table.SelectRows(Carpetas, each
        [Name] <> "Forms" and not Text.StartsWith([Name], "_") and not Text.StartsWith([Name], ".")),

    Renombrada = Table.RenameColumns(SinSistema, {{"Name", "Proyecto"}}),
    Tipada = Table.TransformColumnTypes(Renombrada, {{"Proyecto", type text}}),
    Ordenada = Table.Sort(Tipada, {{"Proyecto", Order.Ascending}})
in
    Ordenada
