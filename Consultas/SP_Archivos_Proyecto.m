let
    ParamProyecto = Text.Trim(ProyectoActual),
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    BasePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. Reportes EDT - Control costos interno/" & ParamProyecto,
    Headers = [Accept="application/json;odata=nometadata"],
    FnEncode = F_Globales[FnEncode],

    // Indice liviano: solo lista carpetas y metadatos. Los binarios se descargan
    // despues de filtrar en cada consulta consumidora.
    FolderResponse = try Json.Document(Web.Contents(SiteUrl, [
        RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(BasePath) & "')/Folders",
        Query = [#"$select" = "Name"],
        Headers = Headers,
        Timeout = #duration(0, 0, 5, 0)
    ])) otherwise null,

    CCFolders =
        if FolderResponse = null or not Record.HasFields(FolderResponse, "value")
        then #table({"Name"}, {})
        else Table.FromRecords(FolderResponse[value]),

    WithFiles = Table.AddColumn(CCFolders, "Archivos", each
        let
            ccActualPath = BasePath & "/" & [Name] & "/Actual",
            result = try Json.Document(Web.Contents(SiteUrl, [
                RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(ccActualPath) & "')/Files",
                Query = [#"$select" = "Name,ServerRelativeUrl,TimeLastModified,Length"],
                Headers = Headers,
                Timeout = #duration(0, 0, 5, 0)
            ])) otherwise null
        in
            if result <> null and Record.HasFields(result, "value") then Table.FromRecords(result[value]) else null
    ),

    ValidCCs = Table.SelectRows(WithFiles, each [Archivos] <> null),
    Expanded = Table.ExpandTableColumn(
        ValidCCs,
        "Archivos",
        {"Name", "ServerRelativeUrl", "TimeLastModified", "Length"},
        {"FileName", "ServerRelativeUrl", "TimeLastModified", "Length"}
    ),

    Relevant = Table.SelectRows(Expanded, each
        not Text.StartsWith([FileName], "~$") and (
            Text.Contains([FileName], "SEGUIMIENTO POR ITEMS",         Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "ANALISIS DE PRECIOS UNITARIOS", Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "INFORMEORDEN",                  Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "ESTADO DE ORDENES",             Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "INFORME ENTRADAS DE ALMACEN",   Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "INFORME ENTRADAS DE ALMACÉN",   Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "MASIVO SALIDAS DETALLADO",      Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "ESTADO DE CONTRATOS",           Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "DESCUENTOS",                    Comparer.OrdinalIgnoreCase)
        )
    ),

    Typed = Table.TransformColumnTypes(Relevant, {{"TimeLastModified", type datetimezone}, {"Length", Int64.Type}}, "en-US"),
    Sorted = Table.Sort(Typed, {{"Name", Order.Ascending}, {"FileName", Order.Ascending}, {"TimeLastModified", Order.Descending}}),
    Final = Table.Buffer(Table.RenameColumns(
        Table.SelectColumns(Sorted, {"Name", "FileName", "ServerRelativeUrl", "TimeLastModified", "Length"}),
        {{"Name", "Centro de Costos"}, {"FileName", "Name"}}
    ))
in
    Final