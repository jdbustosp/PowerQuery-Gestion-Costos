let
    ParamProyecto = Text.Trim(ProyectoActual),
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    BasePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. Reportes EDT - Control costos interno/" & ParamProyecto,
    Headers = [Accept="application/json;odata=nometadata"],
    FnEncode = F_Globales[FnEncode],

    // PASO 1: Listar CCs del proyecto (1 llamada HTTP)
    // Web.Contents con RelativePath/Query: el DataSourcePath queda anclado en SiteUrl,
    // asi Power Query NO cachea URLs especificas que rompen al cambiar el codigo.
    CCFolders = let
        r = Json.Document(Web.Contents(SiteUrl, [
            RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(BasePath) & "')/Folders",
            Query = [#"$select" = "Name"],
            Headers = Headers
        ]))
        in Table.FromRecords(r[value]),

    // PASO 2: Listar archivos en /Actual/ de cada CC (1 llamada por CC)
    // SharePoint REST (OData v3) no soporta $expand anidado, no se puede reducir a una sola llamada.
    WithFiles = Table.AddColumn(CCFolders, "Archivos", each
        let
            ccActualPath = BasePath & "/" & [Name] & "/Actual",
            result = try Json.Document(Web.Contents(SiteUrl, [
                RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(ccActualPath) & "')/Files",
                Query = [#"$select" = "Name,ServerRelativeUrl"],
                Headers = Headers
            ])) otherwise null
        in
            if result <> null then Table.FromRecords(result[value]) else null
    ),
    ValidCCs = Table.SelectRows(WithFiles, each [Archivos] <> null),

    Expanded = Table.ExpandTableColumn(ValidCCs, "Archivos", {"Name", "ServerRelativeUrl"}, {"FileName", "ServerRelativeUrl"}),

    // Filtra archivos relevantes (excluye temporales ~$)
    Relevant = Table.SelectRows(Expanded, each
        not Text.StartsWith([FileName], "~$") and (
        Text.Contains([FileName], "SEGUIMIENTO POR ITEMS",         Comparer.OrdinalIgnoreCase) or
        Text.Contains([FileName], "ANALISIS DE PRECIOS UNITARIOS", Comparer.OrdinalIgnoreCase) or
        Text.Contains([FileName], "INFORMEORDEN",                  Comparer.OrdinalIgnoreCase) or
        Text.Contains([FileName], "ESTADO DE ORDENES",             Comparer.OrdinalIgnoreCase) or
        Text.Contains([FileName], "ESTADO DE CONTRATOS",           Comparer.OrdinalIgnoreCase) or
        Text.Contains([FileName], "DESCUENTOS",                    Comparer.OrdinalIgnoreCase))
    ),

    // Descargar binarios — Binary.Buffer evita re-descargas cuando multiples queries lo consumen
    WithContent = Table.AddColumn(Relevant, "Content", each
        Binary.Buffer(Web.Contents(SiteUrl, [
            RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode([ServerRelativeUrl]) & "')/$value",
            Headers = Headers
        ]))
    ),

    // Table.Buffer materializa todo en memoria para que CONTRATOS, COMPRAS, DESCUENTOS
    // y SP_Seguimiento_Parsed no re-disparen ninguna llamada HTTP
    Final = Table.Buffer(Table.RenameColumns(
        Table.SelectColumns(WithContent, {"Name", "FileName", "Content"}),
        {{"Name", "Centro de Costos"}, {"FileName", "Name"}}
    ))
in
    Final
