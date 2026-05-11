let
    ParamProyecto = Text.Trim(ProyectoActual),
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    BasePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. Reportes EDT - Control costos interno/" & ParamProyecto,
    Headers = [Accept="application/json;odata=nometadata"],
    FnEncode = F_Globales[FnEncode],

    // PASO 1: Listar carpetas del proyecto (Centro de Costos) — 1 llamada HTTP
    FoldersUrl = SiteUrl & "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(BasePath) & "')/Folders?$select=Name",
    CCFolders = let r = Json.Document(Web.Contents(FoldersUrl, [Headers=Headers]))
                in Table.FromRecords(r[value]),

    // PASO 2: Para cada CC, listar archivos en /Actual/ — 1 llamada por CC
    // Nota: SharePoint REST (OData v3) no soporta $expand anidado, no hay forma de reducir estas llamadas
    WithFiles = Table.AddColumn(CCFolders, "Archivos", each
        let
            ccActualPath = BasePath & "/" & [Name] & "/Actual",
            filesUrl = SiteUrl & "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(ccActualPath) & "')/Files?$select=Name,ServerRelativeUrl",
            result = try Json.Document(Web.Contents(filesUrl, [Headers=Headers])) otherwise null
        in
            if result <> null then Table.FromRecords(result[value]) else null
    ),
    ValidCCs = Table.SelectRows(WithFiles, each [Archivos] <> null),

    // PASO 3: Expandir archivos
    Expanded = Table.ExpandTableColumn(ValidCCs, "Archivos", {"Name", "ServerRelativeUrl"}, {"FileName", "ServerRelativeUrl"}),

    // PASO 4: Solo archivos relevantes (excluye temporales ~$)
    Relevant = Table.SelectRows(Expanded, each
        not Text.StartsWith([FileName], "~$") and (
        Text.Contains([FileName], "SEGUIMIENTO POR ITEMS",         Comparer.OrdinalIgnoreCase) or
        Text.Contains([FileName], "ANALISIS DE PRECIOS UNITARIOS", Comparer.OrdinalIgnoreCase) or
        Text.Contains([FileName], "INFORMEORDEN",                  Comparer.OrdinalIgnoreCase) or
        Text.Contains([FileName], "ESTADO DE ORDENES",             Comparer.OrdinalIgnoreCase) or
        Text.Contains([FileName], "ESTADO DE CONTRATOS",           Comparer.OrdinalIgnoreCase) or
        Text.Contains([FileName], "DESCUENTOS",                    Comparer.OrdinalIgnoreCase))
    ),

    // PASO 5: Descargar binarios — Binary.Buffer evita re-descargas cuando multiples queries lo usan
    WithContent = Table.AddColumn(Relevant, "Content", each
        Binary.Buffer(Web.Contents(SiteUrl & "/_api/web/GetFileByServerRelativeUrl('" & FnEncode([ServerRelativeUrl]) & "')/$value"))
    ),

    // PASO 6: Table.Buffer materializa TODO en memoria para que CONTRATOS, COMPRAS, DESCUENTOS
    // y SP_Seguimiento_Parsed no re-disparen ninguna llamada HTTP
    Final = Table.Buffer(Table.RenameColumns(
        Table.SelectColumns(WithContent, {"Name", "FileName", "Content"}),
        {{"Name", "Centro de Costos"}, {"FileName", "Name"}}
    ))
in
    Final
