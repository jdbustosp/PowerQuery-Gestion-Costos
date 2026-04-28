let
    // =========================================================
    // 🚀 API REST DE SHAREPOINT (BYPASS AL BLOQUEO)
    // =========================================================
    ParamProyecto = Text.Trim(ProyectoActual),
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    BasePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/Reportes EDT/" & ParamProyecto,
    Headers = [Accept="application/json;odata=nometadata"],

    // Codificar ruta para URL (solo espacios, que es lo más común)
    FnEncode = (path as text) as text => 
        Text.Combine(List.Transform(Text.Split(path, "/"), each Uri.EscapeDataString(_)), "/"),

    // PASO 1: Listar carpetas del proyecto (Centro de Costos)
    FoldersUrl = SiteUrl & "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(BasePath) & "')/Folders?$select=Name",
    CCFolders = Table.FromRecords(Json.Document(Web.Contents(FoldersUrl, [Headers=Headers]))[value]),

    // PASO 2: Para cada CC, listar archivos en /Actual/
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

    // PASO 4: Solo archivos relevantes
    Relevant = Table.SelectRows(Expanded, each
        not Text.StartsWith([FileName], "~$") and (
        Text.Contains([FileName], "SEGUIMIENTO POR ITEMS", Comparer.OrdinalIgnoreCase) or
        Text.Contains([FileName], "ANALISIS DE PRECIOS UNITARIOS", Comparer.OrdinalIgnoreCase) or
        Text.Contains([FileName], "INFORMEORDEN", Comparer.OrdinalIgnoreCase) or
        Text.Contains([FileName], "ESTADO DE ORDENES", Comparer.OrdinalIgnoreCase) or
        Text.Contains([FileName], "ESTADO DE CONTRATOS", Comparer.OrdinalIgnoreCase) or
        Text.Contains([FileName], "DESCUENTOS", Comparer.OrdinalIgnoreCase))
    ),

    // PASO 5: Descargar contenido y almacenar en memoria (Binary.Buffer evita re-descargas)
    WithContent = Table.AddColumn(Relevant, "Content", each
        Binary.Buffer(Web.Contents(SiteUrl & "/_api/web/GetFileByServerRelativeUrl('" & FnEncode([ServerRelativeUrl]) & "')/$value"))
    ),

    // PASO 6: Table.Buffer materializa TODO el resultado para que CONTRATOS, COMPRAS,
    // DESCUENTOS y SP_Seguimiento_Parsed NO re-disparen las llamadas HTTP
    Final = Table.Buffer(Table.RenameColumns(
        Table.SelectColumns(WithContent, {"Name", "FileName", "Content"}),
        {{"Name", "Centro de Costos"}, {"FileName", "Name"}}
    ))
in
    Final
