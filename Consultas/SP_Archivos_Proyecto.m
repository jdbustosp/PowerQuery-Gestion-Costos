let
    ParamProyecto = Text.Trim(ProyectoActual),
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    BasePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. Reportes EDT - Control costos interno/" & ParamProyecto,
    Headers = [Accept="application/json;odata=nometadata"],
    FnEncode = F_Globales[FnEncode],

    // 1 SOLA LLAMADA HTTP para listar CCs + subcarpeta /Actual/ + archivos.
    // Antes: 1 llamada para CCs + 1 llamada por cada CC = N+1 llamadas.
    // Ahora: 1 llamada total. Para 15 CCs ahorra ~45 segundos solo en latencia de red.
    AllUrl = SiteUrl & "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(BasePath)
        & "')/Folders?$top=500&$select=Name&$expand=Folders($top=20;$select=Name;$expand=Files($top=100;$select=Name,ServerRelativeUrl))",
    CCList = Json.Document(Web.Contents(AllUrl, [Headers=Headers]))[value],

    // Aplanar estructura jerarquica: CC → /Actual/ → archivos
    AllFileRows = List.Combine(List.Transform(CCList, (cc) =>
        let
            ccName = cc[Name],
            subFolders = try cc[Folders][value] otherwise {},
            actualFolder = List.First(List.Select(subFolders, each _[Name] = "Actual"), null),
            files = if actualFolder = null then {} else try actualFolder[Files][value] otherwise {}
        in List.Transform(files, (f) => [#"Centro de Costos" = ccName, #"Name" = f[Name], ServerRelativeUrl = f[ServerRelativeUrl]])
    )),

    FlatTable = if List.Count(AllFileRows) = 0
        then #table({"Centro de Costos", "Name", "ServerRelativeUrl"}, {})
        else Table.FromRecords(AllFileRows),

    // Solo archivos relevantes (excluye temporales ~$)
    Relevant = Table.SelectRows(FlatTable, each
        not Text.StartsWith([Name], "~$") and (
        Text.Contains([Name], "SEGUIMIENTO POR ITEMS",        Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "ANALISIS DE PRECIOS UNITARIOS", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "INFORMEORDEN",                  Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "ESTADO DE ORDENES",             Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "ESTADO DE CONTRATOS",           Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "DESCUENTOS",                    Comparer.OrdinalIgnoreCase))
    ),

    // Descargar binarios — Binary.Buffer evita re-descargas cuando multiples queries lo usan
    WithContent = Table.AddColumn(Relevant, "Content", each
        Binary.Buffer(Web.Contents(SiteUrl & "/_api/web/GetFileByServerRelativeUrl('" & FnEncode([ServerRelativeUrl]) & "')/$value"))
    ),

    // Table.Buffer materializa TODO en memoria para que CONTRATOS, COMPRAS, DESCUENTOS
    // y SP_Seguimiento_Parsed no re-disparen ninguna llamada HTTP
    Final = Table.Buffer(Table.SelectColumns(WithContent, {"Centro de Costos", "Name", "Content"}))
in
    Final
