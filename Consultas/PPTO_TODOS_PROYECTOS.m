let
    // =========================================================
    // FUNCIONES AUXILIARES GLOBALES
    // =========================================================
    FnEncode = F_Globales[FnEncode],
    FxProcesarCentroCosto = F_Globales[FxProcesarCentroCosto],

    // =========================================================
    // API REST DE SHAREPOINT (TODOS LOS PROYECTOS)
    // =========================================================
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    BasePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. Reportes EDT - Control costos interno",
    Headers = [Accept="application/json;odata=nometadata"],

    // =========================================================
    // NIVEL 1: LISTAR TODOS LOS PROYECTOS
    // =========================================================
    ProyectosUrl = SiteUrl & "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(BasePath) & "')/Folders?$select=Name",
    ProyectosFolders = Table.FromRecords(Json.Document(Web.Contents(ProyectosUrl, [Headers=Headers]))[value]),

    // =========================================================
    // NIVEL 2: POR CADA PROYECTO, LISTAR CENTROS DE COSTOS
    // =========================================================
    WithCCs = Table.AddColumn(ProyectosFolders, "CCs", each
        let
            projPath = BasePath & "/" & [Name],
            ccsUrl = SiteUrl & "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(projPath) & "')/Folders?$select=Name",
            result = try Json.Document(Web.Contents(ccsUrl, [Headers=Headers])) otherwise null
        in
            if result <> null and (try List.Count(result[value]) otherwise 0) > 0
            then Table.FromRecords(result[value]) else null
    ),
    ValidProjects = Table.SelectRows(WithCCs, each [CCs] <> null),
    ExpandedCCs = Table.ExpandTableColumn(ValidProjects, "CCs", {"Name"}, {"Centro de Costos"}),
    WithProyecto = Table.RenameColumns(ExpandedCCs, {{"Name", "Proyecto"}}),

    // =========================================================
    // NIVEL 3: POR CADA CC, LISTAR ARCHIVOS EN /Actual/
    // =========================================================
    WithFiles = Table.AddColumn(WithProyecto, "Archivos", each
        let
            ccActualPath = BasePath & "/" & [Proyecto] & "/" & [Centro de Costos] & "/Actual",
            filesUrl = SiteUrl & "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(ccActualPath) & "')/Files?$select=Name,ServerRelativeUrl",
            result = try Json.Document(Web.Contents(filesUrl, [Headers=Headers])) otherwise null
        in
            if result <> null and (try List.Count(result[value]) otherwise 0) > 0
            then Table.FromRecords(result[value]) else null
    ),
    ValidCCs = Table.SelectRows(WithFiles, each [Archivos] <> null),
    ExpandedFiles = Table.ExpandTableColumn(ValidCCs, "Archivos", {"Name", "ServerRelativeUrl"}, {"FileName", "ServerRelativeUrl"}),

    // Solo SEGUIMIENTO POR ITEMS y ANALISIS DE PRECIOS UNITARIOS
    Relevant = Table.SelectRows(ExpandedFiles, each
        not Text.StartsWith([FileName], "~$") and (
        Text.Contains([FileName], "SEGUIMIENTO POR ITEMS", Comparer.OrdinalIgnoreCase) or
        Text.Contains([FileName], "ANALISIS DE PRECIOS UNITARIOS", Comparer.OrdinalIgnoreCase))
    ),

    // Descargar contenido con Binary.Buffer (evita re-descargas)
    WithContent = Table.AddColumn(Relevant, "Content", each
        Binary.Buffer(Web.Contents(SiteUrl & "/_api/web/GetFileByServerRelativeUrl('" & FnEncode([ServerRelativeUrl]) & "')/$value"))
    ),

    // =========================================================
    // AGRUPAR POR PROYECTO + CC Y PROCESAR
    // =========================================================
    Agrupado = Table.Group(WithContent, {"Proyecto", "Centro de Costos"}, {{"Binarios", each
        let
            FilaSeg = Table.SelectRows(_, each Text.Contains([FileName], "SEGUIMIENTO POR ITEMS", Comparer.OrdinalIgnoreCase)),
            FilaPres = Table.SelectRows(_, each Text.Contains([FileName], "ANALISIS DE PRECIOS UNITARIOS", Comparer.OrdinalIgnoreCase))
        in if Table.RowCount(FilaSeg) > 0 and Table.RowCount(FilaPres) > 0
           then [Bin_S = FilaSeg{0}[Content], Bin_P = FilaPres{0}[Content]]
           else null
    }}),
    CentrosCompletos = Table.SelectRows(Agrupado, each [Binarios] <> null),

    // Aplicar funcion compartida desde F_Globales
    TablaConDatos = Table.AddColumn(CentrosCompletos, "Datos", each
        FxProcesarCentroCosto([Binarios][Bin_S], [Binarios][Bin_P])),

    // =========================================================
    // EXPANDIR Y APLICAR LOGICA PPTO_BD
    // =========================================================
    Expandido = Table.ExpandTableColumn(TablaConDatos, "Datos",
        {"Codigo ins", "Ins", "Actividad", "Capitulo", "Subcapitulo", "Cantidad Presupuesto", "VT Presupuesto"}),

    // V/U Presupuesto (VT / Cantidad)
    AddVU = Table.AddColumn(Expandido, "V/U Presupuesto", each
        if [Cantidad Presupuesto] = null or [Cantidad Presupuesto] = 0 or [VT Presupuesto] = null then null
        else [VT Presupuesto] / [Cantidad Presupuesto], Currency.Type),

    Filtered = Table.SelectRows(AddVU, each [VT Presupuesto] <> null and [VT Presupuesto] <> 0),

    Selected = Table.SelectColumns(Filtered,
        {"Proyecto", "Centro de Costos", "Subcapitulo", "Capitulo", "Actividad", "Codigo ins", "Ins",
         "Cantidad Presupuesto", "V/U Presupuesto", "VT Presupuesto"}),

    Typed = Table.TransformColumnTypes(Selected, {
        {"Proyecto", type text}, {"Centro de Costos", type text},
        {"Subcapitulo", type text}, {"Capitulo", type text}, {"Actividad", type text},
        {"Codigo ins", Int64.Type}, {"Ins", type text},
        {"Cantidad Presupuesto", type number}, {"V/U Presupuesto", Currency.Type},
        {"VT Presupuesto", Currency.Type}
    }),

    TablaFinal = Typed
in
    TablaFinal
