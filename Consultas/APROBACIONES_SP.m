let
    // ============================================================
    // FUNCIONES GLOBALES
    // ============================================================
    FxToNumberFlex = F_Globales[FxToNumberFlex],
    FnMatchFolder = F_Globales[FnMatchFolder],
    FnReadSPExcel = F_Globales[FnReadSPExcel],
    FnBuildFolderPrefixMap = F_Globales[FnBuildFolderPrefixMap],
    FnTrimText = F_Globales[FnTrimText],
    FnRemoveAccentsSymbols = F_Globales[FnRemoveAccentsSymbols],

    ParamProyecto = Text.Trim(ProyectoActual),

    // ============================================================
    // CARPETAS DE PROYECTO ACTUAL (sin filtro de archivos)
    // ============================================================
    ListaCarpetas = try List.Distinct(SP_CarpetasCC[Centro de Costos]) otherwise {},
    PrefixMap = FnBuildFolderPrefixMap(ListaCarpetas),

    // ============================================================
    // CONEXION AL ARCHIVO EN SHAREPOINT
    // ============================================================
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    FilePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. CC approvals (solo aprobado) - Control costos interno/0. CONSOLIDADOR APROBACIONES CC SP.xlsx",

    Origen = let t = FnReadSPExcel(SiteUrl, FilePath) in if t = null then #table({"Proyecto:"}, {}) else t,

    // ============================================================
    // FILTRO Y MAPEO
    // ============================================================
    ParamProyectoClean = FnRemoveAccentsSymbols(Text.Upper(ParamProyecto)),
    FiltroProyecto = Table.SelectRows(Origen, each
        try
            let
                proyectoRaw = [#"Proyecto:"],
                proyectoClean = FnRemoveAccentsSymbols(Text.Upper(Text.Trim(Text.From(proyectoRaw))))
            in
                if ParamProyectoClean = "PAYANDE" then
                    Text.StartsWith(proyectoClean, "PAYANDE") and
                    (Text.Contains(proyectoClean, "URB INTERNO") or Text.Contains(proyectoClean, "TORRES"))
                else
                    Text.StartsWith(proyectoClean, ParamProyectoClean) or Text.Contains(proyectoClean, ParamProyectoClean)
        otherwise false
    ),

    ColumnasRenombradas = Table.RenameColumns(FiltroProyecto, {
        {"Desc. - UM", "Ins"},
        {"Nombre del proveedor: ", "Nombre Contratista"},
        {"# CC", "# CC - Comparativo"},
        {"Cant. Total", "Cantidad CC Cons"},
        {"V/U TOTAL", "V/U CC cons"},
        {"VR TOTAL", "VT CC cons"}
    }, MissingField.Ignore),

    TextosLimpios = Table.TransformColumns(ColumnasRenombradas, {
        {"Ins", each FnTrimText(_), type text},
        {"Nombre Contratista", each FnTrimText(_), type text},
        {"# CC - Comparativo", each FnTrimText(_), type text},
        {"Cantidad CC Cons", each FxToNumberFlex(_), type number},
        {"V/U CC cons", each FxToNumberFlex(_), type number},
        {"VT CC cons", each FxToNumberFlex(_), type number}
    }, null, MissingField.Ignore),

    AgregadoTipo = Table.AddColumn(TextosLimpios, "Tipo", each "CC Consolidado", type text),
    AgregadoCC = Table.AddColumn(AgregadoTipo, "Centro de Costos", each
        if Text.StartsWith(Text.Upper(ParamProyecto), "PAMPLONA 1") and [#"# CC - Comparativo"] <> null then
            let
                ccPrefix = Text.Trim(Text.BeforeDelimiter([#"# CC - Comparativo"], "-")),
                fromMap = try Record.Field(PrefixMap, ccPrefix) otherwise null
            in
                if fromMap <> null then fromMap else FnMatchFolder([#"Proyecto:"], ListaCarpetas)
        else
            FnMatchFolder([#"Proyecto:"], ListaCarpetas)
    , type text),

    // ============================================================
    // EXTRACCION DE COLUMNAS PARA BD
    // ============================================================
    TablaFinal = Table.SelectColumns(AgregadoCC,
        {
            "Centro de Costos",
            "Tipo",
            "Ins",
            "Nombre Contratista",
            "# CC - Comparativo",
            "Cantidad CC Cons",
            "V/U CC cons",
            "VT CC cons"
        },
        MissingField.Ignore
    )
in
    TablaFinal
