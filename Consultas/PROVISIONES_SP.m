let
    // ============================================================
    // FUNCIONES GLOBALES
    // ============================================================
    FxToNumberFlex = F_Globales[FxToNumberFlex],
    FnRemoveAccentsSymbols = F_Globales[FnRemoveAccentsSymbols],
    FnMatchFolder = F_Globales[FnMatchFolder],
    FnReadSPExcel = F_Globales[FnReadSPExcel],
    FnBuildFolderPrefixMap = F_Globales[FnBuildFolderPrefixMap],
    FnTrimText = F_Globales[FnTrimText],

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
    FilePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. Provisiones - Control costos interno/0. CONSOLIDADOR PROVISIONES SP.xlsx",

    Origen = let t = FnReadSPExcel(SiteUrl, FilePath) in if t = null then #table({"PROYECTO"}, {}) else t,

    // ============================================================
    // FILTRO Y MAPEO
    // ============================================================
    // 1. Filtrar por el proyecto actual y por Tipo = DIRECTOS
    FiltroProyecto = Table.SelectRows(Origen, each
        try [PROYECTO] <> null and (
            Text.StartsWith(Text.Upper([PROYECTO]), Text.Upper(ParamProyecto)) or
            Text.Contains(FnRemoveAccentsSymbols(Text.Upper([PROYECTO])), FnRemoveAccentsSymbols(Text.Upper(ParamProyecto)))
        ) otherwise false
    ),

    FiltroDirectos = Table.SelectRows(FiltroProyecto, each
        try [Tipo] <> null and Text.Upper(Text.Trim([Tipo])) = "DIRECTOS" otherwise false
    ),

    // 2. Renombrar columnas clave hacia BD.m
    ColumnasRenombradas = Table.RenameColumns(FiltroDirectos, {
        {"Nombre_prov", "Nombre Contratista"},
        {"No_Orden_contrato", "# OC / Contrato"}
    }, MissingField.Ignore),

    // 3. Estandarizacion de tipos de datos
    TextosLimpios = Table.TransformColumns(ColumnasRenombradas, {
        {"PROYECTO", each FnTrimText(_), type text},
        {"Nombre Contratista", each FnTrimText(_), type text},
        {"# OC / Contrato", each FnTrimText(_), type text},
        {"VR_Bruto_con_desc", each FxToNumberFlex(_), type number}
    }, null, MissingField.Ignore),

    // 4. Agregar la etiqueta Tipo y Centro de Costos
    SinTipoOriginal = Table.RemoveColumns(TextosLimpios, {"Tipo"}, MissingField.Ignore),
    AgregadoTipo = Table.AddColumn(SinTipoOriginal, "Tipo", each "PROVISIONES", type text),
    AgregadoCC = Table.AddColumn(AgregadoTipo, "Centro de Costos", each
        if Text.StartsWith(Text.Upper(ParamProyecto), "PAMPLONA 1") and [#"# OC / Contrato"] <> null then
            let
                ccPrefix = Text.Start([#"# OC / Contrato"], 4),
                fromMap = try Record.Field(PrefixMap, ccPrefix) otherwise null
            in
                if fromMap <> null then fromMap else FnMatchFolder([PROYECTO], ListaCarpetas)
        else
            FnMatchFolder([PROYECTO], ListaCarpetas)
    , type text)
in
    AgregadoCC
