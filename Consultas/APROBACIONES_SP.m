let
    // ============================================================
    // FUNCIONES GLOBALES
    // ============================================================
    FnEncode = F_Globales[FnEncode],
    FxToNumberFlex = F_Globales[FxToNumberFlex],
    FnFormatCodigoAct = F_Globales[FnFormatCodigoAct],
    
    ParamProyecto = Text.Trim(ProyectoActual),

    // ============================================================
    // CARPETAS DE PROYECTO ACTUAL (Para Mapeo Exacto de CC)
    // ============================================================
    ListaCarpetas = try List.Distinct(SP_Archivos_Proyecto[Centro de Costos]) otherwise {},
    FolderCount = List.Count(ListaCarpetas),
    
    // Función heurística para encontrar el nombre exacto de la carpeta
    FnMatchFolder = (proyectoExcel as text) as text =>
        if FolderCount = 0 then proyectoExcel
        else if FolderCount = 1 then ListaCarpetas{0}
        else 
            let
                proyUpper = Text.Upper(proyectoExcel),
                // Buscar carpetas cuya última palabra (ej. "ET1", "INTERNO") coincida con el texto
                Match = List.Select(ListaCarpetas, each 
                    let 
                        baseName = if Text.Contains(_, "-") then Text.Trim(Text.AfterDelimiter(_, "-")) else Text.Trim(_),
                        lastWord = List.Last(Text.Split(baseName, " "))
                    in 
                        Text.Contains(proyUpper, lastWord)
                ),
                Result = if List.Count(Match) = 1 then Match{0} else proyectoExcel
            in
                Result,

    // ============================================================
    // CONEXIÓN AL ARCHIVO EN SHAREPOINT
    // ============================================================
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    FilePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. CC approvals (solo aprobado) - Control costos interno/0. CONSOLIDADOR APROBACIONES CC SP.xlsx",
    
    // Descargar y parsear el Excel
    Binario = Web.Contents(SiteUrl & "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(FilePath) & "')/$value"),
    Libro = try Excel.Workbook(Binary.Buffer(Binario), null, true) otherwise null,
    
    // Si falla la lectura, devolvemos tabla vacía, si no, tomamos la primera hoja
    Origen = if Libro = null then #table({"Proyecto:"}, {}) else Table.PromoteHeaders(Libro{0}[Data], [PromoteAllScalars=true]),

    // ============================================================
    // FILTRO Y MAPEO
    // ============================================================
    // 1. Filtrar por el proyecto actual (ignorando lo que esté después del guion)
    FiltroProyecto = Table.SelectRows(Origen, each 
        try [#"Proyecto:"] <> null and Text.StartsWith(Text.Upper([#"Proyecto:"]), Text.Upper(ParamProyecto)) otherwise false
    ),

    // 2. Renombrar las columnas exactas según la captura de error
    ColumnasRenombradas = Table.RenameColumns(FiltroProyecto, {
        {"Desc. - UM", "Ins"},
        {"Nombre del proveedor: ", "Nombre Contratista"},
        {"# CC", "# CC - Comparativo"},
        {"Cant. Total", "Cantidad CC Cons"},
        {"V/U TOTAL", "V/U CC cons"},
        {"VR TOTAL", "VT CC cons"}
    }, MissingField.Ignore),

    // 3. Estandarización de tipos de datos
    TextosLimpios = Table.TransformColumns(ColumnasRenombradas, {
        {"Ins", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Nombre Contratista", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"# CC - Comparativo", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Cantidad CC Cons", each FxToNumberFlex(_), type number},
        {"V/U CC cons", each FxToNumberFlex(_), type number},
        {"VT CC cons", each FxToNumberFlex(_), type number}
    }, null, MissingField.Ignore),

    // 4. Agregar la etiqueta Tipo y Centro de Costos (usando mapeo heurístico)
    AgregadoTipo = Table.AddColumn(TextosLimpios, "Tipo", each "CC Consolidado", type text),
    AgregadoCC = Table.AddColumn(AgregadoTipo, "Centro de Costos", each 
        if Text.StartsWith(Text.Upper(ParamProyecto), "PAMPLONA 1") and [#"# CC - Comparativo"] <> null then
            let
                ccPrefix = Text.Trim(Text.BeforeDelimiter([#"# CC - Comparativo"], "-")),
                matchFolder = List.Select(ListaCarpetas, (x) => Text.StartsWith(x, ccPrefix & "-")),
                result = if List.Count(matchFolder) > 0 then matchFolder{0} else FnMatchFolder([#"Proyecto:"])
            in result
        else 
            FnMatchFolder([#"Proyecto:"])
    , type text),

    // ============================================================
    // EXTRACCIÓN DE COLUMNAS PARA BD
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
