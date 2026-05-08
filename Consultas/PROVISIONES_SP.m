let
    // ============================================================
    // FUNCIONES GLOBALES
    // ============================================================
    FnEncode = F_Globales[FnEncode],
    FxToNumberFlex = F_Globales[FxToNumberFlex],
    FnRemoveAccentsSymbols = F_Globales[FnRemoveAccentsSymbols],
    
    ParamProyecto = Text.Trim(ProyectoActual),

    // ============================================================
    // CARPETAS DE PROYECTO ACTUAL (Para Mapeo Exacto de CC)
    // ============================================================
    ListaCarpetas = try List.Distinct(SP_Archivos_Proyecto[Centro de Costos]) otherwise {},
    FolderCount = List.Count(ListaCarpetas),
    
    // Función heurística para encontrar el nombre exacto de la carpeta sin acentos
    FnMatchFolder = (proyectoExcel as text) as text =>
        if FolderCount = 0 then proyectoExcel
        else if FolderCount = 1 then ListaCarpetas{0}
        else 
            let
                proyClean = FnRemoveAccentsSymbols(Text.Upper(proyectoExcel)),
                // Buscar carpetas cuya última palabra coincida o cuyo nombre sin espacios coincida
                Match = List.Select(ListaCarpetas, each 
                    let 
                        baseName = if Text.Contains(_, "-") then Text.Trim(Text.AfterDelimiter(_, "-")) else Text.Trim(_),
                        baseClean = FnRemoveAccentsSymbols(Text.Upper(baseName)),
                        lastWordFolder = List.Last(Text.Split(baseClean, " ")),
                        lastWordProy = List.Last(Text.Split(proyClean, " "))
                    in 
                        Text.Contains(proyClean, lastWordFolder) or 
                        Text.Contains(baseClean, lastWordProy) or 
                        Text.Replace(baseClean, " ", "") = Text.Replace(proyClean, " ", "")
                ),
                Result = if List.Count(Match) = 1 then Match{0} else proyectoExcel
            in
                Result,

    // ============================================================
    // CONEXIÓN AL ARCHIVO EN SHAREPOINT
    // ============================================================
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    FilePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. Provisiones - Control costos interno/0. CONSOLIDADOR PROVISIONES SP.xlsx",
    
    // Descargar y parsear el Excel
    Binario = Web.Contents(SiteUrl & "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(FilePath) & "')/$value"),
    Libro = try Excel.Workbook(Binary.Buffer(Binario), null, true) otherwise null,
    
    // Si falla la lectura, devolvemos tabla vacía, si no, tomamos la primera hoja
    Origen = if Libro = null then #table({"PROYECTO"}, {}) else Table.PromoteHeaders(Libro{0}[Data], [PromoteAllScalars=true]),

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

    // 3. Estandarización de tipos de datos
    TextosLimpios = Table.TransformColumns(ColumnasRenombradas, {
        {"PROYECTO", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Nombre Contratista", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"# OC / Contrato", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"VR_Bruto_con_desc", each FxToNumberFlex(_), type number}
    }, null, MissingField.Ignore),

    // 4. Agregar la etiqueta Tipo y Centro de Costos
    SinTipoOriginal = Table.RemoveColumns(TextosLimpios, {"Tipo"}, MissingField.Ignore),
    AgregadoTipo = Table.AddColumn(SinTipoOriginal, "Tipo", each "PROVISIONES", type text),
    AgregadoCC = Table.AddColumn(AgregadoTipo, "Centro de Costos", each 
        if Text.StartsWith(Text.Upper(ParamProyecto), "PAMPLONA 1") and [#"# OC / Contrato"] <> null then
            let
                ccPrefix = Text.Start([#"# OC / Contrato"], 4),
                matchFolder = List.Select(ListaCarpetas, (x) => Text.StartsWith(x, ccPrefix & "-")),
                result = if List.Count(matchFolder) > 0 then matchFolder{0} else FnMatchFolder([PROYECTO])
            in result
        else 
            FnMatchFolder([PROYECTO])
    , type text)

in
    AgregadoCC
