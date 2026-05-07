let
    // ============================================================
    // FUNCIONES GLOBALES
    // ============================================================
    FnEncode = F_Globales[FnEncode],
    FxToNumberFlex = F_Globales[FxToNumberFlex],
    FnFormatCodigoAct = F_Globales[FnFormatCodigoAct],
    
    ParamProyecto = Text.Trim(Proyectoactual),

    // ============================================================
    // CONEXIÓN AL ARCHIVO EN SHAREPOINT
    // ============================================================
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    FilePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. CC approvals (solo aprobado) - Control costos interno/0. CONSOLIDADOR APROBACIONES CC SP.xlsx",
    
    // Descargar y parsear el Excel
    Binario = Web.Contents(SiteUrl & "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(FilePath) & "')/$value"),
    Libro = try Excel.Workbook(Binary.Buffer(Binario), null, true) otherwise null,
    
    // Si falla la lectura, devolvemos tabla vacía, si no, tomamos la primera hoja
    Origen = if Libro = null then #table({"Proyecto"}, {}) else Table.PromoteHeaders(Libro{0}[Data], [PromoteAllScalars=true]),

    // ============================================================
    // FILTRO Y MAPEO
    // ============================================================
    // 1. Filtrar por el proyecto actual (ignorando lo que esté después del guion)
    FiltroProyecto = Table.SelectRows(Origen, each 
        [Proyecto] <> null and 
        Text.StartsWith(Text.Upper([Proyecto]), Text.Upper(ParamProyecto))
    ),

    // 2. Renombrar las columnas según tus instrucciones
    // NOTA: Ajusta los nombres exactos si Excel los cargó ligeramente distintos
    ColumnasRenombradas = Table.RenameColumns(FiltroProyecto, {
        {"Desc. - UM", "Ins"},
        {"Nombre del proveedor", "Nombre Contratista"},
        {"# CC", "# CC - Comparativo"}
    }, MissingField.Ignore),

    // 3. Estandarización de tipos de datos
    TextosLimpios = Table.TransformColumns(ColumnasRenombradas, {
        {"Ins", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Nombre Contratista", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"# CC - Comparativo", each if _ = null then null else Text.Trim(Text.From(_)), type text}
        // Agrega aquí "CC" o "Actividad" si es necesario mapearlos para BD
    }, null, MissingField.Ignore),

    // 4. Agregar la etiqueta Tipo
    AgregadoTipo = Table.AddColumn(TextosLimpios, "Tipo", each "CC", type text),

    // ============================================================
    // EXTRACCIÓN DE COLUMNAS PARA BD
    // ============================================================
    // IMPORTANTE: Modifica los nombres "Cant. aprobacion" y "VR total aprobacion" 
    // según cómo se llamen exactamente en las columnas azules de tu Excel.
    TablaFinal = Table.SelectColumns(AgregadoTipo, 
        {
            "Tipo", 
            "Ins", 
            "Nombre Contratista", 
            "# CC - Comparativo"
            // "CC",                    // Descomentar si esta columna va a "Centro de Costos"
            // "Cant. aprobacion",      // <-- REEMPLAZAR por el nombre real de cantidad
            // "VR total aprobacion"    // <-- REEMPLAZAR por el nombre real de valor
        }, 
        MissingField.Ignore
    )
in
    TablaFinal
