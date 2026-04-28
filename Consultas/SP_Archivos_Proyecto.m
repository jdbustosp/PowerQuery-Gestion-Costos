let
    // =========================================================
    // 🚀 CONSULTA MAESTRA - SHAREPOINT.CONTENTS (ULTRA RÁPIDO)
    // =========================================================
    // Navega directamente a la carpeta del proyecto en vez de
    // descargar el catálogo de TODOS los archivos del sitio.
    // =========================================================
    ParamProyecto = Text.Trim(ProyectoActual),

    // PASO 1: Conectar al sitio y navegar por la estructura de carpetas
    Site = SharePoint.Contents("https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv", [ApiVersion=15]),
    
    // PASO 2: Navegar al documento library → carpeta del proyecto
    DeptTec     = Site{[Name="Departamento Tecnico"]}[Content],
    CoordPres   = DeptTec{[Name="COORDINACION DE PRESUPUESTOS"]}[Content],
    ReportesEDT = CoordPres{[Name="Reportes EDT"]}[Content],
    Proyecto    = ReportesEDT{[Name=ParamProyecto]}[Content],

    // PASO 3: Cada fila aquí es un Centro de Costos (carpeta)
    CarpetasCC = Table.SelectRows(Proyecto, each [Kind] = "Folder"),

    // PASO 4: Para cada CC, entrar a /Actual/ y leer sus archivos
    ConArchivos = Table.AddColumn(CarpetasCC, "ArchivosActual", each 
        try 
            let 
                contenidoCC = [Content],
                actualFolder = contenidoCC{[Name="Actual"]}[Content],
                soloArchivos = Table.SelectRows(actualFolder, each 
                    [Kind] = "File" and not Text.StartsWith([Name], "~$")
                )
            in soloArchivos
        otherwise null
    ),

    // PASO 5: Filtrar CCs que no tienen carpeta Actual
    CCValidos = Table.SelectRows(ConArchivos, each [ArchivosActual] <> null),

    // PASO 6: Expandir archivos con el nombre del CC
    Expandido = Table.ExpandTableColumn(
        Table.SelectColumns(CCValidos, {"Name", "ArchivosActual"}),
        "ArchivosActual", 
        {"Name", "Content"}, 
        {"FileName", "Content"}
    ),
    Renombrado = Table.RenameColumns(Expandido, {{"Name", "Centro de Costos"}, {"FileName", "Name"}}),

    // PASO 7: Pre-filtrar SOLO los archivos que usamos (6 tipos)
    SoloRelevantes = Table.SelectRows(Renombrado, each 
        Text.Contains([Name], "SEGUIMIENTO POR ITEMS", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "ANALISIS DE PRECIOS UNITARIOS", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "INFORMEORDEN", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "ESTADO DE ORDENES", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "ESTADO DE CONTRATOS", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "DESCUENTOS", Comparer.OrdinalIgnoreCase)
    ),

    TablaFinal = Table.Buffer(SoloRelevantes)
in
    TablaFinal
