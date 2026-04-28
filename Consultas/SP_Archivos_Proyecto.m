let
    // =========================================================
    // 🔥 CONSULTA MAESTRA DE SHAREPOINT (UNA SOLA LLAMADA)
    // =========================================================
    ParamProyecto = Text.Trim(ProyectoActual),
    RutaBase = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    ArchivosSharePoint = SharePoint.Files(RutaBase, [ApiVersion = 15]),

    // 🔥 PASO 1: Reducir columnas ANTES de filtrar (libera RAM desde el inicio)
    SoloColumnasNecesarias = Table.SelectColumns(ArchivosSharePoint, {"Name", "Content", "Folder Path"}),

    // 🔥 PASO 2: Filtro por proyecto y /Actual/
    ArchivosProyecto = Table.SelectRows(SoloColumnasNecesarias, each 
        Text.Contains([Folder Path], "/" & ParamProyecto & "/", Comparer.OrdinalIgnoreCase) and 
        Text.EndsWith([Folder Path], "/Actual/", Comparer.OrdinalIgnoreCase) and 
        not Text.StartsWith([Name], "~$")
    ),

    // 🔥 PASO 3: Pre-filtrar SOLO los archivos que realmente usamos
    // Esto evita que Table.Buffer guarde archivos irrelevantes
    ArchivosRelevantes = Table.SelectRows(ArchivosProyecto, each 
        Text.Contains([Name], "SEGUIMIENTO POR ITEMS", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "ANALISIS DE PRECIOS UNITARIOS", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "INFORMEORDEN", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "ESTADO DE ORDENES", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "ESTADO DE CONTRATOS", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "DESCUENTOS", Comparer.OrdinalIgnoreCase)
    ),

    // PASO 4: Extraemos el Centro de Costos de la ruta
    ConCentroCosto = Table.AddColumn(ArchivosRelevantes, "Centro de Costos", each 
        Text.Trim(Text.Replace(Text.AfterDelimiter([Folder Path], "/" & ParamProyecto & "/"), "/Actual/", ""))
    ),

    // PASO 5: Eliminamos Folder Path (ya no la necesitamos) y bufferizamos
    ColumnasMinimas = Table.SelectColumns(ConCentroCosto, {"Name", "Content", "Centro de Costos"}),
    TablaFinal = Table.Buffer(ColumnasMinimas)
in
    TablaFinal
