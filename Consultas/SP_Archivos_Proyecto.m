let
    // =========================================================
    // 🔥 CONSULTA MAESTRA DE SHAREPOINT (UNA SOLA LLAMADA)
    // =========================================================
    // Esta consulta se ejecuta UNA sola vez y todas las demás
    // consultas (COMPRAS, CONTRATOS, PPTO_BD, ITEMSINSUMOS,
    // DESCUENTOS) la leen directamente desde memoria.
    // =========================================================
    ParamProyecto = Text.Trim(ProyectoActual),
    RutaBase = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    ArchivosSharePoint = SharePoint.Files(RutaBase, [ApiVersion = 15]),

    // Filtro general: solo archivos del proyecto actual, en carpetas /Actual/
    ArchivosProyecto = Table.SelectRows(ArchivosSharePoint, each 
        Text.Contains([Folder Path], "/" & ParamProyecto & "/", Comparer.OrdinalIgnoreCase) and 
        Text.EndsWith([Folder Path], "/Actual/", Comparer.OrdinalIgnoreCase) and 
        not Text.StartsWith([Name], "~$")
    ),

    // Extraemos el Centro de Costos de la ruta
    ConCentroCosto = Table.AddColumn(ArchivosProyecto, "Centro de Costos", each 
        Text.Trim(Text.Replace(Text.AfterDelimiter([Folder Path], "/" & ParamProyecto & "/"), "/Actual/", ""))
    ),

    // Solo conservamos las columnas que necesitamos (libera RAM)
    ColumnasMinimas = Table.SelectColumns(ConCentroCosto, {"Name", "Content", "Centro de Costos"}),

    // Materializamos una sola vez para que las 4+ consultas lean de la misma tabla en memoria
    TablaFinal = Table.Buffer(ColumnasMinimas)
in
    TablaFinal
