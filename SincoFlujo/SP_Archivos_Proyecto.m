let
    ParamProyecto = Text.Trim(ProyectoActual),
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    BasePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. Reportes EDT - Control costos interno/" & ParamProyecto,
    Headers = [Accept="application/json;odata=nometadata"],
    FnEncode = F_Globales[FnEncode],

    // Reutiliza SP_CarpetasCC (consulta compartida) en vez de repetir la MISMA
    // llamada GetFolderByServerRelativeUrl(...)/Folders por separado - evita un
    // viaje de red redundante (Power Query calcula SP_CarpetasCC una sola vez
    // y la comparte entre todos sus consumidores en el mismo refresco).
    CCFolders =
        try Table.RenameColumns(SP_CarpetasCC, {{"Centro de Costos", "Name"}})
        otherwise #table({"Name"}, {}),

    WithFiles = Table.AddColumn(CCFolders, "Archivos", each
        let
            ccActualPath = BasePath & "/" & [Name] & "/Actual",
            result = try Json.Document(Web.Contents(SiteUrl, [
                RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(ccActualPath) & "')/Files",
                Query = [#"$select" = "Name,ServerRelativeUrl,TimeLastModified,Length"],
                Headers = Headers,
                Timeout = #duration(0, 0, 2, 0)
            ])) otherwise null
        in
            if result <> null and Record.HasFields(result, "value") then Table.FromRecords(result[value]) else null
    ),

    ValidCCs = Table.SelectRows(WithFiles, each [Archivos] <> null),
    Expanded = Table.ExpandTableColumn(
        ValidCCs,
        "Archivos",
        {"Name", "ServerRelativeUrl", "TimeLastModified", "Length"},
        {"FileName", "ServerRelativeUrl", "TimeLastModified", "Length"}
    ),

    Relevant = Table.SelectRows(Expanded, each
        not Text.StartsWith([FileName], "~$") and (
            Text.Contains([FileName], "SEGUIMIENTO POR ITEMS",         Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "ANALISIS DE PRECIOS UNITARIOS", Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "INFORMEORDEN",                  Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "ESTADO DE ORDENES",             Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "INFORME ENTRADAS DE ALMACEN",   Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "INFORME ENTRADAS DE ALMACÉN",   Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "ENTRADAS POR INSUMO",           Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "MASIVO SALIDAS",                Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "ESTADO DE CONTRATOS",           Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "DESCUENTOS",                    Comparer.OrdinalIgnoreCase)
        )
    ),

    Typed = Table.TransformColumnTypes(Relevant, {{"TimeLastModified", type datetimezone}, {"Length", Int64.Type}}, "en-US"),
    Sorted = Table.Sort(Typed, {{"Name", Order.Ascending}, {"FileName", Order.Ascending}, {"TimeLastModified", Order.Descending}}),
    Resultado = Table.Buffer(Table.RenameColumns(
        Table.SelectColumns(Sorted, {"Name", "FileName", "ServerRelativeUrl", "TimeLastModified", "Length"}),
        {{"Name", "Centro de Costos"}, {"FileName", "Name"}}
    )),

    // Sin reportes, COMPRAS/CONTRATOS/BD fallan mas abajo con un error opaco
    // ("The column 'Centro de Costos' of the table wasn't found"). Se corta aqui
    // con un mensaje que dice que proyecto y que carpeta se buscaron.
    CarpetasVistas = Text.Combine(List.Transform(Table.Column(CCFolders, "Name"), Text.From), ", "),
    Final =
        if Table.IsEmpty(Resultado) then
            error Error.Record(
                "Sin reportes SINCO",
                "No se encontraron reportes SINCO para ProyectoActual = """ & ParamProyecto &
                """. Verifique que el parametro ProyectoActual del libro corresponda a este proyecto " &
                "y que los reportes esten en: " & BasePath & "/<Centro de Costos>/Actual",
                "Carpetas de CC encontradas: " & (if CarpetasVistas = "" then "(ninguna)" else CarpetasVistas))
        else Resultado
in
    Final