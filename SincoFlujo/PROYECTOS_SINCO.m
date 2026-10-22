let
    // Lista los proyectos disponibles = archivos .xlsx en la carpeta de
    // descargas (un archivo por proyecto). Se carga en la hoja CONFIG para
    // alimentar el desplegable de ProyectoActual. Agregar un proyecto nuevo
    // = subir "<PROYECTO>.xlsx" a la carpeta; aparece solo en la lista.
    //
    // La carpeta se prueba en 2 ubicaciones posibles (CarpetasCandidatas):
    // la actual y la original, por si vuelve a moverse en SharePoint. Si se
    // reubica a un tercer sitio, hay que agregar esa ruta a la lista.
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    RutaBase = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS",
    CarpetasCandidatas = {
        RutaBase & "/0. Descargas pptos - Control costos interno",
        RutaBase & "/DashBoard/0. Descargas pptos - Control costos interno"
    },
    FnEncode = F_Globales[FnEncode],

    FnListarCarpeta = (ruta as text) as nullable record =>
        try Json.Document(Web.Contents(SiteUrl, [
            RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(ruta) & "')/Files",
            Query = [#"$select" = "Name"],
            Headers = [Accept = "application/json;odata=nometadata"],
            Timeout = #duration(0, 0, 2, 0)
        ])) otherwise null,

    // Perezoso: prueba la 1a candidata y solo intenta la 2a si la 1a falla
    // (mismo fix que DESCARGAS.m - List.Transform llamaba SIEMPRE a ambas).
    Resp1 = FnListarCarpeta(CarpetasCandidatas{0}),
    Resp1Valido = Resp1 <> null and Record.HasFields(Resp1, "value"),
    Resp = if Resp1Valido then Resp1
           else let Resp2 = FnListarCarpeta(CarpetasCandidatas{1}) in
                if Resp2 <> null and Record.HasFields(Resp2, "value") then Resp2 else null,

    Archivos =
        if Resp = null or not Record.HasFields(Resp, "value")
        then #table({"Name"}, {})
        else Table.FromRecords(Resp[value]),

    SoloExcel = Table.SelectRows(Archivos, each
        not Text.StartsWith([Name], "~$") and
        Text.EndsWith(Text.Upper([Name]), ".XLSX") and
        Text.Upper([Name]) <> "DESCARGA PPTO.XLSX"),   // excluir el maestro historico

    ConProyecto = Table.AddColumn(SoloExcel, "Proyecto", each
        Text.Trim(Text.Start([Name], Text.Length([Name]) - 5)), type text),

    SoloProyecto = Table.SelectColumns(ConProyecto, {"Proyecto"}),
    SinDuplicados = Table.Distinct(SoloProyecto),
    Ordenada = Table.Sort(SinDuplicados, {{"Proyecto", Order.Ascending}})
in
    Ordenada
