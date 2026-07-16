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
        RutaBase & "/DashBoard/0. Descargas pptos - Control costos interno",
        RutaBase & "/0. Descargas pptos - Control costos interno"
    },
    FnEncode = F_Globales[FnEncode],

    FnListarCarpeta = (ruta as text) as nullable record =>
        try Json.Document(Web.Contents(SiteUrl, [
            RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(ruta) & "')/Files",
            Query = [#"$select" = "Name"],
            Headers = [Accept = "application/json;odata=nometadata"],
            Timeout = #duration(0, 0, 2, 0)
        ])) otherwise null,

    Intentos = List.Transform(CarpetasCandidatas, each [Resp = FnListarCarpeta(_)]),
    IntentoValido = List.First(List.Select(Intentos, each [Resp] <> null and Record.HasFields([Resp], "value")), null),
    Resp = if IntentoValido = null then null else IntentoValido[Resp],

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
