let
    // ============================================================
    // DESCARGAS: lee la tabla DESCARGA del archivo del proyecto en
    // ".../0. Descargas pptos - Control costos interno/<Proyecto>.xlsx".
    // El archivo se localiza por nombre (exacto primero, luego "contiene"),
    // asi que agregar un proyecto nuevo = subir su archivo, sin tocar codigo.
    //
    // La carpeta se prueba en 2 ubicaciones posibles (CarpetasCandidatas):
    // la actual y la original, por si vuelve a moverse en SharePoint. Si se
    // reubica a un tercer sitio, hay que agregar esa ruta a la lista.
    // ============================================================
    FxToNumberFlex = F_Globales[FxToNumberFlex],
    FnCleanText = F_Globales[FnCleanText],
    FnNormalizeSpaces = F_Globales[FnNormalizeSpaces],
    FnReadSPBinary = F_Globales[FnReadSPBinary],
    FnEncode = F_Globales[FnEncode],

    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    RutaBase = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS",
    CarpetasCandidatas = {
        RutaBase & "/0. Descargas pptos - Control costos interno",
        RutaBase & "/DashBoard/0. Descargas pptos - Control costos interno"
    },
    ParamProyecto = Text.Trim(ProyectoActual),
    ProyUp = Text.Upper(ParamProyecto),

    ColumnasFinales = {
        "Proyecto", "Centro de Costos", "Subcapitulo", "Capitulo", "Actividad", "Codigo ins", "Ins",
        "Cantidad ppto (CC)", "V/U ppto (CC)", "Valor Total ppto (CC)",
        "# CC - Comparativo", "# CC", "Comparativo"
    },
    TablaVacia = #table(ColumnasFinales, {}),

    // ---------- Localizar la carpeta (probando cada candidata) ----------
    FnListarCarpeta = (ruta as text) as nullable record =>
        try Json.Document(Web.Contents(SiteUrl, [
            RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(ruta) & "')/Files",
            Query = [#"$select" = "Name,ServerRelativeUrl"],
            Headers = [Accept = "application/json;odata=nometadata"],
            Timeout = #duration(0, 0, 2, 0)
        ])) otherwise null,

    // Perezoso: prueba la 1a candidata y solo intenta la 2a si la 1a falla,
    // en vez de llamar SIEMPRE a ambas con List.Transform (List.Transform evalua
    // cada elemento sin importar si el primero ya tuvo exito - un viaje de red
    // desperdiciado en cada refresco desde que la carpeta volvio a su ruta original).
    Resp1 = FnListarCarpeta(CarpetasCandidatas{0}),
    Resp1Valido = Resp1 <> null and Record.HasFields(Resp1, "value"),
    Listado = if Resp1Valido then Resp1
              else let Resp2 = FnListarCarpeta(CarpetasCandidatas{1}) in
                   if Resp2 <> null and Record.HasFields(Resp2, "value") then Resp2 else null,

    // ---------- Localizar el archivo del proyecto dentro de esa carpeta ----------
    Archivos =
        if Listado = null or not Record.HasFields(Listado, "value")
        then #table({"Name", "ServerRelativeUrl"}, {})
        else Table.FromRecords(Listado[value]),
    SoloExcel = Table.SelectRows(Archivos, each
        not Text.StartsWith([Name], "~$") and
        Text.EndsWith(Text.Upper([Name]), ".XLSX")),
    NombreBase = (n as text) as text => Text.Upper(Text.Trim(Text.Start(n, Text.Length(n) - 5))),
    Exactos = Table.SelectRows(SoloExcel, each NombreBase([Name]) = ProyUp),
    Contienen = Table.SelectRows(SoloExcel, each Text.Contains(Text.Upper([Name]), ProyUp)),
    Candidatos = if Table.RowCount(Exactos) > 0 then Exactos else Contienen,
    Ruta = if Table.RowCount(Candidatos) = 0 then null else Candidatos{0}[ServerRelativeUrl],

    // ---------- Leer la tabla DESCARGA ----------
    Binario = if Ruta = null then null else FnReadSPBinary(SiteUrl, Ruta),
    Libro = if Binario = null then null else (try Excel.Workbook(Binario, null, true) otherwise null),
    TablaCruda =
        if Libro = null then null
        else try Libro{[Item = "DESCARGA", Kind = "Table"]}[Data]
             otherwise (try Libro{[Item = "DESCARGAS", Kind = "Sheet"]}[Data] otherwise null),
    TablaDescargas = if TablaCruda = null then TablaVacia else TablaCruda,

    // Filtro defensivo: el archivo ya es por proyecto, pero si alguien sube
    // un archivo combinado, igual sale solo el proyecto actual.
    Filtrado =
        if Table.HasColumns(TablaDescargas, "Proyecto")
        then Table.SelectRows(TablaDescargas, each
            Text.Upper(Text.Trim(Text.From(if [Proyecto] = null then "" else [Proyecto]))) = ProyUp)
        else TablaDescargas,
    SinFilasVacias = Table.SelectRows(Filtrado, each
        (try Text.Trim(Text.From(if [Centro de Costos] = null then "" else [Centro de Costos])) otherwise "") <> ""),

    // ---------- Limpieza y tipos ----------
    TextosLimpios = Table.TransformColumns(SinFilasVacias, {
        {"Proyecto", each FnCleanText(_), type text},
        {"Centro de Costos", each FnCleanText(_), type text},
        {"Subcapitulo", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Capitulo", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Actividad", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Ins", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"# CC - Comparativo", each FnNormalizeSpaces(_), type text},
        {"# CC", each FnNormalizeSpaces(_), type text},
        {"Comparativo", each FnNormalizeSpaces(_), type text},

        {"Cantidad ppto (CC)", each FxToNumberFlex(_), type number},
        {"V/U ppto (CC)", each FxToNumberFlex(_), type number},
        {"Valor Total ppto (CC)", each FxToNumberFlex(_), type number}
    }, null, MissingField.Ignore),

    TiposFinales = try Table.TransformColumnTypes(TextosLimpios, {{"Codigo ins", Int64.Type}}) otherwise TextosLimpios,

    TablaFinal = Table.SelectColumns(TiposFinales, ColumnasFinales, MissingField.UseNull),
    Resultado = Table.Buffer(TablaFinal)
in
    Resultado
