// ORA_FUENTES: informes Oracle (Unifier) del proyecto actual, leidos de SharePoint.
// Ruta: ORA_CONFIG[CarpetaReportes]/<ProyectoActual>/<Centro de Costos>/Actual
// (mismo patron REST que SP_CarpetasCC / SP_Archivos_Proyecto de SINCO).
// Devuelve un record perezoso: cada campo se evalua solo si una consulta lo usa.
//   CentroCostos  nombre de la carpeta de CC (en Oracle hay una sola: el proyecto)
//   Archivos      listado de Actual (Name, ServerRelativeUrl, TimeLastModified, Length)
//   PRESUPUESTOS  hoja cruda (sin encabezados): Seguimiento por paquete de trabajo
//   CONTROL       Cost Sheet con encabezados (equivale al Seguimiento por items de SINCO)
//   ASEGURADO, COMPRAS (ordenes OC-), CONTRATOS (contratos CO-): con encabezados
//   ALMACEN (ENTRADAS Y SALIDAS.xls) y FACTURACION (FACTURACION.xls): con encabezados
//   FnNum         numero desde texto en formato espanol o ingles ("3.845.646,84", "-")
// Encabezados duplicados: Power Query los nombra Cantidad, Cantidad_2, Cantidad_3...
let
    Proyecto = Text.Trim(Text.From(ProyectoActual)),
    SiteUrl = ORA_CONFIG[Sitio],
    BasePath = ORA_CONFIG[CarpetaReportes] & "/" & Proyecto,
    Headers = [Accept = "application/json;odata=nometadata"],
    FnEncode = F_Globales[FnEncode],

    CarpetasCC = try List.Distinct(SP_CarpetasCC[Centro de Costos]) otherwise {},
    // Sin carpetas: tabla vacia CON columnas (Table.Combine({}) no trae esquema y el Sort fallaria)
    Listado = if List.IsEmpty(CarpetasCC) then #table({"Name", "ServerRelativeUrl", "TimeLastModified", "Length", "Centro de Costos"}, {})
        else Table.Combine(List.Transform(CarpetasCC, (cc) =>
        let
            r = try Json.Document(Web.Contents(SiteUrl, [
                RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(BasePath & "/" & cc & "/Actual") & "')/Files",
                Query = [#"$select" = "Name,ServerRelativeUrl,TimeLastModified,Length"],
                Headers = Headers,
                Timeout = #duration(0, 0, 2, 0)])) otherwise null,
            t = if r <> null and Record.HasFields(r, "value") then Table.FromRecords(r[value]) else #table({"Name", "ServerRelativeUrl", "TimeLastModified", "Length"}, {})
        in Table.AddColumn(t, "Centro de Costos", each cc, type text))),
    Archivos = Table.Buffer(Table.Sort(
        Table.SelectRows(Listado, each not Text.StartsWith([Name], "~$")),
        {{"TimeLastModified", Order.Descending}})),

    FnBinario = (nombre as text) as binary =>
        let
            f = Table.SelectRows(Archivos, each Text.Upper([Name]) = Text.Upper(nombre)),
            vistos = Text.Combine(List.Transform(Archivos[Name], Text.From), ", ")
        in
            if Table.RowCount(f) = 0 then
                error Error.Record("Falta informe Oracle",
                    "No se encontro " & nombre & " para ProyectoActual = """ & Proyecto & """ en " & BasePath & "/<Centro de Costos>/Actual",
                    "Archivos vistos: " & (if vistos = "" then "(ninguno)" else vistos))
            else Binary.Buffer(F_Globales[FnReadSPBinary](SiteUrl, f{0}[ServerRelativeUrl])),

    // Primera hoja de datos (los .xls de BI Publisher traen una hoja XDO_METADATA)
    FnHoja = (bin as binary, inferir as logical) as table =>
        let
            wb = if inferir then Excel.Workbook(bin, [InferSheetDimensions = true, DelayTypes = true]) else Excel.Workbook(bin, null, true),
            hojas = Table.SelectRows(wb, each Record.FieldOrDefault(_, "Kind", "Sheet") = "Sheet"
                and not Text.Contains(Text.Upper(Record.FieldOrDefault(_, "Name", "")), "XDO_METADATA"))
        in hojas{0}[Data],

    FnNorm = (v as any) as text => Text.Upper(Text.Trim(if v = null then "" else Text.From(v))),
    // Promueve la primera fila (de las 15 primeras) que contenga TODOS los encabezados pedidos
    FnConEncabezado = (tbl as table, requeridos as list, nombre as text) as table =>
        let
            filas = List.FirstN(Table.ToRows(tbl), 15),
            pos = List.PositionOf(List.Transform(filas, (r) =>
                let vs = List.Transform(r, FnNorm) in List.AllTrue(List.Transform(requeridos, (q) => List.Contains(vs, Text.Upper(q))))), true),
            promovida = if pos < 0
                then error Error.Record("Encabezados no encontrados", nombre & ": no se hallo la fila con " & Text.Combine(requeridos, ", "))
                else Table.PromoteHeaders(Table.Skip(tbl, pos), [PromoteAllScalars = true])
        in Table.TransformColumnNames(promovida, Text.Trim),

    FnNum = (v as any) as nullable number =>
        if v = null then null
        else if Value.Is(v, type number) then v
        else
            let
                t = Text.Select(Text.From(v), {"0".."9", ",", ".", "-"}),
                coma = Text.Contains(t, ","), punto = Text.Contains(t, "."),
                limpio =
                    if t = "" or t = "-" then null
                    else if coma and punto then
                        (if Text.PositionOf(t, ",", Occurrence.Last) > Text.PositionOf(t, ".", Occurrence.Last)
                         then Text.Replace(Text.Replace(t, ".", ""), ",", ".")
                         else Text.Replace(t, ",", ""))
                    else if coma then Text.Replace(t, ",", ".")
                    // solo puntos: varios = separador de miles en espanol ("3.845.646")
                    else if List.Count(Text.PositionOfAny(t, {"."}, Occurrence.All)) > 1 then Text.Replace(t, ".", "")
                    else t
            in if limpio = null then null else try Number.FromText(limpio, "en-US") otherwise null,

    Resultado = [
        CentroCostos = if List.IsEmpty(CarpetasCC) then Proyecto else CarpetasCC{0},
        Archivos = Archivos,
        PRESUPUESTOS = FnHoja(FnBinario("PRESUPUESTOS.xlsx"), false),
        CONTROL = FnConEncabezado(FnHoja(FnBinario("CONTROL.xlsx"), true), {"CBS Item"}, "CONTROL.xlsx"),
        ASEGURADO = FnConEncabezado(FnHoja(FnBinario("ASEGURADO.xls"), false), {"Registro", "Proceso"}, "ASEGURADO.xls"),
        COMPRAS = FnConEncabezado(FnHoja(FnBinario("COMPRAS.xls"), false), {"Orden", "CBS", "Articulo"}, "COMPRAS.xls"),
        CONTRATOS = FnConEncabezado(FnHoja(FnBinario("CONTRATOS.xls"), false), {"Orden", "CBS", "Articulo"}, "CONTRATOS.xls"),
        ALMACEN = FnConEncabezado(FnHoja(FnBinario("ENTRADAS Y SALIDAS.xls"), false), {"Registro", "Proceso", "Orden", "Articulo"}, "ENTRADAS Y SALIDAS.xls"),
        FACTURACION = FnConEncabezado(FnHoja(FnBinario("FACTURACION.xls"), false), {"Recepcion", "Orden", "Nit"}, "FACTURACION.xls"),
        FnNum = FnNum
    ]
in
    Resultado
