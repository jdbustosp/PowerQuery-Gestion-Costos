let
    // SINCO_FACTURAS (2026-10-05): facturas de SINCO por OC/contrato, desde los informes
    // "Informe relacion de pagos - pagos" (actas de contratos, valor Total = neto + IVA) y
    // "<centro> - Facturacion de entradas y actas" (entradas de almacen facturadas, Valor Total).
    // Los informes viven en <Reportes EDT>/<proyecto>/<centro>/Actual. Se leen aparte de
    // SP_Archivos_Proyecto para no mezclarlos con los demas reportes.
    FxToNumberFlex = F_Globales[FxToNumberFlex],
    FnDecodeHtml = F_Globales[FnDecodeHtml],
    FnBuildColumnas = F_Globales[FnBuildColumnas],
    FnEncode = F_Globales[FnEncode],
    ParamProyecto = Text.Trim(ProyectoActual),
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    BasePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. Reportes EDT - Control costos interno/" & ParamProyecto,
    FnTexto = (v as any) as text => let t = try Text.Trim(Text.From(v)) otherwise null in if t = null then "" else t,
    FnDigitos = (v as any) as text => Text.Select(FnTexto(v), {"0".."9"}),
    FnFactura = (v as any) as text =>
        let t = Text.Select(Text.Upper(FnTexto(v)), {"A".."Z", "0".."9"}),
            sinCeros = Text.TrimStart(t, "0")
        in sinCeros,

    CCFolders = try Table.RenameColumns(SP_CarpetasCC, {{"Centro de Costos", "Name"}}) otherwise #table({"Name"}, {}),
    ConArchivos = Table.AddColumn(CCFolders, "Archivos", each
        let r = try Json.Document(Web.Contents(SiteUrl, [
                RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(BasePath & "/" & [Name] & "/Actual") & "')/Files",
                Query = [#"$select" = "Name,ServerRelativeUrl"],
                Headers = [Accept = "application/json;odata=nometadata"],
                Timeout = #duration(0, 0, 2, 0)])) otherwise null
        in if r <> null and Record.HasFields(r, "value") then Table.FromRecords(r[value], {"Name", "ServerRelativeUrl"}, MissingField.UseNull) else null),
    Validos = Table.SelectRows(ConArchivos, each [Archivos] <> null),
    Expandido = Table.ExpandTableColumn(Validos, "Archivos", {"Name", "ServerRelativeUrl"}, {"Archivo", "Url"}),
    ConFuente = Table.AddColumn(Expandido, "Fuente", each
        let n = Text.Upper([Archivo]) in
        if Text.StartsWith(n, "~$") then null
        else if Text.Contains(n, "RELACION DE PAGOS") or Text.Contains(n, "RELACI#(00D3)N DE PAGOS") then "Acta"
        else if Text.Contains(n, "FACTURACION DE ENTRADAS") or Text.Contains(n, "FACTURACI#(00D3)N DE ENTRADAS") then "Entrada"
        else null, type text),
    Informes = Table.SelectRows(ConFuente, each [Fuente] <> null),

    FnLeer = (url as text, fuente as text) as table =>
        let
            bin = Binary.Buffer(Web.Contents(SiteUrl, [
                RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(url) & "')/$value",
                Headers = [Accept = "*/*"], Timeout = #duration(0, 0, 5, 0)])),
            n = if fuente = "Acta" then 31 else 12,
            raw = Html.Table(FnDecodeHtml(bin), FnBuildColumnas(n), [RowSelector = "tr"]),
            filas =
                if fuente = "Acta" then
                    Table.SelectRows(raw, each FnDigitos([Columna 2]) <> "" and FnDigitos([Columna 2]) = FnTexto([Columna 2]))
                else
                    Table.SelectRows(raw, each FnDigitos([Columna 1]) <> "" and FnDigitos([Columna 1]) = FnTexto([Columna 1]) and FnDigitos([Columna 2]) = FnTexto([Columna 2])),
            salida = Table.FromRecords(List.Transform(Table.ToRecords(filas), (r) =>
                if fuente = "Acta" then [
                    OC = FnDigitos(r[Columna 2]), Documento = "Acta " & FnTexto(r[Columna 5]),
                    NIT = FnDigitos(r[Columna 4]), Factura = FnTexto(r[Columna 6]),
                    Fecha = FnTexto(r[Columna 1]), Valor = FxToNumberFlex(r[Columna 13])]
                else [
                    OC = FnDigitos(r[Columna 1]), Documento = "EA " & FnTexto(r[Columna 2]),
                    NIT = FnDigitos(r[Columna 3]), Factura = FnTexto(r[Columna 8]),
                    Fecha = FnTexto(r[Columna 9]), Valor = FxToNumberFlex(r[Columna 11])]),
                type table [OC = text, Documento = text, NIT = text, Factura = text, Fecha = text, Valor = number])
        in salida,

    Leidos = Table.AddColumn(Informes, "Datos", each try FnLeer([Url], [Fuente]) otherwise null),
    ConDatos = Table.SelectRows(Leidos, each [Datos] <> null),
    Filas = if Table.IsEmpty(ConDatos)
        then #table(type table [Fuente = text, OC = text, Documento = text, NIT = text, Factura = text, Fecha = text, Valor = number], {})
        else Table.ExpandTableColumn(Table.SelectColumns(ConDatos, {"Fuente", "Datos"}), "Datos",
            {"OC", "Documento", "NIT", "Factura", "Fecha", "Valor"}),
    // El mismo informe puede estar en varias carpetas de centro (informe por proyecto)
    SinDuplicados = Table.Distinct(Filas, {"Fuente", "OC", "Documento", "Factura", "Valor"}),
    ConClave = Table.AddColumn(SinDuplicados, "FacturaClave", each FnFactura([Factura]), type text),
    Tipado = Table.TransformColumnTypes(ConClave, {{"Valor", type number}})
in
    Tipado
