// ENTRADAS_SINCO (hoja ENTRADAS del flujo SINCO): espejo de las entradas de almacen de la constructora.
// Una fila por linea del "Informe masivo de entradas por insumo" del centro del libro, con la factura
// de "Facturacion de entradas y actas" (por OC + entrada). Las lineas en cantidad 0 no van.
// Columnas fijas (subflujos ENTRADAS ALMACEN y FACTURAS ENTRADAS):
//   A Entrada origen (va como remision en SINCO)  B Cod Orden (la OC conserva el numero del original)
//   C Fecha  D INSUMO  E Cod insumo origen  F UM  G CANTIDAD  H Vr. unitario  I Observaciones
//   J Factura  K Fecha factura
//   L ESTADO ENTRADA  M ENTRADA SINCO  N ESTADO FACTURA: los escribe el flujo y se conservan entre
//   refrescos con la clave Entrada origen|Cod insumo origen; las lineas nuevas de una entrada ya creada
//   heredan su ENTRADA SINCO.
//   O NIT proveedor  P Nombre proveedor (el flujo busca el tercero por NIT y luego por nombre)
//   Q Lineas de la entrada (INSUMO^CANTIDAD~INSUMO^CANTIDAD): el flujo crea la entrada de una vez
//   R ESTADO APROBACION: lo escribe el flujo (subflujo APROBACION ENTRADAS SALIDAS)
let
    ColsBase = {"Entrada origen","Cod Orden","Fecha","INSUMO","Cod insumo origen","UM","CANTIDAD","Vr. unitario","Observaciones","Factura","Fecha factura"},
    ColsEstado = {"ESTADO ENTRADA","ENTRADA SINCO","ESTADO FACTURA","ESTADO APROBACION"},
    ColsProv = {"NIT proveedor","Nombre proveedor","Lineas de la entrada"},
    ColsSalida = ColsBase & {"ESTADO ENTRADA","ENTRADA SINCO","ESTADO FACTURA"} & ColsProv & {"ESTADO APROBACION"},
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    Headers = [Accept="application/json;odata=nometadata"],
    FnEncode = F_Globales[FnEncode],
    ArchivosProyecto = Table.Buffer(SP_Archivos_Flujo),
    Limpiar = (v as any) as text =>
        let t = if v = null then "" else Text.From(v)
        in Text.Trim(Text.Replace(Text.Replace(Text.Replace(t, "#(lf)", " "), "#(cr)", " "), "#(00A0)", " ")),
    FxNum = (v as any) as number =>
        if v is number then v else
        let n = try Number.FromText(Text.Replace(Limpiar(v), ",", ""), "en-US") otherwise null
        in if n = null then 0 else n,
    EsNumero = (v as any) as logical => let t = Limpiar(v) in t <> "" and Text.Select(t, {"0".."9"}) = t,
    FnGetBinario = (contiene as text) as nullable binary =>
        let
            Filas = Table.SelectRows(ArchivosProyecto, each Text.Contains([Name], contiene, Comparer.OrdinalIgnoreCase) and not Text.StartsWith([Name], "~$")),
            Ordenadas = Table.Sort(Filas, {{"Centro de Costos", Order.Ascending}, {"TimeLastModified", Order.Descending}})
        in
            if Table.RowCount(Ordenadas) = 0 then null
            else Binary.Buffer(Web.Contents(SiteUrl, [RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(Ordenadas{0}[ServerRelativeUrl]) & "')/$value", Headers = Headers, Timeout = #duration(0, 0, 10, 0)])),
    LeerHtml = (bin as nullable binary, n as number) as table =>
        let Nombres = List.Transform({1..n}, each "Columna" & Text.From(_))
        in if bin = null then #table(Nombres, {})
           else Html.Table(Text.FromBinary(bin, 1252), List.Transform({1..n}, each {"Columna" & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}), [RowSelector="tr"]),

    // Facturacion: Orden | Entrada | NIT | Nombre | ... | No. factura | Fecha factura; una fila por entrada
    Facturacion = Table.Buffer(Table.SelectRows(LeerHtml(FnGetBinario("FACTURACI"), 9), each EsNumero([Columna1]) and EsNumero([Columna2]))),
    MapaFact = Record.FromList(
        List.Transform(Table.ToRecords(Facturacion), each [NIT = Limpiar([Columna3]), Nombre = Limpiar([Columna4]), Factura = Limpiar([Columna8]), Fecha = Limpiar([Columna9])]),
        List.Transform(Table.ToRecords(Facturacion), each "k" & Limpiar([Columna1]) & "|" & Limpiar([Columna2]))),

    // Masivo: Sucursal | Cod | Descripcion | Agrupacion | UM | No. OC | No. EA | Fecha | Cantidad | Vr. Unitario | IVA | Valor Total | Proveedor | Obs.
    Masivo = Table.SelectRows(LeerHtml(FnGetBinario("MASIVO DE ENTRADAS"), 14), each EsNumero([Columna6]) and EsNumero([Columna7]) and FxNum([Columna9]) > 0),
    Vacia = [NIT = "", Nombre = "", Factura = "", Fecha = ""],
    Lineas = Table.FromRecords(Table.TransformRows(Table.AddIndexColumn(Masivo, "__I", 0, 1), (r) =>
        let
            f = Record.FieldOrDefault(MapaFact, "k" & Limpiar(r[Columna6]) & "|" & Limpiar(r[Columna7]), Vacia),
            Prov = Limpiar(r[Columna13]),
            Corte = Text.PositionOf(Prov, "-"),
            NombreMasivo = if Corte < 0 then Prov else Text.Trim(Text.Middle(Prov, Corte + 1)),
            NoVacio = (t as text) as nullable text => if t = "" then null else t
        in [
            #"Entrada origen" = Limpiar(r[Columna7]),
            #"Cod Orden" = Limpiar(r[Columna6]),
            Fecha = Limpiar(r[Columna8]),
            INSUMO = Limpiar(r[Columna3]),
            #"Cod insumo origen" = Limpiar(r[Columna2]),
            UM = Limpiar(r[Columna5]),
            CANTIDAD = FxNum(r[Columna9]),
            #"Vr. unitario" = FxNum(r[Columna10]),
            Observaciones = NoVacio(Text.Replace(Limpiar(r[Columna14]), "|", " ")),
            Factura = NoVacio(f[Factura]),
            #"Fecha factura" = NoVacio(f[Fecha]),
            #"NIT proveedor" = NoVacio(f[NIT]),
            #"Nombre proveedor" = if f[Nombre] <> "" then f[Nombre] else NoVacio(NombreMasivo),
            __I = r[__I]
        ])),

    // Lineas de la entrada: el flujo las pasa juntas al motor, que crea la entrada con todas
    TextoLineas = Table.Group(Lineas, {"Entrada origen"}, {{"Lineas de la entrada", (t) => Text.Combine(List.Transform(Table.ToRecords(t), (x) => x[INSUMO] & "^" & Number.ToText(x[CANTIDAD], "0.######", "en-US")), "~"), type text}}),
    MapaLineas = Record.FromList(TextoLineas[Lineas de la entrada], List.Transform(TextoLineas[Entrada origen], each "k" & _)),
    Ordenado = if Table.RowCount(Lineas) = 0 then Lineas else Table.Sort(Lineas, {{each Number.From([Entrada origen]), Order.Ascending}, {"__I", Order.Ascending}}),
    Nuevo = if Table.RowCount(Ordenado) = 0 then #table(ColsBase & ColsProv, {}) else
        Table.FromRecords(Table.TransformRows(Ordenado, (r) =>
            Record.SelectFields(r, ColsBase & {"NIT proveedor","Nombre proveedor"}) & [#"Lineas de la entrada" = Record.Field(MapaLineas, "k" & r[Entrada origen])])),

    // Lo que escribio el flujo se conserva; una linea nueva de una entrada ya creada hereda su numero
    Conservado = Table.Buffer(FxConservarEstado(Nuevo, "ENTRADAS_SINCO", {"Entrada origen", "Cod insumo origen"}, ColsEstado, ColsSalida)),
    Txt = (v as any) as text => if v = null then "" else Text.Upper(Text.Trim(Text.From(v))),
    PorOrigen = Table.Group(Table.AddColumn(Conservado, "__KE", each Txt([Entrada origen]), type text), {"__KE"}, {{"__EA", each List.First(List.Select([ENTRADA SINCO], (v) => Txt(v) <> "")), type any}}),
    MapaEA = Record.FromList(PorOrigen[__EA], List.Transform(PorOrigen[__KE], each "k" & _)),
    ConEA = Table.FromRecords(Table.TransformRows(Conservado, (r) =>
        if Txt(r[ENTRADA SINCO]) <> "" then r else r & [#"ENTRADA SINCO" = Record.FieldOrDefault(MapaEA, "k" & Txt(r[Entrada origen]), null)])),
    Salida = if Table.RowCount(Conservado) = 0 then #table(ColsSalida, {}) else Table.SelectColumns(ConEA, ColsSalida)
in
    Salida
