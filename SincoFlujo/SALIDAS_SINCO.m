// SALIDAS_SINCO (hoja SALIDAS del flujo SINCO): espejo de las salidas de almacen del origen
// Una fila por linea del "Masivo salidas descriptivas" (HTML) de la carpeta Actual del centro de costos.
// En SINCO se crea UNA salida por salida de origen, con todas sus lineas, el mismo vale y el mismo contratista.
// Columnas fijas (subflujos SALIDAS ALMACEN y APROBACION ENTRADAS SALIDAS):
//   A Salida origen  B Fecha  C Vale  D Descuento  E Observaciones  F INSUMO (descripcion)  G ITEM (codigo)
//   H CANTIDAD  I UM  J Cod insumo origen
//   K..M las escribe el flujo y se conservan (clave Salida origen|INSUMO|ITEM):
//   ESTADO SALIDA, SALIDA SINCO, ESTADO APROBACION.
//   N NIT contratista (facturacion o pagos, por nombre)  O Nombre contratista (el flujo busca el tercero
//   por NIT#NOMBRE). Se recalculan.
// Orden: por salida de origen y linea, para que el flujo agregue las lineas a la misma salida SINCO.
// SALIDA SINCO se copia a todas las filas de la misma salida de origen.
let
    OrigenActual = try Text.Upper(Text.Trim(Origen)) otherwise "SINCO",
    ColsBase = {"Salida origen","Fecha","Vale","Descuento","Observaciones","INSUMO","ITEM","CANTIDAD","UM","Cod insumo origen"},
    ColsEstado = {"ESTADO SALIDA","SALIDA SINCO","ESTADO APROBACION"},
    ColsProv = {"NIT contratista","Nombre contratista"},
    ColsSalida = ColsBase & ColsEstado & ColsProv,

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
            if Table.RowCount(Ordenadas) = 0 then null else
            Binary.Buffer(Web.Contents(SiteUrl, [
                RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(Ordenadas{0}[ServerRelativeUrl]) & "')/$value",
                Headers = Headers,
                Timeout = #duration(0, 0, 10, 0)
            ])),

    // ===== 1. Masivo de salidas: Fecha|Salida No.|Descuento|Contratista|Vale|Obs.|Cod|Descripcion|Item|UM|Cant|Vr.Unit|Vr.Total =====
    BinSalidas = FnGetBinario("MASIVO SALIDAS"),
    ColsHtml = List.Transform({1..13}, each {"Columna" & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}),
    TablaHtml = if BinSalidas = null then #table(List.Transform({1..13}, each "Columna" & Text.From(_)), {})
        else Html.Table(Text.FromBinary(BinSalidas, 1252), ColsHtml, [RowSelector="tr"]),
    Datos = Table.AddIndexColumn(Table.SelectRows(TablaHtml, each EsNumero([Columna2]) and Limpiar([Columna8]) <> ""), "__I", 0, 1),
    Lineas = Table.FromRecords(Table.TransformRows(Datos, (r) =>
        let c = Limpiar(r[Columna4]) in [
        #"Salida origen" = Limpiar(r[Columna2]),
        Fecha = Limpiar(r[Columna1]),
        Vale = Limpiar(r[Columna5]),
        Descuento = Limpiar(r[Columna3]),
        Observaciones = Limpiar(r[Columna6]),
        INSUMO = Limpiar(r[Columna8]),
        ITEM = Limpiar(r[Columna9]),
        CANTIDAD = Number.Round(FxNum(r[Columna11]), 4),
        UM = Limpiar(r[Columna10]),
        #"Cod insumo origen" = Limpiar(r[Columna7]),
        Contratista = if Text.Contains(c, " - ") then Text.Trim(Text.AfterDelimiter(c, " - ")) else c,
        __I = r[__I]
    ])),
    ConCantidad = Table.SelectRows(Lineas, each [CANTIDAD] > 0),

    // ===== 2. NIT del contratista por nombre (facturacion = Orden|Entrada|NIT|Nombre; pagos = Fecha|Contrato|Beneficiario|NIT) =====
    NormNombre = (v as any) as text => Text.Select(Text.Upper(Limpiar(v)), {"A".."Z","0".."9"}),
    LeerInforme = (bin as nullable binary) as list =>
        if bin = null then {} else
        let
            raw = try Excel.Workbook(bin, null, true){0}[Data] otherwise Html.Table(Text.FromBinary(bin, 1252), List.Transform({1..4}, each {"Columna" & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}), [RowSelector="tr"]),
            cols = Table.ColumnNames(raw)
        in Table.ToRecords(Table.RenameColumns(raw, List.Zip({cols, List.Transform({1..List.Count(cols)}, each "Columna" & Text.From(_))}))),
    Campo = (r as record, c as text) as any => Record.FieldOrDefault(r, c, null),
    Fact = List.Select(LeerInforme(FnGetBinario("FACTURACI")), each EsNumero(Campo(_, "Columna1")) and EsNumero(Campo(_, "Columna3"))),
    Pagos = List.Select(LeerInforme(FnGetBinario("RELACION DE PAGOS")), each EsNumero(Campo(_, "Columna2")) and EsNumero(Campo(_, "Columna4"))),
    NitNombre = Table.Distinct(Table.SelectRows(#table(type table [Clave = text, NIT = text],
        List.Transform(Fact, each {NormNombre(Campo(_, "Columna4")), Limpiar(Campo(_, "Columna3"))}) &
        List.Transform(Pagos, each {NormNombre(Campo(_, "Columna3")), Limpiar(Campo(_, "Columna4"))})), each [Clave] <> ""), {"Clave"}),
    MapaNitNombre = Record.FromList(NitNombre[NIT], List.Transform(NitNombre[Clave], each "k" & _)),

    // ===== 3. Una fila por linea, ordenadas por salida de origen y linea =====
    Ordenado = Table.Sort(ConCantidad, {{"Salida origen", Order.Ascending}, {"__I", Order.Ascending}}),
    Nuevo = if Table.RowCount(Ordenado) = 0 then #table(ColsBase & ColsProv, {}) else
        Table.FromRecords(Table.TransformRows(Ordenado, (r) =>
            Record.SelectFields(r, ColsBase) & [
            #"NIT contratista" = if r[Contratista] = "" then null else Record.FieldOrDefault(MapaNitNombre, "k" & NormNombre(r[Contratista]), null),
            #"Nombre contratista" = r[Contratista]
        ])),

    // ===== 4. Conservar lo que escribio el flujo; SALIDA SINCO comun a la salida de origen =====
    Conservado = Table.Buffer(FxConservarEstado(Nuevo, "SALIDAS_SINCO", {"Salida origen", "INSUMO", "ITEM"}, ColsEstado, ColsSalida)),
    Txt = (v as any) as text => if v = null then "" else Text.Upper(Text.Trim(Text.From(v))),
    SaPorOrigen = Table.Group(Table.AddColumn(Conservado, "__KS", each Txt([Salida origen]), type text), {"__KS"},
        {{"__SA", each List.First(List.Select([SALIDA SINCO], (v) => Txt(v) <> "")), type any}}),
    SaMapa = Record.FromList(SaPorOrigen[__SA], List.Transform(SaPorOrigen[__KS], each "k" & _)),
    ConSa = Table.FromRecords(Table.TransformRows(Conservado, (r) =>
        if Txt(r[SALIDA SINCO]) <> "" then r
        else r & [#"SALIDA SINCO" = Record.FieldOrDefault(SaMapa, "k" & Txt(r[Salida origen]), null)])),
    Salida = if OrigenActual = "ORACLE" or Table.RowCount(Conservado) = 0 then #table(ColsSalida, {}) else Table.SelectColumns(ConSa, ColsSalida)
in
    Salida
