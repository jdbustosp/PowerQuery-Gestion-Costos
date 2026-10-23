// ORDENES_COMPRA (hoja ORDENES COMPRA del flujo SINCO)
// Una fila por (orden, insumo) del InformeOrdenDeCompraDetalladoInsumos (xlsx)
// de la carpeta Actual del centro de costos; CONSUMIDO = cantidad recibida
// del insumo en la orden segun "ESTADO DE ORDENES" (HTML).
// Columnas fijas (subflujos ORDENES DE COMPRA, ENTRADAS y SALIDAS):
//   A Cod Orden  B INSUMO  C V/U  D COMPRADO (suma de los items)  E CONSUMIDO
//   F TOTAL (V/U x COMPRADO)  G ITEM SALIDA (codigo del item con mas comprado)
//   H..P los escribe el flujo y se conservan (clave Cod Orden|INSUMO):
//   ESTADO OC, ESTADO ASIGNACION, ESTADO PREAPROBACION, ESTADO APROBACION,
//   ESTADO ENTRADA, ESTADO SALIDA, ESTADO APROBACION E/S, OC SINCO, SALIDA SINCO.
// OC SINCO se copia a todas las filas de la misma orden (una sola OC en SINCO).
// El reparto por item para la asignacion esta en ASIGNACION_OC.
let
    OrigenActual = try Text.Upper(Text.Trim(Origen)) otherwise "SINCO",
    ColsBase = {"Cod Orden","INSUMO","V/U","COMPRADO","CONSUMIDO","TOTAL","ITEM SALIDA"},
    ColsEstado = {"ESTADO OC","ESTADO ASIGNACION","ESTADO PREAPROBACION","ESTADO APROBACION","ESTADO ENTRADA","ESTADO SALIDA","ESTADO APROBACION E/S","OC SINCO","SALIDA SINCO"},
    ColsSalida = ColsBase & ColsEstado,

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

    FnCodigoItem = (t as text) as text => if Text.Contains(t, " - ") then Text.Trim(Text.BeforeDelimiter(t, " - ")) else t,

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

    // ===== 1. Informe detallado de insumos por orden (xlsx) =====
    BinInforme = FnGetBinario("INFORMEORDEN"),
    HojaInforme =
        if BinInforme = null then #table({"Código","Descripción","Ítem","Orden de Compra","Cantidad Comprada","Valor Unitario"}, {})
        else Table.PromoteHeaders(Excel.Workbook(BinInforme, null, true){0}[Data], [PromoteAllScalars=true]),
    BaseInforme = Table.SelectRows(HojaInforme, each Limpiar(Record.FieldOrDefault(_, "Orden de Compra", null)) <> ""),
    Lineas = Table.FromRecords(Table.TransformRows(BaseInforme, (r) => [
        #"Cod Orden" = Limpiar(Record.Field(r, "Orden de Compra")),
        INSUMO = Limpiar(Record.Field(r, "Descripción")),
        CodIns = Limpiar(Record.Field(r, "Código")),
        VU = FxNum(Record.Field(r, "Valor Unitario")),
        Comprado = FxNum(Record.Field(r, "Cantidad Comprada")),
        Item = FnCodigoItem(Limpiar(Record.Field(r, "Ítem")))
    ]), type table [#"Cod Orden" = text, INSUMO = text, CodIns = text, VU = number, Comprado = number, Item = text]),
    ConCompra = Table.Buffer(Table.SelectRows(Lineas, each [Comprado] > 0)),

    // ===== 2. Cantidad recibida por (orden, codigo insumo) =====
    BinEstado = FnGetBinario("ESTADO DE ORDENES"),
    ColsOC = List.Transform({1..7}, each {"Columna" & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}),
    TablaEstado = if BinEstado = null then #table(List.Transform({1..7}, each "Columna" & Text.From(_)), {}) else Html.Table(Text.FromBinary(BinEstado, 1252), ColsOC, [RowSelector="tr"]),
    ConOrden = Table.AddColumn(TablaEstado, "__Orden", each
        let t = Limpiar([Columna1])
        in if Text.Contains(Text.Upper(t), "ORDEN DE COMPRA NO") then Text.Select(t, {"0".."9"}) else null, type text),
    OrdenFD = Table.FillDown(ConOrden, {"__Orden"}),
    Recibidas = Table.SelectRows(OrdenFD, each [__Orden] <> null and Limpiar([Columna1]) <> "" and Text.Select(Limpiar([Columna1]), {"0".."9"}) = Limpiar([Columna1])),
    RecibidasSolo = Table.Buffer(Table.Distinct(Table.FromRecords(Table.TransformRows(Recibidas, (r) => [
        __KOI = r[__Orden] & "|" & Limpiar(r[Columna1]),
        CantRecibida = FxNum(r[Columna7])
    ]), type table [__KOI = text, CantRecibida = number]), {"__KOI"})),

    // ===== 3. Una fila por (orden, insumo) =====
    Agrupado = Table.Group(ConCompra, {"Cod Orden", "INSUMO"}, {
        {"CodIns", each List.First([CodIns]), type text},
        {"VU", each List.First([VU]), type number},
        {"Comprado", each List.Sum([Comprado]), type number},
        {"Item", each Table.Sort(Table.Group(_, {"Item"}, {{"C", each List.Sum([Comprado]), type number}}), {{"C", Order.Descending}}){0}[Item], type text}}),
    ConRecibida = Table.ExpandTableColumn(
        Table.NestedJoin(Table.AddColumn(Agrupado, "__KOI", each [Cod Orden] & "|" & [CodIns], type text), {"__KOI"}, RecibidasSolo, {"__KOI"}, "__R", JoinKind.LeftOuter),
        "__R", {"CantRecibida"}, {"CantRecibida"}),
    Ordenado = Table.Sort(ConRecibida, {{"Cod Orden", Order.Ascending}, {"INSUMO", Order.Ascending}}),
    Nuevo = if Table.RowCount(Ordenado) = 0 then #table(ColsBase, {}) else
        Table.FromRecords(Table.TransformRows(Ordenado, (r) => [
            #"Cod Orden" = r[Cod Orden],
            INSUMO = r[INSUMO],
            #"V/U" = r[VU],
            COMPRADO = Number.Round(r[Comprado], 4),
            CONSUMIDO = if r[CantRecibida] = null then 0 else Number.Round(r[CantRecibida], 4),
            TOTAL = Number.Round(r[VU] * r[Comprado], 2),
            #"ITEM SALIDA" = r[Item]
        ])),

    // ===== 4. Conservar lo que escribio el flujo; OC SINCO comun a la orden =====
    Conservado = Table.Buffer(FxConservarEstado(Nuevo, "ORDENES_COMPRA", {"Cod Orden", "INSUMO"}, ColsEstado, ColsSalida)),
    Txt = (v as any) as text => if v = null then "" else Text.Upper(Text.Trim(Text.From(v))),
    OCPorOrden = Table.Group(Table.AddColumn(Conservado, "__KO", each Txt([Cod Orden]), type text), {"__KO"},
        {{"__OC", each List.First(List.Select([OC SINCO], (v) => Txt(v) <> "")), type any}}),
    OCMapa = Record.FromList(OCPorOrden[__OC], List.Transform(OCPorOrden[__KO], each "k" & _)),
    ConOC = Table.FromRecords(Table.TransformRows(Conservado, (r) =>
        if Txt(r[OC SINCO]) <> "" then r
        else r & [#"OC SINCO" = Record.FieldOrDefault(OCMapa, "k" & Txt(r[Cod Orden]), null)])),
    Salida = if OrigenActual = "ORACLE" or Table.RowCount(Conservado) = 0 then #table(ColsSalida, {}) else Table.SelectColumns(ConOC, ColsSalida)
in
    Salida
