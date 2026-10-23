// ASIGNACION_OC (hoja ASIGNACION del flujo SINCO)
// Reparto de cada (orden, insumo) de ORDENES_COMPRA entre los items del
// presupuesto, segun el InformeOrdenDeCompraDetalladoInsumos (xlsx).
// Columnas fijas (subflujo ASIGNACION ORDENES; filtra por A = Cod Orden y
// D = INSUMO de la hoja ORDENES COMPRA, y toma de alli OC SINCO):
//   A Cod Orden  B CAPITULO  C ACTIVIDAD (codigo del item)  D INSUMO
//   E V/U  F COMPRADO (del item)
//   G CONSUMIDO = recibido del insumo en la orden x COMPRADO del item / COMPRADO total
// El estado de la asignacion queda en ORDENES COMPRA (ESTADO ASIGNACION).
let
    OrigenActual = try Text.Upper(Text.Trim(Origen)) otherwise "SINCO",
    ColsSalida = {"Cod Orden","CAPITULO","ACTIVIDAD","INSUMO","V/U","COMPRADO","CONSUMIDO"},

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
        if BinInforme = null then #table({"Código","Descripción","Ítem","Capitulo","Orden de Compra","Cantidad Comprada","Valor Unitario"}, {})
        else Table.PromoteHeaders(Excel.Workbook(BinInforme, null, true){0}[Data], [PromoteAllScalars=true]),
    BaseInforme = Table.SelectRows(HojaInforme, each Limpiar(Record.FieldOrDefault(_, "Orden de Compra", null)) <> ""),
    Lineas = Table.FromRecords(Table.TransformRows(BaseInforme, (r) => [
        #"Cod Orden" = Limpiar(Record.Field(r, "Orden de Compra")),
        INSUMO = Limpiar(Record.Field(r, "Descripción")),
        CodIns = Limpiar(Record.Field(r, "Código")),
        Capitulo = Limpiar(Record.FieldOrDefault(r, "Capitulo", null)),
        VU = FxNum(Record.Field(r, "Valor Unitario")),
        Comprado = FxNum(Record.Field(r, "Cantidad Comprada")),
        Item = FnCodigoItem(Limpiar(Record.Field(r, "Ítem")))
    ]), type table [#"Cod Orden" = text, INSUMO = text, CodIns = text, Capitulo = text, VU = number, Comprado = number, Item = text]),
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

    // ===== 3. Una fila por (orden, insumo, item) con el recibido prorrateado =====
    PorItem = Table.Group(ConCompra, {"Cod Orden", "INSUMO", "Item"}, {
        {"CodIns", each List.First([CodIns]), type text},
        {"Capitulo", each List.First([Capitulo]), type text},
        {"VU", each List.First([VU]), type number},
        {"Comprado", each List.Sum([Comprado]), type number}}),
    TotalIns = Table.Group(ConCompra, {"Cod Orden", "INSUMO"}, {{"CompradoTotal", each List.Sum([Comprado]), type number}}),
    ConTotal = Table.ExpandTableColumn(
        Table.NestedJoin(PorItem, {"Cod Orden", "INSUMO"}, TotalIns, {"Cod Orden", "INSUMO"}, "__T", JoinKind.LeftOuter),
        "__T", {"CompradoTotal"}, {"CompradoTotal"}),
    ConRecibida = Table.ExpandTableColumn(
        Table.NestedJoin(Table.AddColumn(ConTotal, "__KOI", each [Cod Orden] & "|" & [CodIns], type text), {"__KOI"}, RecibidasSolo, {"__KOI"}, "__R", JoinKind.LeftOuter),
        "__R", {"CantRecibida"}, {"CantRecibida"}),
    Ordenado = Table.Sort(ConRecibida, {{"Cod Orden", Order.Ascending}, {"INSUMO", Order.Ascending}, {"Item", Order.Ascending}}),
    Filas = Table.FromRecords(Table.TransformRows(Ordenado, (r) => [
        #"Cod Orden" = r[Cod Orden],
        CAPITULO = r[Capitulo],
        ACTIVIDAD = r[Item],
        INSUMO = r[INSUMO],
        #"V/U" = r[VU],
        COMPRADO = Number.Round(r[Comprado], 4),
        CONSUMIDO =
            if r[CantRecibida] = null or r[CompradoTotal] = null or r[CompradoTotal] = 0 then 0
            else Number.Round(r[CantRecibida] * r[Comprado] / r[CompradoTotal], 4)
    ])),
    Salida = if OrigenActual = "ORACLE" or Table.RowCount(Ordenado) = 0 then #table(ColsSalida, {}) else Table.SelectColumns(Filas, ColsSalida)
in
    Salida
