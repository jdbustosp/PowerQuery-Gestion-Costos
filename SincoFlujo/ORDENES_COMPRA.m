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
//   Q NIT proveedor (facturacion por orden o por nombre)  R Nombre proveedor (fila del
//   proveedor en "ESTADO DE ORDENES"; el flujo busca el tercero por este nombre). Se recalculan.
// OC SINCO se copia a todas las filas de la misma orden (una sola OC en SINCO).
// El reparto por item para la asignacion esta en ASIGNACION_OC.
let
    OrigenActual = try Text.Upper(Text.Trim(Origen)) otherwise "SINCO",
    ColsBase = {"Cod Orden","INSUMO","V/U","COMPRADO","CONSUMIDO","TOTAL","ITEM SALIDA"},
    ColsEstado = {"ESTADO OC","ESTADO ASIGNACION","ESTADO PREAPROBACION","ESTADO APROBACION","ESTADO ENTRADA","ESTADO SALIDA","ESTADO APROBACION E/S","OC SINCO","SALIDA SINCO"},
    // Al final de la hoja (despues de las del flujo) para no correr las columnas que lee el PAD; se recalculan
    ColsProv = {"NIT proveedor","Nombre proveedor"},
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

    // ===== 2b. Proveedor por orden: fila "codigo - NOMBRE" bajo el encabezado de la orden (Columna1 vacia);
    //          NIT de "Facturacion de entradas y actas" por orden o, si no, por nombre (con los de pagos) =====
    FilasProv = Table.SelectRows(OrdenFD, each [__Orden] <> null and Limpiar([Columna1]) = "" and Limpiar([Columna2]) <> "" and Text.Upper(Limpiar([Columna2])) <> "PROVEEDOR"),
    ProvPorOrden = Table.Distinct(Table.FromRecords(Table.TransformRows(FilasProv, (r) =>
        let p = Limpiar(r[Columna2]) in [__KO = r[__Orden], Prov = if Text.Contains(p, " - ") then Text.Trim(Text.AfterDelimiter(p, " - ")) else p]),
        type table [__KO = text, Prov = text]), {"__KO"}),
    MapaProv = Record.FromList(ProvPorOrden[Prov], List.Transform(ProvPorOrden[__KO], each "k" & _)),
    NormNombre = (v as any) as text => Text.Select(Text.Upper(Limpiar(v)), {"A".."Z","0".."9"}),
    EsNumero = (v as any) as logical => let t = Limpiar(v) in t <> "" and Text.Select(t, {"0".."9"}) = t,
    LeerInforme = (bin as nullable binary) as list =>
        if bin = null then {} else
        let
            raw = try Excel.Workbook(bin, null, true){0}[Data] otherwise Html.Table(Text.FromBinary(bin, 1252), List.Transform({1..4}, each {"Columna" & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}), [RowSelector="tr"]),
            cols = Table.ColumnNames(raw)
        in Table.ToRecords(Table.RenameColumns(raw, List.Zip({cols, List.Transform({1..List.Count(cols)}, each "Columna" & Text.From(_))}))),
    Campo = (r as record, c as text) as any => Record.FieldOrDefault(r, c, null),
    // facturacion = Orden|Entrada|NIT|Nombre; pagos = Fecha|Contrato|Beneficiario|NIT
    Fact = List.Select(LeerInforme(FnGetBinario("FACTURACI")), each EsNumero(Campo(_, "Columna1")) and EsNumero(Campo(_, "Columna3"))),
    Pagos = List.Select(LeerInforme(FnGetBinario("RELACION DE PAGOS")), each EsNumero(Campo(_, "Columna2")) and EsNumero(Campo(_, "Columna4"))),
    NitOrden = Table.Distinct(#table(type table [Doc = text, NIT = text], List.Transform(Fact, each {Limpiar(Campo(_, "Columna1")), Limpiar(Campo(_, "Columna3"))})), {"Doc"}),
    MapaNitOrden = Record.FromList(NitOrden[NIT], List.Transform(NitOrden[Doc], each "k" & _)),
    NitNombre = Table.Distinct(Table.SelectRows(#table(type table [Clave = text, NIT = text],
        List.Transform(Fact, each {NormNombre(Campo(_, "Columna4")), Limpiar(Campo(_, "Columna3"))}) &
        List.Transform(Pagos, each {NormNombre(Campo(_, "Columna3")), Limpiar(Campo(_, "Columna4"))})), each [Clave] <> ""), {"Clave"}),
    MapaNitNombre = Record.FromList(NitNombre[NIT], List.Transform(NitNombre[Clave], each "k" & _)),
    FnProveedor = (orden as text) as record =>
        let
            nombre = Record.FieldOrDefault(MapaProv, "k" & orden, null),
            porOrden = Record.FieldOrDefault(MapaNitOrden, "k" & orden, null),
            nit = if porOrden <> null then porOrden else if nombre = null then null else Record.FieldOrDefault(MapaNitNombre, "k" & NormNombre(nombre), null)
        in [#"NIT proveedor" = nit, #"Nombre proveedor" = nombre],

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
    Nuevo = if Table.RowCount(Ordenado) = 0 then #table(ColsBase & ColsProv, {}) else
        Table.FromRecords(Table.TransformRows(Ordenado, (r) => FnProveedor(r[Cod Orden]) & [
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
