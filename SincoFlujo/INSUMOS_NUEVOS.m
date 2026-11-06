// INSUMOS_NUEVOS (hoja INSUMOS NUEVOS del flujo SINCO)
// Pares (actividad, insumo) que estan en SEGUIMIENTO POR ITEMS pero no en el
// catalogo APU: hay que asociar el insumo a la actividad y, si el insumo no
// existe en ninguna actividad del APU, crearlo antes en el maestro.
// Columnas fijas (subflujos NUEVOS INSUMOS y ASOCIAR INSUMOS):
//   A ACTIVIDAD (codigo - descripcion)  B INSUMO (descripcion)  C TIPO
//   D UM (si SEGUIMIENTO no la trae, "Un")  E GRUPO ("OTROS")  F COD ACTIVIDAD  G ESTADO CREACION
//   H ESTADO ASOCIACION  I ES NUEVO (SI/NO)  J VR UNITARIO (proyectado)
//   K CANT APU = cantidad proyectada del insumo / cantidad de la actividad en
//     el APU (si la actividad no trae cantidad, la cantidad proyectada), 10 decimales
// G y H los escribe el flujo y se conservan (clave COD ACTIVIDAD|INSUMO); el
// OK de creacion se copia a las demas filas del mismo insumo, y las filas cuyo
// par ya esta en el APU salen (paso 4).
// Los codigos de insumo son del SINCO de la constructora: solo se usan para
// comparar sus dos reportes.
let
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    Headers = [Accept="application/json;odata=nometadata"],
    FnEncode = F_Globales[FnEncode],
    ArchivosProyecto = Table.Buffer(SP_Archivos_Flujo),

    Limpiar = (v as any) as text =>
        let t = if v = null then "" else Text.From(v)
        in Text.Trim(Text.Replace(Text.Replace(Text.Replace(t, "#(lf)", " "), "#(cr)", " "), "#(00A0)", " ")),

    FxNum = (v as any) as nullable number =>
        let t = Limpiar(v)
        in if t = "" then null else try Number.FromText(Text.Replace(t, ",", ""), "en-US") otherwise null,

    EsCodigoActividad = (c0 as text) as logical =>
        let
            partes = Text.Split(c0, "."),
            dosPartes = List.Count(partes) = 2,
            p1Num = if dosPartes then (try Number.FromText(partes{0}) otherwise null) else null,
            p2Num = if dosPartes then (try Number.FromText(partes{1}) otherwise null) else null
        in
            dosPartes and p1Num <> null and p2Num <> null and Text.Length(partes{0}) > 0 and Text.Length(partes{1}) > 0,

    EsCodigoInsumo = (c0 as text) as logical =>
        c0 <> "" and Text.Select(c0, {"0".."9"}) = c0,

    FnGetBinario = (contiene as text) as nullable binary =>
        let
            Filas = Table.SelectRows(ArchivosProyecto, each Text.Contains([Name], contiene, Comparer.OrdinalIgnoreCase)),
            Ordenadas = Table.Sort(Filas, {{"Centro de Costos", Order.Ascending}, {"TimeLastModified", Order.Descending}})
        in
            if Table.RowCount(Ordenadas) = 0 then null else
            Binary.Buffer(Web.Contents(SiteUrl, [
                RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(Ordenadas{0}[ServerRelativeUrl]) & "')/$value",
                Headers = Headers,
                Timeout = #duration(0, 0, 5, 0)
            ])),

    // ===== 1. SEGUIMIENTO: pares (actividad, insumo) con la proyeccion =====
    // Columnas: 1 Cod, 2 Desc, 3 Tipo, 4 UM, 5-7 Presupuestado, 8 Cant proyectada, 9 VrU proyectado
    BinSeg = FnGetBinario("SEGUIMIENTO POR ITEMS"),
    NombresSeg = List.Transform({1..9}, each "Columna" & Text.From(_)),
    ColsSeg = List.Transform({1..9}, each {"Columna" & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}),
    TablaSeg = if BinSeg = null then #table(NombresSeg, {}) else Html.Table(Text.FromBinary(BinSeg, 65001), ColsSeg, [RowSelector="tr"]),

    ClasSeg = Table.AddColumn(TablaSeg, "Clase", (r as record) =>
        let
            c0 = Limpiar(Record.Field(r, "Columna1")),
            c2 = Limpiar(Record.Field(r, "Columna3")),
            c3 = Limpiar(Record.Field(r, "Columna4"))
        in
            if c0 = "" or c0 = "Cod" or Text.StartsWith(c0, "SubCapitulo") then "Otro"
            else if c2 <> "" and EsCodigoInsumo(c0) then "Insumo"  // sin UM tambien es insumo
            else if EsCodigoActividad(c0) and Text.EndsWith(c0, ".000") then "Capitulo"
            else if EsCodigoActividad(c0) then "Actividad"
            else "Otro",
        type text),

    ConActCod = Table.AddColumn(ClasSeg, "ActCod", each if [Clase] = "Actividad" then Limpiar([Columna1]) else if [Clase] = "Capitulo" then "" else null, type text),
    ConActDesc = Table.AddColumn(ConActCod, "ActDesc", each if [Clase] = "Actividad" then Limpiar([Columna2]) else if [Clase] = "Capitulo" then "" else null, type text),
    FD = Table.FillDown(ConActDesc, {"ActCod", "ActDesc"}),

    InsumosSeg = Table.SelectRows(FD, each [Clase] = "Insumo" and [ActCod] <> null and [ActCod] <> ""),
    ParesSeg = Table.Distinct(
        Table.FromRecords(Table.TransformRows(InsumosSeg, (r) => [
            ActCod = r[ActCod], ActDesc = r[ActDesc],
            InsCod = Limpiar(r[Columna1]), InsDesc = Text.Trim(Text.Replace(Limpiar(r[Columna2]), "160;", "")),
            Tipo = Limpiar(r[Columna3]), UM = let u = Limpiar(r[Columna4]) in if u = "" then "Un" else u,
            ProyCant = FxNum(r[Columna8]), ProyVrU = FxNum(r[Columna9])
        ]), type table [ActCod = text, ActDesc = text, InsCod = text, InsDesc = text, Tipo = text, UM = text, ProyCant = nullable number, ProyVrU = nullable number]),
        {"ActCod", "InsCod"}
    ),

    // ===== 2. APU: pares, catalogo de insumos y cantidad de cada actividad =====
    BinApu = FnGetBinario("ANALISIS DE PRECIOS UNITARIOS"),
    ColsApu = {{"Columna1", "td:nth-child(1), th:nth-child(1)"}, {"Columna4", "td:nth-child(4), th:nth-child(4)"}},
    TablaApu = if BinApu = null then #table({"Columna1", "Columna4"}, {}) else Html.Table(Text.FromBinary(BinApu, 65001), ColsApu, [RowSelector="tr"]),

    ApuClasificado = Table.AddColumn(TablaApu, "Dato", each
        let
            t = Limpiar([Columna1]),
            antes = if Text.Contains(t, "-") then Text.Trim(Text.BeforeDelimiter(t, "-")) else ""
        in
            if antes = "" then null
            else if EsCodigoActividad(antes) then [K = "A", Cod = antes]
            else if EsCodigoInsumo(antes) then [K = "I", Cod = antes]
            else null),
    ApuConAct = Table.AddColumn(ApuClasificado, "ActApu", each if [Dato] <> null and [Dato][K] = "A" then [Dato][Cod] else null, type text),
    ApuFD = Table.FillDown(ApuConAct, {"ActApu"}),

    ApuActRows = Table.SelectRows(ApuConAct, each [ActApu] <> null),
    CantActividad = Table.Buffer(Table.Distinct(
        Table.FromRecords(Table.TransformRows(ApuActRows, (r) => [ActCod = r[ActApu], CantAct = FxNum(r[Columna4])]),
            type table [ActCod = text, CantAct = nullable number]),
        {"ActCod"})),

    ApuInsRows = Table.SelectRows(ApuFD, each [Dato] <> null and [Dato][K] = "I" and [ActApu] <> null),
    ParesApu = List.Buffer(List.Distinct(Table.TransformRows(ApuInsRows, (r) => r[Dato][Cod] & "|" & r[ActApu]))),
    InsumosApu = List.Buffer(List.Distinct(Table.TransformRows(ApuInsRows, (r) => r[Dato][Cod]))),

    // ===== 3. Pares del seguimiento que no estan en el APU =====
    Nuevos = Table.SelectRows(ParesSeg, each not List.Contains(ParesApu, [InsCod] & "|" & [ActCod])),
    ConCant = Table.ExpandTableColumn(
        Table.NestedJoin(Nuevos, {"ActCod"}, CantActividad, {"ActCod"}, "__C", JoinKind.LeftOuter),
        "__C", {"CantAct"}, {"CantAct"}),
    Ordenado = Table.Sort(ConCant, {{"ActCod", Order.Ascending}, {"InsDesc", Order.Ascending}}),

    // Tipos validos del Maestro Tipos Insumos del SINCO destino; lo demas = "Y" (POR CLASIFICAR)
    TiposValidos = {"E","I","O","M","N","Z","Y","P","S","T"},
    Nuevo = Table.FromRecords(Table.TransformRows(Ordenado, (r) => [
        ACTIVIDAD = r[ActCod] & " - " & r[ActDesc],
        INSUMO = r[InsDesc],
        TIPO = let t = Text.Upper(Text.Trim(r[Tipo])) in if List.Contains(TiposValidos, t) then t else "Y",
        UM = r[UM],
        GRUPO = "OTROS",
        #"COD ACTIVIDAD" = r[ActCod],
        #"ES NUEVO (SI/NO)" = if List.Contains(InsumosApu, r[InsCod]) then "NO" else "SI",
        #"VR UNITARIO" = r[ProyVrU],
        // SINCO guarda CANT APU x cantidad de la actividad: con 6 decimales una cantidad chica se distorsiona
        // (0,000013 x 237.363,65 = 3,09 en vez de 3), por eso va con 10.
        #"CANT APU" =
            if r[ProyCant] = null then null
            else if r[CantAct] = null or r[CantAct] = 0 then Number.Round(r[ProyCant], 10)
            else Number.Round(r[ProyCant] / r[CantAct], 10)
    ])),
    NuevoTabla = if Table.RowCount(Ordenado) = 0
        then #table({"ACTIVIDAD","INSUMO","TIPO","UM","GRUPO","COD ACTIVIDAD","ES NUEVO (SI/NO)","VR UNITARIO","CANT APU"}, {})
        else Nuevo,

    // ===== 4. Estados =====
    // Un par que ya aparece en el APU quedo asociado en SINCO: su fila sale de la
    // hoja (no se conservan huerfanas). Si falta alguno de los dos informes no se
    // puede saber, y se conservan todas para no perder lo hecho.
    ColsSalida = {"ACTIVIDAD","INSUMO","TIPO","UM","GRUPO","COD ACTIVIDAD","ESTADO CREACION","ESTADO ASOCIACION","ES NUEVO (SI/NO)","VR UNITARIO","CANT APU"},
    InformesCompletos = BinSeg <> null and BinApu <> null,
    Conservado = FxConservarEstado(NuevoTabla, "INSUMOS_NUEVOS", {"COD ACTIVIDAD", "INSUMO"}, {"ESTADO CREACION", "ESTADO ASOCIACION"},
        ColsSalida, not InformesCompletos),

    // El insumo se crea una vez en el maestro: si una fila ya dice "OK - ...", las
    // demas filas nuevas del mismo insumo toman ese mismo texto (lleva el codigo
    // SINCO "YA CREADO nnn" / "YA EXISTE nnn" que usa la asociacion).
    TxtK = (v as any) as text => if v = null then "" else Text.Upper(Text.Trim(Text.From(v))),
    EsOK = (v as any) as logical => Text.Contains(TxtK(v), "OK - "),
    ConOK = Table.SelectRows(Conservado, each EsOK([ESTADO CREACION])),
    Candidatos = Table.AddColumn(Table.AddColumn(ConOK, "__I", each TxtK([INSUMO]), type text), "__P", each
        (if Text.Contains(TxtK([ESTADO CREACION]), "PRODUCTIVO") then 0 else 2)
        + (if Text.Contains(TxtK([ESTADO CREACION]), "YA ") then 0 else 1), Int64.Type),
    Mejor = Table.Distinct(Table.Buffer(Table.Sort(Candidatos, {{"__I", Order.Ascending}, {"__P", Order.Ascending}})), {"__I"}),
    OkPorInsumo = Record.FromList(Mejor[ESTADO CREACION], Mejor[__I]),
    ConCopia = Table.AddColumn(Conservado, "__E", (r) =>
        let k = TxtK(r[INSUMO]) in
        if r[#"ES NUEVO (SI/NO)"] = "SI" and not EsOK(r[ESTADO CREACION]) and Record.HasFields(OkPorInsumo, k)
        then Record.Field(OkPorInsumo, k) else r[ESTADO CREACION]),
    Salida = Table.SelectColumns(
        Table.RenameColumns(Table.RemoveColumns(ConCopia, {"ESTADO CREACION"}), {{"__E", "ESTADO CREACION"}}),
        ColsSalida)
in
    Salida
