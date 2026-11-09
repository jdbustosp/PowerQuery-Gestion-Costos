// MAPA_ACT_SINCO: codigo de actividad del ORIGEN -> codigo de la actividad en NUESTRO SINCO.
// Se llama desde las otras consultas con FxGitHub("MAPA_ACT_SINCO") (no necesita consulta en el libro).
// Por que: PLANTILLA_PPTO_SINCO (paso 5b) renumera los codigos de actividad repetidos del APU: la primera
// ocurrencia viva (cantidad <> 0, orden del APU) conserva su codigo y las siguientes pasan a
// <capitulo>.901, .902... (saltando codigos ya usados). Los informes del origen (contratos, OC, salidas,
// actas, seguimiento) siguen trayendo el codigo repetido, asi que cada linea se resuelve aqui con la
// MISMA regla. Un codigo que no se repite se devuelve igual (los proyectos sin repetidos no cambian).
// Resolver(cod, codIns, insumo, desc, cant, campo) para un codigo repetido elige la ocurrencia por:
//   1. la unica que tiene ese insumo (codigo de insumo o, si no hay, nombre sin tildes ni signos);
//   2. la unica con cantidad > 0 del insumo en el campo pedido ("Contratado", "Comprado", "Consumido");
//   3. la unica cuya descripcion empieza por la descripcion dada;
//   4. la unica con la cantidad del insumo igual a cant (margen 0,01).
// Si no queda una sola: "AMBIGUO <cod>" (visible en el libro; el motor no lo ejecuta). Nunca adivina.
// Devuelve un registro: Repetidos (lista), Candidatos(cod) (lista de codigos SINCO), Ocurrencias (tabla),
// PorPosicionSeg (tabla posicion de la fila de actividad en el Seguimiento -> CodSinco) y Resolver.
let
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
    // Igual que PLANTILLA_PPTO_SINCO (cruce de ocurrencias APU <-> Seguimiento)
    FxNormDesc = (t as nullable text) as text =>
        Text.Upper(Text.Remove(Limpiar(t), {" ", "-", ".", ",", "(", ")", "/", "_", ":", ";", "#(00A0)"})),
    // Igual que FnClaveLimpia de CONTRATOS_SINCO (nombre de insumo sin unidad, tildes ni signos)
    FnClave = (t as nullable text) as text =>
        let
            s0 = Limpiar(t),
            s1 = Text.Upper(Text.Trim(if Text.Contains(s0, "(") then Text.BeforeDelimiter(s0, "(") else s0)),
            s2 = List.Accumulate({{"#(00C1)","A"},{"#(00C9)","E"},{"#(00CD)","I"},{"#(00D3)","O"},{"#(00DA)","U"},{"#(00D1)","N"},{"#(00DC)","U"}},
                s1, (s, p) => Text.Replace(s, p{0}, p{1}))
        in Text.Select(s2, {"A".."Z", "0".."9"}),

    FnBin = (contiene as text) as nullable binary =>
        let
            Filas = Table.SelectRows(ArchivosProyecto, each Text.Contains([Name], contiene, Comparer.OrdinalIgnoreCase) and not Text.StartsWith([Name], "~$")),
            Ordenadas = Table.Sort(Filas, {{"Centro de Costos", Order.Ascending}, {"TimeLastModified", Order.Descending}})
        in
            if Table.RowCount(Ordenadas) = 0 then null else
            Binary.Buffer(Web.Contents(SiteUrl, [
                RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(Ordenadas{0}[ServerRelativeUrl]) & "')/$value",
                Headers = Headers,
                Timeout = #duration(0, 0, 5, 0)
            ])),
    FnCols = (n as number) as list => List.Transform({1..n}, each {"Columna" & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}),

    // ===== 1. Actividades del APU y la regla de renumeracion (identica a PLANTILLA_PPTO_SINCO 5b) =====
    BinApu = FnBin("ANALISIS DE PRECIOS UNITARIOS"),
    ApuCruda = if BinApu = null then #table({"Columna1","Columna4"}, {}) else Html.Table(Text.FromBinary(BinApu, 65001), FnCols(4), [RowSelector="tr"]),
    ApuIdx = Table.AddIndexColumn(ApuCruda, "IndiceFila", 1, 1, Int64.Type),
    ApuActs = Table.Buffer(Table.FromRecords(List.RemoveNulls(Table.TransformRows(ApuIdx, (r) =>
        let
            texto = Limpiar(r[Columna1]),
            up = Text.Upper(texto),
            esCap = Text.StartsWith(up, "CAPITULO") or Text.StartsWith(up, "CAPÍTULO"),
            cod = if Text.Contains(texto, "-") then Text.Trim(Text.BeforeDelimiter(texto, "-")) else "",
            dRaw = Text.Trim(Text.AfterDelimiter(texto, "-")),
            d = if Text.StartsWith(dRaw, "-") then Text.Trim(Text.Middle(dRaw, 1)) else dRaw,
            cs = Limpiar(r[Columna4]),
            cant = try Number.FromText(cs, "en-US") otherwise try Number.FromText(cs, "es-ES") otherwise null,
            basura = Text.Contains(up, "ANALISIS DE PRECIOS") or Text.Contains(up, "DETALLE POR ITEMS")
        in
            if texto = "" or basura or esCap or cod = "" or not Text.Contains(cod, ".") then null
            else [IndiceFila = r[IndiceFila], Codigo = cod, DescN = FxNormDesc(d), Cantidad = cant])),
        type table [IndiceFila = Int64.Type, Codigo = text, DescN = text, Cantidad = nullable number])),
    CodigosUsados = List.Buffer(List.Distinct(ApuActs[Codigo])),
    Vivas = Table.SelectRows(ApuActs, each [Cantidad] <> 0),
    VivasOcc = if Table.IsEmpty(Vivas) then Table.AddColumn(Vivas, "Occ", each 1, Int64.Type) else
        Table.Combine(Table.Group(Vivas, {"Codigo"}, {{"t", each Table.AddIndexColumn(Table.Sort(_, {{"IndiceFila", Order.Ascending}}), "Occ", 1, 1, Int64.Type)}})[t]),
    Dup = Table.AddColumn(Table.SelectRows(VivasOcc, each [Occ] > 1), "Prefijo", each Text.BeforeDelimiter([Codigo], "."), type text),
    MapaNuevo = if Table.IsEmpty(Dup) then #table(type table [IndiceFila = Int64.Type, CodigoNuevo = text], {}) else
        Table.SelectColumns(Table.Combine(Table.Group(Dup, {"Prefijo"}, {{"t", (g) =>
            let
                pre = g{0}[Prefijo],
                libres = List.Buffer(List.Select(List.Transform({901..999}, each pre & "." & Text.From(_)), each not List.Contains(CodigosUsados, _))),
                ord = Table.AddIndexColumn(Table.Sort(g, {{"IndiceFila", Order.Ascending}}), "k", 0, 1, Int64.Type)
            in
                Table.AddColumn(ord, "CodigoNuevo", each if [k] < List.Count(libres) then libres{[k]}
                    else error "MAPA_ACT_SINCO: mas de 99 codigos de actividad repetidos en el capitulo " & pre, type text)}})[t]),
            {"IndiceFila", "CodigoNuevo"}),
    Repetidos = List.Buffer(List.Distinct(Dup[Codigo])),
    HayRepetidos = List.Count(Repetidos) > 0,

    // ===== 2. Ocurrencias vivas de los codigos repetidos, con sus insumos del Seguimiento =====
    // Seguimiento (insumo): 1 Cod, 2 Desc, 3 Tipo, 4 UM, 5 Pres cant, 8 Proy cant, 11 Contratado, 14 Comprado, 20 Consumido
    BinSeg = if HayRepetidos then FnBin("SEGUIMIENTO POR ITEMS") else null,
    SegCruda = if BinSeg = null then #table(List.Transform({1..20}, each "Columna" & Text.From(_)), {})
        else Html.Table(Text.FromBinary(BinSeg, 65001), FnCols(20), [RowSelector="tr"]),
    EsAct = (c as text) as logical =>
        let p = Text.Split(c, ".") in List.Count(p) = 2 and p{1} <> "000" and Text.Length(p{0}) > 0 and Text.Length(p{1}) > 0
            and Text.Select(c, {"0".."9", "."}) = c,
    SegClas = Table.AddIndexColumn(Table.AddColumn(SegCruda, "Clase", (r) =>
        let c0 = Limpiar(r[Columna1]) in
            if c0 = "" or Text.StartsWith(c0, "SubCapitulo") then "Otro"
            else if Limpiar(r[Columna3]) <> "" and Text.Select(c0, {"0".."9"}) = c0 then "Insumo"
            else if EsAct(c0) then "Actividad"
            else if Text.EndsWith(c0, ".000") then "Capitulo"
            else "Otro", type text), "PosSeg", 0, 1, Int64.Type),
    SegConAct = Table.FillDown(Table.AddColumn(SegClas, "PosAct", each if [Clase] = "Actividad" then [PosSeg] else if [Clase] = "Capitulo" then -1 else null, Int64.Type), {"PosAct"}),
    SegActs = Table.SelectRows(SegConAct, each [Clase] = "Actividad"),
    SegActsD = Table.AddColumn(Table.AddColumn(SegActs, "Codigo", each Limpiar([Columna1]), type text), "DescN", each FxNormDesc([Columna2]), type text),
    SegOrd = if Table.IsEmpty(SegActsD) then Table.AddColumn(SegActsD, "Ord", each 1, Int64.Type) else
        Table.Combine(Table.Group(SegActsD, {"Codigo", "DescN"}, {{"t", each Table.AddIndexColumn(Table.Sort(_, {{"PosSeg", Order.Ascending}}), "Ord", 1, 1, Int64.Type)}})[t]),
    ApuOrd = if Table.IsEmpty(ApuActs) then Table.AddColumn(ApuActs, "Ord", each 1, Int64.Type) else
        Table.Combine(Table.Group(ApuActs, {"Codigo", "DescN"}, {{"t", each Table.AddIndexColumn(Table.Sort(_, {{"IndiceFila", Order.Ascending}}), "Ord", 1, 1, Int64.Type)}})[t]),
    // Codigo SINCO de cada actividad del APU (null si tiene cantidad 0: PLANTILLA no la crea)
    ApuSinco = Table.AddColumn(Table.ExpandTableColumn(Table.NestedJoin(ApuOrd, {"IndiceFila"}, MapaNuevo, {"IndiceFila"}, "n", JoinKind.LeftOuter), "n", {"CodigoNuevo"}),
        "CodSinco", each if [Cantidad] = 0 then null else if [CodigoNuevo] <> null then [CodigoNuevo] else [Codigo], type nullable text),
    SegSinco = Table.ExpandTableColumn(Table.NestedJoin(SegOrd, {"Codigo", "DescN", "Ord"}, ApuSinco, {"Codigo", "DescN", "Ord"}, "a", JoinKind.LeftOuter), "a", {"CodSinco"}),
    PorPosicionSeg = Table.Buffer(Table.SelectColumns(SegSinco, {"PosSeg", "Codigo", "CodSinco"})),

    InsSeg = Table.Buffer(Table.FromRecords(Table.TransformRows(Table.SelectRows(SegConAct, each [Clase] = "Insumo" and [PosAct] <> null and [PosAct] >= 0), (r) => [
        PosAct = r[PosAct], CodIns = Limpiar(r[Columna1]), Clave = FnClave(r[Columna2]),
        Contratado = FxNum(r[Columna11]), Comprado = FxNum(r[Columna14]), Consumido = FxNum(r[Columna20])]),
        type table [PosAct = Int64.Type, CodIns = text, Clave = text, Contratado = number, Comprado = number, Consumido = number])),
    InsPorAct = Table.Group(InsSeg, {"PosAct"}, {{"Ins", each Table.ToRecords(_), type list}}),

    // Una fila por ocurrencia VIVA de un codigo repetido (lado APU = lo que existe en SINCO)
    OccApu = Table.SelectRows(ApuSinco, each List.Contains(Repetidos, [Codigo]) and [CodSinco] <> null),
    OccConSeg = Table.ExpandTableColumn(Table.NestedJoin(OccApu, {"Codigo", "DescN", "Ord"}, SegOrd, {"Codigo", "DescN", "Ord"}, "s", JoinKind.LeftOuter), "s", {"PosSeg"}),
    OccConIns = Table.ExpandTableColumn(Table.NestedJoin(OccConSeg, {"PosSeg"}, InsPorAct, {"PosAct"}, "i", JoinKind.LeftOuter), "i", {"Ins"}),
    Ocurrencias = Table.Buffer(Table.Sort(Table.SelectColumns(Table.TransformColumns(OccConIns, {{"Ins", each if _ = null then {} else _, type list}}),
        {"Codigo", "CodSinco", "DescN", "IndiceFila", "PosSeg", "Ins"}), {{"IndiceFila", Order.Ascending}})),
    OccRegs = List.Buffer(Table.ToRecords(Ocurrencias)),

    Candidatos = (cod as nullable text) as list =>
        let c = Limpiar(cod) in
            if not List.Contains(Repetidos, c) then {c}
            else List.Transform(List.Select(OccRegs, each [Codigo] = c), each [CodSinco]),

    Resolver = (cod as nullable text, codIns as nullable text, insumo as nullable text, desc as nullable text, cant as nullable number, campo as nullable text) as nullable text =>
        let
            c = Limpiar(cod)
        in
            if c = "" then cod
            else if not List.Contains(Repetidos, c) then c
            else
                let
                    occ = List.Select(OccRegs, each [Codigo] = c),
                    ci = Limpiar(codIns),
                    cl = FnClave(insumo),
                    casa = (i as record) as logical => if ci <> "" then i[CodIns] = ci else cl <> "" and i[Clave] = cl,
                    insDe = (o as record) as list => List.Select(o[Ins], casa),
                    f = Limpiar(campo),
                    qtyDe = (o as record) as number => if f = "" then 0 else List.Sum(List.Transform(insDe(o), each Record.FieldOrDefault(_, f, 0))) ?? 0,
                    p1 = List.Select(occ, each List.Count(insDe(_)) > 0),
                    pool1 = if List.Count(p1) > 0 then p1 else occ,
                    p2 = if f = "" then {} else List.Select(pool1, each qtyDe(_) > 0),
                    pool2 = if List.Count(p2) > 0 then p2 else pool1,
                    d = FxNormDesc(desc),
                    p3 = if d = "" then {} else List.Select(pool2, each Text.StartsWith([DescN], d)),
                    pool3 = if List.Count(p3) > 0 then p3 else pool2,
                    p4 = if cant = null or f = "" then {} else List.Select(pool3, each Number.Abs(qtyDe(_) - cant) <= 0.01),
                    elegido =
                        if List.Count(p1) = 1 then p1{0}
                        else if List.Count(p2) = 1 then p2{0}
                        else if List.Count(p3) = 1 then p3{0}
                        else if List.Count(p4) = 1 then p4{0}
                        else if List.Count(pool3) = 1 then pool3{0}
                        else null
                in
                    if elegido = null then "AMBIGUO " & c else elegido[CodSinco]
in
    [Repetidos = Repetidos, Candidatos = Candidatos, Ocurrencias = Ocurrencias, PorPosicionSeg = PorPosicionSeg, Resolver = Resolver]
