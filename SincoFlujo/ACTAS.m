// ACTAS_SINCO (hoja ACTAS del flujo SINCO): cortes de los contratos por grupo, espejo de la constructora.
// Una acta SINCO por acta de origen ("Detallado actas" del centro del libro); una fila por grupo del acta.
// Columnas fijas (subflujo ACTAS):
//   A Cod Contrato  B Contrato SINCO (No SINCO de la tabla CONTRATOS_SINCO)
//   C Linea (numero o grupo): codigo del item; el motor lo casa con el grupo "item - descripcion"
//   D Cantidad (del grupo en ESA acta)  E ESTADO (lo escribe el flujo)
//   F Acta origen  G Fecha acta  H Valor acta (del grupo)  I Factura ("Informe relacion de pagos")
// El flujo abre un acta nueva cada vez que cambia Cod Contrato + Acta origen.
// Cantidad: si el item del contrato tiene una sola linea en el origen, la cantidad del origen.
// Si tiene varias (nuestro grupo junta varios insumos), la cantidad sale del valor: acumulado del acta
// = Cantidad grupo x valor acumulado / valor contratado (Cantidad grupo x Valor, tabla CONTRATOS_SINCO),
// redondeado a 10 decimales y sin pasar la cantidad del grupo; desde 99,9999 % del valor cierra en ella.
// Cada fila lleva la diferencia con el acumulado anterior. Filas en cantidad 0 no van.
// Clave del estado: Cod Contrato | Acta origen | Linea.
// Codigo de item repetido en el APU (PLANTILLA lo renumera a .901...): la Linea es el codigo SINCO de la
// ocurrencia que el contrato tiene como grupo en CONTRATOS_SINCO; si tiene varias, se elige por la
// descripcion del concepto. Si aun no se distingue (decision de Juan 2026-11-09), la linea va a la
// ocurrencia cuya cantidad del APU es igual a la del acta (y le cabe); si no, llena en orden del APU lo que
// le falta a cada grupo del contrato (lo que sobre va a la ultima); el valor se parte en proporcion.
let
    MapaAct = FxGitHub("MAPA_ACT_SINCO"),
    OccA = List.Buffer(Table.ToRecords(MapaAct[Ocurrencias])),
    CantApu = Record.FromList(List.Transform(OccA, each if [Cantidad] = null then 0 else [Cantidad]), List.Transform(OccA, each "k" & [CodSinco])),
    // Candidatos del item en el contrato: {uno} si se distingue; varios si hay que repartir
    FnCandsActa = (c as text, item as text, desc as nullable text) as list =>
        if not List.Contains(MapaAct[Repetidos], item) then {item}
        else
            let
                cands = List.Select(MapaAct[Candidatos](item), each Record.HasFields(MapaGrupo, "k" & c & "|" & _)),
                rp = MapaAct[Repartir](item, null, null, desc, null, null)
            in if List.Count(cands) = 1 then cands
               else if List.Count(cands) = 0 then {MapaAct[Resolver](item, null, null, desc, null, null)}
               else if List.Count(rp) = 1 and List.Contains(cands, rp{0}[Cod]) then {rp{0}[Cod]}
               else cands,
    Cols = {"Cod Contrato","Contrato SINCO","Linea (numero o grupo)","Cantidad","ESTADO","Acta origen","Fecha acta","Valor acta","Factura"},
    ColsBase = List.RemoveItems(Cols, {"ESTADO"}),
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
    NumX = (v as any) as number =>
        if v is number then v else let n = try Number.From(v) otherwise null in if n = null then 0 else n,
    EsNumero = (v as any) as logical => let t = Limpiar(v) in t <> "" and Text.Select(t, {"0".."9"}) = t,
    TxtNum = (v as any) as text => if v is number then Number.ToText(v, "0") else Limpiar(v),
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

    // Contratos del libro: numero SINCO, cantidad del grupo y valor contratado por (contrato, grupo)
    Contratos =
        let
            T = try Excel.CurrentWorkbook(){[Name = "CONTRATOS_SINCO"]}[Content] otherwise null,
            Ok = T <> null and List.ContainsAll(Table.ColumnNames(T), {"Cod Contrato", "Grupo (actividad)", "Cantidad grupo", "Valor", "No SINCO"}),
            Filas = if not Ok then {} else Table.TransformRows(T, (r) =>
                [C = TxtNum(r[Cod Contrato]), G = Limpiar(r[#"Grupo (actividad)"]), CG = NumX(r[Cantidad grupo]),
                 ValC = NumX(r[Cantidad grupo]) * NumX(r[Valor]), NoSinco = TxtNum(r[No SINCO])]),
            Tab = Table.FromRecords(Filas, type table [C = text, G = text, CG = number, ValC = number, NoSinco = text])
        in Table.Buffer(Tab),
    PorGrupo = Table.Group(Contratos, {"C", "G"}, {{"CG", each List.Max([CG]), type number}, {"ValC", each List.Sum([ValC]), type number}}),
    MapaGrupo = Record.FromList(Table.ToRecords(PorGrupo), List.Transform(Table.ToRecords(PorGrupo), each "k" & [C] & "|" & [G])),
    PorContrato = Table.Group(Table.SelectRows(Contratos, each [NoSinco] <> ""), {"C"}, {{"NoSinco", each List.First([NoSinco]), type text}}),
    MapaSinco = Record.FromList(PorContrato[NoSinco], List.Transform(PorContrato[C], each "k" & _)),

    // Relacion de pagos: Fecha | Contrato No | Beneficiario | NIT | Acta No | Factura No | ...
    Pagos = Table.SelectRows(LeerHtml(FnGetBinario("RELACION DE PAGOS"), 6), each EsNumero([Columna2]) and EsNumero([Columna5])),
    PagosDist = Table.Distinct(Table.SelectRows(Pagos, each Limpiar([Columna6]) <> ""), {"Columna2", "Columna5"}),
    MapaFactura = Record.FromList(List.Transform(PagosDist[Columna6], Limpiar), List.Transform(Table.ToRecords(PagosDist), each "k" & Limpiar([Columna2]) & "|" & Limpiar([Columna5]))),

    // Detallado actas: Concepto (n - item - descripcion - id - ubicacion) | Tipo | UM | Contrato | Acta | Fecha | Cantidad | Vr. Unitario | Vr. IVA | Vr. Total
    Detalle = Table.SelectRows(LeerHtml(FnGetBinario("DETALLADO ACTAS"), 10), each EsNumero([Columna4]) and EsNumero([Columna5])),
    Lineas0 = Table.FromRecords(Table.TransformRows(Table.AddIndexColumn(Detalle, "__I", 0, 1), (r) =>
        let p = List.Transform(Text.Split(Limpiar(r[Columna1]), " - "), Text.Trim)
        in [C = Limpiar(r[Columna4]), Acta = Number.From(Limpiar(r[Columna5])), Fecha = Limpiar(r[Columna6]),
            ItemO = if List.Count(p) > 1 then p{1} else "",
            Cands = if List.Count(p) > 1 then FnCandsActa(Limpiar(r[Columna4]), p{1}, if List.Count(p) > 4 then Text.Combine(List.Range(p, 2, List.Count(p) - 4), " - ") else if List.Count(p) > 2 then p{2} else null) else {""},
            Linea = p{0} & "|" & (if List.Count(p) > 3 then p{List.Count(p) - 2} else ""),
            Cant = FxNum(r[Columna7]), Valor = FxNum(r[Columna10]), __I = r[__I]]),
        type table [C = text, Acta = number, Fecha = text, ItemO = text, Cands = list, Linea = text, Cant = number, Valor = number, __I = number]),
    TipoLin = type table [C = text, Acta = number, Fecha = text, Linea = text, Cant = number, Valor = number, Item = text, ItemO = text],
    Lineas1 = Table.FromRecords(Table.TransformRows(Table.SelectRows(Lineas0, each List.Count([Cands]) = 1), (f) =>
        [C = f[C], Acta = f[Acta], Fecha = f[Fecha], Linea = f[Linea], Cant = f[Cant], Valor = f[Valor], Item = f[Cands]{0}, ItemO = f[ItemO]]), TipoLin),
    // Reparto de las lineas que el contrato tiene en varias ocurrencias, en orden de acta
    GruposN = Table.Group(Table.SelectRows(Lineas0, each List.Count([Cands]) > 1), {"C", "ItemO"},
        {{"T", each Table.Sort(_, {{"Acta", Order.Ascending}, {"__I", Order.Ascending}}), type table}}),
    LineasN = List.Combine(List.Transform(Table.ToRecords(GruposN), (g) =>
        let
            filas = Table.ToRecords(g[T]),
            cands = filas{0}[Cands],
            // item con varias lineas en el origen: se pasa por valor (la cantidad de cada linea no es la del grupo),
            // asi que se llena por el valor que le falta a cada grupo
            porV = List.Count(List.Distinct(List.Transform(filas, each [Linea]))) > 1,
            ini = Record.FromList(List.Transform(cands, each Record.Field(Record.FieldOrDefault(MapaGrupo, "k" & g[C] & "|" & _, [CG = 0, ValC = 0]), if porV then "ValC" else "CG")), cands),
            pasos = List.Accumulate(filas, [Pend = ini, Out = {}], (s, f) =>
                let
                    q = if porV then f[Valor] else f[Cant],
                    exacto = List.Select(cands, each Number.Abs(Record.FieldOrDefault(CantApu, "k" & _, -1) - f[Cant]) <= 0.01 and Record.Field(s[Pend], _) >= q - 0.01),
                    llen = List.Accumulate(cands, [Resto = q, P = {}], (a, cd) =>
                        let t = Number.Round(List.Max({0, List.Min({a[Resto], Record.Field(s[Pend], cd)})}), 10)
                        in if t > 0 then [Resto = Number.Round(a[Resto] - t, 10), P = a[P] & {{cd, t}}] else a),
                    ult = List.Last(cands),
                    porLlenado =
                        if llen[Resto] <= 0 then llen[P]
                        else if List.Count(llen[P]) > 0 and List.Last(llen[P]){0} = ult then List.RemoveLastN(llen[P], 1) & {{ult, List.Last(llen[P]){1} + llen[Resto]}}
                        else llen[P] & {{ult, llen[Resto]}},
                    partes = if q <= 0 then {{cands{0}, q}} else if List.Count(exacto) = 1 then {{exacto{0}, q}} else porLlenado,
                    pend = List.Accumulate(partes, s[Pend], (pp, x) => Record.TransformFields(pp, {{x{0}, each _ - x{1}}})),
                    out = List.Transform(partes, (x) => [C = f[C], Acta = f[Acta], Fecha = f[Fecha], Linea = f[Linea],
                        Cant = if not porV then x{1} else if q = 0 then f[Cant] else f[Cant] * x{1} / q,
                        Valor = if porV then x{1} else if q = 0 then f[Valor] else f[Valor] * x{1} / q, Item = x{0}, ItemO = g[ItemO]])
                in [Pend = pend, Out = s[Out] & out])
        in pasos[Out])),
    Lineas = Table.Combine({Lineas1, Table.FromRecords(LineasN, TipoLin)}),
    LineasOk = Table.Buffer(Table.SelectRows(Lineas, each [Item] <> "")),
    // Items del contrato con mas de una linea de origen (se pasan por valor)
    // (se cuenta sobre el item de origen: sus repartos entre ocurrencias tambien van por valor)
    Varias = Table.SelectRows(Table.Group(LineasOk, {"C", "ItemO"}, {{"N", each List.Count(List.Distinct([Linea])), Int64.Type}, {"Items", each List.Distinct([Item]), type list}}), each [N] > 1),
    SetVarias = List.Buffer(List.Distinct(List.Combine(List.Transform(Table.ToRecords(Varias), (v) => List.Transform(v[Items], each v[C] & "|" & _))))),
    PorActa = Table.Group(LineasOk, {"C", "Item", "Acta"}, {{"Fecha", each List.First([Fecha]), type text}, {"Cant", each List.Sum([Cant]), type number}, {"Valor", each List.Sum([Valor]), type number}}),

    // Cantidad por valor acumulado, item a item del contrato en orden de acta
    PorItem = Table.Group(PorActa, {"C", "Item"}, {{"T", each Table.Sort(_, {{"Acta", Order.Ascending}}), type table}}),
    Filas = List.Combine(List.Transform(Table.ToRecords(PorItem), (g) =>
        let
            k = g[C] & "|" & g[Item],
            grp = Record.FieldOrDefault(MapaGrupo, "k" & k, null),
            porValor = List.Contains(SetVarias, k) and grp <> null and grp[ValC] <> 0,
            filas = Table.ToRecords(g[T]),
            pasos = List.Accumulate(filas, [Acum = 0, Prev = 0, Out = {}], (s, f) =>
                let
                    acum = s[Acum] + f[Valor],
                    r = acum / grp[ValC],
                    obj = if r >= 0.999999 then grp[CG] else List.Min({grp[CG], Number.Round(grp[CG] * r, 10)}),
                    cant = if porValor then Number.Round(obj - s[Prev], 10) else f[Cant]
                in [Acum = acum, Prev = if porValor then obj else s[Prev], Out = s[Out] & {f & [CantSinco = cant]}])
        in List.Transform(pasos[Out], each _ & [Clave = k]))),
    Nuevo = if List.Count(Filas) = 0 then #table(ColsBase, {}) else
        let
            Base = Table.FromRecords(List.Transform(List.Select(Filas, each [CantSinco] <> 0), (f) => [
                #"Cod Contrato" = f[C],
                #"Contrato SINCO" = Record.FieldOrDefault(MapaSinco, "k" & f[C], null),
                #"Linea (numero o grupo)" = f[Item],
                Cantidad = f[CantSinco],
                #"Acta origen" = Text.From(f[Acta]),
                #"Fecha acta" = f[Fecha],
                #"Valor acta" = f[Valor],
                Factura = Record.FieldOrDefault(MapaFactura, "k" & f[C] & "|" & Text.From(f[Acta]), null),
                __A = f[Acta]
            ])),
            Ordenado = Table.Sort(Base, {{"Cod Contrato", Order.Ascending}, {"__A", Order.Ascending}, {"Linea (numero o grupo)", Order.Ascending}})
        in Table.SelectColumns(Ordenado, ColsBase),

    Final = FxConservarEstado(Nuevo, "ACTAS_SINCO", {"Cod Contrato", "Acta origen", "Linea (numero o grupo)"}, {"ESTADO"}, Cols, false)
in
    Final
