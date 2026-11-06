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
let
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
    Lineas = Table.FromRecords(Table.TransformRows(Detalle, (r) =>
        let p = List.Transform(Text.Split(Limpiar(r[Columna1]), " - "), Text.Trim)
        in [C = Limpiar(r[Columna4]), Acta = Number.From(Limpiar(r[Columna5])), Fecha = Limpiar(r[Columna6]),
            Item = if List.Count(p) > 1 then p{1} else "", Linea = p{0} & "|" & (if List.Count(p) > 3 then p{List.Count(p) - 2} else ""),
            Cant = FxNum(r[Columna7]), Valor = FxNum(r[Columna10])]),
        type table [C = text, Acta = number, Fecha = text, Item = text, Linea = text, Cant = number, Valor = number]),
    LineasOk = Table.Buffer(Table.SelectRows(Lineas, each [Item] <> "")),
    // Items del contrato con mas de una linea de origen (se pasan por valor)
    Varias = Table.SelectRows(Table.Group(LineasOk, {"C", "Item"}, {{"N", each List.Count(List.Distinct([Linea])), Int64.Type}}), each [N] > 1),
    SetVarias = List.Buffer(List.Transform(Table.ToRecords(Varias), each [C] & "|" & [Item])),
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
