// FxConservarEstado
// Conserva entre refrescos las columnas que escribe el flujo PAD (estados,
// numeros SINCO, fechas manuales) en una hoja que llena Power Query.
//   Nuevo       tabla recien calculada (sin estados o con ellos vacios)
//   NombreTabla nombre de la tabla de Excel que carga esta misma consulta
//   Claves      columnas que identifican la fila (se comparan sin mayusculas ni espacios)
//   ColsEstado  columnas que se copian de la tabla anterior
//   ColsSalida  columnas y orden de salida
//   ConservarHuerfanas (opcional, por defecto true): las filas anteriores con
//     algun estado cuya clave ya no viene en Nuevo se conservan al final (lo
//     hecho en SINCO no se pierde de la hoja); false las descarta.
// Orden de salida: filas de Nuevo en su orden y luego las conservadas.
(Nuevo as table, NombreTabla as text, Claves as list, ColsEstado as list, ColsSalida as list, optional ConservarHuerfanas as nullable logical) as table =>
let
    Raw = try Excel.CurrentWorkbook(){[Name = NombreTabla]}[Content] otherwise null,
    Prev =
        if Raw = null or not List.Contains(Table.ColumnNames(Raw), Claves{0}) then #table(ColsSalida, {})
        else Table.SelectColumns(Raw, ColsSalida, MissingField.UseNull),

    Txt = (v as any) as text => if v = null then "" else Text.Upper(Text.Trim(Text.From(v))),
    FnK = (r as record) as text => Text.Combine(List.Transform(Claves, (c) => Txt(Record.FieldOrDefault(r, c, null))), "|"),
    TieneEstado = (r as record) as logical => List.AnyTrue(List.Transform(ColsEstado, (c) => Txt(Record.FieldOrDefault(r, c, null)) <> "")),

    PrevK = Table.AddColumn(Prev, "__K", FnK, type text),
    PrevConEstado = Table.Buffer(Table.SelectRows(PrevK, TieneEstado)),
    PrevSolo = Table.Distinct(Table.SelectColumns(PrevConEstado, {"__K"} & ColsEstado), {"__K"}),

    NuevoBase = Table.RemoveColumns(Nuevo, List.Intersect({ColsEstado, Table.ColumnNames(Nuevo)})),
    NuevoK = Table.AddIndexColumn(Table.AddColumn(NuevoBase, "__K", FnK, type text), "__I", 0, 1),
    Unido = Table.NestedJoin(NuevoK, {"__K"}, PrevSolo, {"__K"}, "__P", JoinKind.LeftOuter),
    Expandido = Table.Sort(Table.ExpandTableColumn(Unido, "__P", ColsEstado, ColsEstado), {{"__I", Order.Ascending}}),

    ClavesNuevas = List.Buffer(List.Distinct(NuevoK[__K])),
    Huerfanas = if ConservarHuerfanas = false then Table.FirstN(PrevConEstado, 0)
        else Table.SelectRows(PrevConEstado, each not List.Contains(ClavesNuevas, [__K])),

    Salida = Table.SelectColumns(Table.Combine({Expandido, Huerfanas}), ColsSalida, MissingField.UseNull)
in
    Salida
