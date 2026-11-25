
let
    Norm = (v as any) as text => let t = try Text.Upper(Text.Trim(F_Globales[FnRemoveAccentMarks](Text.From(v)))) otherwise null
        in if t = null then "" else Text.Combine(List.Select(Text.Split(t, " "), each _ <> ""), " "),
    BDGuardada = Excel.CurrentWorkbook(){[Name="BD"]}[Content],
    ClasesGuardadas = Table.SelectColumns(Table.SelectRows(BDGuardada, each [Tipo] = "ITEMS"), {"Centro de Costos", "Codigo ins", "Ins", "Clasificador"}),
    InsumosActuales = Table.Distinct(Table.TransformColumns(Table.SelectColumns(ITEMSINSUMOS, {"Centro de Costos", "Codigo ins", "Ins"}), {
        {"Centro de Costos", each if _ = null then "" else Text.Upper(Text.Trim(Text.From(_))), type text},
        {"Ins", each if _ = null then "" else Text.Upper(Text.Trim(F_Globales[FnRemoveAccentMarks](_))), type text}})),
    Fuente = Table.ExpandTableColumn(Table.NestedJoin(InsumosActuales, {"Centro de Costos", "Codigo ins", "Ins"},
        ClasesGuardadas, {"Centro de Costos", "Codigo ins", "Ins"}, "C", JoinKind.LeftOuter), "C", {"Clasificador"}),
    Reales = Table.SelectRows(Fuente, each Norm([Centro de Costos]) <> "" and Norm([Ins]) <> ""),
    Catalogo = Table.Group(Reales, {"Centro de Costos", "Codigo ins", "Ins"}, {
        {"__ClasesBD", each List.Sort(List.Distinct(List.Select(List.Transform([Clasificador], (v) => if v = null then "" else Text.Trim(Text.From(v))), (v) => v <> "" and v <> "zOtros"))), type list}}),
    // Reglas manuales: tabla MaestroInsumos si existe; si no, lo que ya tiene escrito este catalogo en "Clasificador manual"
    // (se conserva fila por fila: Centro de Costos + Codigo ins + Ins; un insumo nuevo hereda la regla de su Ins).
    TieneMaestro = (try Excel.CurrentWorkbook(){[Name="MaestroInsumos"]}[Content] otherwise null) <> null,
    Guardado = if TieneMaestro then #table({"Centro de Costos", "Codigo ins", "Ins", "Clasificador manual"}, {})
        else try Table.SelectColumns(Excel.CurrentWorkbook(){[Name="CATALOGO_INSUMOS_CC"]}[Content], {"Centro de Costos", "Codigo ins", "Ins", "Clasificador manual"})
        otherwise #table({"Centro de Costos", "Codigo ins", "Ins", "Clasificador manual"}, {}),
    Txt0 = (v as any) as text => if v = null then "" else Text.Trim(Text.From(v)),
    ClaveFila = (r as record) as text => Norm(r[Centro de Costos]) & "|" & Norm(r[Codigo ins]) & "|" & Norm(r[Ins]),
    GuardadoValido = List.Select(Table.ToRecords(Guardado), each Norm([Clasificador manual]) <> ""),
    PorFila = Record.FromList(List.Transform(GuardadoValido, each Text.Trim(Text.From([Clasificador manual]))),
        List.Transform(GuardadoValido, ClaveFila)),
    PorFilaSegura = try PorFila otherwise [],
    Manual = if TieneMaestro then Excel.CurrentWorkbook(){[Name="MaestroInsumos"]}[Content]
        else Table.RenameColumns(Table.FromRecords(GuardadoValido, {"Centro de Costos", "Codigo ins", "Ins", "Clasificador manual"}, MissingField.UseNull), {{"Clasificador manual", "Clasificador"}}),
    Reglas = Table.Group(Table.AddColumn(Table.SelectRows(Manual, each Norm([Ins]) <> "" and Norm([Clasificador]) <> ""), "__Ins", each Norm([Ins]), type text), {"__Ins"},
        {{"__Clases", each List.Distinct(List.Transform([Clasificador], Text.Trim)), type list}}),
    Validas = Table.SelectRows(Reglas, each List.Count([__Clases]) = 1),
    Mapa = Record.FromList(List.Transform(Validas[__Clases], each _{0}), Validas[__Ins]),
    Clasificado = Table.AddColumn(Catalogo, "Clasificador manual", each
        let f = Record.FieldOrDefault(PorFilaSegura, ClaveFila(_), null)
        in if f <> null then f else Record.FieldOrDefault(Mapa, Norm([Ins]), null), type nullable text),
    ClaseBD = Table.AddColumn(Clasificado, "Clasificador", each Text.Combine([__ClasesBD], " | "), type text),
    Estado = Table.AddColumn(ClaseBD, "Estado", each if List.Count([__ClasesBD]) > 1 then "REVISAR CLASIFICADORES"
        else if List.IsEmpty([__ClasesBD]) then "PENDIENTE DE CLASIFICAR" else "CLASIFICADO", type text),
    // Reglas manuales de insumos que no estan en el presupuesto (antes en MaestroInsumos): se conservan como filas
    // "REGLA MANUAL" mientras tengan "Clasificador manual". Borrar el clasificador manual = quitar la regla.
    ClavesActuales = List.Buffer(List.Transform(Table.ToRecords(Catalogo), ClaveFila)),
    InsActuales = List.Buffer(List.Distinct(List.Transform(Catalogo[Ins], Norm))),
    Fuera = (r as record) as logical => not List.Contains(InsActuales, Norm(r[Ins])),
    // Con MaestroInsumos: sus reglas validas cuyo Ins no esta en el presupuesto. Sin el: las filas guardadas que ya no casan.
    ExtraMaestro = if not TieneMaestro then {} else
        List.Select(Table.ToRecords(Table.Distinct(Table.SelectColumns(Table.SelectRows(Table.AddColumn(Manual, "__Ins", each Norm([Ins])),
            each Norm([Ins]) <> "" and Norm([Clasificador]) <> "" and Record.HasFields(Mapa, [__Ins])), {"Ins", "__Ins"}), {"__Ins"})),
            Fuera),
    ExtraLista = if TieneMaestro then List.Transform(ExtraMaestro, each [Ins = [Ins], #"Clasificador manual" = Record.Field(Mapa, [__Ins])])
        else List.Distinct(List.Select(GuardadoValido, each not List.Contains(ClavesActuales, ClaveFila(_)) and Fuera(_)), each Norm([Ins])),
    ExtraTabla = Table.FromRecords(List.Transform(ExtraLista, each [#"Centro de Costos" = "(TODOS)", #"Codigo ins" = "",
        Ins = Txt0([Ins]), Clasificador = "", #"Clasificador manual" = Text.Trim(Text.From([Clasificador manual])), Estado = "REGLA MANUAL"]),
        type table [#"Centro de Costos" = text, #"Codigo ins" = text, Ins = text, Clasificador = text, #"Clasificador manual" = nullable text, Estado = text]),
    ConExtra = Table.Combine({Table.SelectColumns(Estado, {"Centro de Costos", "Codigo ins", "Ins", "Clasificador", "Clasificador manual", "Estado"}), ExtraTabla}),
    Final = Table.Sort(ConExtra,
        {{"Centro de Costos", Order.Ascending}, {"Ins", Order.Ascending}, {"Codigo ins", Order.Ascending}})
in
    Final
