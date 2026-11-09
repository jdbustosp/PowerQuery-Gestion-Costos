// Teorico de mamposteria: de MEM_MAMPOSTERIA (formato de Juan, por tramo) a una fila por insumo.
// Unidades e hilada parada -> Ins del material; morteros (m3) x consumo de CAT_TECNICO segun Tipo pega / Tipo grout.
// Grafil y Hierro son adicionales a la memoria de acero. Unidades y CC: ASIGNACION. M local del libro.
let
    Txt = (v as any) as text => let t = try Text.Trim(Text.From(v)) otherwise null in if t = null then "" else t,
    Num = (v as any) as number => let n = try Number.From(v) otherwise null in if n = null then 0 else n,
    Norm = (v as any) as text => Text.Upper(Txt(v)),
    Libro = Excel.CurrentWorkbook(),
    Mem = Libro{[Name = "MEM_MAMPOSTERIA"]}[Content],
    CatTodo = List.Buffer(List.Select(Table.ToRecords(Libro{[Name = "CAT_TECNICO"]}[Content]), each Txt([Familia]) = "Mamposteria")),
    Asig = List.Buffer(List.Select(Table.ToRecords(Libro{[Name = "ASIGNACION"]}[Content]), each Txt([Estructura]) <> "")),
    CCBase = if List.IsEmpty(Asig) then "" else Txt(Asig{0}[Centro de Costos]),
    // Una estructura puede estar en varios CC (p. ej. TORRE TIPO B en tres etapas): una fila de ASIGNACION por CC.
    Edif = (nombre as text) as list =>
        let f = List.Select(Asig, each Norm([Estructura]) = Norm(nombre))
        in if List.IsEmpty(f) then {[Tipologia = nombre & " (SIN REGISTRAR)", NumUnidades = 1, CC = CCBase]}
           else List.Transform(f, each [Tipologia = Txt([Estructura]), NumUnidades = Num([No Unidades]), CC = Txt([Centro de Costos])]),
    // Repite las filas de una edificacion por cada CC asignado, con sus unidades.
    Repartir = (nombre as text, l as list) as list =>
        List.Combine(List.Transform(Edif(nombre), (e) =>
            List.Transform(l, each _ & [Tipologia = e[Tipologia], NumUnidades = e[NumUnidades], #"Centro de Costos" = e[CC]]))),
    BaseVacia = [Tipologia = "", NumUnidades = 0, #"Centro de Costos" = ""],
    General = (p as text) as number => let f = List.Select(CatTodo, each Txt([Categoria]) = "General" and Txt([Parametro]) = p) in if List.IsEmpty(f) then 0 else Num(f{0}[Valor]),
    SecViga = General("Seccion viga canchada"),
    SecCinta = General("Seccion cinta"),
    Consumos = List.Select(CatTodo, each Txt([Categoria]) = "Consumo" and Txt([Ins]) <> "" and Num([Valor]) <> 0),
    InsDe = (cat as text, tipo as text) as text =>
        let f = List.Select(CatTodo, each Txt([Categoria]) = cat and Txt([Tipo]) = tipo and Txt([Parametro]) = "Insumo") in if List.IsEmpty(f) then "" else Txt(f{0}[Ins]),
    // {TipoElemento, columna, factor, columna de eleccion de mortero, Tipo de consumo}
    Conceptos = {
        {"Unidades", "Piezas est.", 1, null, "M"},
        {"Hilada parada", "Hilada parada (und)", 1, null, "M"},
        {"Pega", "Pega (m³)", 1, "Tipo pega", "Pega"},
        {"Pañete", "Mortero pañete (m³)", 1, "Tipo pega", "Pañete"},
        {"Grout dovelas", "Grouting (m³)", 1, "Tipo grout", "Grout"},
        {"Viga canchada", "Vigas (ml)", SecViga, "Tipo pega", "Viga canchada"},
        {"Cinta", "Cinta ventana (ml)", SecCinta, "Tipo pega", "Cinta"},
        {"Grafil", "Grafil (kg)", 1, null, "Grafil"},
        {"Hierro", "Hierro (kg)", 1, null, "Hierro"}},
    Fila = (r as record) as list =>
        let
            nom = Txt(Record.FieldOrDefault(r, "Edificación", null)),
            tramo = Text.Upper(Txt(Record.FieldOrDefault(r, "Tramo", null))),
            clase = if Text.Contains(tramo, "DIVISORIO") or Text.Contains(tramo, "DINTEL") then "Interno" else "Fachada",
            mat = Txt(Record.FieldOrDefault(r, "Tipo / material", null))
        in
            if nom = "" then {} else Repartir(nom, List.Combine(List.Transform(Conceptos, (c) =>
                let
                    q = Num(Record.FieldOrDefault(r, c{1}, null)) * c{2},
                    base = BaseVacia & [Material = "MAMPOSTERIA", TipoElemento = c{0}, Piso = clase]
                in
                    if q = 0 then {}
                    else if c{3} = null then
                        let ins = if c{4} = "M" then InsDe("Material", mat) else InsDe("Refuerzo", c{4})
                        in if ins = "" then {} else {base & [Ins = ins, CantUnidad = q]}
                    else
                        let mortero = Txt(Record.FieldOrDefault(r, c{3}, null)),
                            usa = List.Select(Consumos, each Txt([Tipo]) = c{4} and Txt([Parametro]) = mortero)
                        in List.Transform(usa, each base & [Ins = Txt([Ins]), CantUnidad = q * Num([Valor])])))),
    Registros = List.Combine(List.Transform(Table.ToRecords(Mem), Fila)),
    Tabla = Table.FromRecords(Registros, type table [Material = text, Tipologia = text, NumUnidades = number,
        #"Centro de Costos" = text, TipoElemento = text, Piso = text, Ins = text, CantUnidad = number]),
    Agrupado = Table.Group(Tabla, {"Material", "Tipologia", "NumUnidades", "Centro de Costos", "TipoElemento", "Piso", "Ins"},
        {{"CantUnidad", each List.Sum([CantUnidad]), type number}}),
    ConTotal = Table.AddColumn(Agrupado, "CantTotal", each [CantUnidad] * [NumUnidades], type number)
in
    ConTotal