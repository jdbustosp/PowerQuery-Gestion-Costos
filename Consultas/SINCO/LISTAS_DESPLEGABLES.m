// Fuentes de los desplegables (validaciones de MEMORIAS, CAT_TECNICO y ASIGNACION) en una sola tabla Lista | Valor,
// ordenada por Lista para que cada lista quede contigua. Los nombres del libro (LISTA_*) son DESREF sobre esta tabla:
//   DESREF(<primera celda Valor>; COINCIDIR("<lista>"; <col Lista>; 0) - 1; 0; MAX(1; CONTAR.SI(<col Lista>; "<lista>")); 1)
// Reemplaza el bloque de formulas FILTRAR/UNICOS (que no se pueden usar directo en una validacion).
// Se actualiza al refrescar (no en vivo).
let
    Libro = Excel.CurrentWorkbook(),
    Tabla = (n as text) as table => try Libro{[Name = n]}[Content] otherwise #table({}, {}),
    Txt = (v as any) as text => if v = null then "" else Text.Trim(Text.From(v)),
    Norm = (v as any) as text => Text.Replace(Text.Replace(Text.Upper(Txt(v)), Character.FromNumber(209), "N"), Character.FromNumber(241), "N"),
    Col = (t as table, c as text) as list => if Table.HasColumns(t, c) then Table.Column(t, c) else List.Repeat({null}, Table.RowCount(t)),
    Filas = (lista as text, valores as list) as list =>
        List.Transform(List.Distinct(List.Select(List.Transform(valores, Txt), each _ <> "")), each {lista, _}),
    // CAT_TECNICO
    Cat = Table.Buffer(Tabla("CAT_TECNICO")),
    CatRec = Table.ToRecords(Table.SelectColumns(Cat, {"Familia", "Categoria", "Tipo", "Parametro"}, MissingField.UseNull)),
    CatSel = (f as text, c as text, t as nullable text, campo as text) as list =>
        List.Transform(List.Select(CatRec, each Norm([Familia]) = Text.Upper(f) and Norm([Categoria]) = Text.Upper(c)
            and (t = null or Norm([Tipo]) = Norm(t))), each Record.Field(_, campo)),
    AceroRef = List.Select(CatRec, each Norm([Familia]) = "ACERO" and Norm([Categoria]) = "REFERENCIA" and Txt([Tipo]) <> ""),
    RefAcero = List.Combine(List.Transform(List.Distinct(List.Transform(AceroRef, each Txt([Tipo]))), (tp) =>
        Filas("REF ACERO|" & tp, List.Transform(List.Select(AceroRef, each Txt([Tipo]) = tp),
            each if Text.StartsWith(Text.Upper(tp), "BARRA") then "#" & Txt([Parametro]) else Txt([Parametro]))))),
    // CATALOGO_INSUMOS_CC: clasificador efectivo = manual si lo hay, si no el de BD
    Catalogo = Table.Buffer(Tabla("CATALOGO_INSUMOS_CC")),
    CatIns = List.Transform(Table.ToRecords(Table.SelectColumns(Catalogo, {"Ins", "Clasificador", "Clasificador manual"}, MissingField.UseNull)),
        each [Ins = Txt([Ins]), Clase = if Txt([Clasificador manual]) <> "" then Txt([Clasificador manual]) else Txt([Clasificador])]),
    InsDe = (clase as text) as list => List.Sort(List.Distinct(List.Transform(List.Select(CatIns, each [Clase] = clase), each [Ins]))),
    InsConcretoMem = List.Sort(List.Distinct(List.Select(List.Transform(CatIns, each Text.Upper([Ins])), each
        (Text.StartsWith(_, "CONCRETO ") or Text.StartsWith(_, "M_CONCRETO ")) and Text.EndsWith(_, "(M3)")))),
    // ASIGNACION: estructuras/edificaciones
    Estructuras = List.Sort(List.Distinct(List.Select(List.Transform(Col(Tabla("ASIGNACION"), "Estructura"), Txt), each _ <> ""))),
    Todo = List.Combine({
        Filas("CONCRETO MEMORIAS", InsConcretoMem),
        Filas("ESTRUCTURAS", Estructuras),
        Filas("MAT MAMP", CatSel("Mamposteria", "Material", null, "Tipo")),
        Filas("MORTERO PEGA", CatSel("Mamposteria", "Consumo", "Pega", "Parametro")),
        Filas("MORTERO PANETE", CatSel("Mamposteria", "Consumo", "Panete", "Parametro")),
        Filas("MORTERO GROUT", CatSel("Mamposteria", "Consumo", "Grout", "Parametro")),
        Filas("REF ENCHAPE", CatSel("Enchape", "Referencia", "Enchape", "Parametro")),
        Filas("REF MORTERO PISO", CatSel("Enchape", "Referencia", "Mortero piso", "Parametro")),
        Filas("TIPO ACERO", List.Sort(List.Transform(AceroRef, each Txt([Tipo])))),
        RefAcero,
        Filas("INS CONCRETO", InsDe("Concreto")),
        Filas("INS ACERO", InsDe("Acero")),
        Filas("INS MORTERO", InsDe("Morteros")),
        Filas("INS BLOQUES", InsDe("Bloques y ladrillos")),
        Filas("INS ENCHAPE", InsDe("Enchapes, pegante y boquilla (SUM)"))}),
    ConOrden = List.Transform(List.Positions(Todo), each Todo{_} & {_}),
    Ordenada = Table.Sort(#table(type table [Lista = text, Valor = text, __i = number], ConOrden), {{"Lista", Order.Ascending}, {"__i", Order.Ascending}}),
    Final = Table.RemoveColumns(Ordenada, {"__i"})
in
    Final
