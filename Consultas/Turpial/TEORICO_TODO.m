// Teorico de concreto, acero y enchapes. Concreto y acero: bloques por edificacion en MEMORIAS (nombre MEMORIA_BLOQUES);
// un bloque se reconoce por los encabezados Familia + Nivel en la fila 5, asi que copiar sus columnas basta.
// Unidades y CC salen de ASIGNACION por el nombre del titulo (despues del punto medio). Sin registro: 1 unidad y
// la tipologia lleva (SIN REGISTRAR), para que se note. Insumos: CAT_TECNICO. M local del libro.
let
    Txt = (v as any) as text => let t = try Text.Trim(Text.From(v)) otherwise null in if t = null then "" else t,
    Num = (v as any) as number => let n = try Number.From(v) otherwise null in if n = null then 0 else n,
    Norm = (v as any) as text => List.Accumulate({{"Á","A"},{"É","E"},{"Í","I"},{"Ó","O"},{"Ú","U"},{"Ñ","N"}}, Text.Upper(Txt(v)), (s, p) => Text.Replace(s, p{0}, p{1})),
    Libro = Excel.CurrentWorkbook(),
    Cat = List.Buffer(Table.ToRecords(Libro{[Name = "CAT_TECNICO"]}[Content])),
    Asig = List.Buffer(List.Select(Table.ToRecords(Libro{[Name = "ASIGNACION"]}[Content]), each Txt([Estructura]) <> "")),
    CCBase = if List.IsEmpty(Asig) then "" else Txt(Asig{0}[Centro de Costos]),
    Edif = (nombre as text) as record =>
        let f = List.Select(Asig, each Norm([Estructura]) = Norm(nombre))
        in if List.IsEmpty(f) then [Tipologia = nombre & " (SIN REGISTRAR)", NumUnidades = 1, CC = CCBase]
           else [Tipologia = Txt(f{0}[Estructura]), NumUnidades = Num(f{0}[No Unidades]), CC = Txt(f{0}[Centro de Costos])],
    SinIns = (x as text) as text => "SIN INSUMO EN CATALOGO - " & x,
    InsCat = (fam as text, tipo as text, par as text) as text =>
        let f = List.Select(Cat, each Txt([Familia]) = fam and Txt([Categoria]) = "Referencia"
                and Text.StartsWith(Norm([Tipo]), Norm(tipo)) and Norm([Parametro]) = Norm(par))
        in if List.IsEmpty(f) then "" else Txt(f{0}[Ins]),

    Filas = List.Buffer(Table.ToRows(Libro{[Name = "MEMORIA_BLOQUES"]}[Content])),
    Titulos = Filas{3},
    Encab = Filas{4},
    Datos = List.Buffer(List.Select(List.Skip(Filas, 5), each List.NonNullCount(_) > 0)),
    N = List.Count(Encab),
    Inicios = List.Buffer(List.Select(List.Numbers(0, N - 1), each Txt(Encab{_}) = "Familia" and Txt(Encab{_ + 1}) = "Nivel")),
    Bloque = (i as number) as list =>
        let
            sig = List.Select(Inicios, each _ > i),
            fin = if List.IsEmpty(sig) then N else sig{0},
            enc = List.Transform(List.Range(Encab, i, fin - i), Txt),
            ColPre = (p as text) as number => let f = List.Select(List.Positions(enc), each Text.StartsWith(enc{_}, p)) in if List.IsEmpty(f) then -1 else i + f{0},
            esAcero = Text.StartsWith(enc{3}, "Pos"),
            t = Txt(Titulos{i}),
            nombre = if Text.Contains(t, "·") then Text.Trim(Text.AfterDelimiter(t, "·")) else t,
            e = Edif(nombre),
            base = [Tipologia = e[Tipologia], NumUnidades = e[NumUnidades], #"Centro de Costos" = e[CC]],
            cCant = ColPre(if esAcero then "Peso total" else "Volumen total"),
            cTipo = ColPre("Tipo"),
            cRef = ColPre("Referencia"),
            FilaC = (f as list) as list =>
                let fam = Txt(f{i}), fc = Txt(f{i + 3}), q = Num(f{cCant})
                in if q = 0 or fc = "" then {}
                   else {base & [Material = "CONCRETO", TipoElemento = fam, Piso = Txt(f{i + 1}),
                                 Ins = if f{i + 3} is number then SinIns(fam & " " & fc) else fc, CantUnidad = q]},
            FilaA = (f as list) as list =>
                let v = f{cCant}
                in if not (v is number) or v = 0 then {}
                   else let r = Text.Replace(Txt(f{cRef}), "#", ""), ins = InsCat("Acero", Txt(f{cTipo}), r)
                        in {base & [Material = "ACERO", TipoElemento = Txt(f{i}), Piso = Txt(f{i + 1}),
                                    Ins = if ins = "" then SinIns(Txt(f{cTipo}) & " " & r) else ins, CantUnidad = v]}
        in
            if cCant < 0 then {} else List.Combine(List.Transform(Datos, if esAcero then FilaA else FilaC)),
    TodasBloques = List.Combine(List.Transform(Inicios, Bloque)),

    // Enchapes: tabla MEM_ENCHAPES; el Ins sale de la referencia; pegante y boquilla solo si tienen consumo en el catalogo.
    Ench = Table.ToRecords(Libro{[Name = "MEM_ENCHAPES"]}[Content]),
    Consumos = List.Select(Cat, each Txt([Familia]) = "Enchape" and Txt([Categoria]) = "Consumo" and Txt([Ins]) <> "" and Num([Valor]) <> 0),
    FilaE = (r as record) as list =>
        let
            nom = Txt(Record.FieldOrDefault(r, "Edificación", null)),
            e = Edif(nom),
            base = [Material = "ENCHAPES", Tipologia = e[Tipologia], NumUnidades = e[NumUnidades], #"Centro de Costos" = e[CC], Piso = Txt(Record.FieldOrDefault(r, "Nivel", null))],
            qE = Num(Record.FieldOrDefault(r, "Enchape total (m²)", null)),
            qM = Num(Record.FieldOrDefault(r, "Mortero total (m³)", null)),
            rE = Txt(Record.FieldOrDefault(r, "Ref. enchape", null)),
            rM = Txt(Record.FieldOrDefault(r, "Ref. mortero", null)),
            iE = InsCat("Enchape", "Enchape", rE),
            iM = InsCat("Enchape", "Mortero piso", rM)
        in
            if nom = "" then {} else
            (if qE = 0 then {} else {base & [TipoElemento = "Enchape", Ins = if iE = "" then SinIns(rE) else iE, CantUnidad = qE]})
            & (if qE = 0 then {} else List.Transform(Consumos, each base & [TipoElemento = Txt([Parametro]), Ins = Txt([Ins]), CantUnidad = qE * Num([Valor])]))
            & (if qM = 0 then {} else {base & [TipoElemento = "Mortero de piso", Ins = if iM = "" then SinIns(rM) else iM, CantUnidad = qM]}),
    TodasEnchapes = List.Combine(List.Transform(Ench, FilaE)),

    Tipo = type table [Material = text, Tipologia = text, NumUnidades = number, #"Centro de Costos" = text, TipoElemento = text, Piso = text, Ins = text, CantUnidad = number],
    Claves = {"Material", "Tipologia", "NumUnidades", "Centro de Costos", "TipoElemento", "Piso", "Ins"},
    Agrupar = (l as list) as table => Table.Group(Table.FromRecords(l, Tipo), Claves, {{"CantUnidad", each List.Sum([CantUnidad]), type number}}),
    Concreto = Table.FromRecords(List.Select(TodasBloques, each [Material] = "CONCRETO"), Tipo),
    Acero = Agrupar(List.Select(TodasBloques, each [Material] = "ACERO")),
    Enchapes = Agrupar(TodasEnchapes),
    Tabla = Table.Combine({Concreto, Acero, Enchapes}),
    ConTotal = Table.AddColumn(Tabla, "CantTotal", each [CantUnidad] * [NumUnidades], type number)
in
    ConTotal
