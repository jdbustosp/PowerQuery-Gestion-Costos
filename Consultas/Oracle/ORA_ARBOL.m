// ORA_ARBOL: estructura del presupuesto Oracle desde PRESUPUESTOS.xlsx
// (Unifier "Seguimiento por paquete de trabajo"). Sirve para cualquier numero de niveles:
// Reglas por fila (columna 1 = codigo con sangria, 2 = descripcion, 3 = UM, 4 = cantidad,
// 5 = V/U, 6 = total):
//   codigo sin "-"                         -> PAQUETE DE TRABAJO (= Subcapitulo)
//   CBS (>= 3 segmentos), ultimo en ceros  -> nodo del arbol (capitulos y niveles)
//   CBS, ultimo segmento distinto de cero  -> ACTIVIDAD
//   codigo de 2 segmentos (Z410300-0076)   -> INSUMO de la actividad anterior
// Niveles de una actividad = sus ancestros que existen en el arbol, poniendo en ceros sus
// segmentos finales (misma regla que PLANTILLA_PPTO_ORACLE de SINCO-SINCO). Nombres de nodo:
// PRESUPUESTOS y, si falta, el Cost Sheet (CONTROL.xlsx).
// Campos: Actividades (CBS, nombre, UM, Nivel 1..4, Capitulo), Insumos (filas de ppto),
//         Nombres (record codigo -> nombre), FnActividad, FnNiveles, FnIns.
let
    Fuentes = ORA_FUENTES,
    FnNum = Fuentes[FnNum],
    Limpiar = (v as any) as text =>
        let t = if v = null then "" else Text.From(v)
        in Text.Trim(Text.Replace(Text.Replace(Text.Replace(t, "#(lf)", " "), "#(cr)", " "), "#(00A0)", " ")),
    EsCero = (s as text) as logical => s <> "" and Text.Select(s, {"0"}) = s,
    Segs = (c as text) as list => Text.Split(c, "-"),
    EsCBS = (c as text) as logical => List.Count(Segs(c)) >= 3,

    Crudo = Fuentes[PRESUPUESTOS],
    Cols = Table.ColumnNames(Crudo),
    Filas = Table.FromColumns({
        List.Transform(Table.Column(Crudo, Cols{0}), Limpiar),
        List.Transform(Table.Column(Crudo, Cols{1}), Limpiar),
        List.Transform(Table.Column(Crudo, Cols{2}), Limpiar),
        List.Transform(Table.Column(Crudo, Cols{3}), FnNum),
        List.Transform(Table.Column(Crudo, Cols{4}), FnNum),
        List.Transform(Table.Column(Crudo, Cols{5}), FnNum)
    }, {"Cod", "Desc", "UM", "Cant", "VU", "Total"}),
    // desde la fila de encabezado "Cod CBS"
    PosEnc = List.PositionOf(Filas[Cod], "Cod CBS"),
    Datos = Table.Skip(Filas, if PosEnc < 0 then 0 else PosEnc + 1),
    Clasif = Table.AddColumn(Datos, "Clase", each
        let c = [Cod] in
        if c = "" or c = "Gran Total" or c = "Cant" then null
        else if not Text.Contains(c, "-") then "PAQUETE"
        else if EsCBS(c) then (if EsCero(List.Last(Segs(c))) then "NODO" else "ACTIVIDAD")
        else if List.Count(Segs(c)) = 2 then "INSUMO"
        else null, type text),
    ConPaquete = Table.FillDown(Table.AddColumn(Clasif, "Paquete", each if [Clase] = "PAQUETE" then [Desc] else null, type text), {"Paquete"}),
    ConAct = Table.FillDown(Table.AddColumn(ConPaquete, "CBS", each if [Clase] = "ACTIVIDAD" then [Cod] else null, type text), {"CBS"}),
    Utiles = Table.Buffer(Table.SelectRows(ConAct, each [Clase] <> null)),

    // ---------- nombres de nodos y actividades ----------
    NodosPpto = Table.SelectRows(Utiles, each [Clase] = "NODO" or [Clase] = "ACTIVIDAD"),
    Control = try Fuentes[CONTROL] otherwise null,
    NodosControl = if Control = null then #table({"Cod", "Desc"}, {})
        else let cc = Table.ColumnNames(Control) in
            Table.FromColumns({List.Transform(Table.Column(Control, cc{0}), Limpiar), List.Transform(Table.Column(Control, cc{1}), Limpiar)}, {"Cod", "Desc"}),
    TodosNodos = Table.Distinct(Table.SelectRows(
        Table.Combine({Table.SelectColumns(NodosPpto, {"Cod", "Desc"}), NodosControl}),
        each [Cod] <> "" and EsCBS([Cod]) and [Desc] <> ""), {"Cod"}),
    Nombres = Record.FromList(TodosNodos[Desc], TodosNodos[Cod]),
    FnNombre = (c as text) as nullable text => Record.FieldOrDefault(Nombres, c, null),

    // ancestros existentes, del mas superficial al mas profundo
    FnAncestros = (cod as text) as list =>
        let
            s = Segs(cod), n = List.Count(s),
            noCero = List.Select({1..n - 1}, (i) => not EsCero(s{i})),
            EnCeros = (k as number) as text =>
                Text.Combine(List.Transform({0..n - 1}, (i) => if i >= k then Text.Repeat("0", Text.Length(s{i})) else s{i}), "-"),
            cands = List.Transform(List.Sort(noCero), EnCeros)
        in List.Select(cands, (a) => a <> cod and FnNombre(a) <> null),
    FnNiveles = (cod as nullable text) as record =>
        let
            anc = if cod = null or cod = "" or not EsCBS(cod) then {} else FnAncestros(cod),
            txt = List.Transform(anc, (a) => a & " - " & FnNombre(a)),
            n = (i as number) => if List.Count(txt) > i then txt{i} else null
        in [#"Nivel 1" = n(0), #"Nivel 2" = n(1), #"Nivel 3" = n(2), #"Nivel 4" = n(3)],

    // ---------- actividades ----------
    ActPpto = Table.Buffer(Table.Group(Table.SelectRows(Utiles, each [Clase] = "ACTIVIDAD"), {"Cod"}, {
        {"Desc", each List.First(List.Select([Desc], (d) => d <> ""), null), type nullable text},
        {"UM", each List.First(List.Select([UM], (u) => u <> ""), null), type nullable text}})),
    // nombre de la actividad: "CBS - NOMBRE (UM)"; si el CBS no esta en el ppto, nombre del Cost Sheet
    FnTextoActividad = (c as text, nom as nullable text, u as nullable text) as text =>
        c & " - " & (if nom = null then "SIN NOMBRE" else nom) & (if u = null or u = "" then "" else " (" & u & ")"),
    Actividades = Table.Buffer(Table.FromRecords(List.Transform(Table.ToRecords(ActPpto), (r) =>
        Record.Combine({[CBS = r[Cod], Actividad = FnTextoActividad(r[Cod], r[Desc], r[UM])], FnNiveles(r[Cod])})))),
    DictActividad = Record.FromList(Actividades[Actividad], Actividades[CBS]),

    // ---------- insumos del presupuesto ----------
    FnIns = (art as nullable text, desc as nullable text, um as nullable text) as nullable text =>
        let a = if art = null then "" else Text.Trim(art), d = if desc = null then "" else Text.Trim(desc)
        in if a = "" and d = "" then null
           else (if a = "" then "" else a & " - ") & d & (if um = null or um = "" then "" else " (" & Text.Upper(um) & ")"),
    Insumos = Table.Buffer(Table.SelectColumns(
        Table.RenameColumns(Table.SelectRows(Utiles, each [Clase] = "INSUMO"), {{"Cod", "Articulo"}}),
        {"Paquete", "CBS", "Articulo", "Desc", "UM", "Cant", "VU", "Total"})),
    // nombre oficial del articulo (el del presupuesto) para que el insumo se llame igual en todas las fuentes
    ArtNombres = Table.Distinct(Table.SelectColumns(Insumos, {"Articulo", "Desc", "UM"}), {"Articulo"}),
    DictArt = Record.FromList(List.Transform(Table.ToRecords(ArtNombres), each [Desc = [Desc], UM = [UM]]), ArtNombres[Articulo]),
    FnInsOficial = (art as nullable text, desc as nullable text, um as nullable text) as nullable text =>
        let o = if art = null then null else Record.FieldOrDefault(DictArt, Text.Trim(art), null)
        in if o <> null then FnIns(art, o[Desc], o[UM]) else FnIns(art, desc, um),
    FnActividadDe = (cbs as nullable text) as nullable text =>
        let c = if cbs = null then "" else Text.Trim(cbs),
            v = if c = "" then null else Record.FieldOrDefault(DictActividad, c, null)
        in if c = "" then null else if v <> null then v else FnTextoActividad(c, FnNombre(c), null),

    // Codigo de actividad para BD: Oracle repite el MISMO CBS en varios paquetes (TORRE 1, TORRE 2...),
    // y BD agrupa por codigo (p. ej. "si la actividad ya tiene asegurado, su POR ADJUDICAR proyecta 0").
    // Como en SINCO cada subcapitulo tiene su propio codigo, aqui el codigo es "CBS | PAQUETE".
    FnCodAct = (cbs as nullable text, paquete as nullable text) as nullable text =>
        let c = if cbs = null then "" else Text.Trim(cbs), p = if paquete = null then "" else Text.Trim(paquete)
        in if c = "" then null else if p = "" then c else c & " | " & p,

    Resultado = [
        FnCodAct = FnCodAct,
        Actividades = Actividades,
        Insumos = Insumos,
        Nombres = Nombres,
        FnNiveles = FnNiveles,
        FnActividad = FnActividadDe,
        FnIns = FnInsOficial,
        Limpiar = Limpiar,
        FnNum = FnNum
    ]
in
    Resultado
