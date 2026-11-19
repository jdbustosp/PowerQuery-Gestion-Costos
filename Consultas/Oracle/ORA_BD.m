// Agrega a BD_BASE el presupuesto V2 por insumo y los niveles de la estructura CBS.
// 2026-11-18 (reglas del usuario, iguales para todos los libros Oracle):
//  - Ins con unidad en todos los libros (ORA_CONFIG MostrarUnidadInsumo = false la apaga) y SIN el codigo
//    del articulo; el codigo queda en Codigo ins (si venia vacio se toma del prefijo del nombre).
//  - Consumido por insumo = recepciones (VT Cortes) de las lineas de OC y contrato. La fila del Cost Sheet
//    queda como (PROYECTADO COST SHEET) con su proyectado, y (AJUSTE COST SHEET) lleva la diferencia
//    Cost Sheet - recepciones de cada actividad, para que el consumido total siga igual al Cost Sheet.
//  - Niveles Nivel 1..N, N = nivel mas profundo del arbol (ORA_ARBOL[NombresNivel]).
let
    Arbol = ORA_ARBOL,
    // Forzar DISPONIBLE antes de BD_BASE: su try/otherwise central no debe ocultar una descarga ilegible.
    Base = if Table.RowCount(Table.Buffer(DISPONIBLE)) >= 0 then BD_BASE
        else error "No se pudo evaluar el presupuesto adjudicado y por adjudicar",
    // V2 es una medida independiente por CBS. La base y sus medidas iniciales se conservan.
    BaseConV2 = Table.AddColumn(Base, "VT Presupuesto v2", each null, type nullable number),
    FilasV2 = Table.SelectColumns(ORA_PPTO_V2, Table.ColumnNames(BaseConV2), MissingField.UseNull),
    Conjunto = Table.Combine({BaseConV2, FilasV2}),
    Columnas = Table.ColumnNames(Conjunto),
    Txt = (v as any) as text => if v = null then "" else (try Text.Trim(Text.From(v)) otherwise ""),
    Num = (v as any) as number => let n = try Number.From(v) otherwise null in if n = null then 0 else n,
    // Reemplaza la columna col por la columna calculada nueva, en la misma posicion.
    Reemplazar = (t as table, col as text, nueva as text) as table =>
        Table.ReorderColumns(Table.RenameColumns(Table.RemoveColumns(t, {col}), {{nueva, col}}), List.RemoveItems(Table.ColumnNames(t), {nueva})),

    // ---------- codigo del articulo ----------
    // "Z410300-0076": dos segmentos, el segundo numerico, sin espacios.
    EsCodigoArt = (c as text) as logical =>
        let s = Text.Split(c, "-")
        in List.Count(s) = 2 and Text.Length(s{0}) > 0 and Text.Length(s{1}) > 0 and not Text.Contains(s{0}, " ")
            and Text.Select(s{1}, {"0".."9"}) = s{1},
    FnPrefijo = (nombre as text) as text =>
        let p = if Text.Contains(nombre, " - ") then Text.Trim(Text.BeforeDelimiter(nombre, " - ")) else ""
        in if EsCodigoArt(Text.Upper(p)) then p else "",
    TablaCodigo = Reemplazar(Table.AddColumn(Conjunto, "__Cod", each
        if Txt([Codigo ins]) <> "" then [Codigo ins]
        else let p = FnPrefijo(Txt([Ins])) in if p = "" then [Codigo ins] else p, type nullable text), "Codigo ins", "__Cod"),

    // ---------- consumido por insumo ----------
    EsLinea = (r as record) as logical => (r[Tipo] = "COMPRAS" or r[Tipo] = "CONTRATO") and Txt(r[Actividad]) <> "",
    ClaveAct = (r as record) as text => Txt(r[Centro de Costos]) & "|" & Txt(r[Actividad]),
    Lineas = Table.SelectRows(TablaCodigo, each EsLinea(_) and Num([VT Cortes]) <> 0),
    CortesAct = Table.Group(Table.AddColumn(Lineas, "__K", ClaveAct, type text), {"__K"}, {
        {"__Cortes", each List.Sum(List.Transform([VT Cortes], Num)), type number},
        {"__Fila", each Table.First(_), type record}}),
    FilasCS = Table.SelectRows(TablaCodigo, each [Tipo] = "ITEMS"),
    CSAct = Table.Group(Table.AddColumn(FilasCS, "__K", ClaveAct, type text), {"__K"}, {
        {"__CS", each List.Sum(List.Transform([VT Consumido], Num)), type number},
        {"__Fila", each Table.First(_), type record}}),
    MapaCortes = Record.FromList(CortesAct[__Cortes], CortesAct[__K]),
    MapaCS = Record.FromList(CSAct[__CS], CSAct[__K]),
    MapaFila = Record.Combine({
        Record.FromList(CortesAct[__Fila], CortesAct[__K]),
        Record.FromList(CSAct[__Fila], CSAct[__K])}),
    ClavesAct = List.Distinct(CortesAct[__K] & CSAct[__K]),
    Ajustes = Table.FromRecords(List.RemoveNulls(List.Transform(ClavesAct, (k) =>
        let
            dif = Num(Record.FieldOrDefault(MapaCS, k, 0)) - Num(Record.FieldOrDefault(MapaCortes, k, 0)),
            f = Record.Field(MapaFila, k)
        in if Number.Abs(dif) < 1 then null else [
            #"Centro de Costos" = f[Centro de Costos], Ins = "(AJUSTE COST SHEET)", Actividad = f[Actividad],
            Capitulo = f[Capitulo], Subcapitulo = null, Tipo = "ITEMS",
            #"# CC - Comparativo" = "SIN CUADRO COMPARATIVO", #"VT Consumido" = dif, Clasificador = "zOtros",
            #"Tipo OC" = "(no aplica)", #"No_Prov seg" = "(otros registros)", #"Fila real" = 0,
            #"Clasif desperdicio" = "(no aplica)", #"Contratista almacen" = "(no aplica)"]))),
    AjustesCompletos = Table.SelectColumns(Ajustes, Columnas, MissingField.UseNull),
    ConsLinea = Reemplazar(Table.AddColumn(TablaCodigo, "__VC", each
        if EsLinea(_) then (if Num([VT Cortes]) = 0 then null else [VT Cortes])
        else if [Tipo] = "ITEMS" then null
        else [VT Consumido], type nullable number), "VT Consumido", "__VC"),
    ConsNombre = Reemplazar(Table.AddColumn(ConsLinea, "__I", each
        if [Tipo] = "ITEMS" and Txt([Ins]) = "(COST SHEET ORACLE)" then "(PROYECTADO COST SHEET)" else [Ins],
        type nullable text), "Ins", "__I"),
    ConAjustes = Table.Combine({ConsNombre, AjustesCompletos}),

    // ---------- unidad del insumo ----------
    FormatearInsumos = (tabla as table) as table =>
        let
            Norm = (v as any) as text => if v = null then "" else Text.Upper(Text.Trim(Text.From(v))),
            EsUnidad = (v as text) as logical => v <> "" and not List.Contains({"-", "N/A", "S/U", "SIN UM"}, v),
            DeAsegurado = Table.TransformColumns(
                Table.SelectColumns(ORA_ASEGURADO, {"Articulo", "Registro", "UM"}),
                {{"Articulo", Norm, type text}, {"Registro", Norm, type text}, {"UM", Norm, type text}}),
            DePpto = Table.TransformColumns(
                Table.SelectColumns(Arbol[Insumos], {"Articulo", "UM"}),
                {{"Articulo", Norm, type text}, {"UM", Norm, type text}}),
            Validar = (t as table) as table => Table.SelectRows(t, each [Articulo] <> "" and EsUnidad([UM])),
            PorRegistro = Table.AddColumn(Validar(DeAsegurado), "Clave", each [Registro] & "|" & [Articulo], type text),
            MapaUnico = (t as table, clave as text) as record =>
                let
                    g = Table.Group(t, {clave}, {{"Unidades", each List.Distinct([UM]), type list}}),
                    u = Table.SelectRows(g, each List.Count([Unidades]) = 1)
                in Record.FromList(List.Transform(u[Unidades], each _{0}), Table.Column(u, clave)),
            MapaRegistro = MapaUnico(PorRegistro, "Clave"),
            MapaArticulo = MapaUnico(Table.Combine({Validar(DeAsegurado), Validar(DePpto)}), "Articulo"),
            Complementos = Record.FieldOrDefault(ORA_CONFIG, "UnidadesComplementarias", []),
            Unidades = List.Buffer(List.Distinct(List.Combine({DeAsegurado[UM], DePpto[UM], Record.FieldValues(Complementos), {"SIN UM"}}))),
            ConNombre = Table.AddColumn(tabla, "__InsUM", each
                let
                    nombre = if [Ins] = null then "" else Text.Trim(Text.From([Ins])),
                    codigo = Norm([Codigo ins]),
                    sufijo = if Text.EndsWith(nombre, ")") then
                        (try Norm(Text.BeforeDelimiter(Text.AfterDelimiter(nombre, "(", {0, RelativePosition.FromEnd}), ")")) otherwise "")
                        else "",
                    unidad = Record.FieldOrDefault(MapaRegistro, Norm([#"# OC / Contrato"]) & "|" & codigo, null)
                        ?? Record.FieldOrDefault(MapaArticulo, codigo, null)
                        ?? Record.FieldOrDefault(Complementos, codigo, "SIN UM")
                in if nombre = "" or not EsCodigoArt(codigo) or (sufijo <> "" and List.Contains(Unidades, sufijo)) then [Ins]
                    else nombre & " (" & unidad & ")", type nullable text)
        in Reemplazar(ConNombre, "Ins", "__InsUM"),
    Etiquetas = if Record.FieldOrDefault(ORA_CONFIG, "MostrarUnidadInsumo", true) then FormatearInsumos(ConAjustes) else ConAjustes,

    // ---------- Ins sin codigo ----------
    SinCodigo = Table.TransformColumns(Etiquetas, {{"Ins", (v) =>
        let t = Txt(v), p = FnPrefijo(t)
        in if p = "" then v else Text.Trim(Text.AfterDelimiter(t, " - ")), type nullable text}}),

    // ---------- niveles ----------
    FnCod = (a as any) as nullable text =>
        let t = try Text.Trim(Text.From(a)) otherwise null
        in if t = null or not Text.Contains(t, " - ") then null else Text.BeforeDelimiter(t, " - "),
    ConNiveles = Table.AddColumn(SinCodigo, "__N", each Arbol[FnNiveles](FnCod([Actividad]))),
    Resultado = Table.ExpandRecordColumn(ConNiveles, "__N", Arbol[NombresNivel])
in
    Table.Buffer(Resultado)
