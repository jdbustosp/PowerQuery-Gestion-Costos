// ORA_COMPRAS: ordenes de compra Oracle con el MISMO esquema que COMPRAS de SINCO (Tipo COMPRAS).
//   VT Comprado = Valor Orden; VT Cortes = Valor Recepcion (invertido).
// Se agregan como COMPRAS los demas costos asegurados de Oracle que no son contrato ni OC
// (ASEGURADO.xls: GESTION RECURSOS HUMANOS, CAJA MENOR, SERVICIO PUBLICO, COSTOS DISTRIBUIBLES),
// para que el asegurado del control cuadre con el "Asegurado Actual" del Cost Sheet. Esas filas
// llevan el proceso en "Descripcion contrato" y como invertido su mismo valor (ya causado).
// Almacen (ENTRADAS Y SALIDAS.xls, solo Estado Activo; cantidades y valores ya traen signo):
//   Recepciones/Devoluciones -> #ENTRADA = Registro, Cantidad Cortes = Cantidad. VT Cortes va vacio:
//     el invertido sigue en la fila de la OC (Valor Recepcion de COMPRAS.xls), asi no se duplica.
//     La fila de la OC deja vacia su Cantidad Cortes cuando la OC tiene entradas (como en SINCO).
//   Salidas/Reintegros -> #SALIDA = Registro, Cantidad Cons Cols = Cantidad, VT Cons Cols = Valor.
//   Oracle no dice quien recibe el material: la entrada lleva el proveedor de la OC y la salida
//   "(sin contratista Oracle)".
let
    Arbol = ORA_ARBOL,
    CC = ORA_FUENTES[CentroCostos],
    Div = (a as nullable number, b as nullable number) as nullable number => if a = null or b = null or b = 0 then null else a / b,

    OC = Table.SelectRows(ORA_MOV, each [Fuente] = "COMPRAS"),
    FnNum = ORA_FUENTES[FnNum],
    Limpiar = Arbol[Limpiar],

    // ---------- Almacen ----------
    A = ORA_FUENTES[ALMACEN],
    FnColA = (nombre as text) as list =>
        if Table.HasColumns(A, nombre) then Table.Column(A, nombre)
        else error Error.Record("Columna no encontrada", "ENTRADAS Y SALIDAS.xls: falta la columna '" & nombre & "'",
            "Columnas: " & Text.Combine(Table.ColumnNames(A), " | ")),
    Alm = Table.SelectRows(Table.FromColumns({
            List.Transform(FnColA("Registro"), Limpiar), List.Transform(FnColA("Proceso"), Limpiar),
            List.Transform(FnColA("Orden"), Limpiar), List.Transform(FnColA("CBS"), Limpiar),
            List.Transform(FnColA("Paquete"), Limpiar), List.Transform(FnColA("Articulo"), Limpiar),
            List.Transform(FnColA("Descripcion_2"), Limpiar), List.Transform(FnColA("Estado"), Limpiar),
            List.Transform(FnColA("Cantidad"), FnNum), List.Transform(FnColA("Valor"), FnNum)},
            {"Registro", "Proceso", "Orden", "CBS", "Paquete", "Articulo", "Desc Articulo", "Estado", "Cantidad", "Valor"}),
        each [Registro] <> "" and [Registro] <> "Registro" and Text.Upper([Estado]) = "ACTIVO"),
    ProveedorOC = let g = Table.Group(OC, {"Orden"}, {{"P", each List.First(List.Select([Razon Social], each _ <> ""), null), type text}})
        in Record.FromList(g[P], g[Orden]),
    EsEntrada = (p as text) as logical => List.Contains({"RECEPCIONES", "DEVOLUCIONES"}, Text.Upper(p)),
    EsSalida = (p as text) as logical => List.Contains({"SALIDAS", "REINTEGROS"}, Text.Upper(p)),
    DeAlm = Table.FromRecords(List.Transform(Table.ToRecords(Table.SelectRows(Alm, each EsEntrada([Proceso]) or EsSalida([Proceso]))), (r) =>
        let
            ent = EsEntrada(r[Proceso]),
            esCbs = List.Count(Text.Split(r[CBS], "-")) >= 3
        in [
        #"Centro de Costos" = CC, #"Codigo ins" = if r[Articulo] = "" then null else r[Articulo],
        Ins = Arbol[FnIns](r[Articulo], r[Desc Articulo], null),
        #"Codigo act" = if ent or not esCbs then null else Arbol[FnCodAct](r[CBS], r[Paquete]),
        Actividad = if ent or not esCbs then null else Arbol[FnActividad](r[CBS]),
        Capitulo = if ent or not esCbs then null else Arbol[FnNiveles](r[CBS])[Nivel 1],
        Subcapitulo = if ent or r[Paquete] = "" then null else r[Paquete],
        #"# OC / Contrato" = r[Orden],
        #"#ENTRADA" = if ent then r[Registro] else null, #"#SALIDA" = if ent then null else r[Registro],
        #"Nombre Contratista" = if ent then Record.FieldOrDefault(ProveedorOC, r[Orden], null) else "(sin contratista Oracle)",
        #"Descripcion contrato" = r[Proceso],
        #"Cantidad Comprado" = null, #"VT Comprado" = null, #"V/U Comprado" = null,
        #"Cantidad Cortes" = if ent then r[Cantidad] else null, #"VT Cortes" = null,
        #"Cantidad Cons Cols" = if ent then null else r[Cantidad], #"VT Cons Cols" = if ent then null else r[Valor],
        Tipo = "COMPRAS"])),
    OCconEntradas = List.Buffer(List.Distinct(Table.SelectRows(Alm, each EsEntrada([Proceso]))[Orden])),
    DeOC = Table.FromRecords(List.Transform(Table.ToRecords(OC), (r) => [
        #"Centro de Costos" = r[Centro de Costos], #"Codigo ins" = r[Codigo ins], Ins = r[Ins],
        #"Codigo act" = r[Codigo act], Actividad = r[Actividad], Capitulo = r[Capitulo], Subcapitulo = r[Subcapitulo],
        #"# OC / Contrato" = r[Orden], #"#ENTRADA" = null, #"#SALIDA" = null,
        #"Nombre Contratista" = r[Razon Social], #"Descripcion contrato" = r[Titulo],
        #"Cantidad Comprado" = r[Cantidad], #"VT Comprado" = r[Valor Orden], #"V/U Comprado" = Div(r[Valor Orden], r[Cantidad]),
        #"Cantidad Cortes" = if List.Contains(OCconEntradas, r[Orden]) then null else r[Cantidad Recepcion], #"VT Cortes" = r[Valor Recepcion],
        #"Cantidad Cons Cols" = null, #"VT Cons Cols" = null, Tipo = "COMPRAS"])),

    OtrosProcesos = {"GESTION RECURSOS HUMANOS", "CAJA MENOR", "SERVICIO PUBLICO", "COSTOS DISTRIBUIBLES"},
    Otros = Table.SelectRows(ORA_ASEGURADO, each List.Contains(OtrosProcesos, [Proceso]) and ([VT] ?? 0) <> 0),
    DeOtros = Table.FromRecords(List.Transform(Table.ToRecords(Otros), (r) =>
        let esCbs = List.Count(Text.Split(r[CBS], "-")) >= 3 in [
        #"Centro de Costos" = CC, #"Codigo ins" = if r[Articulo] = "" then null else r[Articulo],
        Ins = Arbol[FnIns](r[Articulo], r[Desc Articulo], r[UM]),
        #"Codigo act" = if esCbs then Arbol[FnCodAct](r[CBS], r[Paquete]) else null,
        Actividad = if esCbs then Arbol[FnActividad](r[CBS]) else null,
        Capitulo = if esCbs then Arbol[FnNiveles](r[CBS])[Nivel 1] else null,
        Subcapitulo = if r[Paquete] = "" then null else r[Paquete],
        #"# OC / Contrato" = r[Registro], #"#ENTRADA" = null, #"#SALIDA" = null,
        #"Nombre Contratista" = r[Proceso], #"Descripcion contrato" = r[Proceso],
        #"Cantidad Comprado" = r[Cantidad], #"VT Comprado" = r[VT], #"V/U Comprado" = Div(r[VT], r[Cantidad]),
        #"Cantidad Cortes" = r[Cantidad], #"VT Cortes" = r[VT],
        #"Cantidad Cons Cols" = null, #"VT Cons Cols" = null, Tipo = "COMPRAS"])),

    Tipos = type table [#"Centro de Costos" = text, #"Codigo ins" = text, Ins = text, #"Codigo act" = text, Actividad = text,
        Capitulo = text, Subcapitulo = text, #"# OC / Contrato" = text, #"#ENTRADA" = text, #"#SALIDA" = text,
        #"Nombre Contratista" = text, #"Descripcion contrato" = text, #"Cantidad Comprado" = number, #"VT Comprado" = number,
        #"V/U Comprado" = number, #"Cantidad Cortes" = number, #"VT Cortes" = number, #"Cantidad Cons Cols" = number,
        #"VT Cons Cols" = number, Tipo = text],
    Resultado = Value.ReplaceType(Table.Combine({DeOC, DeOtros, DeAlm}), Tipos)
in
    Table.Buffer(Resultado)
