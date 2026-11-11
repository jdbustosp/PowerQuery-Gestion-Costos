// ORA_MOV: lineas de ordenes de compra (COMPRAS.xls, OC-) y contratos (CONTRATOS.xls, CO-) de
// Oracle normalizadas. Columnas POR NOMBRE (sin respaldo por posicion: si Unifier cambia un
// encabezado la consulta falla con el nombre que falta, en vez de leer otra columna).
//   Orden/Contrato = Cantidad y Valor Orden (lo asegurado; igual a ASEGURADO.xls por proceso)
//   Cortes         = Cantidad_2 y Valor Recepcion (recepciones / actas)
// "CCAMBIOS SIN LINEAS" (asi viene escrito de Oracle) es un cambio de contrato sin CBS.
let
    Fuentes = ORA_FUENTES,
    Arbol = ORA_ARBOL,
    FnNum = Fuentes[FnNum],
    Limpiar = Arbol[Limpiar],
    CC = Fuentes[CentroCostos],

    FnCol = (t as table, nombre as text, archivo as text) as list =>
        if Table.HasColumns(t, nombre) then Table.Column(t, nombre)
        else error Error.Record("Columna no encontrada", archivo & ": falta la columna '" & nombre & "'",
            "Columnas: " & Text.Combine(Table.ColumnNames(t), " | ")),

    FnNormalizar = (t as table, fuente as text, archivo as text) as table =>
        let
            base = Table.FromColumns({
                List.Transform(FnCol(t, "Orden", archivo), Limpiar),
                List.Transform(FnCol(t, "Estatus", archivo), Limpiar),
                List.Transform(FnCol(t, "Razon Social", archivo), Limpiar),
                List.Transform(FnCol(t, "CBS", archivo), Limpiar),
                List.Transform(FnCol(t, "Descripcion", archivo), Limpiar),
                List.Transform(FnCol(t, "Paquete de trabajo", archivo), Limpiar),
                List.Transform(FnCol(t, "Articulo", archivo), Limpiar),
                List.Transform(FnCol(t, "Descripcion_1", archivo), Limpiar),
                List.Transform(FnCol(t, "Titulo", archivo), Limpiar),
                List.Transform(FnCol(t, "Cantidad", archivo), FnNum),
                List.Transform(FnCol(t, "Valor Orden", archivo), FnNum),
                List.Transform(FnCol(t, "Cantidad_2", archivo), FnNum),
                List.Transform(FnCol(t, "Valor Recepcion", archivo), FnNum)
            }, {"Orden", "Estatus", "Razon Social", "CBS", "Desc CBS", "Paquete", "Articulo", "Desc Articulo", "Titulo",
                "Cantidad", "Valor Orden", "Cantidad Recepcion", "Valor Recepcion"}),
            filas = Table.SelectRows(base, each [Orden] <> "" and [Orden] <> "Orden")
        in Table.AddColumn(filas, "Fuente", each fuente, type text),

    Lineas = Table.Combine({
        FnNormalizar(Fuentes[COMPRAS], "COMPRAS", "COMPRAS.xls"),
        FnNormalizar(Fuentes[CONTRATOS], "CONTRATOS", "CONTRATOS.xls")}),

    Enriquecido = Table.AddColumn(Lineas, "__R", each
        let
            esCbs = List.Count(Text.Split([CBS], "-")) >= 3,
            niv = Arbol[FnNiveles](if esCbs then [CBS] else null)
        in [
            #"Centro de Costos" = CC,
            #"Codigo ins" = if [Articulo] = "" then null else [Articulo],
            Ins = Arbol[FnIns]([Articulo], [Desc Articulo], null),
            #"Codigo act" = if esCbs then Arbol[FnCodAct]([CBS], [Paquete]) else null,
            Actividad = if esCbs then Arbol[FnActividad]([CBS]) else (if [CBS] = "" then null else Text.Replace([CBS], "CCAMBIOS", "CAMBIOS")),
            Capitulo = niv[Nivel 1],
            Subcapitulo = if [Paquete] = "" then null else [Paquete]
        ]),
    Resultado = Table.Buffer(Table.ExpandRecordColumn(Enriquecido, "__R",
        {"Centro de Costos", "Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo"}))
in
    Resultado
