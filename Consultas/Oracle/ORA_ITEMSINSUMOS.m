// ORA_ITEMSINSUMOS: equivalente Oracle del "Seguimiento por items e insumos" de SINCO, desde el
// Cost Sheet (CONTROL.xlsx), con el MISMO esquema que ITEMSINSUMOS. El Cost Sheet es por
// actividad CBS (sin insumo ni paquete), asi que hay una fila por actividad:
//   VT Proyectado = Asegurado Actual + Por Asegurar (proyeccion del constructor)
//   VT Consumido  = Consumido
// Actividad = fila del Cost Sheet cuyo ultimo segmento del CBS no esta en ceros.
let
    Fuentes = ORA_FUENTES,
    Arbol = ORA_ARBOL,
    FnNum = Fuentes[FnNum],
    Limpiar = Arbol[Limpiar],
    CC = Fuentes[CentroCostos],
    Control = Fuentes[CONTROL],
    Cols = Table.ColumnNames(Control),
    FnCol = (nombre as text) as list =>
        if Table.HasColumns(Control, nombre) then List.Transform(Table.Column(Control, nombre), FnNum)
        else error Error.Record("Columna no encontrada", "CONTROL.xlsx: falta la columna '" & nombre & "'", "Columnas: " & Text.Combine(Cols, " | ")),
    Base = Table.FromColumns({
        List.Transform(Table.Column(Control, Cols{0}), Limpiar),
        FnCol("Asegurado Actual"), FnCol("Por Asegurar"), FnCol("Consumido")},
        {"CBS", "Asegurado", "PorAsegurar", "Consumido"}),
    EsActividad = (c as text) as logical =>
        let s = Text.Split(c, "-") in List.Count(s) >= 3 and Text.Select(List.Last(s), {"0"}) <> List.Last(s),
    Hojas = Table.Group(Table.SelectRows(Base, each EsActividad([CBS])), {"CBS"}, {
        {"VT Proyectado", each List.Sum(List.Transform([Asegurado], (v) => v ?? 0)) + List.Sum(List.Transform([PorAsegurar], (v) => v ?? 0)), type number},
        {"VT Consumido", each List.Sum(List.Transform([Consumido], (v) => v ?? 0)), type number}}),
    ConValor = Table.SelectRows(Hojas, each [VT Proyectado] <> 0 or [VT Consumido] <> 0),
    Filas = Table.FromRecords(List.Transform(Table.ToRecords(ConValor), (r) => [
        #"Centro de Costos" = CC,
        #"Codigo ins" = null,
        Ins = "(COST SHEET ORACLE)",
        #"Codigo act" = r[CBS],
        Actividad = Arbol[FnActividad](r[CBS]),
        Capitulo = Arbol[FnNiveles](r[CBS])[Nivel 1],
        Subcapitulo = null,
        #"Cantidad Proyectado" = null,
        #"VT Proyectado" = r[VT Proyectado],
        #"Cantidad Consumido" = null,
        #"VT Consumido" = r[VT Consumido]]),
        type table [#"Centro de Costos" = text, #"Codigo ins" = text, Ins = text, #"Codigo act" = text, Actividad = text,
            Capitulo = text, Subcapitulo = text, #"Cantidad Proyectado" = number, #"VT Proyectado" = number,
            #"Cantidad Consumido" = number, #"VT Consumido" = number])
in
    Table.Buffer(Filas)
