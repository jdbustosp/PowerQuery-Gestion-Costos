// ORA_PPTO_BD: presupuesto Oracle por insumo (PRESUPUESTOS.xlsx), con el MISMO esquema que
// PPTO_BD de SINCO. Codigo act = CBS de la actividad, Subcapitulo = paquete de trabajo,
// Capitulo = primer nivel del CBS (los demas niveles los agrega ORA_BD).
let
    Arbol = ORA_ARBOL,
    CC = ORA_FUENTES[CentroCostos],
    Agrupado = Table.Group(Arbol[Insumos], {"Paquete", "CBS", "Articulo"}, {
        {"Desc", each List.First(List.Select([Desc], (d) => d <> ""), null), type nullable text},
        {"UM", each List.First(List.Select([UM], (u) => u <> ""), null), type nullable text},
        {"Cantidad Presupuesto", each List.Sum([Cant]), type nullable number},
        {"VT Presupuesto", each List.Sum([Total]), type nullable number}}),
    Filas = Table.FromRecords(List.Transform(Table.ToRecords(Agrupado), (r) => [
        #"Centro de Costos" = CC,
        #"Codigo ins" = r[Articulo],
        Ins = Arbol[FnIns](r[Articulo], r[Desc], r[UM]),
        #"Codigo act" = Arbol[FnCodAct](r[CBS], r[Paquete]),
        Actividad = Arbol[FnActividad](r[CBS]),
        Capitulo = Arbol[FnNiveles](r[CBS])[Nivel 1],
        Subcapitulo = r[Paquete],
        #"Cantidad Presupuesto" = r[Cantidad Presupuesto],
        #"VT Presupuesto" = r[VT Presupuesto],
        #"V/U Presupuesto" = if r[Cantidad Presupuesto] <> null and r[Cantidad Presupuesto] <> 0 and r[VT Presupuesto] <> null
            then r[VT Presupuesto] / r[Cantidad Presupuesto] else null,
        Tipo = "PPTO"]),
        type table [#"Centro de Costos" = text, #"Codigo ins" = text, Ins = text, #"Codigo act" = text, Actividad = text,
            Capitulo = text, Subcapitulo = text, #"Cantidad Presupuesto" = number, #"VT Presupuesto" = number,
            #"V/U Presupuesto" = number, Tipo = text])
in
    Table.Buffer(Filas)
