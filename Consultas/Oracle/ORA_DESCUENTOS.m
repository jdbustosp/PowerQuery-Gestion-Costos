// ORA_DESCUENTOS: descuentos Oracle (ASEGURADO.xls, Proceso DESCUENTOS) con el MISMO esquema que
// DESCUENTOS de SINCO. Oracle los trae negativos; aqui van en positivo como en SINCO.
// Contratista y descripcion: los del contrato/OC del registro (ORA_MOV).
let
    Arbol = ORA_ARBOL,
    CC = ORA_FUENTES[CentroCostos],
    Mov = Table.Distinct(Table.SelectColumns(ORA_MOV, {"Orden", "Razon Social", "Titulo"}), {"Orden"}),
    DictRazon = Record.FromList(Mov[Razon Social], Mov[Orden]),
    DictTitulo = Record.FromList(Mov[Titulo], Mov[Orden]),
    Desc = Table.SelectRows(ORA_ASEGURADO, each [Proceso] = "DESCUENTOS" and ([VT] ?? 0) <> 0),
    Filas = Table.FromRecords(List.Transform(Table.ToRecords(Desc), (r) =>
        let esCbs = List.Count(Text.Split(r[CBS], "-")) >= 3 in [
        #"Centro de Costos" = CC,
        #"Codigo act" = if esCbs then Arbol[FnCodAct](r[CBS], r[Paquete]) else null,
        Actividad = if esCbs then Arbol[FnActividad](r[CBS]) else null,
        Capitulo = if esCbs then Arbol[FnNiveles](r[CBS])[Nivel 1] else null,
        Subcapitulo = if r[Paquete] = "" then null else r[Paquete],
        #"# OC / Contrato" = r[Registro],
        #"Nombre Contratista" = Record.FieldOrDefault(DictRazon, r[Registro], null),
        #"Descripcion contrato" = Record.FieldOrDefault(DictTitulo, r[Registro], null),
        #"Valor descuento" = -r[VT],
        Tipo = "Descuento"]),
        type table [#"Centro de Costos" = text, #"Codigo act" = text, Actividad = text, Capitulo = text, Subcapitulo = text,
            #"# OC / Contrato" = text, #"Nombre Contratista" = text, #"Descripcion contrato" = text,
            #"Valor descuento" = number, Tipo = text])
in
    Table.Buffer(Filas)
