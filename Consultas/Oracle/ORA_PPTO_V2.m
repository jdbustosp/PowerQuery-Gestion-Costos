// ORA_PPTO_V2: filas "PPTO V2" para la BD (columna VT Presupuesto v2), POR INSUMO, desde ORA_PPTO_V2_BD
// (hoja Actualizacion_ppto del archivo de descargas en SharePoint). Antes era por CBS desde el Cost Sheet
// ("Presupuesto Aprobado"), sin detalle de paquete ni insumo.
let
    V2 = ORA_PPTO_V2_BD,
    Res = Table.FromRecords(List.Transform(Table.ToRecords(V2), (r) => [
        #"Centro de Costos" = r[Centro de Costos],
        #"Codigo act" = r[Codigo act],
        #"Codigo ins" = r[Codigo ins],
        Ins = r[Ins],
        Actividad = r[Actividad],
        Capitulo = r[Capitulo],
        Subcapitulo = r[Subcapitulo],
        Tipo = "PPTO V2",
        #"VT Presupuesto v2" = r[VT Presupuesto]]),
        type table [#"Centro de Costos" = text, #"Codigo act" = text, #"Codigo ins" = nullable text, Ins = nullable text, Actividad = text,
            Capitulo = nullable text, Subcapitulo = nullable text, Tipo = text, #"VT Presupuesto v2" = nullable number])
in
    Table.Buffer(Res)
