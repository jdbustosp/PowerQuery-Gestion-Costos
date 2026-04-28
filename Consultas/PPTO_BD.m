let
    // =========================================================
    // PPTO_BD: Lee de SP_Seguimiento_Parsed (sin re-parsear HTML)
    // =========================================================
    Source = SP_Seguimiento_Parsed,
    Selected = Table.SelectColumns(Source, {"Centro de Costos", "Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo", "Cantidad Presupuesto", "VT Presupuesto"}),
    
    // V/U Presupuesto
    AddVU = Table.AddColumn(Selected, "V/U Presupuesto", each 
        if [Cantidad Presupuesto] = null or [Cantidad Presupuesto] = 0 or [VT Presupuesto] = null then null 
        else [VT Presupuesto] / [Cantidad Presupuesto], Currency.Type),
    
    // Tipo y filtro
    AddTipo = Table.AddColumn(AddVU, "Tipo", each "PPTO", type text),
    Filtered = Table.SelectRows(AddTipo, each [VT Presupuesto] <> null and [VT Presupuesto] <> 0),
    
    Typed = Table.TransformColumnTypes(Filtered,{{"Centro de Costos", type text}, {"Codigo ins", Int64.Type}, {"Cantidad Presupuesto", type number}, {"VT Presupuesto", Currency.Type}, {"V/U Presupuesto", Currency.Type}, {"Tipo", type text}}),
    TablaEnMemoria = Table.Buffer(Typed)
in 
    TablaEnMemoria
