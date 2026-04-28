let
    // =========================================================
    // ITEMSINSUMOS: Lee de SP_Seguimiento_Parsed (sin re-parsear HTML)
    // =========================================================
    Source = SP_Seguimiento_Parsed,
    Selected = Table.SelectColumns(Source, {"Centro de Costos", "Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo", "Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido"}),
    Typed = Table.TransformColumnTypes(Selected,{{"Centro de Costos", type text}, {"Codigo ins", Int64.Type}, {"Cantidad Proyectado", type number}, {"VT Proyectado", Currency.Type}, {"Cantidad Consumido", type number}, {"VT Consumido", Currency.Type}}),
    TablaEnMemoria = Table.Buffer(Typed)
in 
    TablaEnMemoria
