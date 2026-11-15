let
    Cache = Table.Buffer(CACHE_SEGUIMIENTO),
    Grupos = Table.Group(Cache, {"Centro de Costos"}, {{"Tabla", each
        let j=Json.Document(Text.Combine(Table.Sort(_, {{"Parte",Order.Ascending}})[Datos], ""))
        in Table.FromRows(j[Filas], j[Columnas]), type table}}),
    Combinado = Table.Combine(Grupos[Tabla]),
    Final = Table.Buffer(Table.TransformColumnTypes(Combinado, {
        {"Centro de Costos",type text},{"Codigo ins",Int64.Type},
        {"Cantidad Presupuesto",type number},{"VT Presupuesto",Currency.Type},
        {"Cantidad Proyectado",type number},{"VT Proyectado",Currency.Type},
        {"Cantidad Consumido",type number},{"VT Consumido",Currency.Type}}))
in
    Final
