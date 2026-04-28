let
    Tol = 0.01,

    // 🔥 MODO CASCADA: Conexión directa a las consultas en memoria.
    T_Items_Raw = ITEMSINSUMOS,
    T_Items = if Table.HasColumns(T_Items_Raw, "Tipo") then T_Items_Raw else Table.AddColumn(T_Items_Raw, "Tipo", each "ITEMS", type text),

    T_Compras = COMPRAS,
    T_Contratos = CONTRATOS,
    T_Ppto = PPTO_BD,
    T_Comp = COMPARATIVOS,
    T_Desc = DESCUENTOS,
    T_Disp = DISPONIBLE,

    Origen = Table.Combine({T_Items, T_Compras, T_Contratos, T_Ppto, T_Comp, T_Desc, T_Disp}),

    ColumnasReordenadas = Table.SelectColumns(Origen, 
        {"Centro de Costos", "Codigo act", "Codigo ins", "Ins", "Actividad", "Capitulo", "Subcapitulo", "Tipo", "# OC / Contrato", "Nombre Contratista", "Descripcion contrato", "# CC - Comparativo", "Clasificador", "Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido", "Cantidad Comprado", "V/U Comprado", "VT Comprado", "Cantidad Contratado", "V/U Contratado", "VT Contratado", "Cantidad Presupuesto", "V/U Presupuesto", "VT Presupuesto", "Cant. aprobacion", "V/U aprobacion", "VR total aprobacion", "Valor Total ppto (CC)", "Cantidad Cortes", "VT Cortes", "Valor descuento", "Cantidad_Calc", "V/U ppto (CC)"}, MissingField.Ignore),

    LlavesLimpias = Table.TransformColumns(ColumnasReordenadas, {
        {"Centro de Costos", each if _ = null then "" else Text.Upper(Text.Trim(Text.From(_))), type text},
        {"Codigo act", each if _ = null then "" else Text.Upper(Text.Trim(Text.From(_))), type text},
        {"Ins", each if _ = null then "" else Text.Upper(Text.Trim(Text.From(_))), type text},
        {"Tipo", each if _ = null then "" else Text.Upper(Text.Trim(Text.From(_))), type text},
        {"# OC / Contrato", each if _ = null then null else Text.Trim(Text.From(_)), type text} 
    }, null, MissingField.Ignore),

    FiltroTipoValido = Table.SelectRows(LlavesLimpias, each [Tipo] <> null and [Tipo] <> ""),

    // 🚀 Lookup por Record para Clasificadores (O(1) por fila en vez de JOIN O(N))
    ClasificadorRows = Table.SelectRows(
        Table.Distinct(
            Table.SelectColumns(FiltroTipoValido, {"Centro de Costos", "Codigo act", "Ins", "Clasificador"}),
            {"Centro de Costos", "Codigo act", "Ins"}
        ),
        each [Clasificador] <> null and Text.Trim(Text.From([Clasificador])) <> ""
    ),
    ClasificadorKeys = List.Transform(Table.ToRecords(Table.SelectColumns(ClasificadorRows, {"Centro de Costos", "Codigo act", "Ins"})), each [Centro de Costos] & "|" & [Codigo act] & "|" & [Ins]),
    ClasificadorMap = Record.FromList(ClasificadorRows[Clasificador], ClasificadorKeys),
    BaseClasificada = Table.AddColumn(Table.RemoveColumns(FiltroTipoValido, {"Clasificador"}, MissingField.Ignore), "Clasificador", each 
        let key = [Centro de Costos] & "|" & [Codigo act] & "|" & [Ins]
        in try Record.Field(ClasificadorMap, key) otherwise null, type text),

    NumCols = {"Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido", "Cantidad Comprado", "V/U Comprado", "VT Comprado", "Cantidad Contratado", "V/U Contratado", "VT Contratado", "Cantidad Presupuesto", "V/U Presupuesto", "VT Presupuesto", "Cant. aprobacion", "V/U aprobacion", "VR total aprobacion", "Valor Total ppto (CC)", "Cantidad Cortes", "VT Cortes", "Valor descuento", "Cantidad_Calc", "V/U ppto (CC)"},
    NumerosSeguros = Table.TransformColumns(BaseClasificada, List.Transform(NumCols, each {_, (v) => let n = try Number.From(v) otherwise 0 in if n = null then 0 else n, type number}), null, MissingField.Ignore),

    AddCantAseg = Table.AddColumn(NumerosSeguros, "Cantidad asegurada", each [Cantidad Contratado] + [Cantidad Comprado], type number),
    AddVTAseg = Table.AddColumn(AddCantAseg, "VT Asegurada", each [VT Contratado] + [VT Comprado], type number),
    AddVUAseg = Table.AddColumn(AddVTAseg, "V/U asegurada", each if [Cantidad asegurada] <> 0 then [VT Asegurada] / [Cantidad asegurada] else 0, type number),

    AgrupadoResumen = Table.Group(AddVUAseg, {"Centro de Costos", "Codigo act", "Ins"}, {{"vtProj", each List.Sum([VT Proyectado]), type number}, {"vtCons", each List.Sum([VT Consumido]), type number}, {"vtAseg", each List.Sum([VT Asegurada]), type number}, {"vtAprb", each List.Sum([VR total aprobacion]), type number}}),

    // 🔥 EL MOTOR DE ESCENARIOS ACTUALIZADO CON ESCENARIO 5
    ResumenEscenarios = Table.AddColumn(AgrupadoResumen, "Motor", each let 
        vAseg = [vtAseg], 
        vAprb = [vtAprb], 
        vProj = [vtProj], 
        vCons = [vtCons], 
        
        E1 = (vAseg > 0) and (vAprb > 0) and (Number.Abs(vAseg - vAprb) <= Tol), 
        MaxA = if vAseg > vAprb then vAseg else vAprb, 
        E3 = (vProj <> 0) and (Number.Abs(vProj - vCons) <= Tol) and (vProj < MaxA), 
        E2 = (vAseg > 0),
        E4 = (vCons > vAseg),
        
        // NUEVO ESCENARIO 5: Si hay valor asegurado pero el consumido es exactamente 0.
        E5 = (vAseg > 0) and (vCons = 0)

    in [
        Esc1 = if E1 = null then false else E1, 
        Esc3 = if E3 = null then false else E3, 
        Esc2 = if E2 = null then false else E2,
        Esc4 = if E4 = null then false else E4,
        Esc5 = if E5 = null then false else E5
    ]),
    
    ExpandirBanderas = Table.ExpandRecordColumn(ResumenEscenarios, "Motor", {"Esc1", "Esc3", "Esc2", "Esc4", "Esc5"}),

    // 🚀 Buffer las banderas para acelerar el JOIN de escenarios
    BanderasBuffer = Table.Buffer(Table.SelectColumns(ExpandirBanderas, {"Centro de Costos", "Codigo act", "Ins", "Esc1", "Esc2", "Esc3", "Esc4", "Esc5"})),

    CruceConBase = Table.NestedJoin(AddVUAseg, {"Centro de Costos", "Codigo act", "Ins"}, BanderasBuffer, {"Centro de Costos", "Codigo act", "Ins"}, "B", JoinKind.Inner),
    BaseConBanderas = Table.ExpandTableColumn(CruceConBase, "B", {"Esc1", "Esc3", "Esc2", "Esc4", "Esc5"}),

    // LA REGLA DE APLICACIÓN: Orden de prioridad estricto
    AplicarProyeccion = Table.AddColumn(BaseConBanderas, "VT Proyectado Colsubsidio", each 
        if [Tipo] = "POR ADJUDICAR" then [#"Valor Total ppto (CC)"] 
        else if [Esc4] = true then [VT Consumido] 
        else if [Esc5] = true then 0   // <-- Prioridad: Si no hay consumo, la proyección baja a 0
        else if [Esc1] = true then [VR total aprobacion] 
        else if [Esc3] = true then [VT Proyectado] 
        else if [Esc2] = true then (if [VT Asegurada] <> 0 then [VT Asegurada] else null) 
        else null, 
    type number),

    FinalClean = Table.RemoveColumns(AplicarProyeccion, {"Esc1", "Esc3", "Esc2", "Esc4", "Esc5"}),
    TablaMaestraFinal = FinalClean
in
    TablaMaestraFinal
