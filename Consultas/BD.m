let
    Tol = 0.01,
    FnRemoveAccentsSymbols = F_Globales[FnRemoveAccentsSymbols],
    FnCleanContratista = F_Globales[FnCleanContratista],

    // ============================================================
    // CONSTANTES DE COLUMNAS
    // ============================================================
    ColumnasOrden = {
        "Centro de Costos", "Codigo act", "Codigo ins", "Ins", "Actividad", "Capitulo", "Subcapitulo", "Tipo",
        "# OC / Contrato", "Nombre Contratista", "Descripcion contrato", "# CC - Comparativo", "Clasificador",
        "Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido",
        "Cantidad Comprado", "V/U Comprado", "VT Comprado",
        "Cantidad Contratado", "V/U Contratado", "VT Contratado",
        "Cantidad Presupuesto", "V/U Presupuesto", "VT Presupuesto",
        "Cant. aprobacion", "V/U aprobacion", "VR total aprobacion",
        "Valor Total ppto (CC)", "Cantidad Cortes", "VT Cortes", "Valor descuento",
        "Cantidad_Calc", "V/U ppto (CC)", "Cantidad CC Cons", "V/U CC cons", "VT CC cons",
        "VR_Bruto_con_desc", "Estado", "Fecha_de_pago", "# Prov._(descue", "No_Prov",
        "Centros_de_costos", "Clasificador_Actividad", "Capitulo_Costo directo", "Capitulo_Centro_Costos",
        "NIT", "No_Factura", "Fecha_Factura"
    },
    NumCols = {
        "Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido",
        "Cantidad Comprado", "V/U Comprado", "VT Comprado",
        "Cantidad Contratado", "V/U Contratado", "VT Contratado",
        "Cantidad Presupuesto", "V/U Presupuesto", "VT Presupuesto",
        "Cant. aprobacion", "V/U aprobacion", "VR total aprobacion",
        "Valor Total ppto (CC)", "Cantidad Cortes", "VT Cortes", "Valor descuento",
        "Cantidad_Calc", "V/U ppto (CC)", "Cantidad CC Cons", "V/U CC cons", "VT CC cons",
        "VR_Bruto_con_desc"
    },
    ColsBanderas = {"Esc1", "Esc3", "Esc2", "Esc4", "Esc5"},

    // 🔥 MODO CASCADA: Conexion directa a las consultas en memoria.
    T_Items_Raw = ITEMSINSUMOS,
    T_Items = if Table.HasColumns(T_Items_Raw, "Tipo") then T_Items_Raw else Table.AddColumn(T_Items_Raw, "Tipo", each "ITEMS", type text),

    T_Compras = COMPRAS,
    T_Contratos = CONTRATOS,
    T_Ppto = PPTO_BD,
    T_Comp = COMPARATIVOS,
    T_Aprob = try APROBACIONES_SP otherwise #table({"Tipo"}, {}),
    T_Prov = try PROVISIONES_SP otherwise #table({"Tipo"}, {}),
    T_Desc = DESCUENTOS,
    T_Disp = DISPONIBLE,

    Origen = Table.Combine({T_Items, T_Compras, T_Contratos, T_Ppto, T_Comp, T_Aprob, T_Prov, T_Desc, T_Disp}),

    ColumnasReordenadas = Table.SelectColumns(Origen, ColumnasOrden, MissingField.Ignore),

    LlavesLimpias = Table.TransformColumns(ColumnasReordenadas, {
        {"Centro de Costos", each if _ = null then "" else Text.Upper(Text.Trim(Text.From(_))), type text},
        {"Codigo act", each if _ = null then "" else Text.Upper(Text.Trim(Text.From(_))), type text},
        {"Ins", each if _ = null then "" else Text.Upper(Text.Trim(Text.From(_))), type text},
        {"Tipo", each if _ = null then "" else Text.Upper(Text.Trim(Text.From(_))), type text},
        {"# OC / Contrato", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Nombre Contratista", each FnCleanContratista(_), type text},
        {"Descripcion contrato", each FnRemoveAccentsSymbols(if _ = null then null else Text.Trim(Text.From(_))), type text}
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

    // 🔥 UNIFICACION DE CONTRATISTAS (Prioridad: COMPRAS y CONTRATOS)
    ContratistasPrioridad = Table.SelectRows(BaseClasificada, each ([Tipo] = "COMPRAS" or [Tipo] = "CONTRATOS") and [Nombre Contratista] <> null and Text.Trim([Nombre Contratista]) <> ""),
    ContratistasMaestros = Table.Group(ContratistasPrioridad, {"Centro de Costos", "Codigo act", "Ins"}, {{"Nombre Maestro", each List.First([Nombre Contratista]), type text}}),

    CruceContratistas = Table.NestedJoin(BaseClasificada, {"Centro de Costos", "Codigo act", "Ins"}, ContratistasMaestros, {"Centro de Costos", "Codigo act", "Ins"}, "Maestro", JoinKind.LeftOuter),
    BaseExpandidaC = Table.ExpandTableColumn(CruceContratistas, "Maestro", {"Nombre Maestro"}),

    BaseConNombreUnificado = Table.AddColumn(BaseExpandidaC, "Nombre Contratista Final", each if [Nombre Maestro] <> null then [Nombre Maestro] else [Nombre Contratista], type text),
    BaseClasificadaFinal = Table.RenameColumns(
        Table.RemoveColumns(BaseConNombreUnificado, {"Nombre Contratista", "Nombre Maestro"}, MissingField.Ignore),
        {{"Nombre Contratista Final", "Nombre Contratista"}}
    ),

    // El "otherwise 0" del try ya cubre el caso null; no necesitamos doble guarda
    NumerosSeguros = Table.TransformColumns(BaseClasificadaFinal, List.Transform(NumCols, each {_, (v) => try Number.From(v) otherwise 0, type number}), null, MissingField.Ignore),

    AddCantAseg = Table.AddColumn(NumerosSeguros, "Cantidad asegurada", each [Cantidad Contratado] + [Cantidad Comprado], type number),
    AddVTAseg = Table.AddColumn(AddCantAseg, "VT Asegurada", each [VT Contratado] + [VT Comprado], type number),
    // 🚀 Buffer aqui: AddVUAseg es referenciado dos veces (en AgrupadoResumen y CruceConBase)
    AddVUAseg = Table.Buffer(Table.AddColumn(AddVTAseg, "V/U asegurada", each if [Cantidad asegurada] <> 0 then [VT Asegurada] / [Cantidad asegurada] else 0, type number)),

    AgrupadoResumen = Table.Group(AddVUAseg, {"Centro de Costos", "Codigo act", "Ins"}, {{"vtProj", each List.Sum([VT Proyectado]), type number}, {"vtCons", each List.Sum([VT Consumido]), type number}, {"vtAseg", each List.Sum([VT Asegurada]), type number}, {"vtAprb", each List.Sum([VR total aprobacion]), type number}}),

    // 🔥 EL MOTOR DE ESCENARIOS
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
        E5 = (vAseg > 0) and (vCons = 0)
    in [
        Esc1 = if E1 = null then false else E1,
        Esc3 = if E3 = null then false else E3,
        Esc2 = if E2 = null then false else E2,
        Esc4 = if E4 = null then false else E4,
        Esc5 = if E5 = null then false else E5
    ]),

    ExpandirBanderas = Table.ExpandRecordColumn(ResumenEscenarios, "Motor", ColsBanderas),
    BanderasBuffer = Table.Buffer(Table.SelectColumns(ExpandirBanderas, {"Centro de Costos", "Codigo act", "Ins"} & ColsBanderas)),

    // 🛡️ LeftOuter + coalesce de banderas a false: defensivo si AgrupadoResumen perdiera alguna llave
    CruceConBase = Table.NestedJoin(AddVUAseg, {"Centro de Costos", "Codigo act", "Ins"}, BanderasBuffer, {"Centro de Costos", "Codigo act", "Ins"}, "B", JoinKind.LeftOuter),
    BaseConBanderas = Table.ExpandTableColumn(CruceConBase, "B", ColsBanderas),
    BaseConBanderasSafe = Table.ReplaceValue(BaseConBanderas, null, false, Replacer.ReplaceValue, ColsBanderas),

    // LA REGLA DE APLICACION: Orden de prioridad estricto
    AplicarProyeccion = Table.AddColumn(BaseConBanderasSafe, "VT Proyectado Colsubsidio", each
        if [Tipo] = "POR ADJUDICAR" then [#"Valor Total ppto (CC)"]
        else if [Esc4] = true then [VT Consumido]
        else if [Esc5] = true then 0
        else if [Esc1] = true then [VR total aprobacion]
        else if [Esc3] = true then [VT Proyectado]
        else if [Esc2] = true then (if [VT Asegurada] <> 0 then [VT Asegurada] else null)
        else null,
    type number),

    FinalClean = Table.RemoveColumns(AplicarProyeccion, ColsBanderas, MissingField.Ignore),
    TablaMaestraFinal = FinalClean
in
    TablaMaestraFinal
