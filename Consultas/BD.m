let
    Tol = 0.01,
    FnRemoveAccentsSymbols = F_Globales[FnRemoveAccentsSymbols],
    FnCleanContratista = F_Globales[FnCleanContratista],
    FnRemoveAccentMarks = F_Globales[FnRemoveAccentMarks],
    ToNumber0 = (v as any) as number =>
        let n = try Number.From(v) otherwise null
        in if n = null then 0 else n,

    // ============================================================
    // CONSTANTES DE COLUMNAS
    // ============================================================
    ColumnasOrden = {
        "Centro de Costos", "Codigo act", "Codigo ins", "Ins", "Actividad", "Capitulo", "Subcapitulo", "Tipo",
        "# OC / Contrato", "#ENTRADA", "#SALIDA", "Nombre Contratista", "Descripcion contrato", "# CC - Comparativo", "Clasificador",
        "Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido",
        "Cantidad Comprado", "V/U Comprado", "VT Comprado",
        "Cantidad Contratado", "V/U Contratado", "VT Contratado",
        "Cantidad Presupuesto", "V/U Presupuesto", "VT Presupuesto",
        "Cant. aprobacion", "V/U aprobacion", "VR total aprobacion",
        "Valor Total ppto (CC)", "Cantidad Cortes", "VT Cortes", "Cantidad Cons Cols", "VT Cons Cols", "Valor descuento",
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
        "Valor Total ppto (CC)", "Cantidad Cortes", "VT Cortes", "Cantidad Cons Cols", "VT Cons Cols", "Valor descuento",
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
    T_Comp = try COMPARATIVOS otherwise #table({"Tipo"}, {}),
    T_Aprob = try APROBACIONES_SP otherwise #table({"Tipo"}, {}),
    T_Prov = try PROVISIONES_SP otherwise #table({"Tipo"}, {}),
    T_Desc = try DESCUENTOS otherwise #table({"Tipo"}, {}),
    T_Disp = try DISPONIBLE otherwise #table({"Tipo"}, {}),

    Origen = Table.Combine({T_Items, T_Compras, T_Contratos, T_Ppto, T_Comp, T_Aprob, T_Prov, T_Desc, T_Disp}),

    // RED DE SEGURIDAD CRITICA: cualquier valor Error que venga de las fuentes (COMPRAS, CONTRATOS,
    // DESCUENTOS, COMPARATIVOS, etc) lo convertimos a null antes de procesar. Esto neutraliza
    // los miles de errores que Power Query rastrea (cada uno es costoso) y permite que el
    // resto del pipeline use sus rutas null-safe normales.
    OrigenLimpio = Table.ReplaceErrorValues(Origen,
        List.Transform(Table.ColumnNames(Origen), each {_, null})),

    ColumnasReordenadas = Table.SelectColumns(OrigenLimpio, ColumnasOrden, MissingField.Ignore),

    // Recorte TEMPRANO (solo La Arboleda): estas 16 columnas no se usan en NINGUN calculo
    // intermedio de todo el pipeline (verificado: 0 referencias [Campo] en BD.m), asi que
    // se quitan aqui mismo en vez de esperar al final. Esto reduce el ancho de tabla que
    // procesan TODOS los pasos siguientes (TransformColumns, ReplaceErrorValues, Table.Group,
    // NestedJoin), no solo la escritura final a la hoja.
    ColumnasReordenadas_Trim = Table.RemoveColumns(ColumnasReordenadas, {
        "Codigo ins", "V/U Comprado", "V/U Contratado", "V/U Presupuesto",
        "Cantidad Cortes", "Cantidad Cons Cols", "VT Cons Cols", "Cantidad_Calc",
        "V/U ppto (CC)", "Estado", "Fecha_de_pago", "Clasificador_Actividad",
        "Capitulo_Costo directo", "NIT", "No_Factura", "Fecha_Factura"
    }, MissingField.Ignore),

    // try/otherwise en cada transformacion: si Text.From recibe algo raro (record, list, etc) no rompe
    LlavesLimpias = Table.TransformColumns(ColumnasReordenadas_Trim, {
        {"Centro de Costos",     each try (if _ = null then "" else Text.Upper(Text.Trim(Text.From(_)))) otherwise "", type text},
        {"Codigo act",           each try (if _ = null then "" else Text.Upper(Text.Trim(Text.From(_)))) otherwise "", type text},
        {"Ins",                  each try (if _ = null then "" else Text.Upper(Text.Trim(FnRemoveAccentMarks(_)))) otherwise "", type text},
        {"Actividad",            each try (if _ = null then null else Text.Upper(Text.Trim(FnRemoveAccentMarks(_)))) otherwise null, type text},
        {"Tipo",                 each try (if _ = null then "" else Text.Upper(Text.Trim(Text.From(_)))) otherwise "", type text},
        {"# OC / Contrato",      each try (if _ = null then null else Text.Trim(Text.From(_))) otherwise null, type text},
        {"Nombre Contratista",   each try FnCleanContratista(_) otherwise null, type text},
        {"Descripcion contrato", each try FnRemoveAccentsSymbols(if _ = null then null else Text.Trim(Text.From(_))) otherwise null, type text}
    }, null, MissingField.Ignore),

    // Buffer: FiltroTipoValido se usa 2 veces (ClasificadorRows y BaseClasificada) — sin buffer,
    // cada referencia puede forzar recalcular toda la cadena previa (Origen: Combine de 9 fuentes
    // + ReplaceErrorValues sobre 53 columnas, el paso mas caro conocido de BD).
    FiltroTipoValido = Table.Buffer(Table.SelectRows(LlavesLimpias, each [Tipo] <> null and [Tipo] <> "")),

    // 🚀 Lookup por Record para Clasificadores (O(1) por fila en vez de JOIN O(N))
    ClasificadorRows = Table.SelectRows(
        Table.Distinct(
            Table.SelectColumns(FiltroTipoValido, {"Centro de Costos", "Codigo act", "Ins", "Clasificador"}, MissingField.Ignore),
            {"Centro de Costos", "Codigo act", "Ins"}
        ),
        each try ([Clasificador] <> null and Text.Trim(Text.From([Clasificador])) <> "") otherwise false
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
    NumerosSeguros = Table.TransformColumns(BaseClasificadaFinal, List.Transform(NumCols, each {_, ToNumber0, type number}), null, MissingField.Ignore),

    // try/otherwise 0 aqui porque NumerosSeguros puede dejar errores de conversion si la celda
    // fuente traia un valor no numerico que Number.From no pudo convertir
    AddCantAseg = Table.AddColumn(NumerosSeguros, "Cantidad asegurada",
        each ToNumber0([Cantidad Contratado]) + ToNumber0([Cantidad Comprado]), type number),
    AddVTAseg = Table.AddColumn(AddCantAseg, "VT Asegurada",
        each ToNumber0([VT Contratado]) + ToNumber0([VT Comprado]), type number),
    AddVUAseg = Table.Buffer(Table.AddColumn(AddVTAseg, "V/U asegurada",
        each let qa = ToNumber0([Cantidad asegurada]), vt = ToNumber0([VT Asegurada])
            in if qa <> 0 then vt / qa else 0, type number)),

    // List.Transform con try protege List.Sum de valores Error en celdas individuales
    AgrupadoResumen = Table.Group(AddVUAseg, {"Centro de Costos", "Codigo act", "Ins"}, {
        {"vtProj", each List.Sum(List.Transform([VT Proyectado],       each ToNumber0(_))), type number},
        {"vtCons", each List.Sum(List.Transform([VT Consumido],        each ToNumber0(_))), type number},
        {"vtAseg", each List.Sum(List.Transform([VT Asegurada],        each ToNumber0(_))), type number},
        {"vtAprb", each List.Sum(List.Transform([VR total aprobacion], each ToNumber0(_))), type number}
    }),

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
    FinalSinErrores = Table.ReplaceErrorValues(FinalClean, List.Transform(Table.ColumnNames(FinalClean), each {_, null})),

    // Campo para tablas dinamicas: relaciona No_Prov solo por OC normalizada.
    // No replica por centro de costos porque eso infla las sumas.
    CleanKeyText = (v as any) as nullable text =>
        let t = try Text.Upper(Text.Trim(Text.From(v))) otherwise null
        in if t = null or t = "" then null else t,
    CleanOCKey = (v as any) as nullable text =>
        let
            t = CleanKeyText(v),
            digits = if t = null then null else Text.Select(t, {"0".."9"})
        in
            if digits <> null and digits <> "" then digits else t,

    RelKeys1 = Table.AddColumn(FinalSinErrores, "__OC_Key", each CleanOCKey([#"# OC / Contrato"]), type text),
    // Buffer: RelKeys se usa 2 veces (ProvisionesPorOC y CruceNoProvOC) — mismo patron que
    // FiltroTipoValido arriba.
    RelKeys = Table.Buffer(Table.AddColumn(RelKeys1, "__NoProv_Key", each CleanKeyText([No_Prov]), type text)),

    ProvisionesPorOC = Table.Buffer(Table.Group(
        Table.SelectRows(RelKeys, each [__OC_Key] <> null and [__NoProv_Key] <> null),
        {"__OC_Key"},
        {{"No_Prov OC", each
            let nums = List.Sort(List.Distinct(List.RemoveNulls([__NoProv_Key])))
            in if List.IsEmpty(nums) then null else nums{0},
            type text
        }}
    )),

    CruceNoProvOC = Table.NestedJoin(RelKeys, {"__OC_Key"}, ProvisionesPorOC, {"__OC_Key"}, "NPOC", JoinKind.LeftOuter),
    ExpandNoProvOC = Table.ExpandTableColumn(CruceNoProvOC, "NPOC", {"No_Prov OC"}, {"No_Prov OC"}),
    AddNoProvFiltro = Table.AddColumn(ExpandNoProvOC, "No_Prov Filtro", each
        if [__NoProv_Key] <> null and [__NoProv_Key] <> "" then [__NoProv_Key] else [No_Prov OC],
        type text
    ),
    FinalConNoProvFiltro = Table.RemoveColumns(AddNoProvFiltro, {"__OC_Key", "__NoProv_Key", "No_Prov OC"}, MissingField.Ignore),
    FinalOCLimpia = Table.TransformColumns(FinalConNoProvFiltro, {
        {"# OC / Contrato", each
            let oc = try Text.Trim(Text.From(_)) otherwise ""
            in if oc = "" then "SIN OC / CONTRATO" else oc,
            type text
        }
    }, null, MissingField.Ignore),

    // ============================================================
    // RECORTE DE COLUMNAS (SOLO LA ARBOLEDA): las 5 tablas dinamicas de
    // este libro especificamente solo usan 31 de las 55 columnas de BD.
    // Escribir menos columnas a la hoja reduce proporcionalmente el costo
    // de escritura de celdas + reconstruccion de pivotCache (el cuello de
    // botella dominante para este proyecto, el mas grande de todos).
    // Lista confirmada y ajustada por el usuario el 2026-08-06.
    // OJO: este recorte es especifico de La Arboleda - NO aplicar el mismo
    // archivo a otros proyectos sin revisar que sus dinamicas usen las
    // mismas columnas.
    // ============================================================
    ColumnasFinalesArboleda = {
        "Centro de Costos", "Ins", "Actividad", "Capitulo", "Subcapitulo", "Tipo",
        "# OC / Contrato", "#ENTRADA", "#SALIDA", "Descripcion contrato", "# CC - Comparativo",
        "Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido",
        "Cantidad Presupuesto", "VT Presupuesto",
        "Cant. aprobacion", "V/U aprobacion", "VR total aprobacion",
        "Valor Total ppto (CC)", "VT Cortes", "Valor descuento",
        "Cantidad CC Cons", "V/U CC cons", "VT CC cons",
        "VR_Bruto_con_desc", "No_Prov", "Nombre Contratista",
        "VT Asegurada", "VT Proyectado Colsubsidio"
    },
    FinalRecortada = Table.SelectColumns(FinalOCLimpia, ColumnasFinalesArboleda, MissingField.UseNull),

    TablaMaestraFinal = Table.Buffer(FinalRecortada)
in
    TablaMaestraFinal
