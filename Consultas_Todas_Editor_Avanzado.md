# Consultas Power Query - Editor avanzado

Documento GENERADO desde los archivos .m de /Consultas (no editar a mano: editar el .m y regenerar). Copia cada bloque en la consulta con el mismo nombre.

## APROBACIONES_SP

```powerquery
let
    // ============================================================
    // FUNCIONES GLOBALES
    // ============================================================
    FxToNumberFlex = F_Globales[FxToNumberFlex],
    FnMatchFolder = F_Globales[FnMatchFolder],
    FnReadSPExcel = F_Globales[FnReadSPExcel],
    FnBuildFolderPrefixMap = F_Globales[FnBuildFolderPrefixMap],
    FnTrimText = F_Globales[FnTrimText],
    FnNormalizeSpaces = F_Globales[FnNormalizeSpaces],
    FnRemoveAccentsSymbols = F_Globales[FnRemoveAccentsSymbols],

    ParamProyecto = Text.Trim(ProyectoActual),

    // ============================================================
    // CARPETAS DE PROYECTO ACTUAL (sin filtro de archivos)
    // ============================================================
    ListaCarpetas = try List.Distinct(SP_CarpetasCC[Centro de Costos]) otherwise {},
    PrefixMap = FnBuildFolderPrefixMap(ListaCarpetas),

    // ============================================================
    // CONEXION AL ARCHIVO EN SHAREPOINT
    // ============================================================
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    FilePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. CC approvals (solo aprobado) - Control costos interno/0. CONSOLIDADOR APROBACIONES CC SP.xlsx",

    Origen = let t = FnReadSPExcel(SiteUrl, FilePath) in if t = null then #table({"Proyecto:"}, {}) else t,

    // ============================================================
    // FILTRO Y MAPEO
    // ============================================================
    ParamProyectoClean = FnRemoveAccentsSymbols(Text.Upper(ParamProyecto)),
    FiltroProyecto = Table.SelectRows(Origen, each
        try
            let
                proyectoRaw = [#"Proyecto:"],
                proyectoClean = FnRemoveAccentsSymbols(Text.Upper(Text.Trim(Text.From(proyectoRaw))))
            in
                if ParamProyectoClean = "PAYANDE" then
                    Text.StartsWith(proyectoClean, "PAYANDE") and
                    (Text.Contains(proyectoClean, "URB INTERNO") or Text.Contains(proyectoClean, "TORRES"))
                else
                    Text.StartsWith(proyectoClean, ParamProyectoClean) or Text.Contains(proyectoClean, ParamProyectoClean)
        otherwise false
    ),

    ColumnasRenombradas = Table.RenameColumns(FiltroProyecto, {
        {"Desc. - UM", "Ins"},
        {"Nombre del proveedor: ", "Nombre Contratista"},
        {"# CC", "# CC - Comparativo"},
        {"Cant. Total", "Cantidad CC Cons"},
        {"V/U TOTAL", "V/U CC cons"},
        {"VR TOTAL", "VT CC cons"}
    }, MissingField.Ignore),

    TextosLimpios = Table.TransformColumns(ColumnasRenombradas, {
        {"Ins", each FnTrimText(_), type text},
        {"Nombre Contratista", each FnTrimText(_), type text},
        // Clave de cruce contra Det_CC/COMPARATIVOS: normalizar espacios dobles/duros que
        // vienen del consolidador de SharePoint, no solo Trim de extremos
        {"# CC - Comparativo", each FnNormalizeSpaces(_), type text},
        {"Cantidad CC Cons", each FxToNumberFlex(_), type number},
        {"V/U CC cons", each FxToNumberFlex(_), type number},
        {"VT CC cons", each FxToNumberFlex(_), type number}
    }, null, MissingField.Ignore),

    AgregadoTipo = Table.AddColumn(TextosLimpios, "Tipo", each "CC Consolidado", type text),
    AgregadoCC = Table.AddColumn(AgregadoTipo, "Centro de Costos", each
        if Text.StartsWith(Text.Upper(ParamProyecto), "PAMPLONA 1") and [#"# CC - Comparativo"] <> null then
            let
                ccPrefix = Text.Trim(Text.BeforeDelimiter([#"# CC - Comparativo"], "-")),
                fromMap = try Record.Field(PrefixMap, ccPrefix) otherwise null
            in
                if fromMap <> null then fromMap else FnMatchFolder([#"Proyecto:"], ListaCarpetas)
        else
            FnMatchFolder([#"Proyecto:"], ListaCarpetas)
    , type text),

    // ============================================================
    // EXTRACCION DE COLUMNAS PARA BD
    // ============================================================
    TablaFinal = Table.SelectColumns(AgregadoCC,
        {
            "Centro de Costos",
            "Tipo",
            "Ins",
            "Nombre Contratista",
            "# CC - Comparativo",
            "Cantidad CC Cons",
            "V/U CC cons",
            "VT CC cons"
        },
        MissingField.Ignore
    )
in
    TablaFinal
```

## BD

```powerquery
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

    // REGLA SIMPLE (2026-09-02, definida por el usuario): la Proyeccion Colsubsidio es
    // lo asegurado (contratos + ordenes de compra) mas el presupuesto de lo que FALTA
    // por adjudicar. "Falta" = la actividad NO tiene nada asegurado todavia: si la
    // actividad ya tiene contratos/OC (asegurado > 0), su fila POR ADJUDICAR proyecta 0
    // (evita el doble conteo ppto + asegurado en items ya adjudicados). Sin escenarios:
    // el "motor" de banderas quedo sin uso (M es perezoso, no se evalua) y puede
    // retirarse en una limpieza futura.
    AsegPorActividad = Table.Buffer(Table.Group(
        Table.SelectRows(AddVUAseg, each [Tipo] = "CONTRATO" or [Tipo] = "COMPRAS"),
        {"Centro de Costos", "Codigo act"},
        {{"vtAsegAct", each List.Sum(List.Transform([VT Asegurada], each ToNumber0(_))), type number}})),
    ConAsegActividad = Table.ExpandTableColumn(
        Table.NestedJoin(AddVUAseg, {"Centro de Costos", "Codigo act"}, AsegPorActividad, {"Centro de Costos", "Codigo act"}, "__AA", JoinKind.LeftOuter),
        "__AA", {"vtAsegAct"}),
    AplicarProyeccion = Table.AddColumn(ConAsegActividad, "VT Proyectado Colsubsidio", each
        if [Tipo] = "POR ADJUDICAR" then (if ToNumber0([vtAsegAct]) > 0 then 0 else [#"Valor Total ppto (CC)"])
        else if [Tipo] = "CONTRATO" or [Tipo] = "COMPRAS" then (if [VT Asegurada] <> 0 then [VT Asegurada] else null)
        else null,
    type number),

    FinalClean = Table.RemoveColumns(AplicarProyeccion, ColsBanderas & {"vtAsegAct"}, MissingField.Ignore),
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
        "VT Asegurada", "VT Proyectado Colsubsidio",
        // Agregadas 2026-08-27: las calcula AddCantAseg/AddVUAseg mas arriba pero se
        // perdian aqui en el recorte final. Las necesita SINCO.m (Bosque de Turpial)
        // para armar Cantidad/V.U. asegurada = Cantidad y VT Contratado + Comprado.
        "Cantidad asegurada", "V/U asegurada"
    },
    FinalRecortada = Table.SelectColumns(FinalOCLimpia, ColumnasFinalesArboleda, MissingField.UseNull),

    TablaMaestraFinal = Table.Buffer(FinalRecortada)
in
    TablaMaestraFinal
```

## COMPARATIVOS

```powerquery
let
    // ============================================================
    // FUNCIONES AUXILIARES GLOBALES (Desde F_Globales)
    // ============================================================
    FnFormatCodigoAct = F_Globales[FnFormatCodigoAct],
    FxToNumberFlex = F_Globales[FxToNumberFlex],
    FnNormalizeSpaces = F_Globales[FnNormalizeSpaces],

    // ============================================================
    // PROCESAMIENTO DE LA TABLA MANUAL (Det_CC sin columnas PPTO)
    // ============================================================
    Origen = Excel.CurrentWorkbook(){[Name="Det_CC"]}[Content],

    // 1. FILTRO DE FILAS VACÍAS
    FilasValidas = Table.SelectRows(Origen, each [Ins] <> null or [Actividad] <> null),

    // 2. EXTRAER Y ESTANDARIZAR CÓDIGO DE ACTIVIDAD
    AgregadoCodAct = Table.AddColumn(
        FilasValidas, 
        "Codigo act", 
        each 
            let
                txt = Text.Trim(Text.From(if [Actividad] = null then "" else [Actividad])),
                cod = Text.BeforeDelimiter(txt, "-", 0)
            in 
                if txt = "" then null else FnFormatCodigoAct(cod), 
        type text
    ),

    // 3. ETIQUETA DE TIPO
    AgregadoTipo = Table.AddColumn(AgregadoCodAct, "Tipo", each "CC", type text),

    // 4. LIMPIEZA Y TIPOS DE DATOS ROBUSTOS
    TextosLimpios = Table.TransformColumns(AgregadoTipo, {
        {"Centro de Costos", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Ins", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Actividad", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Capitulo", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Subcapitulo", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"# OC / Contrato", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Nombre Contratista", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        // Claves de cruce contra APROBACIONES_SP: normalizar espacios (dobles/duros), no solo Trim
        {"# CC - Comparativo", each FnNormalizeSpaces(_), type text},
        {"# CC", each FnNormalizeSpaces(_), type text},
        {"Comparativo", each FnNormalizeSpaces(_), type text},
        {"Clasificador", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        
        {"Cant. aprobacion", each FxToNumberFlex(_), type number},
        {"V/U aprobacion", each FxToNumberFlex(_), type number},
        {"VR total aprobacion", each FxToNumberFlex(_), type number}
    }, null, MissingField.Ignore),

    TiposFinales = try Table.TransformColumnTypes(TextosLimpios, {{"Codigo ins", Int64.Type}}) otherwise TextosLimpios,

    // 5. SELECCIÓN Y ORDEN FINAL DE COLUMNAS
    TablaFinal = Table.SelectColumns(TiposFinales, 
        {"Centro de Costos", "Subcapitulo", "Capitulo", "Actividad", "Codigo ins", "Ins", 
         "# OC / Contrato", "Nombre Contratista", "Cant. aprobacion", "V/U aprobacion", 
         "VR total aprobacion", "# CC - Comparativo", "# CC", "Comparativo", "Clasificador",
         "Codigo act", "Tipo"}, MissingField.Ignore)
in
    TablaFinal
```

## COMPRAS

```powerquery
let
    // ============================================================
    // FUNCIONES AUXILIARES GLOBALES
    // ============================================================
    FnFormatCodigoAct = F_Globales[FnFormatCodigoAct],
    FnDecodeHtml = F_Globales[FnDecodeHtml],
    FnPrepareTableWithHeader = F_Globales[FnPrepareTableWithHeader],
    FxToNumberFlex = F_Globales[FxToNumberFlex],
    FnClaveLimpia = F_Globales[FnClaveLimpia],
    FnMapColumn = F_Globales[FnMapColumn],
    FnReadSPBinary = F_Globales[FnReadSPBinary],
    FnRemoveAccentMarks = F_Globales[FnRemoveAccentMarks],
    FnRemoveAccentsSymbols = F_Globales[FnRemoveAccentsSymbols],
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    Columnas_OC = F_Globales[FnBuildColumnas](10),
    Columnas_Entradas = F_Globales[FnBuildColumnas](10),
    Columnas_Salidas = F_Globales[FnBuildColumnas](12),

    ColumnasBase = {
        "Codigo ins", "Ins", "Actividad", "Codigo act", "InsClave", "# OC / Contrato",
        "Cantidad Comprado", "VT Comprado", "VU_Crudo", "IVA_Crudo", "Nombre Contratista",
        "#ENTRADA", "Cantidad Cortes", "VT Cortes", "#SALIDA", "Cantidad Cons Cols", "VT Cons Cols"
    },

    EmptyCompras = #table(ColumnasBase, {}),

    FnText = (v as any) as text =>
        try Text.Trim(Text.From(if v = null then "" else v)) otherwise "",

    FnCleanDisplay = (v as any) as nullable text =>
        let
            t = FnText(v),
            clean = if t = "" then null else Text.Upper(Text.Trim(FnRemoveAccentMarks(t)))
        in
            clean,

    FnBuildInsUM = (desc as any, um as any) as nullable text =>
        let
            d = FnCleanDisplay(desc),
            u = FnCleanDisplay(um)
        in
            if d = null then null else if u = null or u = "" then d else d & " (" & u & ")",

    FnCleanContratistaFromDash = (v as any) as nullable text =>
        let
            t = FnText(v),
            afterDash = if Text.Contains(t, "-") then Text.Trim(Text.AfterDelimiter(t, "-")) else t,
            clean = FnCleanDisplay(afterDash)
        in
            clean,

    FnRenameSequential = (tbl as table) as table =>
        let
            cols = Table.ColumnNames(tbl),
            renamed = Table.RenameColumns(tbl, List.Zip({cols, List.Transform({1..List.Count(cols)}, each "Columna" & Text.From(_))}))
        in
            renamed,

    // ============================================================
    // PROCESAR INFORMEORDEN + ESTADO DE ORDENES
    // ============================================================
    FxProcesarCompras = (BinDetalles as binary, BinOC as binary) as table => let
        RawOC_Raw = try Excel.Workbook(BinOC, null, true){0}[Data]
                otherwise Html.Table(FnDecodeHtml(BinOC), Columnas_OC, [RowSelector="tr"]),
        RawOC = FnRenameSequential(RawOC_Raw),

        AddOCKey = Table.AddColumn(RawOC, "OC_Key_Temp", each let v = FnText([Columna1]) in if Text.StartsWith(v, "Orden de Compra No.") then Text.Trim(Text.Replace(v, "Orden de Compra No.", "")) else null, type text),
        Ordenes_Agrupadas = Table.RenameColumns(Table.Group(Table.SelectRows(Table.FillDown(AddOCKey, {"OC_Key_Temp"}), each [OC_Key_Temp] <> null), {"OC_Key_Temp"}, {{"Proveedor_Raw", each let l = List.RemoveNulls([Columna2]), l2 = List.Select(l, (x) => let t = FnText(x) in t <> "Proveedor" and t <> "Insumo") in if List.IsEmpty(l2) then null else List.First(l2), type text}}), {{"OC_Key_Temp", "OC_Key"}}),

        LibroExcel = Excel.Workbook(Binary.Buffer(BinDetalles), null, true),
        DetallesCrudos = FnPrepareTableWithHeader(LibroExcel{0}[Data]),
        Cols = Table.ColumnNames(DetallesCrudos),
        // FnMapColumn no depende de la fila, solo de Cols+keywords (siempre fijos aqui) —
        // resolver el nombre de columna UNA sola vez en vez de re-buscarlo en cada fila
        // (antes: 8 busquedas completas de texto por fila x N filas, todas con el mismo resultado).
        // Misma logica de match que FnMapColumn (F_Globales), pero devolviendo el NOMBRE de
        // columna en vez del valor, para poder resolverla una sola vez fuera del loop por fila.
        FnResolverCol = (keywords as list) as nullable text =>
            let
                norm = (x as any) as text =>
                    let
                        txt = try Text.From(x) otherwise "",
                        clean = FnRemoveAccentsSymbols(txt)
                    in Text.Upper(if clean = null then "" else clean),
                match = List.First(List.Select(Cols, (c) => List.AnyTrue(List.Transform(keywords, (k) => Text.Contains(norm(c), norm(k))))), null)
            in match,
        ColCodigoIns = FnResolverCol({"CÓDIGO", "CODIGO", "COD."}),
        ColIns = FnResolverCol({"INSUMO", "DESCRIPCIÓN", "DESCRIPCION"}),
        ColAct = FnResolverCol({"ACTIVIDAD", "DESTINO", "FRENTE", "ITEM", "ÍTEM"}),
        ColCant = FnResolverCol({"CANTIDAD", "CANT."}),
        ColVU = FnResolverCol({"VALOR UNITARIO", "VLR UNIT", "UNITARIO"}),
        ColIVA = FnResolverCol({"IVA %", "IVA", "% IVA"}),
        ColVT = FnResolverCol({"VALOR TOTAL", "VLR TOTAL", "TOTAL"}),
        ColOC = FnResolverCol({"ORDEN", "PEDIDO", "O.C"}),
        MapStd = Table.AddColumn(DetallesCrudos, "Std", each [
            Codigo_ins = if ColCodigoIns = null then null else Record.Field(_, ColCodigoIns),
            Ins = if ColIns = null then null else Record.Field(_, ColIns),
            Act = if ColAct = null then null else Record.Field(_, ColAct),
            Cant = if ColCant = null then null else Record.Field(_, ColCant),
            VU_Crudo = try Record.FieldValues(_){10} otherwise (if ColVU = null then null else Record.Field(_, ColVU)),
            IVA_Crudo = try Record.FieldValues(_){11} otherwise (if ColIVA = null then null else Record.Field(_, ColIVA)),
            VT = try Record.FieldValues(_){12} otherwise (if ColVT = null then null else Record.Field(_, ColVT)),
            OC = if ColOC = null then null else Record.Field(_, ColOC)
        ]),
        DetallesStd = Table.ExpandRecordColumn(MapStd, "Std", {"Codigo_ins", "Ins", "Act", "Cant", "VT", "VU_Crudo", "IVA_Crudo", "OC"}, {"Codigo ins", "Ins", "Actividad", "Cantidad Comprado", "VT Comprado", "VU_Crudo", "IVA_Crudo", "# OC / Contrato"}),
        DetConKeyOC = Table.AddColumn(DetallesStd, "OC_Key", each FnText([#"# OC / Contrato"]), type text),
        DetConCodAct = Table.AddColumn(DetConKeyOC, "Codigo act", each let c = Text.Trim(Text.BeforeDelimiter(FnText([Actividad]), "-", 0)) in if c = "" then null else c, type text),
        DetConClave = Table.AddColumn(DetConCodAct, "InsClave", each FnClaveLimpia([Ins]), type text),
        MergedOC = Table.NestedJoin(DetConClave, {"OC_Key"}, Ordenes_Agrupadas, {"OC_Key"}, "ORD", JoinKind.LeftOuter),
        ExpandedOC = Table.ExpandTableColumn(MergedOC, "ORD", {"Proveedor_Raw"}, {"Proveedor_Raw"}),
        AddedNombreContratista = Table.AddColumn(ExpandedOC, "Nombre Contratista", each FnCleanContratistaFromDash([Proveedor_Raw]), type text),
        Selected = Table.SelectColumns(AddedNombreContratista, ColumnasBase, MissingField.UseNull)
    in Selected,

    // ============================================================
    // PROCESAR INFORME ENTRADAS DE ALMACEN DETALLADAS
    // ============================================================
    FxProcesarEntradas = (BinEntradas as binary) as table => let
        Raw_Raw = try Excel.Workbook(Binary.Buffer(BinEntradas), null, true){0}[Data]
                otherwise Html.Table(FnDecodeHtml(BinEntradas), Columnas_Entradas, [RowSelector="tr"]),
        // FillDown O(N): se marca cada fila de encabezado una sola vez y se arrastra
        // hacia abajo, en lugar de re-escanear toda la tabla por cada fila (O(N^2)).
        Raw = Table.Buffer(FnRenameSequential(Raw_Raw)),
        ConFlagMeta = Table.AddColumn(Raw, "__esMeta", each
            FnText([Columna1]) <> "" and FnText([Columna2]) <> "" and FnText([Columna4]) <> "" and
            (try Date.From([Columna1]) otherwise null) <> null and
            (try Number.FromText(FnText([Columna2])) otherwise null) <> null and
            (try Number.FromText(FnText([Columna4])) otherwise null) <> null, type logical),
        ConMetaRec = Table.AddColumn(ConFlagMeta, "__meta", each if [__esMeta] then _ else null, type nullable record),
        WithMeta = Table.FillDown(ConMetaRec, {"__meta"}),
        Items = Table.SelectRows(WithMeta, each
            let
                cod = FnText([Columna1]),
                codNum = try Number.FromText(cod) otherwise null,
                ins = FnText([Columna2]),
                cant = FxToNumberFlex([Columna4]),
                vt = FxToNumberFlex([Columna7])
            in
                codNum <> null and ins <> "" and (cant <> null or vt <> null)
        ),
        AddStd = Table.AddColumn(Items, "Std", each
            let
                m = [__meta],
                entrada = if m = null then null else FnText(Record.Field(m, "Columna2")),
                oc = if m = null then null else FnText(Record.Field(m, "Columna4")),
                proveedor = if m = null then null else Record.Field(m, "Columna3"),
                insFinal = FnBuildInsUM([Columna2], [Columna3])
            in [
                #"Codigo ins" = FnText([Columna1]),
                Ins = insFinal,
                Actividad = null,
                #"Codigo act" = null,
                InsClave = FnClaveLimpia(insFinal),
                #"# OC / Contrato" = oc,
                #"Cantidad Comprado" = null,
                #"VT Comprado" = null,
                VU_Crudo = null,
                IVA_Crudo = null,
                #"Nombre Contratista" = FnCleanContratistaFromDash(proveedor),
                #"#ENTRADA" = entrada,
                #"Cantidad Cortes" = FxToNumberFlex([Columna4]),
                #"VT Cortes" = FxToNumberFlex([Columna7]),
                #"#SALIDA" = null,
                #"Cantidad Cons Cols" = null,
                #"VT Cons Cols" = null
            ]),
        Expanded = Table.ExpandRecordColumn(AddStd, "Std", ColumnasBase, ColumnasBase),
        Selected = Table.SelectColumns(Expanded, ColumnasBase, MissingField.UseNull)
    in Selected,

    // ============================================================
    // PROCESAR MASIVO SALIDAS DETALLADO
    // ============================================================
    FxProcesarSalidas = (BinSalidas as binary) as table => let
        Raw_Raw = Html.Table(FnDecodeHtml(BinSalidas), Columnas_Salidas, [RowSelector="tr"]),
        // FillDown O(N): mismo patron que en Entradas para evitar el re-escaneo O(N^2).
        Raw = Table.Buffer(FnRenameSequential(Raw_Raw)),
        ConFlagMeta = Table.AddColumn(Raw, "__esMeta", each
            FnText([Columna1]) <> "" and FnText([Columna2]) <> "" and FnText([Columna3]) <> "" and
            (try Date.From([Columna1]) otherwise null) <> null and
            (try Number.FromText(FnText([Columna2])) otherwise null) <> null, type logical),
        ConMetaRec = Table.AddColumn(ConFlagMeta, "__meta", each if [__esMeta] then _ else null, type nullable record),
        WithMeta = Table.FillDown(ConMetaRec, {"__meta"}),
        Items = Table.SelectRows(WithMeta, each
            let
                cod = FnText([Columna1]),
                codNum = try Number.FromText(cod) otherwise null,
                ins = FnText([Columna2]),
                cant = FxToNumberFlex([Columna5]),
                vt = FxToNumberFlex([Columna8])
            in
                codNum <> null and ins <> "" and (cant <> null or vt <> null)
        ),
        AddStd = Table.AddColumn(Items, "Std", each
            let
                m = [__meta],
                salida = if m = null then null else FnText(Record.Field(m, "Columna2")),
                contratista = if m = null then null else Record.Field(m, "Columna3"),
                insFinal = FnBuildInsUM([Columna2], [Columna4]),
                codAct = FnFormatCodigoAct([Columna3])
            in [
                #"Codigo ins" = FnText([Columna1]),
                Ins = insFinal,
                Actividad = null,
                #"Codigo act" = codAct,
                InsClave = FnClaveLimpia(insFinal),
                #"# OC / Contrato" = null,
                #"Cantidad Comprado" = null,
                #"VT Comprado" = null,
                VU_Crudo = null,
                IVA_Crudo = null,
                #"Nombre Contratista" = FnCleanContratistaFromDash(contratista),
                #"#ENTRADA" = null,
                #"Cantidad Cortes" = null,
                #"VT Cortes" = null,
                #"#SALIDA" = salida,
                #"Cantidad Cons Cols" = FxToNumberFlex([Columna5]),
                #"VT Cons Cols" = FxToNumberFlex([Columna8])
            ]),
        Expanded = Table.ExpandRecordColumn(AddStd, "Std", ColumnasBase, ColumnasBase),
        Selected = Table.SelectColumns(Expanded, ColumnasBase, MissingField.UseNull)
    in Selected,

    // ============================================================
    // CONEXIÓN A SHAREPOINT (LECTURA DESDE CONSULTA COMPARTIDA)
    // ============================================================
    ArchivosProyecto = Table.SelectRows(SP_Archivos_Proyecto, each
        Text.Contains([Name], "INFORMEORDEN", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "ESTADO DE ORDENES", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "INFORME ENTRADAS DE ALMACEN", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "INFORME ENTRADAS DE ALMACÉN", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "ENTRADAS POR INSUMO", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "MASIVO SALIDAS", Comparer.OrdinalIgnoreCase)
    ),
    ConCentroCosto = ArchivosProyecto,

    PickLatestBinary = (t as table, containsText as text) as nullable binary =>
        let
            candidatos = Table.Sort(
                Table.SelectRows(t, each Text.Contains([Name], containsText, Comparer.OrdinalIgnoreCase)),
                {{"TimeLastModified", Order.Descending}, {"Name", Order.Ascending}}
            ),
            path = if Table.RowCount(candidatos) = 0 then null else candidatos{0}[ServerRelativeUrl]
        in
            if path = null then null else FnReadSPBinary(SiteUrl, path),

    PickLatestBinaryAny = (t as table, containsTexts as list) as nullable binary =>
        let
            candidatos = Table.Sort(
                Table.SelectRows(t, each List.AnyTrue(List.Transform(containsTexts, (needle) => Text.Contains([Name], needle, Comparer.OrdinalIgnoreCase)))),
                {{"TimeLastModified", Order.Descending}, {"Name", Order.Ascending}}
            ),
            path = if Table.RowCount(candidatos) = 0 then null else candidatos{0}[ServerRelativeUrl]
        in
            if path = null then null else FnReadSPBinary(SiteUrl, path),

    // === Preferir formato liviano (plano, sin fill-down) si existe en SharePoint,
    // con respaldo automatico al formato "detallado" viejo si el proyecto aun no
    // tiene el reporte nuevo subido. Nunca rompe: si no hay ninguno, retorna null.
    PickPreferido = (t as table, containsLiviano as text, containsViejo as text) as record =>
        let
            candLiviano = Table.Sort(Table.SelectRows(t, each Text.Contains([Name], containsLiviano, Comparer.OrdinalIgnoreCase)), {{"TimeLastModified", Order.Descending}, {"Name", Order.Ascending}}),
            hayLiviano = Table.RowCount(candLiviano) > 0,
            candViejo = Table.Sort(Table.SelectRows(t, each Text.Contains([Name], containsViejo, Comparer.OrdinalIgnoreCase) and not Text.Contains([Name], containsLiviano, Comparer.OrdinalIgnoreCase)), {{"TimeLastModified", Order.Descending}, {"Name", Order.Ascending}}),
            path = if hayLiviano then candLiviano{0}[ServerRelativeUrl] else if Table.RowCount(candViejo) > 0 then candViejo{0}[ServerRelativeUrl] else null,
            bin = if path = null then null else FnReadSPBinary(SiteUrl, path)
        in [Binario = bin, EsLiviano = hayLiviano],

    // === List.Transform en vez de Table.Group + AddColumn anidado ===
    // Confirmado por prueba: el mismo trabajo via Table.Group tardaba 402s vs
    // 133s via List.Transform (3x mas rapido), para el mismo Centro de Costo.
    CCsConArchivos = List.Distinct(ConCentroCosto[Centro de Costos]),
    ResultadosPorCC = List.Transform(CCsConArchivos, (cc) =>
        let
            filtrado = Table.SelectRows(ConCentroCosto, each [Centro de Costos] = cc),
            binDet = PickLatestBinary(filtrado, "INFORMEORDEN"),
            binOC = PickLatestBinary(filtrado, "ESTADO DE ORDENES"),
            resEntradas = PickPreferido(filtrado, "ENTRADAS POR INSUMO", "INFORME ENTRADAS DE ALMACEN"),
            resSalidas = PickPreferido(filtrado, "DESCRIPTIVAS", "MASIVO SALIDAS"),
            tCompras = if binDet <> null and binOC <> null then FxProcesarCompras(binDet, binOC) else EmptyCompras,
            tEntradas = if resEntradas[Binario] = null then EmptyCompras
                        else if resEntradas[EsLiviano] then F_Globales[FxProcesarEntradasPorInsumo](resEntradas[Binario], ColumnasBase)
                        else FxProcesarEntradas(resEntradas[Binario]),
            tSalidas = if resSalidas[Binario] = null then EmptyCompras
                       else if resSalidas[EsLiviano] then F_Globales[FxProcesarSalidasDescriptivas](resSalidas[Binario], ColumnasBase)
                       else FxProcesarSalidas(resSalidas[Binario]),
            combinado = Table.Combine({tCompras, tEntradas, tSalidas}),
            conCC = Table.AddColumn(combinado, "Centro de Costos", each cc, type text)
        in
            conCC
    ),
    Expandido = Table.Combine(ResultadosPorCC),

    Expandido_Clean = Table.TransformColumns(Expandido, {
        {"Centro de Costos", each if _ = null then null else Text.Upper(Text.Trim(Text.From(_))), type text},
        {"Codigo act", each FnFormatCodigoAct(_), type text}
    }, null, MissingField.Ignore),

    Compras_Unicas = Table.Buffer(Table.Distinct(Expandido_Clean)),

    // ============================================================
    // LECTURA DIRECTA DE CONSULTAS (Memoria)
    // ============================================================
    ITEMS_TablaLocal = ITEMSINSUMOS,
    ITEMS_Clean = Table.TransformColumns(ITEMS_TablaLocal, {
        {"Centro de Costos", each if _ = null then null else Text.Upper(Text.Trim(Text.From(_))), type text},
        {"Codigo act", each FnFormatCodigoAct(_), type text}
    }, null, MissingField.Ignore),
    // Buffer: ITEMS_Base se usa 3 veces (ITEMS_Insumos_Dist, ItemsPorCodigo_Estricto,
    // ItemsPorCodigo_Generico) — sin buffer, cada uno recalcula TransformColumns+SelectColumns
    // desde cero en vez de reusar el resultado ya calculado.
    ITEMS_Base = Table.Buffer(Table.SelectColumns(ITEMS_Clean, {"Centro de Costos", "Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo"})),

    ITEMS_Insumos_Dist = Table.Buffer(Table.Distinct(Table.AddColumn(ITEMS_Base, "InsClave", each FnClaveLimpia([Ins]), type text), {"Centro de Costos", "Codigo act", "InsClave"})),
    ITEMS_Respaldo = Table.Buffer(Table.Distinct(ITEMS_Insumos_Dist, {"Centro de Costos", "InsClave"})),

    ItemsPorCodigo_Estricto = Table.Buffer(Table.Group(ITEMS_Base, {"Centro de Costos", "Codigo act"}, {
        {"Act_Estricto", each List.First(List.RemoveNulls([Actividad])), type text},
        {"Cap_Estricto", each List.First(List.RemoveNulls([Capitulo])), type text},
        {"Sub_Estricto", each List.First(List.RemoveNulls([Subcapitulo])), type text}
    })),

    ItemsPorCodigo_Generico = Table.Buffer(Table.Group(ITEMS_Base, {"Codigo act"}, {
        {"Act_Gen", each List.First(List.RemoveNulls([Actividad])), type text},
        {"Cap_Gen", each List.First(List.RemoveNulls([Capitulo])), type text},
        {"Sub_Gen", each List.First(List.RemoveNulls([Subcapitulo])), type text}
    })),

    // ============================================================
    // CRUCES FINALES
    // ============================================================
    MergedExacto = Table.NestedJoin(Compras_Unicas, {"Centro de Costos", "Codigo act", "InsClave"}, ITEMS_Insumos_Dist, {"Centro de Costos", "Codigo act", "InsClave"}, "EXACTO", JoinKind.LeftOuter),
    ExpandedExacto = Table.ExpandTableColumn(MergedExacto, "EXACTO", {"Ins"}, {"Ex.Ins"}),

    MergedRescate = Table.NestedJoin(ExpandedExacto, {"Centro de Costos", "InsClave"}, ITEMS_Respaldo, {"Centro de Costos", "InsClave"}, "RESCATE", JoinKind.LeftOuter),
    ExpandedRescate = Table.ExpandTableColumn(MergedRescate, "RESCATE", {"Codigo act", "Actividad", "Ins"}, {"Rs.Codigo act", "Rs.Actividad", "Rs.Ins"}),

    MergedEstricto = Table.NestedJoin(ExpandedRescate, {"Centro de Costos", "Codigo act"}, ItemsPorCodigo_Estricto, {"Centro de Costos", "Codigo act"}, "EST", JoinKind.LeftOuter),
    ExpandedEstricto = Table.ExpandTableColumn(MergedEstricto, "EST", {"Act_Estricto", "Cap_Estricto", "Sub_Estricto"}, {"Act_Estricto", "Cap_Estricto", "Sub_Estricto"}),

    MergedGenerico = Table.NestedJoin(ExpandedEstricto, {"Codigo act"}, ItemsPorCodigo_Generico, {"Codigo act"}, "GEN", JoinKind.LeftOuter),
    ExpandedGenerico = Table.ExpandTableColumn(MergedGenerico, "GEN", {"Act_Gen", "Cap_Gen", "Sub_Gen"}, {"Act_Gen", "Cap_Gen", "Sub_Gen"}),

    AddedCoalesced = Table.AddColumn(ExpandedGenerico, "FinalCols", each
        let
            e = [Ex.Ins] <> null,
            ca = if [#"#ENTRADA"] <> null then null else if e then [Codigo act] else (if [Rs.Codigo act] <> null then [Rs.Codigo act] else [Codigo act]),
            a0 = FnText([Actividad]),
            aOrig = if a0 = "" then null else if ca <> null and not Text.StartsWith(a0, Text.From(ca)) then Text.From(ca) & " - " & a0 else a0,

            ActOficial = if [#"#ENTRADA"] <> null then null else if [Act_Estricto] <> null then [Act_Estricto] else if [Act_Gen] <> null then [Act_Gen] else aOrig,
            CapFinal = if [#"#ENTRADA"] <> null then null else if [Cap_Estricto] <> null then [Cap_Estricto] else [Cap_Gen],
            SubCapFinal = if [#"#ENTRADA"] <> null then null else if [Sub_Estricto] <> null then [Sub_Estricto] else [Sub_Gen]
        in [
            InsFinal = if e then [Ex.Ins] else (if [Rs.Ins] <> null then [Rs.Ins] else [Ins]),
            CodActFinal = ca,
            ActFinal = ActOficial,
            CapFinal = CapFinal,
            SubCapFinal = SubCapFinal
        ]),

    ExpandedFinalCols = Table.ExpandRecordColumn(Table.RemoveColumns(AddedCoalesced, {"Ins", "Actividad", "Codigo act", "Ex.Ins", "Rs.Codigo act", "Rs.Actividad", "Rs.Ins", "Act_Estricto", "Cap_Estricto", "Sub_Estricto", "Act_Gen", "Cap_Gen", "Sub_Gen"}), "FinalCols", {"InsFinal", "CodActFinal", "ActFinal", "CapFinal", "SubCapFinal"}, {"Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo"}),

    NumericColumns = Table.TransformColumns(ExpandedFinalCols, {
        {"#ENTRADA", each FxToNumberFlex(_), Int64.Type},
        {"#SALIDA", each FxToNumberFlex(_), Int64.Type},
        {"Cantidad Comprado", each FxToNumberFlex(_), type number},
        {"VT Comprado", each FxToNumberFlex(_), type number},
        {"VU_Crudo", each FxToNumberFlex(_), type number},
        {"IVA_Crudo", each FxToNumberFlex(_), type number},
        {"Cantidad Cortes", each FxToNumberFlex(_), type number},
        {"VT Cortes", each FxToNumberFlex(_), type number},
        {"Cantidad Cons Cols", each FxToNumberFlex(_), type number},
        {"VT Cons Cols", each FxToNumberFlex(_), type number}
    }),
    Added_VU = Table.AddColumn(NumericColumns, "V/U Comprado", each let vb = [VU_Crudo], iva = [IVA_Crudo], p = if iva = null then 0 else if iva >= 1 then iva / 100 else iva, vc = if vb = null then null else vb * (1 + p) in if vc = null then null else Number.Round(vc, 0), type number),
    FilteredZeros = Table.SelectRows(Added_VU, each try
        ([VT Comprado] <> null and [VT Comprado] <> 0) or
        ([VT Cortes] <> null and [VT Cortes] <> 0) or
        ([VT Cons Cols] <> null and [VT Cons Cols] <> 0)
    otherwise false),

    SelectedFinal = Table.SelectColumns(Table.AddColumn(Table.AddColumn(FilteredZeros, "Tipo", each "COMPRAS", type text), "Descripcion contrato", each "pedido obra", type text), {
        "Centro de Costos", "Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo",
        "# OC / Contrato", "#ENTRADA", "#SALIDA", "Nombre Contratista", "Descripcion contrato",
        "Cantidad Comprado", "VT Comprado", "V/U Comprado", "Cantidad Cortes", "VT Cortes", "Cantidad Cons Cols", "VT Cons Cols", "Tipo"
    }, MissingField.Ignore),
    TypedFinal = Table.TransformColumnTypes(SelectedFinal, {
        {"Centro de Costos", type text}, {"Codigo ins", Int64.Type}, {"Cantidad Comprado", type number},
        {"#ENTRADA", Int64.Type}, {"#SALIDA", Int64.Type}, {"VT Comprado", type number}, {"V/U Comprado", type number}, {"Cantidad Cortes", type number},
        {"VT Cortes", type number}, {"Cantidad Cons Cols", type number}, {"VT Cons Cols", type number}
    }),
    TablaFinal = TypedFinal
in
    TablaFinal
```

## CONTRATOS

```powerquery
let
    // ============================================================
    // 1. FUNCIONES AUXILIARES GLOBALES
    // ============================================================
    FnFormatCodigoAct = F_Globales[FnFormatCodigoAct],
    FnDecodeHtml = F_Globales[FnDecodeHtml],
    FxToNumberFlex = F_Globales[FxToNumberFlex],
    FnClaveLimpia = F_Globales[FnClaveLimpia],
    FnReadSPBinary = F_Globales[FnReadSPBinary],
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    Columnas_HTML = F_Globales[FnBuildColumnas](15),

    // ============================================================
    // 2. FUNCIÓN MÁGICA: PROCESAR CORTES
    // ============================================================
    FxProcesarCortes = (BinarioCortes as binary) =>
        let
            // 🚀 Excel.Workbook es más rápido que Html.Table
            Source_Raw = try Excel.Workbook(BinarioCortes, null, true){0}[Data]
                     otherwise Html.Table(FnDecodeHtml(BinarioCortes), Columnas_HTML, [RowSelector="tr"]),
            // Estandarizar nombres de columnas
            Source_ColNames = Table.ColumnNames(Source_Raw),
            Source = Table.RenameColumns(Source_Raw, List.Zip({Source_ColNames, List.Transform({1..List.Count(Source_ColNames)}, each "Columna" & Text.From(_))})),

            AddFilaTexto = Table.AddColumn(Source, "FilaTexto", each let vals = Record.FieldValues(_), soloTexto = List.Transform(List.Select(vals, each _ <> null and _ <> ""), Text.From) in Text.Trim(Text.Combine(soloTexto, " ")), type text),
            AddOC = Table.AddColumn(AddFilaTexto, "# OC / Contrato", each let txt = [FilaTexto] in if txt <> null and Text.Contains(Text.Upper(txt), "CONTRATO NO") then let after = Text.TrimStart(Text.Replace(Text.Range(txt, Text.PositionOf(Text.Upper(txt), "CONTRATO NO") + 11), "#(00A0)", " "), {".", ":", " "}), first = Text.BeforeDelimiter(after, " "), num = Text.Select(if first = "" then after else first, {"0".."9"}) in if num = "" then null else num else null, type text),
            AddDesc = Table.AddColumn(AddOC, "Descripcion contrato", each let txt = [FilaTexto] in if txt <> null and Text.Contains(Text.Upper(txt), "CONTRATO NO") then let after = Text.TrimStart(Text.Range(txt, Text.PositionOf(Text.Upper(txt), "CONTRATO NO") + 11), {".", ":", " "}), idx = Text.PositionOfAny(after, {"A".."Z","a".."z"}), desc = if idx = -1 then null else Text.Range(after, idx), lim = if desc = null then null else if Text.Contains(Text.Upper(desc), "CONTRATISTA") then Text.BeforeDelimiter(Text.Upper(desc), "CONTRATISTA") else desc in if lim = null then null else Text.Trim(lim) else null, type text),
            AddNombre = Table.AddColumn(AddDesc, "Nombre Contratista", each let txt = [FilaTexto] in if txt <> null and Text.Contains(Text.Upper(txt), "CONTRATISTA") then Text.Trim(Text.TrimStart(Text.AfterDelimiter(Text.Upper(txt), "CONTRATISTA"), {":","-"," "})) else null, type text),
            FillDown1 = Table.FillDown(AddNombre, {"# OC / Contrato","Descripcion contrato","Nombre Contratista"}),
            AddCodAct = Table.AddColumn(FillDown1, "CodigoAct", each let c = [Columna1], t = if c = null then null else Text.Trim(Text.From(c)) in if t <> null and t <> "" and (try Number.From(Text.Replace(t, ".", "")) otherwise null) <> null then FnFormatCodigoAct(t) else null, type text),
            AddActFuente = Table.AddColumn(AddCodAct, "ActividadFuente", each if [CodigoAct] <> null then [Columna2] else null, type text),
            FillDown2 = Table.FillDown(AddActFuente, {"CodigoAct", "ActividadFuente"}),
            
            AddCantC = Table.AddColumn(FillDown2, "Cantidades contrato", each FxToNumberFlex([Columna4]), type number),
            AddVTC = Table.AddColumn(AddCantC, "VT contrato", each FxToNumberFlex([Columna5]), type number),
            AddCantCortes = Table.AddColumn(AddVTC, "Cantidad Cortes", each FxToNumberFlex([Columna10]), type number),
            AddNums = Table.AddColumn(AddCantCortes, "VT Cortes", each FxToNumberFlex([Columna11]), type number),
            
            Filtered = Table.SelectRows(AddNums, each 
                [Columna2] <> null and 
                [CodigoAct] <> null and 
                ([Columna1] = null or Text.Trim(Text.From([Columna1])) = "") and 
                not Text.Contains(Text.Upper(Text.From(if [Columna1] = null then "" else [Columna1])), "TOTAL") and
                not Text.Contains(Text.Upper(Text.From(if [Columna2] = null then "" else [Columna2])), "TOTAL")
            ),
            
            // Creamos la clave robusta para el cruce en SINCO
            AddClave = Table.AddColumn(Filtered, "InsClave_Cruce", each FnClaveLimpia([Columna2]), type text)
        in AddClave, 

    // ============================================================
    // CONEXIÓN A SHAREPOINT (LECTURA DESDE CONSULTA COMPARTIDA)
    // ============================================================
    ArchivosProyecto = Table.SelectRows(SP_Archivos_Proyecto, each
        Text.Contains([Name], "ESTADO DE CONTRATOS", Comparer.OrdinalIgnoreCase)
    ),

    PickLatestBinary = (t as table, containsText as text) as nullable binary =>
        let
            candidatos = Table.Sort(
                Table.SelectRows(t, each Text.Contains([Name], containsText, Comparer.OrdinalIgnoreCase)),
                {{"TimeLastModified", Order.Descending}, {"Name", Order.Ascending}}
            ),
            path = if Table.RowCount(candidatos) = 0 then null else candidatos{0}[ServerRelativeUrl]
        in
            if path = null then null else FnReadSPBinary(SiteUrl, path),

    Agrupado = Table.Group(ArchivosProyecto, {"Centro de Costos"}, {{"Binario", each PickLatestBinary(_, "ESTADO DE CONTRATOS")}}),
    CentrosConArchivo = Table.SelectRows(Agrupado, each [Binario] <> null),
    TablaConDatos = Table.AddColumn(CentrosConArchivo, "Datos", each FxProcesarCortes([Binario])),
    SoloDatos = Table.RemoveColumns(TablaConDatos, {"Binario"}),
    
    Expandido = Table.ExpandTableColumn(SoloDatos, "Datos", {"# OC / Contrato", "Descripcion contrato", "Nombre Contratista", "CodigoAct", "ActividadFuente", "Cantidades contrato", "VT contrato", "Cantidad Cortes", "VT Cortes", "Columna2", "InsClave_Cruce"}),
    
    Expandido_Clean = Table.TransformColumns(Expandido, {
        {"Centro de Costos", each if _ = null then null else Text.Upper(Text.Trim(Text.From(_))), type text},
        {"CodigoAct", each FnFormatCodigoAct(_), type text}
    }, null, MissingField.Ignore),

    Expandido_Unico = Table.Buffer(Table.Distinct(Expandido_Clean)),

    // ============================================================
    // 4. LECTURA DE LA CONSULTA MAESTRA (EL DICCIONARIO OFICIAL)
    // ============================================================
    ITEMS_Raw = ITEMSINSUMOS,
    
    ITEMS_Clean = Table.Buffer(Table.TransformColumns(ITEMS_Raw, {
        {"Centro de Costos", each if _ = null then null else Text.Upper(Text.Trim(Text.From(_))), type text},
        {"Codigo act", each FnFormatCodigoAct(_), type text}
    }, null, MissingField.Ignore)),

    ITEMS_Jerarquia = Table.Buffer(Table.Group(ITEMS_Clean, {"Centro de Costos", "Codigo act"}, {
        {"Ref.Act", each List.First(List.RemoveNulls([Actividad])), type text},
        {"Ref.Cap", each List.First(List.RemoveNulls([Capitulo])), type text}, 
        {"Ref.Sub", each List.First(List.RemoveNulls([Subcapitulo])), type text}
    })),

    // 🔥 PREPARAMOS TBITEMS
    ITEMS_Insumos_Base = Table.AddColumn(ITEMS_Clean, "InsClave_Cruce", each FnClaveLimpia([Ins]), type text),
    ITEMS_Insumos_Dist = Table.Buffer(Table.Group(ITEMS_Insumos_Base, {"Centro de Costos", "Codigo act", "InsClave_Cruce"}, {
        {"Ref.InsOficial", each List.First([Ins]), type text},
        {"Ref.CodIns", each List.First([Codigo ins]), type any}
    })),

    // ============================================================
    // 5. CRUCES FINALES Y REEMPLAZO DE NOMBRES
    // ============================================================
    MergedJerarquia = Table.NestedJoin(Expandido_Unico, {"Centro de Costos", "CodigoAct"}, ITEMS_Jerarquia, {"Centro de Costos", "Codigo act"}, "JER", JoinKind.LeftOuter),
    ExpandedJerarquia = Table.ExpandTableColumn(MergedJerarquia, "JER", {"Ref.Act", "Ref.Cap", "Ref.Sub"}, {"Ref.Act", "Ref.Cap", "Ref.Sub"}),
    
    MergedInsumos = Table.NestedJoin(ExpandedJerarquia, {"Centro de Costos", "CodigoAct", "InsClave_Cruce"}, ITEMS_Insumos_Dist, {"Centro de Costos", "Codigo act", "InsClave_Cruce"}, "INS", JoinKind.LeftOuter),
    ExpandedInsumos = Table.ExpandTableColumn(MergedInsumos, "INS", {"Ref.CodIns", "Ref.InsOficial"}, {"Ref.CodIns", "Ref.InsOficial"}),

    AddFinalCols = Table.AddColumn(ExpandedInsumos, "FinalCols", each [
        // Si cruzó, toma el nombre OFICIAL de TbItems. Si no, deja el de SINCO.
        I = if [Ref.InsOficial] <> null then [Ref.InsOficial] else (if [Columna2] = null or Text.Trim([Columna2]) = "" then "SIN DESCRIPCION" else Text.Trim([Columna2])),
        
        A_Original = let a0 = Text.Trim(Text.From(if [ActividadFuente] = null then "" else [ActividadFuente])) in if a0 = "" then null else if [CodigoAct] <> null and not Text.StartsWith(a0, Text.From([CodigoAct])) then Text.From([CodigoAct]) & " - " & a0 else a0,
        A = if [Ref.Act] <> null then [Ref.Act] else A_Original
    ]),
    ExpandedFinal = Table.ExpandRecordColumn(AddFinalCols, "FinalCols", {"I", "A"}, {"Ins", "Actividad"}),
    
    AddCodIns = Table.AddColumn(ExpandedFinal, "Codigo ins_Final", each [Ref.CodIns]),
    
    Selected = Table.SelectColumns(Table.AddColumn(AddCodIns, "Tipo", each "Contrato"), {"Centro de Costos", "Codigo ins_Final", "Ins", "CodigoAct", "Actividad", "Ref.Cap", "Ref.Sub", "# OC / Contrato", "Nombre Contratista", "Descripcion contrato", "Cantidades contrato", "VT contrato", "Cantidad Cortes", "VT Cortes", "Tipo"}),
    Renamed = Table.RenameColumns(Selected, {{"Codigo ins_Final", "Codigo ins"}, {"CodigoAct", "Codigo act"}, {"Ref.Cap", "Capitulo"}, {"Ref.Sub", "Subcapitulo"}, {"Cantidades contrato", "Cantidad Contratado"}, {"VT contrato", "VT Contratado"}}),
    
    Typed = Table.TransformColumnTypes(Renamed, {{"Centro de Costos", type text}, {"Codigo ins", Int64.Type}, {"Cantidad Contratado", type number}, {"VT Contratado", type number}, {"Cantidad Cortes", type number}, {"VT Cortes", type number}}),
    
    FilteredZeros = Table.SelectRows(Typed, each ([VT Contratado] <> 0 and [VT Contratado] <> null)),
    
    TablaFinal = FilteredZeros
in
    TablaFinal
```

## DESCARGAS

```powerquery
let
    // ============================================================
    // DESCARGAS: lee la tabla DESCARGA del archivo del proyecto en
    // ".../0. Descargas pptos - Control costos interno/<Proyecto>.xlsx".
    // El archivo se localiza por nombre (exacto primero, luego "contiene"),
    // asi que agregar un proyecto nuevo = subir su archivo, sin tocar codigo.
    //
    // La carpeta se prueba en 2 ubicaciones posibles (CarpetasCandidatas):
    // la actual y la original, por si vuelve a moverse en SharePoint. Si se
    // reubica a un tercer sitio, hay que agregar esa ruta a la lista.
    // ============================================================
    FxToNumberFlex = F_Globales[FxToNumberFlex],
    FnCleanText = F_Globales[FnCleanText],
    FnNormalizeSpaces = F_Globales[FnNormalizeSpaces],
    FnReadSPBinary = F_Globales[FnReadSPBinary],
    FnEncode = F_Globales[FnEncode],

    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    RutaBase = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS",
    CarpetasCandidatas = {
        RutaBase & "/0. Descargas pptos - Control costos interno",
        RutaBase & "/DashBoard/0. Descargas pptos - Control costos interno"
    },
    ParamProyecto = Text.Trim(ProyectoActual),
    ProyUp = Text.Upper(ParamProyecto),

    ColumnasFinales = {
        "Proyecto", "Centro de Costos", "Subcapitulo", "Capitulo", "Actividad", "Codigo ins", "Ins",
        "Cantidad ppto (CC)", "V/U ppto (CC)", "Valor Total ppto (CC)",
        "# CC - Comparativo", "# CC", "Comparativo"
    },
    TablaVacia = #table(ColumnasFinales, {}),

    // ---------- Localizar la carpeta (probando cada candidata) ----------
    FnListarCarpeta = (ruta as text) as nullable record =>
        try Json.Document(Web.Contents(SiteUrl, [
            RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(ruta) & "')/Files",
            Query = [#"$select" = "Name,ServerRelativeUrl"],
            Headers = [Accept = "application/json;odata=nometadata"],
            Timeout = #duration(0, 0, 2, 0)
        ])) otherwise null,

    // Perezoso: prueba la 1a candidata y solo intenta la 2a si la 1a falla,
    // en vez de llamar SIEMPRE a ambas con List.Transform (List.Transform evalua
    // cada elemento sin importar si el primero ya tuvo exito - un viaje de red
    // desperdiciado en cada refresco desde que la carpeta volvio a su ruta original).
    Resp1 = FnListarCarpeta(CarpetasCandidatas{0}),
    Resp1Valido = Resp1 <> null and Record.HasFields(Resp1, "value"),
    Listado = if Resp1Valido then Resp1
              else let Resp2 = FnListarCarpeta(CarpetasCandidatas{1}) in
                   if Resp2 <> null and Record.HasFields(Resp2, "value") then Resp2 else null,

    // ---------- Localizar el archivo del proyecto dentro de esa carpeta ----------
    Archivos =
        if Listado = null or not Record.HasFields(Listado, "value")
        then #table({"Name", "ServerRelativeUrl"}, {})
        else Table.FromRecords(Listado[value]),
    SoloExcel = Table.SelectRows(Archivos, each
        not Text.StartsWith([Name], "~$") and
        Text.EndsWith(Text.Upper([Name]), ".XLSX")),
    NombreBase = (n as text) as text => Text.Upper(Text.Trim(Text.Start(n, Text.Length(n) - 5))),
    Exactos = Table.SelectRows(SoloExcel, each NombreBase([Name]) = ProyUp),
    Contienen = Table.SelectRows(SoloExcel, each Text.Contains(Text.Upper([Name]), ProyUp)),
    Candidatos = if Table.RowCount(Exactos) > 0 then Exactos else Contienen,
    Ruta = if Table.RowCount(Candidatos) = 0 then null else Candidatos{0}[ServerRelativeUrl],

    // ---------- Leer la tabla DESCARGA ----------
    Binario = if Ruta = null then null else FnReadSPBinary(SiteUrl, Ruta),
    Libro = if Binario = null then null else (try Excel.Workbook(Binario, null, true) otherwise null),
    TablaCruda =
        if Libro = null then null
        else try Libro{[Item = "DESCARGA", Kind = "Table"]}[Data]
             otherwise (try Libro{[Item = "DESCARGAS", Kind = "Sheet"]}[Data] otherwise null),
    TablaDescargas = if TablaCruda = null then TablaVacia else TablaCruda,

    // Filtro defensivo: el archivo ya es por proyecto, pero si alguien sube
    // un archivo combinado, igual sale solo el proyecto actual.
    Filtrado =
        if Table.HasColumns(TablaDescargas, "Proyecto")
        then Table.SelectRows(TablaDescargas, each
            Text.Upper(Text.Trim(Text.From(if [Proyecto] = null then "" else [Proyecto]))) = ProyUp)
        else TablaDescargas,
    SinFilasVacias = Table.SelectRows(Filtrado, each
        (try Text.Trim(Text.From(if [Centro de Costos] = null then "" else [Centro de Costos])) otherwise "") <> ""),

    // ---------- Limpieza y tipos ----------
    TextosLimpios = Table.TransformColumns(SinFilasVacias, {
        {"Proyecto", each FnCleanText(_), type text},
        {"Centro de Costos", each FnCleanText(_), type text},
        {"Subcapitulo", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Capitulo", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Actividad", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Ins", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"# CC - Comparativo", each FnNormalizeSpaces(_), type text},
        {"# CC", each FnNormalizeSpaces(_), type text},
        {"Comparativo", each FnNormalizeSpaces(_), type text},

        {"Cantidad ppto (CC)", each FxToNumberFlex(_), type number},
        {"V/U ppto (CC)", each FxToNumberFlex(_), type number},
        {"Valor Total ppto (CC)", each FxToNumberFlex(_), type number}
    }, null, MissingField.Ignore),

    TiposFinales = try Table.TransformColumnTypes(TextosLimpios, {{"Codigo ins", Int64.Type}}) otherwise TextosLimpios,

    TablaFinal = Table.SelectColumns(TiposFinales, ColumnasFinales, MissingField.UseNull),
    Resultado = Table.Buffer(TablaFinal)
in
    Resultado
```

## DESCUENTOS

```powerquery
let
    // ============================================================
    // FUNCIONES AUXILIARES GLOBALES
    // ============================================================
    FnFormatCodigoAct = F_Globales[FnFormatCodigoAct],
    FnDecodeHtml = F_Globales[FnDecodeHtml],
    FxToNumberFlex = F_Globales[FxToNumberFlex],
    FnReadSPBinary = F_Globales[FnReadSPBinary],
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    Columnas_HTML = F_Globales[FnBuildColumnas](15),

    // ============================================================
    // FUNCIÓN MÁGICA: PROCESAR DESCUENTOS
    // ============================================================
    FxProcesarDescuentos = (Binario as binary) =>
        let
            RawTable_0 = try Excel.Workbook(Binario, null, true){0}[Data]
                       otherwise Html.Table(FnDecodeHtml(Binario), Columnas_HTML, [RowSelector="tr"]),
            // Estandarizar nombres de columnas
            RawTable_ColNames = Table.ColumnNames(RawTable_0),
            RawTable = Table.RenameColumns(RawTable_0, List.Zip({RawTable_ColNames, List.Transform({1..List.Count(RawTable_ColNames)}, each "Columna" & Text.From(_))})),
            
            FilasLimpias = Table.SelectRows(RawTable, each [Columna1] <> "GRAN TOTAL" and [Columna1] <> "DESCUENTOS SALIDAS" and [Columna1] <> null),
            FillContrato = Table.FillDown(Table.AddColumn(FilasLimpias, "ContratoInfo", each let txt = Text.Trim(Text.From(if [Columna1] = null then "" else [Columna1])) in if Text.StartsWith(Text.Upper(txt), "CONTRATO") then txt else null), {"ContratoInfo"}),
            AddOCContrato = Table.AddColumn(FillContrato, "# OC / Contrato", each let raw = Text.Select(Text.From(if [ContratoInfo] = null then "" else [ContratoInfo]), {"0".."9"}) in if raw = "" then null else raw, type text),
            AddCodigoAct = Table.AddColumn(AddOCContrato, "Codigo act", each let txt = Text.Trim(Text.From(if [Columna3] = null then "" else [Columna3])), baseCod = if Text.Contains(txt, " ") then Text.BeforeDelimiter(txt, " ") else txt in if baseCod = "" then null else FnFormatCodigoAct(baseCod), type text),
            
            AddValorDescuento = Table.AddColumn(AddCodigoAct, "Valor descuento", each let rawTxt = Text.Remove(Text.Trim(Text.From(if [Columna7] = null then "" else [Columna7])), {"$", " "}) in FxToNumberFlex(rawTxt), Currency.Type),
            
            BaseFinal = Table.SelectColumns(Table.SelectRows(AddValorDescuento, each [#"# OC / Contrato"] <> null and [Codigo act] <> null and [Valor descuento] <> null and [Valor descuento] <> 0), {"# OC / Contrato", "Codigo act", "Valor descuento"})
        in BaseFinal,

    // ============================================================
    // CONEXIÓN A SHAREPOINT (LECTURA DESDE CONSULTA COMPARTIDA)
    // ============================================================
    ArchivosProyecto = Table.SelectRows(SP_Archivos_Proyecto, each
        Text.Contains([Name], "DESCUENTOS", Comparer.OrdinalIgnoreCase)
    ),
    ConCentroCosto = ArchivosProyecto,

    PickLatestBinary = (t as table, containsText as text) as nullable binary =>
        let
            candidatos = Table.Sort(
                Table.SelectRows(t, each Text.Contains([Name], containsText, Comparer.OrdinalIgnoreCase)),
                {{"TimeLastModified", Order.Descending}, {"Name", Order.Ascending}}
            ),
            path = if Table.RowCount(candidatos) = 0 then null else candidatos{0}[ServerRelativeUrl]
        in
            if path = null then null else FnReadSPBinary(SiteUrl, path),

    Agrupado = Table.Group(ConCentroCosto, {"Centro de Costos"}, {{"Binario", each PickLatestBinary(_, "DESCUENTOS")}}),
    CentrosConArchivo = Table.SelectRows(Agrupado, each [Binario] <> null),
    TablaConDatos = Table.AddColumn(CentrosConArchivo, "Datos", each FxProcesarDescuentos([Binario])),
    
    SinBinario = Table.RemoveColumns(TablaConDatos, {"Binario"}),
    Expandido = Table.ExpandTableColumn(SinBinario, "Datos", {"# OC / Contrato", "Codigo act", "Valor descuento"}),
    
    Descuentos_Clean = Table.TransformColumns(Expandido, {
        {"Centro de Costos", each if _ = null then null else Text.Upper(Text.Trim(Text.From(_))), type text},
        {"Codigo act", each FnFormatCodigoAct(_), type text},
        {"# OC / Contrato", each if _ = null then null else Text.Trim(Text.From(_)), type text}
    }, null, MissingField.Ignore),
    
    BaseDescuentos_EnMemoria = Descuentos_Clean,

    // ============================================================
    // LECTURA DIRECTA DE CONSULTAS (Memoria)
    // ============================================================
    SourceContratos = CONTRATOS,
    CONTRATOS_Clean = Table.TransformColumns(SourceContratos, {
        {"# OC / Contrato", each if _ = null then null else Text.Trim(Text.From(_)), type text}
    }, null, MissingField.Ignore),
    // 🚀 Buffer para que el JOIN no re-evalúe toda la cadena de CONTRATOS
    ContratosPorOC = Table.Buffer(Table.Group(CONTRATOS_Clean, {"# OC / Contrato"}, {{"Nombre Contratista", each List.First([Nombre Contratista]), type text}, {"Descripcion contrato", each List.First([Descripcion contrato]), type text}})),

    SourceItems = ITEMSINSUMOS,
    ITEMS_Clean = Table.TransformColumns(SourceItems, {
        {"Codigo act", each FnFormatCodigoAct(_), type text}
    }, null, MissingField.Ignore),
    ItemsPorCodigo = Table.Buffer(Table.Group(ITEMS_Clean, {"Codigo act"}, {{"Actividad", each List.First([Actividad]), type text}, {"Capitulo", each List.First([Capitulo]), type text}, {"Subcapitulo", each List.First([Subcapitulo]), type text}})),

    // ============================================================
    // CRUCES FINALES Y SELECCIÓN ESTRICTA
    // ============================================================
    MergeContratos = Table.NestedJoin(BaseDescuentos_EnMemoria, {"# OC / Contrato"}, ContratosPorOC, {"# OC / Contrato"}, "C", JoinKind.LeftOuter),
    ExpandContratos = Table.ExpandTableColumn(MergeContratos, "C", {"Nombre Contratista", "Descripcion contrato"}, {"Nombre Contratista", "Descripcion contrato"}),

    MergeItems = Table.NestedJoin(ExpandContratos, {"Codigo act"}, ItemsPorCodigo, {"Codigo act"}, "I", JoinKind.LeftOuter),
    ExpandItems = Table.ExpandTableColumn(MergeItems, "I", {"Actividad", "Capitulo", "Subcapitulo"}, {"Actividad", "Capitulo", "Subcapitulo"}),

    AgregadoTipo = Table.AddColumn(ExpandItems, "Tipo", each "Descuento", type text),
    
    SelectedFinal = Table.SelectColumns(AgregadoTipo, {"Centro de Costos", "Codigo act", "Actividad", "Capitulo", "Subcapitulo", "# OC / Contrato", "Nombre Contratista", "Descripcion contrato", "Valor descuento", "Tipo"}),
    TypedFinal = Table.TransformColumnTypes(SelectedFinal, {{"Centro de Costos", type text}, {"Codigo act", type text}, {"Actividad", type text}, {"Capitulo", type text}, {"Subcapitulo", type text}, {"# OC / Contrato", type text}, {"Nombre Contratista", type text}, {"Descripcion contrato", type text}, {"Valor descuento", Currency.Type}, {"Tipo", type text}}),
    
    FilteredZeros = Table.SelectRows(TypedFinal, each [Valor descuento] <> 0 and [Valor descuento] <> null),

    TablaFinal = FilteredZeros
in
    TablaFinal
```

## DIAGNOSTICO

```powerquery
let
    // Query de diagnostico: pegala como una consulta nueva en Power Query y cargala como tabla.
    // Te muestra cuantas filas trae cada query del modelo, si tiene errores y cuanto tarda.
    // Util para aislar el query roto o lento sin tener que abrir cada uno por separado.

    Medir = (nombre as text, fn as function) =>
        let
            t0 = DateTime.LocalNow(),
            // El "if t0 = null" fuerza a evaluar t0 ANTES de la consulta; Table.Buffer
            // materializa el resultado una sola vez (RowCount + errores no re-evaluan).
            res = if t0 = null then null else (try Table.Buffer(fn()) otherwise null),
            filas = if res = null then -1 else try Table.RowCount(res) otherwise -1,
            errores = if res = null then -1 else try Table.RowCount(Table.SelectRowsWithErrors(res)) otherwise -1,
            // La dependencia artificial en filas/errores fuerza a que t1 se evalue DESPUES
            // del trabajo; sin esto la evaluacion perezosa marca 0 segundos.
            t1 = DateTime.LocalNow() + #duration(0, 0, 0, (filas - filas) + (errores - errores)),
            segundos = Duration.TotalSeconds(t1 - t0)
        in
            [Query = nombre, Filas = filas, Errores = errores, Segundos = Number.Round(segundos, 1)],

    Filas = {
        Medir("SP_CarpetasCC",          () => SP_CarpetasCC),
        Medir("SP_Archivos_Proyecto",   () => SP_Archivos_Proyecto),
        Medir("SP_Seguimiento_Parsed",  () => SP_Seguimiento_Parsed),
        Medir("ITEMSINSUMOS",           () => ITEMSINSUMOS),
        Medir("PPTO_BD",                () => PPTO_BD),
        Medir("DESCARGAS",              () => DESCARGAS),
        Medir("CONTRATOS",              () => CONTRATOS),
        Medir("COMPRAS",                () => COMPRAS),
        Medir("DESCUENTOS",             () => DESCUENTOS),
        Medir("APROBACIONES_SP",        () => APROBACIONES_SP),
        Medir("PROVISIONES_SP",         () => PROVISIONES_SP),
        Medir("COMPARATIVOS",           () => COMPARATIVOS),
        Medir("DISPONIBLE",             () => DISPONIBLE),
        Medir("BD",                     () => BD)
    },

    Resultado = Table.FromRecords(Filas)
in
    Resultado
```

## DISPONIBLE

```powerquery
let
    // ============================================================
    // 1. FUNCIONES DE LIMPIEZA (Centralizadas desde F_Globales)
    // ============================================================
    FnCleanText = F_Globales[FnCleanText],
    FnNormalizeSpaces = F_Globales[FnNormalizeSpaces],
    FnRemoveAccentsSymbols = F_Globales[FnRemoveAccentsSymbols],

    // ============================================================
    // 2. LECTURA DE BASES (Conexión Directa en Memoria)
    // ============================================================
    FuentePPTO  = PPTO_BD,
    FuenteDetCC = DESCARGAS,

    // ============================================================
    // 3. PROCESAMIENTO PPTO (Con escudo anti-errores de texto)
    // ============================================================
    PPTO_Slim = Table.SelectColumns(FuentePPTO, {"Centro de Costos", "Codigo act", "Capitulo", "Actividad", "Subcapitulo", "Ins", "VT Presupuesto", "V/U Presupuesto"}, MissingField.Ignore),
    PPTO_Typed = Table.TransformColumns(PPTO_Slim, {
        {"Centro de Costos", each FnCleanText(_), type text}, 
        {"Codigo act", each FnCleanText(_), type text}, 
        {"Capitulo", each FnCleanText(_), type text}, 
        {"Actividad", each FnCleanText(_), type text}, 
        {"Subcapitulo", each FnCleanText(_), type text}, 
        {"Ins", each FnCleanText(_), type text}, 
        {"VT Presupuesto", each try Number.From(_) otherwise 0, type number}, 
        {"V/U Presupuesto", each try Number.From(_) otherwise 0, type number}
    }, null, MissingField.Ignore),
    
    PPTO_WithStdIns = Table.AddColumn(PPTO_Typed, "InsNorm", each FnRemoveAccentsSymbols([Ins]), type text),
    // Claves normalizadas (sin tildes) para el cruce con descargas: el seguimiento
    // trae "SALON SOCIAL" con tilde y las descargas sin ella (o viceversa).
    PPTO_WithNorm = Table.AddColumn(Table.AddColumn(PPTO_WithStdIns,
        "ActNorm", each FnRemoveAccentsSymbols([Actividad]), type text),
        "SubcapNorm", each FnRemoveAccentsSymbols(if [Subcapitulo] = null then "" else [Subcapitulo]), type text),
    PPTO_Grouped_Buffer = Table.Buffer(Table.Group(PPTO_WithNorm, {"Centro de Costos", "Codigo act", "Capitulo", "Actividad", "Subcapitulo", "InsNorm"}, {{"Ins_Oficial", each List.First(List.RemoveNulls([Ins])), type text}, {"ValorTotal_PPTO_Bloque", each List.Sum([VT Presupuesto]), type number}, {"Unitario_PPTO_Bloque", each List.First(List.RemoveNulls([#"V/U Presupuesto"])), type number}, {"ActNorm", each List.First(List.RemoveNulls([ActNorm])), type text}, {"SubcapNorm", each List.First([SubcapNorm]), type text}})),

    // ============================================================
    // 4. PROCESAMIENTO ADJUDICADOS (COMPARATIVOS)
    // ============================================================
    DetCC_Selected = Table.SelectColumns(FuenteDetCC, {"Centro de Costos", "Capitulo", "Actividad", "Subcapitulo", "Ins", "# CC - Comparativo", "Valor Total ppto (CC)", "V/U ppto (CC)"}, MissingField.Ignore),
    DetCC_Typed = Table.TransformColumns(DetCC_Selected, {
        {"Centro de Costos", each FnCleanText(_), type text}, 
        {"Capitulo", each FnCleanText(_), type text}, 
        {"Actividad", each FnCleanText(_), type text}, 
        {"Subcapitulo", each FnCleanText(_), type text}, 
        {"Ins", each FnCleanText(_), type text}, 
        {"# CC - Comparativo", each FnCleanText(FnNormalizeSpaces(_)), type text}, 
        {"Valor Total ppto (CC)", each try Number.From(_) otherwise null, type number}, 
        {"V/U ppto (CC)", each try Number.From(_) otherwise null, type number}
    }, null, MissingField.Ignore),
    
    DetCC_WithStdIns = Table.AddColumn(DetCC_Typed, "InsNorm", each FnRemoveAccentsSymbols([Ins]), type text),
    DetCC_Valid = Table.SelectRows(DetCC_WithStdIns, each [#"# CC - Comparativo"] <> null),

    // Las descargas traen el subcapitulo PEGADO al nombre de la actividad
    // ("3.03-LOSA INFERIOR 0.10M - TORRES (M2)") y la columna Subcapitulo vacia,
    // mientras que el ppto (seguimiento) ya viene con nombre limpio + Subcapitulo
    // lleno. Se deriva aqui igual (helper compartido) para que las claves crucen.
    FnSepararSubcapDD = F_Globales[FnSepararSubcapDeNombre],
    DetCC_Sep = Table.AddColumn(DetCC_Valid, "__Sep", each
        let s = if [Subcapitulo] = null then "" else Text.Trim(Text.From([Subcapitulo]))
        in if s <> "" then null else FnSepararSubcapDD([Actividad])),
    // El override tambien aplica a subcapitulos que YA vienen llenos en el archivo
    // de descargas (traen truncaduras propias: "GEN", "SALON SOCIA", "URBANISMO").
    FnOverrideSubcapDD = F_Globales[FnAplicarOverrideSubcap],
    DetCC_Derivado = Table.AddColumn(Table.AddColumn(DetCC_Sep,
        "ActClave", each if [__Sep] <> null and [__Sep][Subcap] <> null then [__Sep][Nombre] else [Actividad], type text),
        "SubcapClave", each
            let s = if [Subcapitulo] = null then "" else Text.Trim(Text.From([Subcapitulo]))
            in if s <> "" then FnOverrideSubcapDD([Subcapitulo]) else (if [__Sep] <> null then [__Sep][Subcap] else null), type text),
    DetCC_ConNorm = Table.AddColumn(Table.AddColumn(DetCC_Derivado,
        "ActNorm", each FnRemoveAccentsSymbols([ActClave]), type text),
        "SubcapNorm", each FnRemoveAccentsSymbols(if [SubcapClave] = null then "" else [SubcapClave]), type text),

    // ============================================================
    // 5. CRUCE 1: Alinear la base Adjudicada contra la estructura oficial
    // (por claves normalizadas: sin tildes, con subcapitulo derivado)
    // ============================================================
    DetCC_JoinPPTOBlock = Table.NestedJoin(DetCC_ConNorm, {"Centro de Costos", "Capitulo", "ActNorm", "SubcapNorm", "InsNorm"}, PPTO_Grouped_Buffer, {"Centro de Costos", "Capitulo", "ActNorm", "SubcapNorm", "InsNorm"}, "PPTOBlock", JoinKind.LeftOuter),
    DetCC_Expanded = Table.ExpandTableColumn(DetCC_JoinPPTOBlock, "PPTOBlock", {"Codigo act", "Actividad", "Subcapitulo", "Ins_Oficial", "ValorTotal_PPTO_Bloque", "Unitario_PPTO_Bloque"}, {"Codigo act", "Act_Oficial", "Subcap_Oficial", "Ins_Oficial", "ValorTotal_PPTO_Bloque", "Unitario_PPTO_Bloque"}),
    // Adoptar los nombres OFICIALES del ppto cuando hubo match, para que el
    // Cruce 2 (que une por texto) siempre coincida.
    DetCC_WithFinalIns0 = Table.AddColumn(DetCC_Expanded, "Ins_Final", each if [Ins_Oficial] <> null then [Ins_Oficial] else [Ins], type text),
    DetCC_WithFinalIns1 = Table.AddColumn(DetCC_WithFinalIns0, "Act_Final", each if [Act_Oficial] <> null then [Act_Oficial] else [ActClave], type text),
    DetCC_WithFinalIns = Table.AddColumn(DetCC_WithFinalIns1, "Subcap_Final", each if [Act_Oficial] <> null then [Subcap_Oficial] else [SubcapClave], type text),

    DetCC_WithCantidad = Table.AddColumn(DetCC_WithFinalIns, "Cantidad_Calc", each let total = [#"Valor Total ppto (CC)"], unit = [#"V/U ppto (CC)"] in if unit <> null and unit <> 0 then total / unit else null, type number),
    DetCC_ReportShape0 = Table.SelectColumns(DetCC_WithCantidad, {"Centro de Costos", "Codigo act", "Capitulo", "Act_Final", "Subcap_Final", "Ins_Final", "# CC - Comparativo", "Valor Total ppto (CC)", "V/U ppto (CC)", "Cantidad_Calc"}, MissingField.Ignore),
    DetCC_FinalAdjudicados_Renamed = Table.RenameColumns(DetCC_ReportShape0, {{"Ins_Final", "Ins"}, {"Act_Final", "Actividad"}, {"Subcap_Final", "Subcapitulo"}}),
    DetCC_FinalAdjudicados = Table.AddColumn(DetCC_FinalAdjudicados_Renamed, "Tipo", each "Adjudicado", type text),

    // ============================================================
    // 6. CRUCE 2: Restar lo adjudicado al PPTO para hallar Saldo
    // ============================================================
    Adj_Grouped_Buffer = Table.Buffer(Table.Group(DetCC_FinalAdjudicados, {"Centro de Costos", "Codigo act", "Capitulo", "Actividad", "Subcapitulo", "Ins"}, {{"ValorAdjudicado_Bloque", each List.Sum([#"Valor Total ppto (CC)"]), type number}, {"CantAdjudicada_Bloque", each List.Sum([Cantidad_Calc]), type number}})),

    Bloques_Merge = Table.NestedJoin(PPTO_Grouped_Buffer, {"Centro de Costos", "Codigo act", "Capitulo", "Actividad", "Subcapitulo", "Ins_Oficial"}, Adj_Grouped_Buffer, {"Centro de Costos", "Codigo act", "Capitulo", "Actividad", "Subcapitulo", "Ins"}, "AdjTot", JoinKind.LeftOuter),
    Bloques_ExpandedAdj = Table.ExpandTableColumn(Bloques_Merge, "AdjTot", {"ValorAdjudicado_Bloque", "CantAdjudicada_Bloque"}, {"ValorAdjudicado_Bloque", "CantAdjudicada_Bloque"}),
    Bloques_Filled = Table.TransformColumns(Bloques_ExpandedAdj, {{"ValorTotal_PPTO_Bloque", each if _ = null then 0 else _, type number}, {"Unitario_PPTO_Bloque", each if _ = null then 0 else _, type number}, {"ValorAdjudicado_Bloque", each if _ = null then 0 else _, type number}, {"CantAdjudicada_Bloque", each if _ = null then 0 else _, type number}}),
    
    // Hallamos el valor pendiente real
    Bloques_WithSaldoValor = Table.AddColumn(Bloques_Filled, "Pendiente_Valor", each [ValorTotal_PPTO_Bloque] - [ValorAdjudicado_Bloque], type number),
    Bloques_WithSaldoCant = Table.AddColumn(Bloques_WithSaldoValor, "Pendiente_Cantidad", each let vTot = [ValorTotal_PPTO_Bloque], u = [Unitario_PPTO_Bloque], cObj = if u <> null and u <> 0 then vTot / u else null, cAdj = [CantAdjudicada_Bloque] in if cObj <> null and cAdj <> null then cObj - cAdj else null, type number),

    // ============================================================
    // 7. ARMAR LA TABLA FINAL DE SALDOS Y UNIR
    // ============================================================
    PorAdj_BaseRows = Table.SelectColumns(Bloques_WithSaldoCant, {"Centro de Costos", "Codigo act", "Capitulo", "Actividad", "Subcapitulo", "Ins_Oficial", "Pendiente_Valor", "Unitario_PPTO_Bloque", "Pendiente_Cantidad"}, MissingField.Ignore),
    PorAdj_WithColsRenamed = Table.RenameColumns(PorAdj_BaseRows, {{"Ins_Oficial", "Ins"}, {"Pendiente_Valor", "Valor Total ppto (CC)"}, {"Unitario_PPTO_Bloque", "V/U ppto (CC)"}, {"Pendiente_Cantidad", "Cantidad_Calc"}}, MissingField.Ignore),
    PorAdj_AddComparativo = Table.AddColumn(PorAdj_WithColsRenamed, "# CC - Comparativo", each "Por adjudicar", type text),
    PorAdj_AddTipo = Table.AddColumn(PorAdj_AddComparativo, "Tipo", each "Por adjudicar", type text),

    UnionFullRaw = Table.Combine({DetCC_FinalAdjudicados, PorAdj_AddTipo}),
    
    // 🔥 LA BARREDORA VITAL: Si el saldo pendiente queda en 0 (con tolerancia de centavos), lo elimina para no estorbar.
    UnionFiltered = Table.SelectRows(UnionFullRaw, each 
        ([Tipo] = "Por adjudicar" and (Number.Round([#"Valor Total ppto (CC)"], 2) <> 0)) or 
        ([Tipo] = "Adjudicado" and [#"Valor Total ppto (CC)"] <> null)
    ),
    
    // Subcapitulo embebido en el nombre (proyectos tipo TURPIAL): las filas que
    // llegan sin Subcapitulo pero con el patron "ACTIVIDAD - SUBCAP (UM)" en el
    // nombre lo derivan con el helper compartido, y el nombre queda limpio.
    FnSepararSubcap = F_Globales[FnSepararSubcapDeNombre],
    ConSubcapDerivado0 = Table.AddColumn(UnionFiltered, "__Sep", each
        let s = if [Subcapitulo] = null then "" else Text.Trim(Text.From([Subcapitulo]))
        in if s <> "" then null else FnSepararSubcap([Actividad])),
    ConSubcapDerivado1 = Table.AddColumn(ConSubcapDerivado0, "SubcapFinal", each
        let s = if [Subcapitulo] = null then "" else Text.Trim(Text.From([Subcapitulo]))
        in if s <> "" then FnOverrideSubcapDD([Subcapitulo])
           else if [__Sep] <> null then [__Sep][Subcap] else null, type text),
    ConSubcapDerivado2 = Table.AddColumn(ConSubcapDerivado1, "ActividadFinal", each
        if [__Sep] <> null and [__Sep][Subcap] <> null then [__Sep][Nombre] else [Actividad], type text),
    ConSubcapDerivado = Table.RenameColumns(
        Table.RemoveColumns(ConSubcapDerivado2, {"Subcapitulo", "Actividad", "__Sep"}),
        {{"SubcapFinal", "Subcapitulo"}, {"ActividadFinal", "Actividad"}}),

    Final_Ordered = Table.ReorderColumns(ConSubcapDerivado, {"Centro de Costos", "Codigo act", "Capitulo", "Actividad", "Subcapitulo", "Ins", "# CC - Comparativo", "Tipo", "Cantidad_Calc", "V/U ppto (CC)", "Valor Total ppto (CC)"}, MissingField.Ignore),

    TablaFinal = Final_Ordered
in
    TablaFinal
```

## F_Globales

```powerquery
let
    Funciones = [
        FnFormatCodigoAct = (raw as any) as nullable text =>
            let
                // El #(00A0) (espacio no separable) llega invisible desde los reportes de
                // SharePoint; si no se limpia aqui, "2.041" y "2.041 " (con NBSP) quedan
                // como codigos distintos y rompen el fill-down/join por codigo de actividad.
                txtRaw  = if raw = null then null else Text.Trim(Text.Replace(Text.From(raw), "#(00A0)", " ")),
                result =
                    if txtRaw = null or txtRaw = "" then null
                    else
                        let
                            txtNorm = Text.Replace(Text.Replace(txtRaw, ",", "."), " ", ""),
                            hasDot  = Text.Contains(txtNorm, ".")
                        in
                            if hasDot then txtNorm
                            else
                                let
                                    digits = Text.Select(txtNorm, {"0".."9"}),
                                    len    = Text.Length(digits)
                                in
                                    if len <= 3 then null
                                    else Text.Range(digits, 0, len - 3) & "." & Text.Range(digits, len - 3, 3)
            in result,

        FxToNumberFlex = (value as any) as nullable number =>
            let
                v              = value,
                isNum          = Value.Is(v, type number),
                numeroDirecto  = if isNum then Number.From(v) else null,
                t0             = if v = null then "" else Text.From(v),
                t              = Text.Trim(Text.Replace(Text.Replace(t0, "#(00A0)", ""), " ", "")),
                tryUS          = try Number.FromText(t, "en-US"),
                valUS          = if tryUS[HasError] then null else tryUS[Value],
                tryES          = try Number.FromText(t, "es-ES"),
                valES          = if tryES[HasError] then null else tryES[Value],
                result         = if numeroDirecto <> null then numeroDirecto
                                 else if t = "" then null
                                 else if valUS <> null then valUS
                                 else valES
            in result,

        FnRemoveAccentsSymbols = (t as any) as nullable text =>
            let
                initial = try (if t = null then null else Text.From(t)) otherwise null,
                replacements = {
                    {"#(00E1)","a"},{"#(00C1)","A"},
                    {"#(00E9)","e"},{"#(00C9)","E"},
                    {"#(00ED)","i"},{"#(00CD)","I"},
                    {"#(00F3)","o"},{"#(00D3)","O"},
                    {"#(00FA)","u"},{"#(00DA)","U"},
                    {"#(00F1)","n"},{"#(00D1)","N"},
                    {"#(00BA)",""},{"#(00B0)",""},{"#(00A8)",""},
                    {"#(lf)", " "}, {"#(cr)", " "}
                },
                result = if initial = null then null
                         else List.Accumulate(replacements, initial, (state, current) => Text.Replace(state, current{0}, current{1}))
            in result,

        FnRemoveAccentMarks = (t as any) as nullable text =>
            let
                initial = try (if t = null then null else Text.From(t)) otherwise null,
                replacements = {
                    {"#(00E1)","a"},{"#(00C1)","A"},
                    {"#(00E9)","e"},{"#(00C9)","E"},
                    {"#(00ED)","i"},{"#(00CD)","I"},
                    {"#(00F3)","o"},{"#(00D3)","O"},
                    {"#(00FA)","u"},{"#(00DA)","U"},
                    {"#(00DC)","U"},{"#(00FC)","u"},
                    {"#(00BA)",""},{"#(00B0)",""},{"#(00A8)",""},
                    {"#(lf)", " "}, {"#(cr)", " "}
                },
                result = if initial = null then null
                         else List.Accumulate(replacements, initial, (state, current) => Text.Replace(state, current{0}, current{1}))
            in result,
        FnClaveLimpia = (t as nullable text) as nullable text =>
            let
                sinUnidad = if t = null then null
                            else if Text.Contains(t, "(") then Text.BeforeDelimiter(t, "(")
                            else t,
                t1 = if sinUnidad = null then null else Text.Upper(Text.Trim(sinUnidad)),
                repl = {
                    {"#(00C1)","A"},{"#(00C9)","E"},{"#(00CD)","I"},
                    {"#(00D3)","O"},{"#(00DA)","U"},{"#(00D1)","N"},{"#(00DC)","U"}
                },
                t2 = if t1 = null then null
                     else List.Accumulate(repl, t1, (state, current) => Text.Replace(state, current{0}, current{1})),
                t3 = if t2 = null then null else Text.Select(t2, {"A".."Z", "0".."9"}),
                result = if t3 = null or t3 = "" then null else t3
            in result,

        FnCleanText = (t as any) as nullable text =>
            try (if t = null then null else let txt = Text.Trim(Text.From(t)) in if txt = "" then null else Text.Upper(txt)) otherwise null,

        FnTrimText = (t as any) as nullable text =>
            try (if t = null then null else Text.Trim(Text.From(t))) otherwise null,

        // Normaliza claves de texto para cruces entre fuentes (ej. "# CC - Comparativo"):
        // 1) espacios duros (nbsp) -> normales, 2) colapsa espacios repetidos internos,
        // 3) quita espacios alrededor de guiones ("002 - NOMBRE" -> "002-NOMBRE", forma
        //    compacta que usa la tabla manual Det_CC), 4) recorta extremos.
        // Debe aplicarse EN AMBOS LADOS de un join con la misma funcion.
        FnNormalizeSpaces = (t as any) as nullable text =>
            try (
                if t = null then null
                else
                    let
                        txt = Text.Replace(Text.From(t), "#(00A0)", " "),
                        partes = List.Select(Text.Split(txt, " "), each _ <> ""),
                        unido = Text.Combine(partes, " "),
                        sinEspGuion = Text.Replace(Text.Replace(unido, " -", "-"), "- ", "-")
                    in if sinEspGuion = "" then null else sinEspGuion
            ) otherwise null,

        // Decodifica un binario HTML/texto de los reportes SINCO detectando la codificacion:
        // intenta UTF-8 y, si el resultado trae el caracter de reemplazo U+FFFD (tipico de
        // decodificar Latin-1/Windows-1252 como UTF-8: la enie y tildes se vuelven "?"),
        // re-decodifica como ISO-8859-1. Los reportes de SINCO/Oracle vienen mezclados en
        // ambas codificaciones segun el modulo que los exporta — NUNCA usar un codepage
        // fijo en Text.FromBinary para estos archivos, usar siempre esta funcion.
        FnDecodeHtml = (bin as binary) as text =>
            let
                buf = Binary.Buffer(bin),
                utf8 = try Text.FromBinary(buf, TextEncoding.Utf8) otherwise null,
                usarLatin1 = utf8 = null or Text.Contains(utf8, "#(FFFD)"),
                result = if usarLatin1 then Text.FromBinary(buf, 28591) else utf8
            in
                result,

        // ============================================================
        // Subcapitulo embebido en el nombre (proyectos tipo TURPIAL)
        // ============================================================
        // Sufijos de "frente"/especialidad que el reporte agrega DESPUES del
        // subcapitulo real ("... - CUARTO DE BASURAS - ELECTRICO"): NO son
        // subcapitulos. Se reconocen tambien sus truncaduras (ELE, ELEC...).
        SubcapSufijosIgnorados = {"ELECTRICO"},
        // Truncaduras irrecuperables (texto cortado en el reporte de origen sin
        // forma de reconstruirlo alli mismo), confirmadas por el usuario.
        // Clave: valor derivado NORMALIZADO (mayusculas, sin tildes).
        // Valor: subcapitulo real (con tildes via escape #(00D3) para mantener
        // este archivo en ASCII puro y a salvo de problemas de decodificacion).
        SubcapOverrides = [
            #"APTOS (U"   = "TORRES",
            GEN           = "GENERALES",
            #"SALON SOCIA"= "SAL#(00D3)N SOCIAL",
            URBANISMO     = "URBANISMO INTERIOR"
        ],

        FnEsSufijoSubcapIgnorado = (t as text) as logical =>
            let norm = Text.Upper(FnRemoveAccentsSymbols(t))
            in List.AnyTrue(List.Transform(SubcapSufijosIgnorados, (s) =>
                norm = s or (Text.Length(norm) >= 3 and Text.StartsWith(s, norm)))),

        FnAplicarOverrideSubcap = (v as nullable text) as nullable text =>
            if v = null then null
            else let norm = Text.Upper(FnRemoveAccentsSymbols(v)),
                     o = try Record.Field(SubcapOverrides, norm) otherwise null
                 in if o <> null then o else v,

        // Quita guiones sueltos colgando al final de un texto (recursivo).
        FnQuitarGuionFinal = (t as text) as text =>
            let r = Text.Trim(t)
            in if Text.EndsWith(r, "-") then @FnQuitarGuionFinal(Text.Range(r, 0, Text.Length(r) - 1)) else r,

        // Extrae la cola tras el ultimo " - " (con limpieza de guion colgante),
        // saltando sufijos ignorados de forma recursiva. null si no hay cola valida.
        FnExtraerSubcapDeTexto = (txt as text) as nullable text =>
            let
                pos    = Text.PositionOf(txt, " - ", Occurrence.Last),
                cola   = if pos < 0 then "" else FnQuitarGuionFinal(Text.Trim(Text.Range(txt, pos + 3))),
                cabeza = if pos < 0 then "" else Text.Trim(Text.Range(txt, 0, pos)),
                valida = cola <> "" and cabeza <> "" and Text.Length(cola) <= 60
            in
                if not valida then null
                else if FnEsSufijoSubcapIgnorado(cola) then @FnExtraerSubcapDeTexto(cabeza)
                else cola,

        // Quita del final de un nombre los sufijos ignorados (" - ELECTRICO").
        FnQuitarSufijosSubcapIgnorados = (txt as text) as text =>
            let
                pos    = Text.PositionOf(txt, " - ", Occurrence.Last),
                cola   = if pos < 0 then "" else FnQuitarGuionFinal(Text.Trim(Text.Range(txt, pos + 3))),
                cabeza = if pos < 0 then "" else Text.Trim(Text.Range(txt, 0, pos))
            in if pos >= 0 and cola <> "" and FnEsSufijoSubcapIgnorado(cola) then @FnQuitarSufijosSubcapIgnorados(cabeza) else txt,

        // Separa "NOMBRE - SUBCAP (UM)" en [Nombre, Subcap]: desprende la unidad final
        // "(UM)" si existe, quita sufijos ignorados, extrae el subcapitulo (con overrides)
        // y devuelve el nombre sin el subcapitulo, re-anexando la unidad si no quedo ya.
        // Para nombres cuyo subcapitulo viene DESPUES de la unidad ("X (M3) - TANQUE")
        // o antes ("X - TANQUE (M3)") funciona en ambos ordenes.
        FnSepararSubcapDeNombre = (nombreRaw as nullable text) as record =>
            let
                txt0   = if nombreRaw = null then "" else Text.Trim(Text.Replace(Text.From(nombreRaw), "#(00A0)", " ")),
                txt    = Text.Combine(List.Select(Text.Split(txt0, " "), each _ <> ""), " "),
                // desprender "(UM)" final si existe
                tieneUM = Text.EndsWith(txt, ")") and Text.PositionOf(txt, "(", Occurrence.Last) >= 0,
                posPar  = if tieneUM then Text.PositionOf(txt, "(", Occurrence.Last) else -1,
                um      = if tieneUM then Text.Range(txt, posPar) else "",
                cuerpo0 = if tieneUM then Text.Trim(Text.Range(txt, 0, posPar)) else txt,
                cuerpo  = FnQuitarGuionFinal(FnQuitarSufijosSubcapIgnorados(cuerpo0)),
                subcapX = FnExtraerSubcapDeTexto(cuerpo),
                subcap  = FnAplicarOverrideSubcap(subcapX),
                posTail = if subcapX = null then -1 else Text.PositionOf(cuerpo, " - ", Occurrence.Last),
                cabeza  = if posTail < 0 then cuerpo else FnQuitarGuionFinal(Text.Trim(Text.Range(cuerpo, 0, posTail))),
                nombreF0 = if subcapX = null then cuerpo else cabeza,
                nombreF  = if um = "" or Text.EndsWith(nombreF0, um) then nombreF0 else nombreF0 & " " & um
            in
                [Nombre = nombreF, Subcap = subcap],

        FnPrepareTableWithHeader = (tbl as table) as table =>
            let
                firstColName   = Table.ColumnNames(tbl){0},
                firstColValues = Table.Column(tbl, firstColName),
                headerFlags    = List.Transform(firstColValues, (x) =>
                    let
                        txt     = Text.Upper(if x = null then "" else Text.From(x)),
                        txtNorm = Text.Replace(txt, "#(00D3)", "O")
                    in Text.Contains(txtNorm, "COD")),
                hasHeader = List.Contains(headerFlags, true),
                promoted  = if hasHeader then
                    let
                        headerIndex = List.PositionOf(headerFlags, true),
                        skipped     = Table.Skip(tbl, headerIndex)
                    in Table.PromoteHeaders(skipped, [PromoteAllScalars = true])
                    else tbl
            in promoted,

        FnEncode = (path as nullable text) as nullable text =>
            if path = null then null
            else Text.Combine(List.Transform(Text.Split(path, "/"), each Uri.EscapeDataString(_)), "/"),

        FnBuildColumnas = (n as number) as list =>
            List.Transform({1..n}, each {"Columna " & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}),

        FnCleanContratista = (t as any) as nullable text =>
            let
                safe       = try (if t = null then null else Text.From(t)) otherwise null,
                t2         = if safe = null then null else Text.Replace(safe, Character.FromNumber(65533), Character.FromNumber(78)),
                t3         = if t2 = null then null else Text.Trim(Text.Upper(t2)),
                repl       = {
                    {Character.FromNumber(193),"A"},{Character.FromNumber(201),"E"},
                    {Character.FromNumber(205),"I"},{Character.FromNumber(211),"O"},
                    {Character.FromNumber(218),"U"},{Character.FromNumber(209),"N"}
                },
                t3_clean   = if t3 = null then null
                             else List.Accumulate(repl, t3, (state, current) => Text.Replace(state, current{0}, current{1})),
                suffixes   = {" S.A.S.", " S.A.S", " SAS.", " SAS", " S.A.", " S.A", " SA.", " SA", " LTDA.", " LTDA", " S EN C", " S. EN C."},
                t4         = if t3_clean = null then null
                             else List.Accumulate(suffixes, t3_clean, (state, suffix) =>
                                 if Text.EndsWith(state, suffix)
                                 then Text.Trim(Text.Range(state, 0, Text.Length(state) - Text.Length(suffix)))
                                 else state),
                result     = if t4 = null or t4 = "" then null else t4
            in result,

        FnMapColumn = (rec as record, cols as list, keywords as list) =>
            let
                norm = (x as any) as text =>
                    let
                        txt = try Text.From(x) otherwise "",
                        clean = FnRemoveAccentsSymbols(txt)
                    in Text.Upper(if clean = null then "" else clean),
                match = List.First(
                    List.Select(cols, (c) =>
                        List.AnyTrue(List.Transform(keywords, (k) => Text.Contains(norm(c), norm(k))))
                    ),
                    null
                )
            in if match = null then null else Record.Field(rec, match),

        FnBuildFolderPrefixMap = (carpetas as list) as record =>
            let
                pares = List.Transform(carpetas, (x) =>
                    let
                        nombre = try Text.From(x) otherwise "",
                        prefix = if Text.Contains(nombre, "-") then Text.Trim(Text.BeforeDelimiter(nombre, "-")) else Text.Trim(nombre)
                    in {prefix, nombre}),
                validos = List.Select(pares, each _{0} <> null and _{0} <> ""),
                tabla = Table.Distinct(Table.FromRows(validos, {"Clave", "Valor"}), {"Clave"})
            in Record.FromList(tabla[Valor], tabla[Clave]),

        FnMatchFolder = (proyectoExcel as text, listaCarpetas as list) as text =>
            let
                count = List.Count(listaCarpetas)
            in
                if count = 0 then proyectoExcel
                else if count = 1 then listaCarpetas{0}
                else
                    let
                        proyClean = FnRemoveAccentsSymbols(Text.Upper(proyectoExcel)),
                        matches   = List.Select(listaCarpetas, each
                            let
                                baseName       = if Text.Contains(_, "-") then Text.Trim(Text.AfterDelimiter(_, "-")) else Text.Trim(_),
                                baseClean      = FnRemoveAccentsSymbols(Text.Upper(baseName)),
                                lastWordFolder = List.Last(Text.Split(baseClean, " ")),
                                lastWordProy   = List.Last(Text.Split(proyClean, " "))
                            in
                                Text.Contains(proyClean, lastWordFolder) or
                                Text.Contains(baseClean, lastWordProy) or
                                Text.Replace(baseClean, " ", "") = Text.Replace(proyClean, " ", "")
                        )
                    in if List.Count(matches) = 1 then matches{0} else proyectoExcel,

        FnReadSPBinary = (siteUrl as text, filePath as text) as nullable binary =>
            let
                raw = try Web.Contents(siteUrl, [
                    RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(filePath) & "')/$value",
                    Headers = [Accept = "*/*"],
                    Timeout = #duration(0, 0, 2, 0),
                    ManualStatusHandling = {404, 429, 500, 502, 503, 504}
                ]) otherwise null,
                status = if raw = null then null else try Value.Metadata(raw)[Response.Status] otherwise 200,
                result = if raw = null or status >= 400 then null else Binary.Buffer(raw)
            in
                result,

        FnReadSPExcel = (siteUrl as text, filePath as text) as nullable table =>
            let
                binario = FnReadSPBinary(siteUrl, filePath),
                libro = if binario = null then null else try Excel.Workbook(binario, null, true) otherwise null,
                data = if libro = null or Table.RowCount(libro) = 0 then null else try libro{0}[Data] otherwise null,
                result = if data = null then null else try Table.PromoteHeaders(data, [PromoteAllScalars=true]) otherwise null
            in
                result,

        FxProcesarCentroCosto = (BinarioSeguimiento as binary, BinarioPresupuesto as binary) as table =>
            let
                Columnas_HTML = FnBuildColumnas(25),
                Columnas_APU  = FnBuildColumnas(3),

                OrigenItems   = try Excel.Workbook(BinarioSeguimiento, null, true){0}[Data]
                                otherwise Html.Table(FnDecodeHtml(BinarioSeguimiento), Columnas_HTML, [RowSelector="tr"]),
                ItemsPrepared = Table.Buffer(FnPrepareTableWithHeader(OrigenItems)),

                ItemsColNames     = Table.ColumnNames(ItemsPrepared),
                ItemsCodColName   = ItemsColNames{0},
                ItemsDescColName  = ItemsColNames{1},
                ItemsTipoColName  = ItemsColNames{2},
                ItemsUMColName    = ItemsColNames{3},

                ItemsWithTipoFila = Table.AddColumn(ItemsPrepared, "TipoFila", (r as record) =>
                    let
                        codValue  = Record.Field(r, ItemsCodColName),
                        descValue = Record.Field(r, ItemsDescColName),
                        tipoValue = Record.Field(r, ItemsTipoColName),
                        umValue   = Record.Field(r, ItemsUMColName),
                        codText   = if codValue  = null then "" else Text.Trim(Text.Replace(Text.From(codValue), "#(00A0)", " ")),
                        descText  = if descValue = null then "" else Text.Trim(Text.From(descValue)),
                        tipoText  = if tipoValue = null then "" else Text.Trim(Text.From(tipoValue)),
                        umText    = if umValue   = null then "" else Text.Trim(Text.From(umValue)),
                        codUpper  = Text.Upper(codText),
                        descUpper = Text.Upper(descText),
                        // Sin este replace, un codigo de Actividad con NBSP pegado hace fallar
                        // Number.FromText -> la fila cae en "Otro" -> el fill-down de abajo
                        // arrastra el codigo de la actividad ANTERIOR sobre un bloque de insumos
                        // que en realidad pertenece a otra actividad (insumos "ajenos").
                        codTextNum = Text.Replace(codText, " ", ""),
                        tryNum    = try Number.FromText(codTextNum),
                        isNumeric = not tryNum[HasError],
                        numValue  = if isNumeric then tryNum[Value] else 0,
                        tipoFila  =
                            if codText = "" then "Otro"
                            else if Text.StartsWith(codUpper, "SUBCAP") or Text.StartsWith(descUpper, "SUBCAP") then "SubCapitulo"
                            else if Text.Contains(codUpper, "CAPITULO") or Text.Contains(descUpper, "CAPITULO") then "Capitulo"
                            else if isNumeric and tipoText = "" and umText = "" and (Text.Length(codText) <= 2 or (numValue >= 1000 and Number.Mod(numValue, 1000) = 0)) then "Capitulo"
                            else if isNumeric and tipoText = "" and umText = "" then "Actividad"
                            else if isNumeric then "Insumo"
                            else "Otro"
                    in tipoFila, type text),

                ItemsWithCapitulo = Table.AddColumn(ItemsWithTipoFila, "Capitulo", (r as record) =>
                    let
                        tipo   = Record.Field(r, "TipoFila"),
                        codRaw = Record.Field(r, ItemsCodColName),
                        descRaw= Record.Field(r, ItemsDescColName),
                        codTxt = if codRaw  = null then "" else Text.Trim(Text.From(codRaw)),
                        descTxt= if descRaw = null then "" else Text.Trim(Text.From(descRaw)),
                        tryN   = try Number.FromText(codTxt),
                        codCap = if codTxt = "00" then codTxt
                                 else if not tryN[HasError] and tryN[Value] >= 1000 and Number.Mod(tryN[Value], 1000) = 0
                                 then Text.From(tryN[Value] / 1000)
                                 else codTxt,
                        capTxt = if descTxt = "" then codCap else codCap & "-" & descTxt
                    in if tipo = "Capitulo" then capTxt else null, type text),

                ItemsCapituloFillDown   = Table.FillDown(ItemsWithCapitulo, {"Capitulo"}),

                ItemsWithSubcapitulo = Table.AddColumn(ItemsCapituloFillDown, "Subcapitulo", (r as record) =>
                    let
                        tipo      = Record.Field(r, "TipoFila"),
                        codRaw    = Record.Field(r, ItemsCodColName),
                        descRaw   = Record.Field(r, ItemsDescColName),
                        codTxt    = if codRaw  = null then "" else Text.From(codRaw),
                        descTxt   = if descRaw = null then "" else Text.From(descRaw),
                        fuenteRaw = if Text.Contains(Text.Upper(codTxt), "SUBCAP") then codTxt
                                    else if Text.Contains(Text.Upper(descTxt), "SUBCAP") then descTxt
                                    else "",
                        subTxt    = if tipo <> "SubCapitulo" or fuenteRaw = "" then null
                                    else let baseTxt = if Text.Contains(fuenteRaw, ":") then Text.AfterDelimiter(fuenteRaw, ":") else fuenteRaw
                                         in Text.Trim(baseTxt)
                    in subTxt, type text),

                ItemsSubcapituloFillDown  = Table.FillDown(ItemsWithSubcapitulo, {"Subcapitulo"}),

                // true si el SEGUIMIENTO trae al menos una fila explicita "SUBCAPITULO:".
                // Proyectos como TURPIAL no las traen: alli el subcapitulo viene pegado
                // como sufijo del nombre de la actividad ("... - GENERALES") y se deriva
                // mas abajo (SubcapDerivado). El gate evita aplicar esa heuristica en
                // proyectos que si declaran subcapitulos (alli un " - X" final puede ser
                // parte legitima del nombre y no un subcapitulo).
                TieneSubcapExplicito = List.Contains(List.Buffer(Table.Column(ItemsWithTipoFila, "TipoFila")), "SubCapitulo"),
                ItemsWithCodActRaw        = Table.AddColumn(ItemsSubcapituloFillDown, "CodigoActRaw", (r as record) =>
                    let tipo = Record.Field(r, "TipoFila") in if tipo = "Actividad" then Text.From(Record.Field(r, ItemsCodColName)) else null, type text),
                // Descripcion real de la fila-Actividad en SEGUIMIENTO POR ITEMS, capturada y
                // arrastrada junto con el codigo. El codigo de APU es una numeracion
                // independiente que puede coincidir con el de SEGUIMIENTO por pura casualidad
                // sin ser la misma actividad; esta descripcion (que SI viene del mismo
                // reporte que trae los insumos) es la fuente confiable del nombre.
                ItemsWithDescActRaw       = Table.AddColumn(ItemsWithCodActRaw, "DescActRaw", (r as record) =>
                    let tipo = Record.Field(r, "TipoFila") in if tipo = "Actividad" then Record.Field(r, ItemsDescColName) else null, type text),
                ItemsCodActRawFillDown    = Table.FillDown(ItemsWithDescActRaw, {"CodigoActRaw", "DescActRaw"}),
                ItemsWithCodigoAct        = Table.AddColumn(ItemsCodActRawFillDown, "Codigo act", each FnFormatCodigoAct([CodigoActRaw]), type text),
                ItemsSoloInsumos          = Table.SelectRows(ItemsWithCodigoAct, each [TipoFila] = "Insumo"),
                ItemsColsInsumos          = Table.ColumnNames(ItemsSoloInsumos),

                CantPresCol = if List.Count(ItemsColsInsumos) > 4  then ItemsColsInsumos{4}  else null,
                VTPresCol   = if List.Count(ItemsColsInsumos) > 6  then ItemsColsInsumos{6}  else null,
                CantProyCol = if List.Count(ItemsColsInsumos) > 7  then ItemsColsInsumos{7}  else null,
                VTProyCol   = if List.Count(ItemsColsInsumos) > 9  then ItemsColsInsumos{9}  else null,
                CantConsCol = if List.Count(ItemsColsInsumos) > 19 then ItemsColsInsumos{19} else null,
                VTConsCol   = if List.Count(ItemsColsInsumos) > 21 then ItemsColsInsumos{21} else null,

                A1 = Table.AddColumn(ItemsSoloInsumos, "Cantidad Presupuesto", (r) => if CantPresCol = null then null else Record.Field(r, CantPresCol)),
                A2 = Table.AddColumn(A1, "VT Presupuesto",    (r) => if VTPresCol   = null then null else Record.Field(r, VTPresCol)),
                A3 = Table.AddColumn(A2, "Cantidad Proyectado",(r) => if CantProyCol = null then null else Record.Field(r, CantProyCol)),
                A4 = Table.AddColumn(A3, "VT Proyectado",     (r) => if VTProyCol   = null then null else Record.Field(r, VTProyCol)),
                A5 = Table.AddColumn(A4, "Cantidad Consumido", (r) => if CantConsCol = null then null else Record.Field(r, CantConsCol)),
                A6 = Table.AddColumn(A5, "VT Consumido",      (r) => if VTConsCol   = null then null else Record.Field(r, VTConsCol)),

                ItemsWithCodigoIns = Table.AddColumn(A6, "Codigo ins", each Text.From(Record.Field(_, ItemsCodColName)), type text),
                ItemsWithIns = Table.AddColumn(ItemsWithCodigoIns, "Ins", (r as record) =>
                    let
                        descIns = Record.Field(r, ItemsDescColName),
                        umIns   = Record.Field(r, ItemsUMColName),
                        dTxt0   = if descIns = null then "" else Text.Trim(Text.From(descIns)),
                        umTxt   = if umIns   = null then "" else Text.Trim(Text.From(umIns)),
                        baseTxt = if umTxt = "" then dTxt0 else dTxt0 & " (" & umTxt & ")"
                    in baseTxt, type text),

                OrigenAPU_Raw = try Excel.Workbook(BinarioPresupuesto, null, true){0}[Data]
                                otherwise Html.Table(FnDecodeHtml(BinarioPresupuesto), Columnas_APU, [RowSelector="tr"]),
                OrigenAPU_Cols = Table.SelectColumns(OrigenAPU_Raw, List.FirstN(Table.ColumnNames(OrigenAPU_Raw), 3)),
                OrigenAPU = Table.RenameColumns(OrigenAPU_Cols, List.Zip({Table.ColumnNames(OrigenAPU_Cols), {"Columna 1", "Columna 2", "Columna 3"}})),

                APU_Paso1 = Table.AddColumn(OrigenAPU, "Cod_Temp", each
                    let
                        c1Value = if [#"Columna 1"] = null then "" else [#"Columna 1"],
                        c1      = Text.Trim(Text.From(c1Value)),
                        hasDash = Text.Contains(c1, "-"),
                        preDash = if hasDash then Text.Trim(Text.BeforeDelimiter(c1, "-")) else "",
                        // El archivo APU mezcla, en la misma columna, filas de Actividad
                        // (codigo YA con punto, ej. "10.002") con filas de detalle de
                        // material/insumo del catalogo (codigo entero SIN punto, ej. "11016",
                        // "6400"). FnFormatCodigoAct le inserta un punto a los enteros sueltos
                        // ("11016" -> "11.016"), lo que hace que un codigo de MATERIAL choque
                        // por pura casualidad numerica con un codigo de ACTIVIDAD distinto.
                        // Por eso solo se acepta como codigo de actividad si YA trae el punto
                        // en el texto original: descarta los codigos de material del catalogo.
                        tienePunto = Text.Contains(preDash, "."),
                        esNum   = try Number.FromText(preDash) otherwise null
                    in if hasDash and esNum <> null and tienePunto then FnFormatCodigoAct(preDash) else null),

                APU_Paso2 = Table.SelectRows(APU_Paso1, each [Cod_Temp] <> null),

                APU_Diccionario = Table.AddColumn(APU_Paso2, "NombreActAPU", each
                    let
                        c1Value  = if [#"Columna 1"] = null then "" else [#"Columna 1"],
                        rawName  = Text.AfterDelimiter(Text.From(c1Value), "-"),
                        cleanName= Text.Trim(Text.Replace(Text.Replace(Text.Replace(rawName, "#(lf)", " "), "#(cr)", " "), "#(00A0)", " "))
                    in cleanName, type text),

                APU_DiccionarioLimpio = Table.SelectColumns(APU_Diccionario, {"Cod_Temp", "NombreActAPU", "Columna 3"}, MissingField.Ignore),
                APU_DiccionarioRenombrado = Table.RenameColumns(APU_DiccionarioLimpio,
                    List.Select({{"Cod_Temp", "CodigoActAPU"}, {"Columna 3", "UM_Actividad"}}, each Table.HasColumns(APU_DiccionarioLimpio, _{0}))),
                // Un mismo codigo puede tener MAS DE UNA actividad real distinta en el APU
                // (ej. "2.041" = "Esmaltado Parqueadero" (m2) en un bloque y "Cinta PVC" (ml)
                // en otro). No se reduce a un solo candidato por codigo (Table.Distinct no
                // garantiza cual "gana" de forma confiable ni estable entre refrescos); se
                // agrupan TODOS los candidatos por codigo y, para cada insumo, se elige el
                // candidato cuyo nombre realmente coincide con la descripcion propia de
                // SEGUIMIENTO en esa fila (misma fuente que el insumo, asi que es la senal
                // mas confiable para saber cual de las actividades es).
                DiccionarioAPU_Candidatos = Table.Buffer(Table.Group(APU_DiccionarioRenombrado, {"CodigoActAPU"}, {
                    {"CandidatosAPU", each Table.SelectColumns(_, {"NombreActAPU", "UM_Actividad"}), type table}
                })),

                ItemsJoinAPU       = Table.NestedJoin(ItemsWithIns, {"Codigo act"}, DiccionarioAPU_Candidatos, {"CodigoActAPU"}, "APU", JoinKind.LeftOuter),
                ItemsExpandedAPU0  = Table.ExpandTableColumn(ItemsJoinAPU, "APU", {"CandidatosAPU"}, {"CandidatosAPU"}),
                // Normaliza para comparar: SEGUIMIENTO trae el separador con guion normal
                // ("Cinta PVC - PARQUEADERO BLOQUE A") mientras que APU lo trae con NBSP y
                // sin guion ("Cinta PVC[NBSP]PARQUEADERO BLOQUE A"). Sin normalizar esa
                // diferencia, ninguno de los 2 textos "contiene" al otro y la coincidencia
                // nunca se detecta.
                FnNormalizarParaComparar = (t as any) as text =>
                    let
                        base      = Text.Upper(Text.Trim(Text.From(if t = null then "" else t))),
                        sinNBSP   = Text.Replace(base, "#(00A0)", " "),
                        sinGuion  = Text.Replace(sinNBSP, " - ", " "),
                        colapsado = Text.Combine(List.Select(Text.Split(sinGuion, " "), each _ <> ""), " ")
                    in colapsado,

                // Algunos nombres de actividad vienen del reporte con un guion final
                // colgando y nada detras (p.ej. "Remate muros - Aptos -"), sin que exista
                // Subcapitulo alguno que explique ese guion (viene asi de crudo en
                // DescActRaw/NombreActAPU). Se quita cualquier guion suelto al final del
                // nombre, de forma recursiva por si quedara mas de uno.
                FnQuitarGuionColgante = (t as text) as text =>
                    let
                        recortado = Text.Trim(t),
                        limpio    = if Text.EndsWith(recortado, "-")
                                    then @FnQuitarGuionColgante(Text.Range(recortado, 0, Text.Length(recortado) - 1))
                                    else recortado
                    in limpio,

                // Mismo problema que el guion final, pero al INICIO del nombre
                // (p.ej. "-SC - TOPELLANTAS (Un) - URBANISMO INTERIOR -" trae un
                // guion colgando antes de "SC" sin nada delante que lo explique).
                FnQuitarGuionInicial = (t as text) as text =>
                    let
                        recortado = Text.Trim(t),
                        limpio    = if Text.StartsWith(recortado, "-")
                                    then @FnQuitarGuionInicial(Text.Range(recortado, 1))
                                    else recortado
                    in limpio,

                // Algunos nombres traen la unidad ya incrustada en el texto libre
                // (ej. "TOPELLANTAS (Un) - URBANISMO INTERIOR"), redundante con la
                // unidad que esta misma consulta agrega al final entre parentesis.
                // Si el parentesis embebido coincide (sin distinguir mayusculas) con
                // la unidad final, se quita para no duplicarla; si no coincide (ej.
                // "(bloque fachada)" cuando la unidad final es "M2") se deja intacto
                // porque es contenido real del nombre, no una unidad repetida.
                FnQuitarUnidadEmbebida = (t as text, um as text) as text =>
                    let
                        patron    = "(" & um & ")",
                        tUpper    = Text.Upper(t),
                        pos       = if um = "" then -1 else Text.PositionOf(tUpper, Text.Upper(patron)),
                        sinUnidad = if pos < 0 then t else Text.RemoveRange(t, pos, Text.Length(patron)),
                        colapsado = Text.Combine(List.Select(Text.Split(sinUnidad, " "), each _ <> ""), " ")
                    in colapsado,

                ItemsConAPUElegido = Table.AddColumn(ItemsExpandedAPU0, "APUElegido", each
                    let
                        candidatos      = [CandidatosAPU],
                        hayCandidatos   = candidatos <> null and Table.RowCount(candidatos) > 0,
                        descPropia      = FnNormalizarParaComparar(if [DescActRaw] = null then "" else [DescActRaw]),
                        conCoincidencia = if not hayCandidatos or descPropia = "" then null
                            else Table.SelectRows(candidatos, each
                                let nombreNorm = FnNormalizarParaComparar([NombreActAPU])
                                in Text.Contains(descPropia, nombreNorm) or Text.Contains(nombreNorm, descPropia)),
                        tieneCoincidencia = conCoincidencia <> null and Table.RowCount(conCoincidencia) > 0
                    in
                        if not hayCandidatos then [NombreActAPU = null, UM_Actividad = null]
                        else if tieneCoincidencia then conCoincidencia{0}
                        else candidatos{0}),
                ItemsExpandedAPU   = Table.ExpandRecordColumn(ItemsConAPUElegido, "APUElegido", {"NombreActAPU", "UM_Actividad"}, {"NombreActAPU", "UM_Actividad"}),

                // Subcapitulo DERIVADO del nombre de la actividad, solo cuando el proyecto
                // no declara subcapitulos en el SEGUIMIENTO (ej. TURPIAL): el reporte de
                // "Analisis De Precios Unitarios Items Presupuesto Detallado" agrega el
                // subcapitulo como sufijo tras el ultimo " - " ("1.01 - COMISION ... - GENERALES").
                // Se toma la cola despues del ULTIMO " - " como subcapitulo; el recorte del
                // nombre lo hace la logica ya existente en ItemsWithActividad (patron conGuion).
                // Helpers de subcapitulo embebido: definidos a nivel de F_Globales
                // (los usa tambien DISPONIBLE para los POR ADJUDICAR); aqui solo un alias.
                FnQuitarSufijosIgnorados = FnQuitarSufijosSubcapIgnorados,

                ItemsConSubcapDerivado = Table.AddColumn(ItemsExpandedAPU, "SubcapDerivado", each
                    let
                        subcapSeg = if [Subcapitulo] = null then "" else Text.Trim(Text.From([Subcapitulo])),
                        aplica    = (TieneSubcapExplicito = false) and subcapSeg = "",
                        descRaw   = Text.Trim(Text.Replace(Text.From(if [DescActRaw] = null then "" else [DescActRaw]), "#(00A0)", " ")),
                        descTxt   = Text.Combine(List.Select(Text.Split(descRaw, " "), each _ <> ""), " "),
                        apuRaw    = Text.Trim(Text.Replace(Text.From(if [NombreActAPU] = null then "" else [NombreActAPU]), "#(00A0)", " ")),
                        apuTxt    = Text.Combine(List.Select(Text.Split(apuRaw, " "), each _ <> ""), " "),
                        desdeDesc = if (not aplica) or descTxt = "" then null else FnExtraerSubcapDeTexto(descTxt),
                        // Un "(" sin ")" en la cola delata texto TRUNCADO por el reporte
                        // ("APTOS (U" en vez de "...APTOS (Un) - TORRES"): en ese caso se
                        // intenta con el nombre del APU (otro reporte, normalmente completo).
                        truncado  = desdeDesc <> null and Text.Contains(desdeDesc, "(") and not Text.Contains(desdeDesc, ")"),
                        desdeApu  = if (not aplica) or apuTxt = "" then null else FnExtraerSubcapDeTexto(apuTxt),
                        elegido   = if not aplica then null
                                    else if desdeDesc = null or truncado then (if desdeApu <> null then desdeApu else desdeDesc)
                                    else desdeDesc
                    in FnAplicarOverrideSubcap(elegido), type text),

                // Canonicaliza subcapitulos derivados TRUNCADOS por el reporte (SINCO corta
                // el texto en algunas filas: "ELE", "ELEC", "ELECTRIC" en vez de "ELECTRICO";
                // "SALON SOCIA" en vez de "SALON SOCIAL"). Regla: si un valor derivado es
                // PREFIJO (sin tildes, sin mayusculas) de otro valor derivado MAS LARGO y
                // MAS FRECUENTE del mismo archivo, se reemplaza por el completo. Un subcapitulo
                // real corto no se toca salvo que exista uno mas largo que empiece igual y
                // tenga mas filas (las truncaduras son artefactos de pocas filas).
                SubcapDerivadosLista = List.Buffer(
                    let
                        vals = List.RemoveNulls(Table.Column(ItemsConSubcapDerivado, "SubcapDerivado")),
                        dist = List.Distinct(vals)
                    in List.Transform(dist, (d) => [V = d, N = List.Count(List.Select(vals, (x) => x = d)), Norm = Text.Upper(FnRemoveAccentsSymbols(d))])),

                FnCanonSubcap = (v as nullable text) as nullable text =>
                    if v = null then null else
                    let
                        normV  = Text.Upper(FnRemoveAccentsSymbols(v)),
                        propio = List.First(List.Select(SubcapDerivadosLista, each [Norm] = normV), [N = 0]),
                        cands  = List.Select(SubcapDerivadosLista, each [Norm] <> normV and Text.StartsWith([Norm], normV) and [N] > propio[N]),
                        mejor  = if List.Count(cands) = 0 then null
                                 else List.Accumulate(cands, null, (s, c) => if s = null or Text.Length(c[Norm]) > Text.Length(s[Norm]) then c else s)
                    in if mejor = null then v else mejor[V],

                ItemsWithActividad = Table.AddColumn(ItemsConSubcapDerivado, "Actividad", each
                    let
                        codTxt        = if [Codigo act]  = null then "" else [Codigo act],
                        // Prioridad: primero la descripcion real de SEGUIMIENTO (misma fuente
                        // que el insumo, siempre correcta para ese codigo); si viene vacia,
                        // el nombre de APU; si tampoco hay, un texto generico.
                        // Se limpia el NBSP y se colapsan espacios aqui mismo (no solo en el
                        // codigo) para que el patron " - Subcapitulo" se detecte de forma
                        // fiable mas abajo, sin importar espacios/caracteres invisibles sueltos
                        // pegados en el texto libre del reporte.
                        descSegRaw    = Text.Trim(Text.Replace(Text.From(if [DescActRaw] = null then "" else [DescActRaw]), "#(00A0)", " ")),
                        descSegTxt    = Text.Combine(List.Select(Text.Split(descSegRaw, " "), each _ <> ""), " "),
                        nombreExtraido= Text.Trim(Text.From(if [NombreActAPU] = null then "" else [NombreActAPU])),
                        nombreReal0   = if descSegTxt <> "" then descSegTxt
                                        else if nombreExtraido <> "" then nombreExtraido
                                        else "Actividad " & codTxt,
                        // Los sufijos ignorados ("- ELECTRICO") tampoco van en el nombre.
                        nombreReal    = FnQuitarSufijosIgnorados(nombreReal0),
                        // Subcapitulo normalizado igual que la descripcion (NBSP->espacio,
                        // espacios colapsados) para que la busqueda de mas abajo no falle por
                        // una diferencia de espacios/caracteres invisibles entre los 2 campos
                        // (vienen de columnas distintas del mismo reporte, no siempre coinciden
                        // caracter por caracter aunque se vean iguales).
                        subcapFuenteSeg = if [Subcapitulo] = null then "" else Text.Trim(Text.From([Subcapitulo])),
                        subcapFuente  = if subcapFuenteSeg <> "" then subcapFuenteSeg
                                        else Text.From(if [SubcapDerivado] = null then "" else [SubcapDerivado]),
                        subcapRaw     = Text.Trim(Text.Replace(subcapFuente, "#(00A0)", " ")),
                        subcapTxt     = Text.Combine(List.Select(Text.Split(subcapRaw, " "), each _ <> ""), " "),
                        // Si el Subcapitulo viene pegado con un guion separador ("Texto - SUBCAP"),
                        // se quita el bloque completo (guion incluido) para no dejar un guion
                        // huerfano colgando antes de la unidad. Si no aparece con ese patron
                        // exacto, se cae al comportamiento anterior (quitar solo el texto).
                        // La busqueda es insensible a mayusculas (Subcapitulo y la descripcion
                        // no siempre coinciden en mayusculas/minusculas), pero el recorte se
                        // hace sobre el texto ORIGINAL para no alterar su capitalizacion real.
                        conGuion      = if subcapTxt = "" then "" else " - " & subcapTxt,
                        nombreRealUpper = Text.Upper(nombreReal),
                        posConGuion   = if conGuion = "" then -1 else Text.PositionOf(nombreRealUpper, Text.Upper(conGuion)),
                        posSubcap     = if subcapTxt = "" then -1 else Text.PositionOf(nombreRealUpper, Text.Upper(subcapTxt)),
                        nombreSinSub  = if subcapTxt = "" then nombreReal
                                        else if posConGuion >= 0 then Text.RemoveRange(nombreReal, posConGuion, Text.Length(conGuion))
                                        else if posSubcap >= 0 then Text.RemoveRange(nombreReal, posSubcap, Text.Length(subcapTxt))
                                        else nombreReal,
                        umTxt         = Text.Trim(Text.From(if [UM_Actividad] = null then "" else [UM_Actividad])),
                        nombreColapsado = Text.Combine(List.Select(Text.Split(nombreSinSub, " "), each _ <> ""), " "),
                        // Limpieza en 2 pasadas del texto crudo de la actividad (nada de
                        // esto tiene relacion con Subcapitulo, ya se quito arriba si aplicaba):
                        // 1) quitar la unidad si ya viene incrustada en el nombre, redundante
                        //    con la que se agrega mas abajo entre parentesis;
                        // 2) quitar guiones sueltos al inicio y/o al final que a veces
                        //    vienen asi de crudo en el reporte (ver funciones mas arriba).
                        nombreSinUnidad  = FnQuitarUnidadEmbebida(nombreColapsado, umTxt),
                        nombreLimpio  = FnQuitarGuionInicial(FnQuitarGuionColgante(nombreSinUnidad)),
                        actTxt        = if umTxt = "" then codTxt & "-" & nombreLimpio
                                        else codTxt & "-" & nombreLimpio & " (" & umTxt & ")"
                    in actTxt, type text),

                // Unifica el Subcapitulo: el explicito del SEGUIMIENTO gana; si no hay,
                // usa el derivado del sufijo del nombre (proyectos tipo TURPIAL).
                SubcapUnificado0 = Table.AddColumn(ItemsWithActividad, "SubcapituloFinal", each
                    let s = if [Subcapitulo] = null then "" else Text.Trim(Text.From([Subcapitulo]))
                    in if s <> "" then s else FnCanonSubcap([SubcapDerivado]), type text),
                SubcapUnificado = Table.RenameColumns(
                    Table.RemoveColumns(SubcapUnificado0, {"Subcapitulo", "SubcapDerivado"}),
                    {{"SubcapituloFinal", "Subcapitulo"}}),

                NumsTyped = Table.TransformColumns(SubcapUnificado, {
                    {"Cantidad Presupuesto", each FxToNumberFlex(_), type number},
                    {"VT Presupuesto",       each FxToNumberFlex(_), Currency.Type},
                    {"Cantidad Proyectado",  each FxToNumberFlex(_), type number},
                    {"VT Proyectado",        each FxToNumberFlex(_), Currency.Type},
                    {"Cantidad Consumido",   each FxToNumberFlex(_), type number},
                    {"VT Consumido",         each FxToNumberFlex(_), Currency.Type}
                }),

                Final = Table.SelectColumns(NumsTyped, {
                    "Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo",
                    "Cantidad Presupuesto", "VT Presupuesto",
                    "Cantidad Proyectado",  "VT Proyectado",
                    "Cantidad Consumido",   "VT Consumido"
                })
            in Final,

        // Procesa "Masivo salidas DESCRIPTIVAS" (formato plano nuevo: una fila completa
        // por insumo, sin bloques repetidos meta/subheader/total, no necesita fill-down).
        // ColumnasBase debe venir del llamador (mismo shape que usa FxProcesarSalidas).
        FxProcesarSalidasDescriptivas = (BinSalidas as binary, ColumnasBase as list) as table =>
            let
                FnText = (v as any) as text => try Text.Trim(Text.From(if v = null then "" else v)) otherwise "",
                FnCleanDisplay = (v as any) as nullable text =>
                    let t = FnText(v), clean = if t = "" then null else Text.Upper(Text.Trim(FnRemoveAccentMarks(t)))
                    in clean,
                FnBuildInsUM = (desc as any, um as any) as nullable text =>
                    let d = FnCleanDisplay(desc), u = FnCleanDisplay(um)
                    in if d = null then null else if u = null or u = "" then d else d & " (" & u & ")",
                FnCleanContratistaFromDash = (v as any) as nullable text =>
                    let t = FnText(v), afterDash = if Text.Contains(t, "-") then Text.Trim(Text.AfterDelimiter(t, "-")) else t, clean = FnCleanDisplay(afterDash)
                    in clean,
                Columnas = FnBuildColumnas(13),
                Raw = Table.Buffer(Html.Table(FnDecodeHtml(BinSalidas), Columnas, [RowSelector="tr"])),
                AddStd = Table.AddColumn(Raw, "Std", each
                    let
                        salidaNo = FnText(Record.Field(_, "Columna 2")),
                        contratista = Record.Field(_, "Columna 4"),
                        codigoIns = FnText(Record.Field(_, "Columna 7")),
                        descripcion = Record.Field(_, "Columna 8"),
                        item = Record.Field(_, "Columna 9"),
                        um = Record.Field(_, "Columna 10"),
                        cant = Record.Field(_, "Columna 11"),
                        vrTotal = Record.Field(_, "Columna 13"),
                        insFinal = FnBuildInsUM(descripcion, um),
                        codAct = FnFormatCodigoAct(item)
                    in [
                        #"Codigo ins" = codigoIns,
                        Ins = insFinal,
                        Actividad = null,
                        #"Codigo act" = codAct,
                        InsClave = FnClaveLimpia(insFinal),
                        #"# OC / Contrato" = null,
                        #"Cantidad Comprado" = null,
                        #"VT Comprado" = null,
                        VU_Crudo = null,
                        IVA_Crudo = null,
                        #"Nombre Contratista" = FnCleanContratistaFromDash(contratista),
                        #"#ENTRADA" = null,
                        #"Cantidad Cortes" = null,
                        #"VT Cortes" = null,
                        #"#SALIDA" = salidaNo,
                        #"Cantidad Cons Cols" = FxToNumberFlex(cant),
                        #"VT Cons Cols" = FxToNumberFlex(vrTotal)
                    ]),
                Expanded = Table.ExpandRecordColumn(AddStd, "Std", ColumnasBase, ColumnasBase),
                Filtrado = Table.SelectRows(Expanded, each try (Number.FromText([#"Codigo ins"]) <> null) otherwise false),
                Selected = Table.SelectColumns(Filtrado, ColumnasBase, MissingField.UseNull)
            in Selected,

        // Procesa "Informe entradas por insumo" (formato plano nuevo: una fila completa
        // por entrada, sin bloques repetidos meta/item que hay que reconstruir con
        // FillDown como en FxProcesarEntradas). Mismo patron que FxProcesarSalidasDescriptivas.
        // Columnas reales del reporte: 1 Sucursal, 2 Cod, 3 Descripcion, 4 Agrupacion, 5 UM,
        // 6 No. OC, 7 No. EA, 8 Fecha, 9 Cantidad, 10 Vr. Unitario, 11 IVA, 12 Valor Total,
        // 13 Proveedor, 14 Obs.
        FxProcesarEntradasPorInsumo = (BinEntradas as binary, ColumnasBase as list) as table =>
            let
                FnText = (v as any) as text => try Text.Trim(Text.From(if v = null then "" else v)) otherwise "",
                FnCleanDisplay = (v as any) as nullable text =>
                    let t = FnText(v), clean = if t = "" then null else Text.Upper(Text.Trim(FnRemoveAccentMarks(t)))
                    in clean,
                FnBuildInsUM = (desc as any, um as any) as nullable text =>
                    let d = FnCleanDisplay(desc), u = FnCleanDisplay(um)
                    in if d = null then null else if u = null or u = "" then d else d & " (" & u & ")",
                FnCleanContratistaFromDash = (v as any) as nullable text =>
                    let t = FnText(v), afterDash = if Text.Contains(t, "-") then Text.Trim(Text.AfterDelimiter(t, "-")) else t, clean = FnCleanDisplay(afterDash)
                    in clean,
                Columnas = FnBuildColumnas(14),
                Raw = Table.Buffer(Html.Table(FnDecodeHtml(BinEntradas), Columnas, [RowSelector="tr"])),
                AddStd = Table.AddColumn(Raw, "Std", each
                    let
                        codigoIns = FnText(Record.Field(_, "Columna 2")),
                        descripcion = Record.Field(_, "Columna 3"),
                        um = Record.Field(_, "Columna 5"),
                        ocNo = FnText(Record.Field(_, "Columna 6")),
                        eaNo = FnText(Record.Field(_, "Columna 7")),
                        cantidad = Record.Field(_, "Columna 9"),
                        valorTotal = Record.Field(_, "Columna 12"),
                        proveedor = Record.Field(_, "Columna 13"),
                        insFinal = FnBuildInsUM(descripcion, um)
                    in [
                        #"Codigo ins" = codigoIns,
                        Ins = insFinal,
                        Actividad = null,
                        #"Codigo act" = null,
                        InsClave = FnClaveLimpia(insFinal),
                        #"# OC / Contrato" = ocNo,
                        #"Cantidad Comprado" = null,
                        #"VT Comprado" = null,
                        VU_Crudo = null,
                        IVA_Crudo = null,
                        #"Nombre Contratista" = FnCleanContratistaFromDash(proveedor),
                        #"#ENTRADA" = eaNo,
                        #"Cantidad Cortes" = FxToNumberFlex(cantidad),
                        #"VT Cortes" = FxToNumberFlex(valorTotal),
                        #"#SALIDA" = null,
                        #"Cantidad Cons Cols" = null,
                        #"VT Cons Cols" = null
                    ]),
                Expanded = Table.ExpandRecordColumn(AddStd, "Std", ColumnasBase, ColumnasBase),
                Filtrado = Table.SelectRows(Expanded, each try (Number.FromText([#"Codigo ins"]) <> null) otherwise false),
                Selected = Table.SelectColumns(Filtrado, ColumnasBase, MissingField.UseNull)
            in Selected
    ]
in
    Funciones
```

## ITEMSINSUMOS

```powerquery
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
```

## PPTO_BD

```powerquery
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
```

## PROVISIONES_SP

```powerquery
let
    // ============================================================
    // FUNCIONES GLOBALES
    // ============================================================
    FxToNumberFlex = F_Globales[FxToNumberFlex],
    FnRemoveAccentsSymbols = F_Globales[FnRemoveAccentsSymbols],
    FnMatchFolder = F_Globales[FnMatchFolder],
    FnReadSPExcel = F_Globales[FnReadSPExcel],
    FnBuildFolderPrefixMap = F_Globales[FnBuildFolderPrefixMap],
    FnTrimText = F_Globales[FnTrimText],

    ParamProyecto = Text.Trim(ProyectoActual),

    // ============================================================
    // CARPETAS DE PROYECTO ACTUAL (sin filtro de archivos)
    // ============================================================
    ListaCarpetas = try List.Distinct(SP_CarpetasCC[Centro de Costos]) otherwise {},
    PrefixMap = FnBuildFolderPrefixMap(ListaCarpetas),

    // ============================================================
    // CONEXION AL ARCHIVO EN SHAREPOINT
    // ============================================================
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    FilePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. Provisiones - Control costos interno/0. CONSOLIDADOR PROVISIONES SP.xlsx",

    Origen = let t = FnReadSPExcel(SiteUrl, FilePath) in if t = null then #table({"PROYECTO"}, {}) else t,

    // ============================================================
    // FILTRO Y MAPEO
    // ============================================================
    // 1. Filtrar por el proyecto actual y por Tipo = DIRECTOS
    FiltroProyecto = Table.SelectRows(Origen, each
        try [PROYECTO] <> null and (
            Text.StartsWith(Text.Upper([PROYECTO]), Text.Upper(ParamProyecto)) or
            Text.Contains(FnRemoveAccentsSymbols(Text.Upper([PROYECTO])), FnRemoveAccentsSymbols(Text.Upper(ParamProyecto)))
        ) otherwise false
    ),

    FiltroDirectos = Table.SelectRows(FiltroProyecto, each
        try [Tipo] <> null and Text.Upper(Text.Trim([Tipo])) = "DIRECTOS" otherwise false
    ),

    // 2. Renombrar columnas clave hacia BD.m
    ColumnasRenombradas = Table.RenameColumns(FiltroDirectos, {
        {"Nombre_prov", "Nombre Contratista"},
        {"No_Orden_contrato", "# OC / Contrato"}
    }, MissingField.Ignore),

    // 3. Estandarizacion de tipos de datos
    TextosLimpios = Table.TransformColumns(ColumnasRenombradas, {
        {"PROYECTO", each FnTrimText(_), type text},
        {"Nombre Contratista", each FnTrimText(_), type text},
        {"# OC / Contrato", each FnTrimText(_), type text},
        {"VR_Bruto_con_desc", each FxToNumberFlex(_), type number}
    }, null, MissingField.Ignore),

    // 4. Agregar la etiqueta Tipo y Centro de Costos
    SinTipoOriginal = Table.RemoveColumns(TextosLimpios, {"Tipo"}, MissingField.Ignore),
    AgregadoTipo = Table.AddColumn(SinTipoOriginal, "Tipo", each "PROVISIONES", type text),
    AgregadoCC = Table.AddColumn(AgregadoTipo, "Centro de Costos", each
        if Text.StartsWith(Text.Upper(ParamProyecto), "PAMPLONA 1") and [#"# OC / Contrato"] <> null then
            let
                ccPrefix = Text.Start([#"# OC / Contrato"], 4),
                fromMap = try Record.Field(PrefixMap, ccPrefix) otherwise null
            in
                if fromMap <> null then fromMap else FnMatchFolder([PROYECTO], ListaCarpetas)
        else
            FnMatchFolder([PROYECTO], ListaCarpetas)
    , type text)
in
    AgregadoCC
```

## SINCO

```powerquery
let
    SourceRaw = BD,
    TablaComparativo = COMPARATIVOS,
    FnRemoveAccentsSymbols = F_Globales[FnRemoveAccentsSymbols],

    Source = Table.ReplaceErrorValues(SourceRaw, List.Transform(Table.ColumnNames(SourceRaw), each {_, null})),

    ToTextClean = (v as any) as text =>
        let t = try Text.Trim(Text.From(v)) otherwise ""
        in if t = null then "" else t,

    ToNumber0 = (v as any) as number =>
        let n = try Number.From(v) otherwise null
        in if n = null then 0 else n,

    GetField = (r as record, fieldName as text, defaultValue as any) as any =>
        let value = try Record.Field(r, fieldName) otherwise defaultValue
        in if value = null then defaultValue else value,

    ListaOC_Excluir = List.Distinct(
        List.RemoveNulls(
            List.Transform(
                try TablaComparativo[#"# OC / Contrato"] otherwise {},
                each let oc = ToTextClean(_) in if oc = "" then null else oc
            )
        )
    ),

    SetOC =
        if List.Count(ListaOC_Excluir) = 0
        then []
        else Record.FromList(List.Repeat({true}, List.Count(ListaOC_Excluir)), ListaOC_Excluir),

    BaseConValor = Table.SelectRows(Source, each
        Text.Upper(ToTextClean(GetField(_, "Tipo", ""))) <> "PPTO" and
        ToNumber0(GetField(_, "VT Asegurada", 0)) <> 0
    ),

    FiltradoPorOC = Table.SelectRows(BaseConValor, each
        let ocText = ToTextClean(GetField(_, "# OC / Contrato", ""))
        in ocText = "" or not Record.HasFields(SetOC, {ocText})
    ),

    // Si COMPARATIVOS excluye todo, se conserva BaseConValor para evitar SINCO en 0 filas.
    BaseSINCO = if Table.RowCount(FiltradoPorOC) > 0 then FiltradoPorOC else BaseConValor,

    LimpiezaTextos = Table.TransformColumns(BaseSINCO, {
        {"Nombre Contratista", each FnRemoveAccentsSymbols(if _ = null then null else Text.Trim(Text.From(_))), type text},
        {"Descripcion contrato", each FnRemoveAccentsSymbols(if _ = null then null else Text.Trim(Text.From(_))), type text}
    }, null, MissingField.Ignore),

    ColumnasFinales = Table.SelectColumns(LimpiezaTextos,
        {"Centro de Costos", "Subcapitulo", "Capitulo", "Actividad", "Codigo ins", "Ins",
         "# OC / Contrato", "Nombre Contratista", "Cantidad asegurada", "V/U asegurada",
         "VT Asegurada", "Descripcion contrato", "Tipo"}, MissingField.Ignore),

    SinErrores = Table.ReplaceErrorValues(ColumnasFinales, List.Transform(Table.ColumnNames(ColumnasFinales), each {_, null})),
    ResultadoFinal = Table.Buffer(SinErrores)
in
    ResultadoFinal
```

## SP_Archivos_Proyecto

```powerquery
let
    ParamProyecto = Text.Trim(ProyectoActual),
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    BasePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. Reportes EDT - Control costos interno/" & ParamProyecto,
    Headers = [Accept="application/json;odata=nometadata"],
    FnEncode = F_Globales[FnEncode],

    // Reutiliza SP_CarpetasCC (consulta compartida) en vez de repetir la MISMA
    // llamada GetFolderByServerRelativeUrl(...)/Folders por separado - evita un
    // viaje de red redundante (Power Query calcula SP_CarpetasCC una sola vez
    // y la comparte entre todos sus consumidores en el mismo refresco).
    CCFolders =
        try Table.RenameColumns(SP_CarpetasCC, {{"Centro de Costos", "Name"}})
        otherwise #table({"Name"}, {}),

    WithFiles = Table.AddColumn(CCFolders, "Archivos", each
        let
            ccActualPath = BasePath & "/" & [Name] & "/Actual",
            result = try Json.Document(Web.Contents(SiteUrl, [
                RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(ccActualPath) & "')/Files",
                Query = [#"$select" = "Name,ServerRelativeUrl,TimeLastModified,Length"],
                Headers = Headers,
                Timeout = #duration(0, 0, 2, 0)
            ])) otherwise null
        in
            if result <> null and Record.HasFields(result, "value") then Table.FromRecords(result[value]) else null
    ),

    ValidCCs = Table.SelectRows(WithFiles, each [Archivos] <> null),
    Expanded = Table.ExpandTableColumn(
        ValidCCs,
        "Archivos",
        {"Name", "ServerRelativeUrl", "TimeLastModified", "Length"},
        {"FileName", "ServerRelativeUrl", "TimeLastModified", "Length"}
    ),

    Relevant = Table.SelectRows(Expanded, each
        not Text.StartsWith([FileName], "~$") and (
            Text.Contains([FileName], "SEGUIMIENTO POR ITEMS",         Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "ANALISIS DE PRECIOS UNITARIOS", Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "INFORMEORDEN",                  Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "ESTADO DE ORDENES",             Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "INFORME ENTRADAS DE ALMACEN",   Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "INFORME ENTRADAS DE ALMACÉN",   Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "ENTRADAS POR INSUMO",           Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "MASIVO SALIDAS",                Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "ESTADO DE CONTRATOS",           Comparer.OrdinalIgnoreCase) or
            Text.Contains([FileName], "DESCUENTOS",                    Comparer.OrdinalIgnoreCase)
        )
    ),

    Typed = Table.TransformColumnTypes(Relevant, {{"TimeLastModified", type datetimezone}, {"Length", Int64.Type}}, "en-US"),
    Sorted = Table.Sort(Typed, {{"Name", Order.Ascending}, {"FileName", Order.Ascending}, {"TimeLastModified", Order.Descending}}),
    Final = Table.Buffer(Table.RenameColumns(
        Table.SelectColumns(Sorted, {"Name", "FileName", "ServerRelativeUrl", "TimeLastModified", "Length"}),
        {{"Name", "Centro de Costos"}, {"FileName", "Name"}}
    ))
in
    Final
```

## SP_CarpetasCC

```powerquery
let
    // Lista TODAS las carpetas (Centro de Costos) del proyecto actual,
    // sin filtrar por presencia de archivos. Usado por APROBACIONES_SP
    // y PROVISIONES_SP para mapear nombres de proyecto -> CC sin perder
    // CCs que no tengan archivos SEGUIMIENTO/INFORMEORDEN/etc.
    ParamProyecto = Text.Trim(ProyectoActual),
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    BasePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. Reportes EDT - Control costos interno/" & ParamProyecto,
    Headers = [Accept="application/json;odata=nometadata"],
    FnEncode = F_Globales[FnEncode],

    // Web.Contents con RelativePath/Query: DataSourcePath estable, sin problemas de cache
    Respuesta = Json.Document(Web.Contents(SiteUrl, [
        RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(BasePath) & "')/Folders",
        Query = [#"$select" = "Name"],
        Headers = Headers
    ])),
    AsTable = Table.FromRecords(Respuesta[value]),
    Renamed = Table.RenameColumns(AsTable, {{"Name", "Centro de Costos"}}),
    Final = Table.Buffer(Renamed)
in
    Final
```

## SP_Seguimiento_Parsed

```powerquery
let
    // =========================================================
    // Parseo unificado Seguimiento + APU. La logica vive en
    // F_Globales[FxProcesarCentroCosto] para no duplicarse con
    // PPTO_TODOS_PROYECTOS.
    // =========================================================
    FxProcesarCentroCosto = F_Globales[FxProcesarCentroCosto],
    FnReadSPBinary = F_Globales[FnReadSPBinary],
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",

    ConCentroCosto = SP_Archivos_Proyecto,

    PickLatestBinary = (t as table, containsText as text) as nullable binary =>
        let
            candidatos = Table.Sort(
                Table.SelectRows(t, each Text.Contains([Name], containsText, Comparer.OrdinalIgnoreCase)),
                {{"TimeLastModified", Order.Descending}, {"Name", Order.Ascending}}
            ),
            path = if Table.RowCount(candidatos) = 0 then null else candidatos{0}[ServerRelativeUrl]
        in
            if path = null then null else FnReadSPBinary(SiteUrl, path),

    Agrupado = Table.Group(ConCentroCosto, {"Centro de Costos"}, {{"Binarios", each
        let
            binPres = PickLatestBinary(_, "ANALISIS DE PRECIOS UNITARIOS"),
            binSeg = PickLatestBinary(_, "SEGUIMIENTO POR ITEMS")
        in
            if binPres <> null and binSeg <> null then [Bin_P = binPres, Bin_S = binSeg] else null
    }}),
    CentrosCompletos = Table.SelectRows(Agrupado, each [Binarios] <> null),
    TablaConDatos = Table.AddColumn(CentrosCompletos, "Datos", each FxProcesarCentroCosto([Binarios][Bin_S], [Binarios][Bin_P])),
    Expandido = Table.ExpandTableColumn(TablaConDatos, "Datos", {"Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo", "Cantidad Presupuesto", "VT Presupuesto", "Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido"}),
    ColumnasUtiles = Table.SelectColumns(Expandido, {"Centro de Costos", "Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo", "Cantidad Presupuesto", "VT Presupuesto", "Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido"}),
    TiposFinales = Table.TransformColumnTypes(ColumnasUtiles,{{"Centro de Costos", type text}, {"Codigo ins", Int64.Type}, {"Cantidad Presupuesto", type number}, {"VT Presupuesto", Currency.Type}, {"Cantidad Proyectado", type number}, {"VT Proyectado", Currency.Type}, {"Cantidad Consumido", type number}, {"VT Consumido", Currency.Type}}),
    // BUFFER MAESTRO: ITEMSINSUMOS y PPTO_BD comparten este resultado
    Resultado = Table.Buffer(TiposFinales)
in
    Resultado
```
