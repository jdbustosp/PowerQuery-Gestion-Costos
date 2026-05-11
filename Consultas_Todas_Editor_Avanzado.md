# Consultas Power Query - Editor avanzado

Copia cada bloque en la consulta con el mismo nombre.

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
    FiltroProyecto = Table.SelectRows(Origen, each
        try [#"Proyecto:"] <> null and Text.StartsWith(Text.Upper([#"Proyecto:"]), Text.Upper(ParamProyecto)) otherwise false
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
        {"# CC - Comparativo", each FnTrimText(_), type text},
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
    ToNumber0 = (v as any) as number =>
        let n = try Number.From(v) otherwise null
        in if n = null then 0 else n,

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

    // try/otherwise en cada transformacion: si Text.From recibe algo raro (record, list, etc) no rompe
    LlavesLimpias = Table.TransformColumns(ColumnasReordenadas, {
        {"Centro de Costos",     each try (if _ = null then "" else Text.Upper(Text.Trim(Text.From(_)))) otherwise "", type text},
        {"Codigo act",           each try (if _ = null then "" else Text.Upper(Text.Trim(Text.From(_)))) otherwise "", type text},
        {"Ins",                  each try (if _ = null then "" else Text.Upper(Text.Trim(Text.From(_)))) otherwise "", type text},
        {"Tipo",                 each try (if _ = null then "" else Text.Upper(Text.Trim(Text.From(_)))) otherwise "", type text},
        {"# OC / Contrato",      each try (if _ = null then null else Text.Trim(Text.From(_))) otherwise null, type text},
        {"Nombre Contratista",   each try FnCleanContratista(_) otherwise null, type text},
        {"Descripcion contrato", each try FnRemoveAccentsSymbols(if _ = null then null else Text.Trim(Text.From(_))) otherwise null, type text}
    }, null, MissingField.Ignore),

    FiltroTipoValido = Table.SelectRows(LlavesLimpias, each [Tipo] <> null and [Tipo] <> ""),

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
    TablaMaestraFinal = Table.Buffer(FinalSinErrores)
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
        {"# CC - Comparativo", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"# CC", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Comparativo", each if _ = null then null else Text.Trim(Text.From(_)), type text},
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
    FnPrepareTableWithHeader = F_Globales[FnPrepareTableWithHeader],
    FxToNumberFlex = F_Globales[FxToNumberFlex],
    FnClaveLimpia = F_Globales[FnClaveLimpia],
    FnMapColumn = F_Globales[FnMapColumn],
    FnReadSPBinary = F_Globales[FnReadSPBinary],
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    Columnas_OC = F_Globales[FnBuildColumnas](10),

    // ============================================================
    // FUNCIÓN MÁGICA: PROCESAR COMPRAS
    // ============================================================
    FxProcesarCompras = (BinDetalles as binary, BinOC as binary) => let
        // 🚀 Excel.Workbook es más rápido que Html.Table
        RawOC_Raw = try Excel.Workbook(BinOC, null, true){0}[Data]
                otherwise Html.Table(Text.FromBinary(Binary.Buffer(BinOC), 65001), Columnas_OC, [RowSelector="tr"]),
        // Estandarizar nombres de columnas
        RawOC_ColNames = Table.ColumnNames(RawOC_Raw),
        RawOC = Table.RenameColumns(RawOC_Raw, List.Zip({RawOC_ColNames, List.Transform({1..List.Count(RawOC_ColNames)}, each "Columna" & Text.From(_))})),

        AddOCKey = Table.AddColumn(RawOC, "OC_Key_Temp", each let v = Text.From(if [Columna1] = null then "" else [Columna1]) in if Text.StartsWith(v, "Orden de Compra No.") then Text.Trim(Text.Replace(v, "Orden de Compra No.", "")) else null, type text),
        Ordenes_Agrupadas = Table.RenameColumns(Table.Group(Table.SelectRows(Table.FillDown(AddOCKey, {"OC_Key_Temp"}), each [OC_Key_Temp] <> null), {"OC_Key_Temp"}, {{"Proveedor_Raw", each let l = List.RemoveNulls([Columna2]), l2 = List.Select(l, (x) => let t = Text.Trim(Text.From(if x = null then "" else x)) in t <> "Proveedor" and t <> "Insumo") in if List.IsEmpty(l2) then null else List.First(l2), type text}}), {{"OC_Key_Temp", "OC_Key"}}),

        LibroExcel = Excel.Workbook(Binary.Buffer(BinDetalles), null, true),
        DetallesCrudos = FnPrepareTableWithHeader(LibroExcel{0}[Data]),
        Cols = Table.ColumnNames(DetallesCrudos),
        MapStd = Table.AddColumn(DetallesCrudos, "Std", each [ Codigo_ins = FnMapColumn(_, Cols, {"CÓDIGO", "CODIGO", "COD."}), Ins = FnMapColumn(_, Cols, {"INSUMO", "DESCRIPCIÓN", "DESCRIPCION"}), Act = FnMapColumn(_, Cols, {"ACTIVIDAD", "DESTINO", "FRENTE", "ITEM", "ÍTEM"}), Cant = FnMapColumn(_, Cols, {"CANTIDAD", "CANT."}), VU_Crudo = try Record.FieldValues(_){10} otherwise FnMapColumn(_, Cols, {"VALOR UNITARIO", "VLR UNIT", "UNITARIO"}), IVA_Crudo = try Record.FieldValues(_){11} otherwise FnMapColumn(_, Cols, {"IVA %", "IVA", "% IVA"}), VT = try Record.FieldValues(_){12} otherwise FnMapColumn(_, Cols, {"VALOR TOTAL", "VLR TOTAL", "TOTAL"}), OC = FnMapColumn(_, Cols, {"ORDEN", "PEDIDO", "O.C"}) ]),
        DetallesStd = Table.ExpandRecordColumn(MapStd, "Std", {"Codigo_ins", "Ins", "Act", "Cant", "VT", "VU_Crudo", "IVA_Crudo", "OC"}, {"Codigo ins", "Ins", "Actividad", "Cantidad Comprado", "VT Comprado", "VU_Crudo", "IVA_Crudo", "# OC / Contrato"}),
        DetConKeyOC = Table.AddColumn(DetallesStd, "OC_Key", each Text.Trim(Text.From(if [#"# OC / Contrato"] = null then "" else [#"# OC / Contrato"])), type text),
        DetConCodAct = Table.AddColumn(DetConKeyOC, "Codigo act", each let c = Text.Trim(Text.BeforeDelimiter(Text.Trim(Text.From(if [Actividad] = null then "" else [Actividad])), "-", 0)) in if c = "" then null else c, type text),
        DetConClave = Table.AddColumn(DetConCodAct, "InsClave", each FnClaveLimpia([Ins]), type text),
        MergedOC = Table.NestedJoin(DetConClave, {"OC_Key"}, Ordenes_Agrupadas, {"OC_Key"}, "ORD", JoinKind.LeftOuter),
        ExpandedOC = Table.ExpandTableColumn(MergedOC, "ORD", {"Proveedor_Raw"}, {"Proveedor_Raw"}),
        AddedNombreContratista = Table.AddColumn(ExpandedOC, "Nombre Contratista", each let p = try Text.From([Proveedor_Raw]) otherwise null, t = if p = null then null else let pos = Text.PositionOf(p, "-") in if pos < 0 then Text.Trim(p) else Text.Trim(Text.Range(p, pos + 1)) in t, type text)
    in AddedNombreContratista,

    // ============================================================
    // CONEXIÓN A SHAREPOINT (LECTURA DESDE CONSULTA COMPARTIDA)
    // ============================================================
    ArchivosProyecto = Table.SelectRows(SP_Archivos_Proyecto, each
        Text.Contains([Name], "INFORMEORDEN", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "ESTADO DE ORDENES", Comparer.OrdinalIgnoreCase)
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

    Agrupado = Table.Group(ConCentroCosto, {"Centro de Costos"}, {{"Binarios", each
        let
            binDet = PickLatestBinary(_, "INFORMEORDEN"),
            binOC = PickLatestBinary(_, "ESTADO DE ORDENES")
        in
            if binDet <> null and binOC <> null then [Bin_Det = binDet, Bin_OC = binOC] else null
    }}),
    CentrosCompletos = Table.SelectRows(Agrupado, each [Binarios] <> null),
    TablaConDatos = Table.AddColumn(CentrosCompletos, "Datos", each FxProcesarCompras([Binarios][Bin_Det], [Binarios][Bin_OC])),
    Expandido = Table.ExpandTableColumn(TablaConDatos, "Datos", {"Codigo ins", "Ins", "Actividad", "Codigo act", "InsClave", "# OC / Contrato", "Cantidad Comprado", "VT Comprado", "VU_Crudo", "IVA_Crudo", "Nombre Contratista"}),

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
    ITEMS_Base = Table.SelectColumns(ITEMS_Clean, {"Centro de Costos", "Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo"}),

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
            ca = if e then [Codigo act] else (if [Rs.Codigo act] <> null then [Rs.Codigo act] else [Codigo act]), 
            a0 = Text.Trim(Text.From(if [Actividad] = null then "" else [Actividad])), 
            aOrig = if a0 = "" then null else if ca <> null and not Text.StartsWith(a0, Text.From(ca)) then Text.From(ca) & " - " & a0 else a0,
            
            ActOficial = if [Act_Estricto] <> null then [Act_Estricto] else if [Act_Gen] <> null then [Act_Gen] else aOrig,
            CapFinal = if [Cap_Estricto] <> null then [Cap_Estricto] else [Cap_Gen],
            SubCapFinal = if [Sub_Estricto] <> null then [Sub_Estricto] else [Sub_Gen]
        in [ 
            InsFinal = if e then [Ex.Ins] else (if [Rs.Ins] <> null then [Rs.Ins] else [Ins]), 
            CodActFinal = ca, 
            ActFinal = ActOficial, 
            CapFinal = CapFinal, 
            SubCapFinal = SubCapFinal 
        ]),
    
    ExpandedFinalCols = Table.ExpandRecordColumn(Table.RemoveColumns(AddedCoalesced, {"Ins", "Actividad", "Codigo act", "Ex.Ins", "Rs.Codigo act", "Rs.Actividad", "Rs.Ins", "Act_Estricto", "Cap_Estricto", "Sub_Estricto", "Act_Gen", "Cap_Gen", "Sub_Gen"}), "FinalCols", {"InsFinal", "CodActFinal", "ActFinal", "CapFinal", "SubCapFinal"}, {"Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo"}),

    NumericColumns = Table.TransformColumns(ExpandedFinalCols, {{"Cantidad Comprado", each FxToNumberFlex(_), type number}, {"VT Comprado", each FxToNumberFlex(_), type number}, {"VU_Crudo", each FxToNumberFlex(_), type number}, {"IVA_Crudo", each FxToNumberFlex(_), type number}}),
    Added_VU = Table.AddColumn(NumericColumns, "V/U Comprado", each let vb = [VU_Crudo], iva = [IVA_Crudo], p = if iva = null then 0 else if iva >= 1 then iva / 100 else iva, vc = if vb = null then null else vb * (1 + p) in if vc = null then null else Number.Round(vc, 0), type number),
    FilteredZeros = Table.SelectRows(Added_VU, each try [VT Comprado] <> null and [VT Comprado] <> 0 otherwise false),

    SelectedFinal = Table.SelectColumns(Table.AddColumn(Table.AddColumn(FilteredZeros, "Tipo", each "COMPRAS", type text), "Descripcion contrato", each "pedido obra", type text), {"Centro de Costos", "Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo", "# OC / Contrato", "Nombre Contratista", "Descripcion contrato", "Cantidad Comprado", "VT Comprado", "V/U Comprado", "Tipo"}),
    TypedFinal = Table.TransformColumnTypes(SelectedFinal, {{"Centro de Costos", type text}, {"Codigo ins", Int64.Type}, {"Cantidad Comprado", type number}, {"VT Comprado", type number}, {"V/U Comprado", type number}}),
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
                     otherwise Html.Table(Text.FromBinary(BinarioCortes, 65001), Columnas_HTML, [RowSelector="tr"]),
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
    // FUNCIONES AUXILIARES GLOBALES
    // ============================================================
    FxToNumberFlex = F_Globales[FxToNumberFlex],
    FnCleanText = F_Globales[FnCleanText],
    FnEncode = F_Globales[FnEncode],

    // ============================================================
    // CONEXIÓN A SHAREPOINT: Archivo "Descarga ppto"
    // Ruta: /Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. Descargas pptos - Control costos interno/
    // ============================================================
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    FilePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. Descargas pptos - Control costos interno/Descarga ppto.xlsx",

    // Descargar el archivo Excel desde SharePoint
    // Web.Contents con RelativePath: DataSourcePath estable (SiteUrl), evita problemas de cache
    BinarioArchivo = Binary.Buffer(Web.Contents(SiteUrl, [
        RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(FilePath) & "')/$value"
    ])),

    // Abrir el libro y buscar la tabla DESCARGAS
    Libro = Excel.Workbook(BinarioArchivo, null, true),
    TablaDescargas = Libro{[Item="DESCARGA", Kind="Table"]}[Data],

    // ============================================================
    // FILTRAR POR PROYECTO ACTUAL
    // ============================================================
    ParamProyecto = Text.Trim(ProyectoActual),
    FiltradoPorProyecto = Table.SelectRows(TablaDescargas, each 
        Text.Upper(Text.Trim(Text.From(if [Proyecto] = null then "" else [Proyecto]))) = Text.Upper(ParamProyecto)
    ),

    // ============================================================
    // LIMPIEZA Y TIPOS DE DATOS
    // ============================================================
    TextosLimpios = Table.TransformColumns(FiltradoPorProyecto, {
        {"Proyecto", each FnCleanText(_), type text},
        {"Centro de Costos", each FnCleanText(_), type text},
        {"Subcapitulo", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Capitulo", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Actividad", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Ins", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"# CC - Comparativo", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"# CC", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Comparativo", each if _ = null then null else Text.Trim(Text.From(_)), type text},

        {"Cantidad", each FxToNumberFlex(_), type number},
        {"V/U ppto (CC)", each FxToNumberFlex(_), type number},
        {"Valor Total ppto (CC)", each FxToNumberFlex(_), type number}
    }, null, MissingField.Ignore),

    TiposFinales = try Table.TransformColumnTypes(TextosLimpios, {{"Codigo ins", Int64.Type}}) otherwise TextosLimpios,

    // ============================================================
    // SELECCIÓN Y ORDEN FINAL DE COLUMNAS
    // ============================================================
    TablaFinal = Table.SelectColumns(TiposFinales, 
        {"Proyecto", "Centro de Costos", "Subcapitulo", "Capitulo", "Actividad", "Codigo ins", "Ins", 
         "Cantidad", "V/U ppto (CC)", "Valor Total ppto (CC)", 
         "# CC - Comparativo", "# CC", "Comparativo"}, MissingField.Ignore),

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
                       otherwise Html.Table(Text.FromBinary(Binario, 65001), Columnas_HTML, [RowSelector="tr"]),
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
    // Te muestra cuantas filas trae cada query del modelo y si tiene errores.
    // Util para aislar el query roto sin tener que abrir cada uno por separado.

    Medir = (nombre as text, fn as function) =>
        let
            t0 = DateTime.LocalNow(),
            res = try fn() otherwise null,
            t1 = DateTime.LocalNow(),
            segundos = Duration.TotalSeconds(t1 - t0),
            filas = if res = null then -1 else try Table.RowCount(res) otherwise -1,
            errores = if res = null then -1 else try
                let conErr = Table.SelectRowsWithErrors(res) in Table.RowCount(conErr)
                otherwise -1
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
        Medir("BD",                     () => BD),
        Medir("SINCO",                  () => SINCO)
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
    PPTO_Grouped_Buffer = Table.Buffer(Table.Group(PPTO_WithStdIns, {"Centro de Costos", "Codigo act", "Capitulo", "Actividad", "Subcapitulo", "InsNorm"}, {{"Ins_Oficial", each List.First(List.RemoveNulls([Ins])), type text}, {"ValorTotal_PPTO_Bloque", each List.Sum([VT Presupuesto]), type number}, {"Unitario_PPTO_Bloque", each List.First(List.RemoveNulls([#"V/U Presupuesto"])), type number}})),

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
        {"# CC - Comparativo", each FnCleanText(_), type text}, 
        {"Valor Total ppto (CC)", each try Number.From(_) otherwise null, type number}, 
        {"V/U ppto (CC)", each try Number.From(_) otherwise null, type number}
    }, null, MissingField.Ignore),
    
    DetCC_WithStdIns = Table.AddColumn(DetCC_Typed, "InsNorm", each FnRemoveAccentsSymbols([Ins]), type text),
    DetCC_Valid = Table.SelectRows(DetCC_WithStdIns, each [#"# CC - Comparativo"] <> null),

    // ============================================================
    // 5. CRUCE 1: Alinear la base Adjudicada contra la estructura oficial
    // ============================================================
    DetCC_JoinPPTOBlock = Table.NestedJoin(DetCC_Valid, {"Centro de Costos", "Capitulo", "Actividad", "Subcapitulo", "InsNorm"}, PPTO_Grouped_Buffer, {"Centro de Costos", "Capitulo", "Actividad", "Subcapitulo", "InsNorm"}, "PPTOBlock", JoinKind.LeftOuter),
    DetCC_Expanded = Table.ExpandTableColumn(DetCC_JoinPPTOBlock, "PPTOBlock", {"Codigo act", "Ins_Oficial", "ValorTotal_PPTO_Bloque", "Unitario_PPTO_Bloque"}, {"Codigo act", "Ins_Oficial", "ValorTotal_PPTO_Bloque", "Unitario_PPTO_Bloque"}),
    DetCC_WithFinalIns = Table.AddColumn(DetCC_Expanded, "Ins_Final", each if [Ins_Oficial] <> null then [Ins_Oficial] else [Ins], type text),

    DetCC_WithCantidad = Table.AddColumn(DetCC_WithFinalIns, "Cantidad_Calc", each let total = [#"Valor Total ppto (CC)"], unit = [#"V/U ppto (CC)"] in if unit <> null and unit <> 0 then total / unit else null, type number),
    DetCC_ReportShape = Table.SelectColumns(DetCC_WithCantidad, {"Centro de Costos", "Codigo act", "Capitulo", "Actividad", "Subcapitulo", "Ins_Final", "# CC - Comparativo", "Valor Total ppto (CC)", "V/U ppto (CC)", "Cantidad_Calc"}, MissingField.Ignore),
    DetCC_FinalAdjudicados_Renamed = Table.RenameColumns(DetCC_ReportShape, {{"Ins_Final", "Ins"}}),
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
    
    Final_Ordered = Table.ReorderColumns(UnionFiltered, {"Centro de Costos", "Codigo act", "Capitulo", "Actividad", "Subcapitulo", "Ins", "# CC - Comparativo", "Tipo", "Cantidad_Calc", "V/U ppto (CC)", "Valor Total ppto (CC)"}, MissingField.Ignore),
    
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
                txtRaw  = if raw = null then null else Text.Trim(Text.From(raw)),
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
                    Timeout = #duration(0, 0, 10, 0),
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
                                otherwise Html.Table(Text.FromBinary(BinarioSeguimiento, 65001), Columnas_HTML, [RowSelector="tr"]),
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
                        codText   = if codValue  = null then "" else Text.Trim(Text.From(codValue)),
                        descText  = if descValue = null then "" else Text.Trim(Text.From(descValue)),
                        tipoText  = if tipoValue = null then "" else Text.Trim(Text.From(tipoValue)),
                        umText    = if umValue   = null then "" else Text.Trim(Text.From(umValue)),
                        codUpper  = Text.Upper(codText),
                        descUpper = Text.Upper(descText),
                        tryNum    = try Number.FromText(codText),
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
                    in if tipo = "Capitulo" then FnRemoveAccentsSymbols(capTxt) else null, type text),

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
                                         in FnRemoveAccentsSymbols(Text.Trim(baseTxt))
                    in subTxt, type text),

                ItemsSubcapituloFillDown  = Table.FillDown(ItemsWithSubcapitulo, {"Subcapitulo"}),
                ItemsWithCodActRaw        = Table.AddColumn(ItemsSubcapituloFillDown, "CodigoActRaw", (r as record) =>
                    let tipo = Record.Field(r, "TipoFila") in if tipo = "Actividad" then Text.From(Record.Field(r, ItemsCodColName)) else null, type text),
                ItemsCodActRawFillDown    = Table.FillDown(ItemsWithCodActRaw, {"CodigoActRaw"}),
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
                    in FnRemoveAccentsSymbols(baseTxt), type text),

                OrigenAPU_Raw = try Excel.Workbook(BinarioPresupuesto, null, true){0}[Data]
                                otherwise Html.Table(Text.FromBinary(BinarioPresupuesto, 65001), Columnas_APU, [RowSelector="tr"]),
                OrigenAPU_Cols = Table.SelectColumns(OrigenAPU_Raw, List.FirstN(Table.ColumnNames(OrigenAPU_Raw), 3)),
                OrigenAPU = Table.RenameColumns(OrigenAPU_Cols, List.Zip({Table.ColumnNames(OrigenAPU_Cols), {"Columna 1", "Columna 2", "Columna 3"}})),

                APU_Paso1 = Table.AddColumn(OrigenAPU, "Cod_Temp", each
                    let
                        c1Value = if [#"Columna 1"] = null then "" else [#"Columna 1"],
                        c1      = Text.Trim(Text.From(c1Value)),
                        hasDash = Text.Contains(c1, "-"),
                        preDash = if hasDash then Text.Trim(Text.BeforeDelimiter(c1, "-")) else "",
                        esNum   = try Number.FromText(preDash) otherwise null
                    in if hasDash and esNum <> null then FnFormatCodigoAct(preDash) else null),

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
                DiccionarioAPU_Unico = Table.Buffer(Table.Distinct(APU_DiccionarioRenombrado, {"CodigoActAPU"})),

                ItemsJoinAPU     = Table.NestedJoin(ItemsWithIns, {"Codigo act"}, DiccionarioAPU_Unico, {"CodigoActAPU"}, "APU", JoinKind.LeftOuter),
                ItemsExpandedAPU = Table.ExpandTableColumn(ItemsJoinAPU, "APU", {"NombreActAPU", "UM_Actividad"}, {"NombreActAPU", "UM_Actividad"}),

                ItemsWithActividad = Table.AddColumn(ItemsExpandedAPU, "Actividad", each
                    let
                        codTxt        = if [Codigo act]  = null then "" else [Codigo act],
                        nombreExtraido= Text.Trim(Text.From(if [NombreActAPU] = null then "" else [NombreActAPU])),
                        nombreReal    = if nombreExtraido = "" then "Actividad " & codTxt else nombreExtraido,
                        subcapTxt     = Text.Trim(Text.From(if [Subcapitulo] = null then "" else [Subcapitulo])),
                        nombreSinSub  = if subcapTxt <> "" then Text.Replace(nombreReal, subcapTxt, "") else nombreReal,
                        umTxt         = Text.Trim(Text.From(if [UM_Actividad] = null then "" else [UM_Actividad])),
                        nombreLimpio  = Text.Combine(List.Select(Text.Split(nombreSinSub, " "), each _ <> ""), " "),
                        actTxt        = if umTxt = "" then codTxt & "-" & nombreLimpio
                                        else codTxt & "-" & nombreLimpio & " (" & umTxt & ")"
                    in FnRemoveAccentsSymbols(actTxt), type text),

                NumsTyped = Table.TransformColumns(ItemsWithActividad, {
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
            in Final
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
        Text.Upper(ToTextClean(Record.FieldOrDefault(_, "Tipo", ""))) <> "PPTO" and
        ToNumber0(Record.FieldOrDefault(_, "VT Asegurada", 0)) <> 0
    ),

    FiltradoPorOC = Table.SelectRows(BaseConValor, each
        let ocText = ToTextClean(Record.FieldOrDefault(_, "# OC / Contrato", ""))
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

    // Indice liviano: solo lista carpetas y metadatos. Los binarios se descargan
    // despues de filtrar en cada consulta consumidora.
    FolderResponse = try Json.Document(Web.Contents(SiteUrl, [
        RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(BasePath) & "')/Folders",
        Query = [#"$select" = "Name"],
        Headers = Headers,
        Timeout = #duration(0, 0, 5, 0)
    ])) otherwise null,

    CCFolders =
        if FolderResponse = null or not Record.HasFields(FolderResponse, "value")
        then #table({"Name"}, {})
        else Table.FromRecords(FolderResponse[value]),

    WithFiles = Table.AddColumn(CCFolders, "Archivos", each
        let
            ccActualPath = BasePath & "/" & [Name] & "/Actual",
            result = try Json.Document(Web.Contents(SiteUrl, [
                RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(ccActualPath) & "')/Files",
                Query = [#"$select" = "Name,ServerRelativeUrl,TimeLastModified,Length"],
                Headers = Headers,
                Timeout = #duration(0, 0, 5, 0)
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

