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

    // Recorte TEMPRANO (solo La Arboleda): estas 14 columnas no se usan en NINGUN calculo
    // intermedio de todo el pipeline (verificado: 0 referencias [Campo] en BD.m), asi que
    // se quitan aqui mismo en vez de esperar al final. Esto reduce el ancho de tabla que
    // procesan TODOS los pasos siguientes (TransformColumns, ReplaceErrorValues, Table.Group,
    // NestedJoin), no solo la escritura final a la hoja.
    ColumnasReordenadas_Trim = Table.RemoveColumns(ColumnasReordenadas, {
        "Codigo ins", "V/U Comprado", "V/U Contratado", "V/U Presupuesto",
        "Cantidad Cortes", "Cantidad Cons Cols", "VT Cons Cols",
        "Estado", "Fecha_de_pago", "Clasificador_Actividad",
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
    // Primero se filtran las filas CON clasificador y despues se deja una por llave:
    // al reves, Table.Distinct se quedaba con la primera fila de cada llave (casi
    // siempre ITEMS/PPTO, sin clasificador) y el de Det_CC se perdia.
    // CLASIFICADOR (2026-10-05, regla del usuario: "siempre manda aprobaciones").
    // Cada fila toma el primero que exista:
    //   1. APROBACIONES (filas CC CONSOLIDADO) por su comparativo (nombre completo
    //      normalizado: el numero solo se repite, ej. 002 topografia / 002 ensayos), o por el
    //      comparativo de su OC/contrato segun el puente comparativo<->OC de Det_CC.
    //   2. PROVISIONES: el de la propia fila de provision o el de las provisiones de su OC.
    //   3. Det_CC (respaldo; la columna Clasificador de Det_CC se retira de los libros).
    //   4. Propagacion por CC|codigo act|insumo desde las filas ya clasificadas en 1-3
    //      (ITEMS, PPTO, POR ADJUDICAR, compras de insumos ya clasificados).
    //   5. En blanco.
    FnBuscar = (rec as record, key as nullable text) as any =>
        if key = null or key = "" then null else Record.FieldOrDefault(rec, key, null),
    FnTieneTexto = (v as any) as logical =>
        let t = try Text.Trim(Text.From(v)) otherwise null in t <> null and t <> "",
    FnOCKeyValida = (v as any) as nullable text =>
        let k = CleanOCKey(v)
        in if k = null or Text.Select(k, {"0".."9"}) = "" or Text.Remove(k, {"0"}) = "" then null else k,
    // Sin insumo no hay llave (provisiones y compras sueltas): evita que "CC||" clasifique filas ajenas
    FnClaveIns = (r as record) as nullable text =>
        if r[Ins] = "" then null else r[Centro de Costos] & "|" & r[Codigo act] & "|" & r[Ins],
    FnClaveComp = (r as record) as text => FnTextoODef(r[#"# CC - Comparativo"]),
    FiltroConClasif = if Table.HasColumns(FiltroTipoValido, "Clasificador") then FiltroTipoValido else Table.AddColumn(FiltroTipoValido, "Clasificador", each null),
    // Mapa clave -> valor solo para las claves con UN unico valor distinto: una OC puede
    // incluir varios comparativos (o clasificadores) y no se debe repartir uno de ellos a todo.
    FnMapaUnico = (tbl as table, claveFn as function, valorCol as text) as record =>
        let
            conClave = Table.SelectRows(Table.AddColumn(tbl, "__K", each claveFn(_), type text), each [__K] <> null and [__K] <> ""),
            grupos = Table.Group(conClave, {"__K"}, {{"__V", each List.Distinct(List.RemoveNulls(Table.Column(_, valorCol))), type list}}),
            unicos = Table.SelectRows(grupos, each List.Count([__V]) = 1)
        in Record.FromList(List.Transform(unicos[__V], each _{0}), unicos[__K]),
    FnFilasClasif = (tipo as text) as table =>
        Table.TransformColumns(
            Table.SelectRows(FiltroConClasif, each [Tipo] = tipo and FnTieneTexto([Clasificador])),
            {{"Clasificador", each Text.Trim(Text.From(_)), type text}}),
    FnClaveOC = (r as record) as nullable text => FnOCKeyValida(r[#"# OC / Contrato"]),
    FnClaveOCIns = (r as record) as nullable text =>
        let k = FnOCKeyValida(r[#"# OC / Contrato"]) in if k = null or r[Ins] = "" then null else k & "|" & r[Ins],
    ClasifAprobComp = FnMapaModa(FnFilasClasif("CC CONSOLIDADO"), FnClaveComp, "Clasificador"),
    ClasifProvOC    = FnMapaUnico(FnFilasClasif("PROVISIONES"), FnClaveOC, "Clasificador"),
    ClasifDetComp   = FnMapaModa(FnFilasClasif("CC"), FnClaveComp, "Clasificador"),
    ClasifDetOCIns  = FnMapaModa(FnFilasClasif("CC"), FnClaveOCIns, "Clasificador"),
    ClasifDetOC     = FnMapaUnico(FnFilasClasif("CC"), FnClaveOC, "Clasificador"),
    // Puente Det_CC: OC + insumo -> comparativo (exacto); OC sola solo si tiene un unico comparativo
    FilasPuente = Table.SelectRows(FiltroTipoValido, each [Tipo] = "CC" and FnTextoODef([#"# CC - Comparativo"]) <> ""),
    CompPorOCIns = FnMapaModa(FilasPuente, FnClaveOCIns, "# CC - Comparativo"),
    CompPorOC    = FnMapaUnico(FilasPuente, FnClaveOC, "# CC - Comparativo"),
    Etapa1 = Table.AddColumn(FiltroConClasif, "__ClasifFinal", each
        let
            comp = FnClaveComp(_),
            ocK = FnOCKeyValida([#"# OC / Contrato"]),
            ocIns = FnClaveOCIns(_),
            compOCIns = FnBuscar(CompPorOCIns, ocIns),
            compOC = FnTextoODef(if compOCIns <> null then compOCIns else FnBuscar(CompPorOC, ocK)),
            propio = if FnTieneTexto([Clasificador]) then Text.Trim(Text.From([Clasificador])) else null,
            // Las filas de aprobaciones conservan el de su linea (17 de 1.070 comparativos
            // tienen lineas con clasificadores distintos); las demas toman el mayoritario
            a0 = if [Tipo] = "CC CONSOLIDADO" then propio else null,
            a1 = if a0 <> null then a0 else FnBuscar(ClasifAprobComp, comp),
            a2 = if a1 <> null then a1 else FnBuscar(ClasifAprobComp, compOC),
            p1 = if a2 <> null then a2 else if [Tipo] = "PROVISIONES" then propio else null,
            p2 = if p1 <> null then p1 else FnBuscar(ClasifProvOC, ocK),
            d1 = if p2 <> null then p2 else FnBuscar(ClasifDetComp, comp),
            d2 = if d1 <> null then d1 else FnBuscar(ClasifDetOCIns, ocIns),
            d3 = if d2 <> null then d2 else FnBuscar(ClasifDetOC, ocK)
        in d3, type text),
    ClasifPorIns = FnMapaModa(Table.SelectRows(Etapa1, each [__ClasifFinal] <> null), FnClaveIns, "__ClasifFinal"),
    BaseClasificada0 = Table.AddColumn(Etapa1, "__Clasif2", each
        if [__ClasifFinal] <> null then [__ClasifFinal] else FnBuscar(ClasifPorIns, FnClaveIns(_)), type text),
    BaseClasificada = Table.RenameColumns(
        Table.RemoveColumns(BaseClasificada0, {"Clasificador", "__ClasifFinal"}), {{"__Clasif2", "Clasificador"}}),

    // CONTRATISTAS (2026-10-05): el nombre maestro es el de SINCO (filas CONTRATO/COMPRAS).
    // Cada fila toma, en cascada: (1) el nombre SINCO de su OC/contrato, (2) el de la OC
    // de su comparativo en Det_CC, (3) el nombre SINCO con la misma clave normalizada
    // (sin tildes, puntuacion, espacios ni SAS/SA/LTDA/SCA/CIA/Y), (4) su propio nombre.
    // Al final se aplica la tabla manual ALIAS_CONTRATISTAS (Nombre fuente -> Nombre
    // unificado) si existe en el libro. Antes se unificaba por insumo y se pisaban
    // proveedores distintos del mismo insumo.
    SufijosSociales = {"SAS", "SA", "SCA", "LTDA", "LIMITADA", "CIA", "Y", "EU", "SENC", "ZOMAC", "BIC", "SUCURSAL"},
    FnClaveNombre = (v as any) as nullable text =>
        let
            t = FnTextoODef(v),
            t1 = Text.Replace(Text.Replace(Text.Upper(FnRemoveAccentMarks(t)), ".", ""), "&", "Y"),
            partes = List.Select(Text.SplitAny(t1, " ,;:-_/\()'""#" & "#(tab)"), each _ <> "" and not List.Contains(SufijosSociales, _)),
            k = Text.Select(Text.Combine(partes, ""), {"A".."Z", "0".."9"})
        in if k = "" then null else k,
    FilasSinco = Table.SelectRows(FiltroTipoValido, each ([Tipo] = "CONTRATO" or [Tipo] = "COMPRAS") and FnTieneTexto([Nombre Contratista])),
    FnMapaModa = (tbl as table, claveFn as function, valorCol as text) as record =>
        let
            conClave = Table.SelectRows(Table.AddColumn(tbl, "__K", each claveFn(_), type text), each [__K] <> null and [__K] <> ""),
            grupos = Table.Group(conClave, {"__K"}, {{"__V", each List.Mode(List.RemoveNulls(Table.Column(_, valorCol))), type text}})
        in Record.FromList(grupos[__V], grupos[__K]),
    NombrePorOC    = FnMapaModa(FilasSinco, each FnOCKeyValida([#"# OC / Contrato"]), "Nombre Contratista"),
    NombrePorClave = FnMapaModa(FilasSinco, each FnClaveNombre([Nombre Contratista]), "Nombre Contratista"),
    OCPorComp = FnMapaModa(
        Table.SelectRows(
            Table.AddColumn(Table.SelectRows(FiltroTipoValido, each [Tipo] = "CC"), "__OC", each FnOCKeyValida([#"# OC / Contrato"]), type text),
            each [__OC] <> null),
        each FnTextoODef([#"# CC - Comparativo"]), "__OC"),
    AliasTabla = try Excel.CurrentWorkbook(){[Name = "ALIAS_CONTRATISTAS"]}[Content] otherwise #table({"Nombre fuente", "Nombre unificado"}, {}),
    AliasFilas = Table.SelectRows(AliasTabla, each try FnTieneTexto([Nombre unificado]) otherwise false),
    AliasMapa = try FnMapaModa(AliasFilas, each try FnClaveNombre([Nombre fuente]) otherwise null, "Nombre unificado") otherwise [],
    BaseClasificadaFinal = Table.AddColumn(BaseClasificada, "__NombreFinal", each
            let
                propio = if FnTieneTexto([Nombre Contratista]) then [Nombre Contratista] else null,
                n1 = FnBuscar(NombrePorOC, FnOCKeyValida([#"# OC / Contrato"])),
                n2 = if n1 <> null then n1 else FnBuscar(NombrePorOC, FnBuscar(OCPorComp, FnTextoODef([#"# CC - Comparativo"]))),
                n3 = if n2 <> null then n2 else FnBuscar(NombrePorClave, FnClaveNombre(propio)),
                cand = if n3 <> null then n3 else propio,
                a1 = FnBuscar(AliasMapa, FnClaveNombre(cand)),
                a2 = if a1 <> null then a1 else FnBuscar(AliasMapa, FnClaveNombre(propio))
            in if a2 <> null then FnCleanContratista(a2) else cand, type text),
    BaseClasificadaFinal2 = Table.RenameColumns(
        Table.RemoveColumns(BaseClasificadaFinal, {"Nombre Contratista"}), {{"__NombreFinal", "Nombre Contratista"}}),

    // El "otherwise 0" del try ya cubre el caso null; no necesitamos doble guarda
    NumerosSeguros = Table.TransformColumns(BaseClasificadaFinal2,List.Transform(NumCols, each {_, ToNumber0, type number}), null, MissingField.Ignore),

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
    // Comparativos YA contratados: su # OC / Contrato (tabla Det_CC, filas Tipo CC)
    // tiene valor asegurado en SINCO (filas CONTRATO/COMPRAS con esa OC).
    FnOCKey = (v as any) as text => let k = CleanOCKey(v) in if k = null then "" else k,
    OCsAseguradas = List.Buffer(List.Distinct(List.Transform(
        Table.SelectRows(AddVUAseg, each ([Tipo] = "CONTRATO" or [Tipo] = "COMPRAS") and ToNumber0([VT Asegurada]) <> 0)[#"# OC / Contrato"],
        FnOCKey))),
    ComparativosContratados = List.Buffer(List.Distinct(List.Transform(
        Table.SelectRows(AddVUAseg, each [Tipo] = "CC" and List.Contains(OCsAseguradas, FnOCKey([#"# OC / Contrato"])) and FnOCKey([#"# OC / Contrato"]) <> "")[#"# CC - Comparativo"],
        FnTextoODef))),
    // PROYECTADO SINCO EN 0 (2026-10-05, regla del usuario): si el insumo tiene ppto > 0
    // pero su proyectado en el seguimiento de SINCO (filas ITEMS, VT Proyectado) es 0,
    // POR ADJUDICAR y ADJUDICADO proyectan 0. Llave CC|codigo act|insumo; si el insumo no
    // esta en el seguimiento (ADJUDICADO con el nombre del archivo de descargas), se mira
    // la actividad completa (CC|codigo act). Sin fila ITEMS no se aplica.
    FnSumaPor = (tipo as text, valorCol as text, claveFn as function) as record =>
        let
            filas = Table.SelectRows(AddVUAseg, each [Tipo] = tipo),
            conClave = Table.SelectRows(Table.AddColumn(filas, "__K", each claveFn(_), type text), each [__K] <> null),
            grupos = Table.Group(conClave, {"__K"}, {{"__S", each List.Sum(List.Transform(Table.Column(_, valorCol), ToNumber0)), type number}})
        in Record.FromList(grupos[__S], grupos[__K]),
    FnClaveAct = (r as record) as nullable text =>
        if r[Codigo act] = "" then null else r[Centro de Costos] & "|" & r[Codigo act],
    ProjSincoIns = FnSumaPor("ITEMS", "VT Proyectado", FnClaveIns),
    ProjSincoAct = FnSumaPor("ITEMS", "VT Proyectado", FnClaveAct),
    PptoIns = FnSumaPor("PPTO", "VT Presupuesto", FnClaveIns),
    PptoAct = FnSumaPor("PPTO", "VT Presupuesto", FnClaveAct),
    FnSincoEnCero = (r as record) as logical =>
        let
            kI = FnClaveIns(r), kA = FnClaveAct(r),
            usarIns = kI <> null and Record.HasFields(ProjSincoIns, kI),
            ppto = if usarIns then FnBuscar(PptoIns, kI) else FnBuscar(PptoAct, kA),
            proj = if usarIns then FnBuscar(ProjSincoIns, kI) else FnBuscar(ProjSincoAct, kA)
        in proj <> null and ppto <> null and ppto > 0 and Number.Abs(proj) < 1,
    AplicarProyeccion = Table.AddColumn(ConAsegActividad, "VT Proyectado Colsubsidio", each
        if [Tipo] = "POR ADJUDICAR" then (if ToNumber0([vtAsegAct]) > 0 or FnSincoEnCero(_) then 0 else [#"Valor Total ppto (CC)"])
        // ADJUDICADO = descargado. Proyecta su valor SOLO mientras el comparativo no
        // tenga contrato/OC en SINCO (ej. pintura: comparativo aprobado, contrato sin
        // montar). Si el comparativo ya esta contratado, lo que no quedo en el contrato
        // no se va a gastar: proyecta 0 (regla del usuario: asegurado + por adjudicar).
        // Tambien 0 si la actividad ya tiene asegurado por otro contrato.
        else if [Tipo] = "ADJUDICADO" then
            (if List.Contains(ComparativosContratados, FnTextoODef([#"# CC - Comparativo"])) then 0
             else if ToNumber0([vtAsegAct]) > 0 then 0
             else if FnSincoEnCero(_) then 0
             else [#"Valor Total ppto (CC)"])
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
        "Cantidad asegurada", "V/U asegurada",
        // Agregadas 2026-09-21: presupuesto descargado por comparativo (filas
        // ADJUDICADO/POR ADJUDICAR de DISPONIBLE), para comparar ppto vs aprobado
        // en la dinamica de comparativos.
        "Cantidad ppto (CC)", "V/U ppto (CC)",
        // Vuelve 2026-09-21 a pedido del usuario (se calcula en BaseClasificada desde Det_CC).
        "Clasificador"
    },
    // ============================================================
    // RELLENO DE SUBCAPITULO (2026-09-04): muchas filas (COMPRAS, CONTRATO, CC,
    // CC CONSOLIDADO) llegan sin Subcapitulo aunque su actividad SI lo tiene en
    // el ppto/seguimiento. Reglas, en orden:
    //   1. Si ya tiene Subcapitulo, se respeta.
    //   2. Capitulo o Actividad con "SAN AGUSTIN" -> URBANISMO SAN AGUSTIN
    //      (regla del usuario: el capitulo 31 es urbanismo san agustin).
    //   3. Propagacion por (Centro de Costos + codigo de actividad): el codigo
    //      sale de la columna "Codigo act" o, si viene vacia (filas CC de
    //      aprobaciones), se parsea del prefijo numerico del nombre ("1.01-...").
    //   4. El nombre de la actividad termina con un subcapitulo ya conocido del
    //      proyecto ("2.21-BASE GRANULAR URBANISMO INTERIOR (M3)").
    // Filas sin actividad (COMPRAS sueltas, CC CONSOLIDADO, PROVISIONES) no se
    // pueden clasificar y quedan en blanco.
    // ============================================================
    // OJO: Text.From(null) devuelve null (NO error), asi que el try no lo captura
    // y null = "" es false en M — hay que chequear null explicitamente.
    EsSubcapVacio = (v as any) as logical =>
        let t = try Text.Trim(Text.From(v)) otherwise "" in t = null or t = "",
    FnTextoODef = (v as any) as text =>
        let t = try Text.Trim(Text.From(v)) otherwise "" in if t = null then "" else t,
    FnParseCodAct = (cod as any, act as any) as text =>
        let c = FnTextoODef(cod)
        in if c <> "" then c
           else
             let a = FnTextoODef(act),
                 pre = if a = "" then "" else FnTextoODef(Text.BeforeDelimiter(a, "-")),
                 esCodigo = pre <> "" and Text.Remove(pre, {"0".."9", "."}) = ""
             in if esCodigo then pre else "",
    FnNormSubcap = (v as any) as text =>
        let t = try Text.Upper(FnRemoveAccentsSymbols(Text.From(v))) otherwise "" in if t = null then "" else t,

    BaseRelleno = Table.Buffer(Table.AddColumn(FinalOCLimpia, "__CodKey",
        each FnParseCodAct([Codigo act], [Actividad]), type text)),

    MapaSubcapAct = Table.Buffer(Table.Group(
        Table.SelectRows(BaseRelleno, each (not EsSubcapVacio([Subcapitulo])) and [__CodKey] <> ""),
        {"Centro de Costos", "__CodKey"},
        {{"subcapAct", each List.First(List.RemoveNulls([Subcapitulo])), type text}})),

    // Subcapitulos conocidos del proyecto (por CC), con su forma normalizada,
    // ordenados del mas largo al mas corto para que gane la coincidencia mayor.
    SubcapConocidos = Table.Buffer(Table.Group(
        Table.SelectRows(BaseRelleno, each not EsSubcapVacio([Subcapitulo])),
        {"Centro de Costos"},
        {{"__ListaSubcaps", each
            let vs = List.Distinct(List.Transform(List.RemoveNulls([Subcapitulo]), each Text.Trim(Text.From(_))))
            in List.Sort(
                List.Transform(vs, (v) => [Orig = v, Norm = FnNormSubcap(v)]),
                each - Text.Length([Norm])), type list}})),

    RellenoJoin1 = Table.ExpandTableColumn(
        Table.NestedJoin(BaseRelleno, {"Centro de Costos", "__CodKey"}, MapaSubcapAct, {"Centro de Costos", "__CodKey"}, "__SP", JoinKind.LeftOuter),
        "__SP", {"subcapAct"}),
    RellenoJoin2 = Table.ExpandTableColumn(
        Table.NestedJoin(RellenoJoin1, {"Centro de Costos"}, SubcapConocidos, {"Centro de Costos"}, "__SC", JoinKind.LeftOuter),
        "__SC", {"__ListaSubcaps"}),

    RellenoCalc = Table.AddColumn(RellenoJoin2, "__SubcapFill", each
        if not EsSubcapVacio([Subcapitulo]) then [Subcapitulo]
        else
          let
            capN = FnNormSubcap([Capitulo]),
            actTxt = FnTextoODef([Actividad]),
            actN = FnNormSubcap(actTxt)
          in
            if Text.Contains(capN, "SAN AGUSTIN") or Text.Contains(actN, "SAN AGUSTIN") then "URBANISMO SAN AGUSTIN"
            else if not EsSubcapVacio([subcapAct]) then [subcapAct]
            else
              let
                sinUM = if Text.EndsWith(actTxt, ")") and Text.Contains(actTxt, "(")
                        then Text.Trim(Text.BeforeDelimiter(actTxt, "(", {0, RelativePosition.FromEnd}))
                        else actTxt,
                sinUMN = FnNormSubcap(sinUM),
                lista = if [__ListaSubcaps] = null then {} else [__ListaSubcaps],
                cand = List.First(List.Select(lista, each [Norm] <> "" and Text.EndsWith(sinUMN, [Norm])), null)
              in if cand <> null then cand[Orig] else null,
        type text),

    RellenoFinal = Table.RenameColumns(
        Table.RemoveColumns(RellenoCalc, {"Subcapitulo", "subcapAct", "__CodKey", "__ListaSubcaps"}, MissingField.Ignore),
        {{"__SubcapFill", "Subcapitulo"}}),

    // ============================================================
    // NOMBRE UNICO POR ACTIVIDAD (2026-09-21): la misma actividad llegaba con
    // textos distintos segun la fuente (Det_CC con " - TORRES", ppto sin el, etc.)
    // y salia duplicada en las dinamicas. Cada fila toma el nombre OFICIAL del
    // presupuesto (filas PPTO/ITEMS) de su Centro de Costos + codigo de actividad.
    // Solo se aplica cuando ese codigo tiene UN unico nombre en el ppto; si tiene
    // varios (el mismo codigo en distintos subcapitulos), se deja cada fila como viene.
    // ============================================================
    FnCodDeNombre = (act as any, cod as any) as text =>
        let
            a = FnTextoODef(act),
            pre = if a = "" then "" else FnTextoODef(Text.BeforeDelimiter(a, "-")),
            esCodigo = pre <> "" and Text.Remove(pre, {"0".."9", "."}) = ""
        in if esCodigo then pre else FnTextoODef(cod),
    BaseNombres = Table.Buffer(Table.AddColumn(RellenoFinal, "__CodNom",
        each FnCodDeNombre([Actividad], [Codigo act]), type text)),
    NombresOficiales = Table.Buffer(Table.SelectRows(Table.Group(
        Table.SelectRows(BaseNombres, each ([Tipo] = "PPTO" or [Tipo] = "ITEMS")
            and [__CodNom] <> "" and FnTextoODef([Actividad]) <> ""),
        {"Centro de Costos", "__CodNom"},
        {{"__Nombres", each List.Distinct(List.RemoveNulls([Actividad])), type list}}),
        each List.Count([__Nombres]) = 1)),
    UnirNombres = Table.ExpandTableColumn(
        Table.NestedJoin(BaseNombres, {"Centro de Costos", "__CodNom"}, NombresOficiales, {"Centro de Costos", "__CodNom"}, "__NO", JoinKind.LeftOuter),
        "__NO", {"__Nombres"}),
    ActUnificada0 = Table.AddColumn(UnirNombres, "__ActFinal",
        each if [__Nombres] <> null then [__Nombres]{0} else [Actividad], type text),
    ActUnificada = Table.RenameColumns(
        Table.RemoveColumns(ActUnificada0, {"Actividad", "__CodNom", "__Nombres"}),
        {{"__ActFinal", "Actividad"}}),

    RellenoConPpto = Table.RenameColumns(ActUnificada, {{"Cantidad_Calc", "Cantidad ppto (CC)"}}, MissingField.Ignore),
    FinalRecortada = Table.SelectColumns(RellenoConPpto, ColumnasFinalesArboleda, MissingField.UseNull),

    TablaMaestraFinal = Table.Buffer(FinalRecortada)
in
    TablaMaestraFinal
