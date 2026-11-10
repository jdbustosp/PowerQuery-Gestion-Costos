let
    Tol = 0.01,
    FnRemoveAccentsSymbols = F_Globales[FnRemoveAccentsSymbols],
    FnCleanContratista = F_Globales[FnCleanContratista],
    FnRemoveAccentMarks = F_Globales[FnRemoveAccentMarks],
    ToNumber0 = (v as any) as number =>
        let n = try Number.From(v) otherwise null
        in if n = null then 0 else n,

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

    OrigenLimpio = Table.ReplaceErrorValues(Origen,
        List.Transform(Table.ColumnNames(Origen), each {_, null})),

    ColumnasReordenadas = Table.SelectColumns(OrigenLimpio, ColumnasOrden, MissingField.Ignore),

    ColumnasReordenadas_Trim = Table.RemoveColumns(ColumnasReordenadas, {
        "V/U Comprado", "V/U Contratado", "V/U Presupuesto",
        "VT Cons Cols",
        "Estado", "Fecha_de_pago", "Clasificador_Actividad",
        "Capitulo_Costo directo", "Fecha_Factura"
    }, MissingField.Ignore),

    LlavesLimpias = Table.TransformColumns(ColumnasReordenadas_Trim, {
        {"Centro de Costos",     each try (if _ = null then "" else Text.Upper(Text.Trim(Text.From(_)))) otherwise "", type text},
        {"Codigo act",           each try (if _ = null then "" else Text.Upper(Text.Trim(Text.From(_)))) otherwise "", type text},
        {"Ins",                  each try (if _ = null then "" else Text.Upper(Text.Trim(FnRemoveAccentMarks(_)))) otherwise "", type text},
        {"Actividad",            each try (if _ = null then null else Text.Upper(Text.Trim(FnRemoveAccentMarks(_)))) otherwise null, type text},
        {"Tipo",                 each try (if _ = null then "" else Text.Upper(Text.Trim(Text.From(_)))) otherwise "", type text},
        {"# OC / Contrato",      each try (if _ = null then null else Text.Trim(Text.From(_))) otherwise null, type text},
        {"Nombre Contratista",   each try FnCleanContratista(_) otherwise null, type text},
        {"Descripcion contrato", each try FnRemoveAccentsSymbols(if _ = null then null else Text.Trim(Text.From(_))) otherwise null, type text},

        {"# CC - Comparativo",   each try (if _ = null then null else Text.Upper(Text.Trim(Text.From(_)))) otherwise null, type text}
    }, null, MissingField.Ignore),

    FiltroTipoValido = Table.Buffer(Table.SelectRows(LlavesLimpias, each [Tipo] <> null and [Tipo] <> "")),

    FnClaveMaestro = (v as any) as text =>
        let t = try Text.Upper(Text.Trim(FnRemoveAccentMarks(Text.From(v)))) otherwise null
        in if t = null then "" else Text.Combine(List.Select(Text.Split(t, " "), each _ <> ""), " "),
    MaestroFilas = try Table.SelectRows(
            Table.SelectColumns(Excel.CurrentWorkbook(){[Name = "MaestroInsumos"]}[Content], {"Ins", "Clasificador"}),
            each FnTieneTexto([Clasificador]) and FnClaveMaestro([Ins]) <> "")
        otherwise #table({"Ins", "Clasificador"}, {}),
    ClasifMaestro = FnMapaModa(
        Table.TransformColumns(MaestroFilas, {{"Clasificador", each Text.Trim(Text.From(_)), type text}}),
        each FnClaveMaestro([Ins]), "Clasificador"),
    FnBuscar = (rec as record, key as nullable text) as any =>
        if key = null or key = "" then null else Record.FieldOrDefault(rec, key, null),
    FnTieneTexto = (v as any) as logical =>
        let t = try Text.Trim(Text.From(v)) otherwise null in t <> null and t <> "",
    FnOCKeyValida = (v as any) as nullable text =>
        let k = CleanOCKey(v)
        in if k = null or Text.Select(k, {"0".."9"}) = "" or Text.Remove(k, {"0"}) = "" then null else k,

    FnClaveIns = (r as record) as nullable text =>
        if r[Ins] = "" then null else r[Centro de Costos] & "|" & r[Codigo act] & "|" & r[Ins],
    FnClaveComp = (r as record) as text => FnTextoODef(r[#"# CC - Comparativo"]),
    FiltroConClasif = if Table.HasColumns(FiltroTipoValido, "Clasificador") then FiltroTipoValido else Table.AddColumn(FiltroTipoValido, "Clasificador", each null),

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
    ClasifDetComp   = FnMapaModa(FnFilasClasif("CC"), FnClaveComp, "Clasificador"),
    ClasifDetOCIns  = FnMapaModa(FnFilasClasif("CC"), FnClaveOCIns, "Clasificador"),
    ClasifDetOC     = FnMapaUnico(FnFilasClasif("CC"), FnClaveOC, "Clasificador"),

    FilasPuente = Table.SelectRows(FiltroTipoValido, each [Tipo] = "CC" and FnTextoODef([#"# CC - Comparativo"]) <> ""),
    CompPorOCIns = FnMapaModa(FilasPuente, FnClaveOCIns, "# CC - Comparativo"),
    CompPorOC    = FnMapaUnico(FilasPuente, FnClaveOC, "# CC - Comparativo"),
    Etapa0 = Table.AddColumn(FiltroConClasif, "__ClasifComp", each
        let
            comp = FnClaveComp(_),
            ocK = FnOCKeyValida([#"# OC / Contrato"]),
            ocIns = FnClaveOCIns(_),
            compOCIns = FnBuscar(CompPorOCIns, ocIns),
            compOC = FnTextoODef(if compOCIns <> null then compOCIns else FnBuscar(CompPorOC, ocK)),
            propio = if FnTieneTexto([Clasificador]) then Text.Trim(Text.From([Clasificador])) else null,

            a0 = if [Tipo] = "CC CONSOLIDADO" then propio else null,
            a1 = if a0 <> null then a0 else FnBuscar(ClasifAprobComp, comp),
            a2 = if a1 <> null then a1 else FnBuscar(ClasifAprobComp, compOC),
            p1 = if a2 <> null then a2 else if [Tipo] = "PROVISIONES" then propio else null,
            d1 = if p1 <> null then p1 else FnBuscar(ClasifDetComp, comp),
            d2 = if d1 <> null then d1 else FnBuscar(ClasifDetOCIns, ocIns),
            d3 = if d2 <> null then d2 else FnBuscar(ClasifDetOC, ocK)
        in d3, type text),
    Etapa1 = Table.AddColumn(Etapa0, "__ClasifFinal", each
        let m0 = FnBuscar(ClasifMaestro, FnClaveMaestro([Ins]))
        in if [Tipo] = "CC CONSOLIDADO" then (if FnTieneTexto([Clasificador]) then Text.Trim(Text.From([Clasificador])) else "zOtros")
           else if List.Contains({"CC", "ADJUDICADO"}, [Tipo]) then
               (let a = FnBuscar(ClasifAprobComp, FnClaveComp(_)) in if a <> null then a else "zOtros")
           else if m0 <> null then m0 else [__ClasifComp], type text),
    FnEsAsegurado = (r as record) as logical =>
        r[Tipo] = "CONTRATO" or (r[Tipo] = "COMPRAS" and not FnTieneTexto(r[#"#ENTRADA"]) and not FnTieneTexto(r[#"#SALIDA"])),

    ClasifAsegOC = FnMapaModa(Table.SelectRows(Etapa1, each [__ClasifComp] <> null and FnEsAsegurado(_)), FnClaveOC, "__ClasifComp"),
    ClasifPorIns = FnMapaModa(Table.SelectRows(Etapa1, each [__ClasifFinal] <> null and [Tipo] <> "PROVISIONES"), FnClaveIns, "__ClasifFinal"),
    BaseClasificada0 = Table.AddColumn(Etapa1, "__Clasif2", each
        if [__ClasifFinal] <> null then [__ClasifFinal]
        else if FnEsAsegurado(_) then
            (let oc = FnBuscar(ClasifAsegOC, FnClaveOC(_)) in if oc <> null then oc else "zOtros")
        else FnBuscar(ClasifPorIns, FnClaveIns(_)), type text),
    BaseClasificada = Table.RenameColumns(
        Table.RemoveColumns(BaseClasificada0, {"Clasificador", "__ClasifComp", "__ClasifFinal"}), {{"__Clasif2", "Clasificador"}}),

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

    NumerosSeguros = Table.TransformColumns(BaseClasificadaFinal2,List.Transform(NumCols, each {_, ToNumber0, type number}), null, MissingField.Ignore),

    AddCantAseg = Table.AddColumn(NumerosSeguros, "Cantidad asegurada",
        each ToNumber0([Cantidad Contratado]) + ToNumber0([Cantidad Comprado]), type number),
    AddVTAseg = Table.AddColumn(AddCantAseg, "VT Asegurada",
        each ToNumber0([VT Contratado]) + ToNumber0([VT Comprado]), type number),
    AddVUAseg = Table.Buffer(Table.AddColumn(AddVTAseg, "V/U asegurada",
        each let qa = ToNumber0([Cantidad asegurada]), vt = ToNumber0([VT Asegurada])
            in if qa <> 0 then vt / qa else 0, type number)),

    AgrupadoResumen = Table.Group(AddVUAseg, {"Centro de Costos", "Codigo act", "Ins"}, {
        {"vtProj", each List.Sum(List.Transform([VT Proyectado],       each ToNumber0(_))), type number},
        {"vtCons", each List.Sum(List.Transform([VT Consumido],        each ToNumber0(_))), type number},
        {"vtAseg", each List.Sum(List.Transform([VT Asegurada],        each ToNumber0(_))), type number},
        {"vtAprb", each List.Sum(List.Transform([VR total aprobacion], each ToNumber0(_))), type number}
    }),

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

    CruceConBase = Table.NestedJoin(AddVUAseg, {"Centro de Costos", "Codigo act", "Ins"}, BanderasBuffer, {"Centro de Costos", "Codigo act", "Ins"}, "B", JoinKind.LeftOuter),
    BaseConBanderas = Table.ExpandTableColumn(CruceConBase, "B", ColsBanderas),
    BaseConBanderasSafe = Table.ReplaceValue(BaseConBanderas, null, false, Replacer.ReplaceValue, ColsBanderas),

    AsegPorActividad = Table.Buffer(Table.Group(
        Table.SelectRows(AddVUAseg, each [Tipo] = "CONTRATO" or [Tipo] = "COMPRAS"),
        {"Centro de Costos", "Codigo act"},
        {{"vtAsegAct", each List.Sum(List.Transform([VT Asegurada], each ToNumber0(_))), type number}})),
    ConAsegActividad = Table.ExpandTableColumn(
        Table.NestedJoin(AddVUAseg, {"Centro de Costos", "Codigo act"}, AsegPorActividad, {"Centro de Costos", "Codigo act"}, "__AA", JoinKind.LeftOuter),
        "__AA", {"vtAsegAct"}),

    FnOCKey = (v as any) as text => let k = CleanOCKey(v) in if k = null then "" else k,
    OCsAseguradas = List.Buffer(List.Distinct(List.Transform(
        Table.SelectRows(AddVUAseg, each ([Tipo] = "CONTRATO" or [Tipo] = "COMPRAS") and ToNumber0([VT Asegurada]) <> 0)[#"# OC / Contrato"],
        FnOCKey))),
    ComparativosContratados = List.Buffer(List.Distinct(List.Transform(
        Table.SelectRows(AddVUAseg, each [Tipo] = "CC" and List.Contains(OCsAseguradas, FnOCKey([#"# OC / Contrato"])) and FnOCKey([#"# OC / Contrato"]) <> "")[#"# CC - Comparativo"],
        FnTextoODef))),

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

        "Cantidad asegurada", "V/U asegurada",

        "Cantidad ppto (CC)", "V/U ppto (CC)",

        "Clasificador"
    },

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

    FnOCDig = (v as any) as text => Text.Select(FnTextoODef(v), {"0".."9"}),
    BaseAjustes = Table.Buffer(Table.AddColumn(RellenoConPpto, "__OCd", each FnOCDig([#"# OC / Contrato"]), type text)),
    ModaClasifOC = Table.SelectRows(Table.Group(Table.SelectRows(BaseAjustes, each [__OCd] <> "" and [Tipo] <> "PROVISIONES"), {"__OCd"},
        {{"Value", each let x = List.Select(List.Transform([Clasificador], FnTextoODef), each _ <> "")
            in if List.IsEmpty(x) then null else List.Mode(x), type nullable text}}), each [Value] <> null),
    MapaClasifOC = Record.FromTable(Table.RenameColumns(ModaClasifOC, {{"__OCd", "Name"}})),
    ClasifCompleto0 = Table.AddColumn(BaseAjustes, "__Clasif3", each
        if FnTextoODef([Clasificador]) <> "" then [Clasificador]
        else Record.FieldOrDefault(MapaClasifOC, [__OCd], "zOtros"), type text),
    ClasifCompleto = Table.TransformColumns(
        Table.RenameColumns(Table.RemoveColumns(ClasifCompleto0, {"Clasificador", "__OCd"}), {{"__Clasif3", "Clasificador"}}),
        {{"# CC - Comparativo", each if FnTextoODef(_) = "" then "SIN CUADRO COMPARATIVO" else _, type text}},
        null, MissingField.Ignore),
    AdicTabla = try Excel.CurrentWorkbook(){[Name = "ADICIONALES"]}[Content] otherwise #table({"Centro de Costos"}, {}),
    AdicValidas = Table.SelectRows(Table.AddColumn(Table.AddColumn(AdicTabla, "__Cod", each
            FnTextoODef(Text.BeforeDelimiter(FnTextoODef(Record.FieldOrDefault(_, "Centro de Costos", null)), "-")), type text),
            "__Valor", each try Number.From(Record.FieldOrDefault(_, "Valor", null)) otherwise null, type nullable number),
        each [__Cod] <> "" and [__Valor] <> null and [__Valor] <> 0),

    CCsBD = Table.Buffer(Table.Group(Table.AddColumn(ClasifCompleto, "__Cod", each
            FnTextoODef(Text.BeforeDelimiter(FnTextoODef([Centro de Costos]), "-")), type text), {"__Cod"}, {
        {"__CC", each List.First(List.Select([Centro de Costos], each FnTextoODef(_) <> ""), null), type text},
        {"__MaxCap", each List.Max(List.Transform([Capitulo], each
            try Number.From(Text.BeforeDelimiter(FnTextoODef(_), "-")) otherwise null), 0), type number}})),
    AdicConCC0 = Table.ExpandTableColumn(Table.NestedJoin(AdicValidas, {"__Cod"}, CCsBD, {"__Cod"}, "__B", JoinKind.LeftOuter),
        "__B", {"__CC", "__MaxCap"}),
    AdicSinRegistro = Table.SelectColumns(
        Table.NestedJoin(CCsBD, {"__Cod"}, AdicConCC0, {"__Cod"}, "__Existe", JoinKind.LeftAnti),
        {"__Cod", "__CC", "__MaxCap"}),
    AdicConCC = Table.Combine({AdicConCC0, Table.AddColumn(AdicSinRegistro, "__Valor", each 0, type number)}),
    AdicFilas = Table.FromRecords(List.Transform(Table.ToRecords(AdicConCC), (r) => [
        #"Centro de Costos" = if r[__CC] <> null then r[__CC] else FnTextoODef(r[#"Centro de Costos"]),
        Ins = "ADICIONAL",
        Actividad = FnTextoODef(Record.FieldOrDefault(r, "Descripci#(00F3)n", null)),
        Capitulo = Text.From((if r[__MaxCap] = null then 0 else r[__MaxCap]) + 1) & "-ADICIONALES",
        Subcapitulo = "ADICIONALES",
        Tipo = "ADICIONAL",
        #"# CC - Comparativo" = "SIN CUADRO COMPARATIVO",
        Clasificador = "zOtros",
        #"VT Proyectado Colsubsidio" = r[__Valor],
        #"VT Adicionales" = r[__Valor]]),
        type table [#"Centro de Costos" = text, Ins = text, Actividad = text, Capitulo = text, Subcapitulo = text,
            Tipo = text, #"# CC - Comparativo" = text, Clasificador = text,
            #"VT Proyectado Colsubsidio" = nullable number, #"VT Adicionales" = nullable number]),
    ConAdicionales = Table.Combine({ClasifCompleto, AdicFilas}),

    FnNumONull = (v as any) as nullable number => try Number.From(v) otherwise null,
    ConAlmacen = Table.AddColumn(Table.AddColumn(ConAdicionales, "Cant. entradas", each
            if [Tipo] = "COMPRAS" and FnTieneTexto([#"#ENTRADA"]) then FnNumONull([Cantidad Cortes]) else null, type nullable number),
        "Cant. salidas", each
            if [Tipo] = "COMPRAS" and FnTieneTexto([#"#SALIDA"]) then FnNumONull([Cantidad Cons Cols]) else null, type nullable number),
    TeoVacio = #table({"Material", "Tipologia", "Centro de Costos", "TipoElemento", "Piso", "Ins", "CantTotal"}, {}),
    TeoConcreto = try Table.Buffer(TEORICO_TODO) otherwise TeoVacio,
    TeoMamposteria = try Table.Buffer(TEORICO_MAMPOSTERIA) otherwise TeoVacio,

    TeoRepartido = Table.Combine({TeoConcreto, TeoMamposteria}),
    TeoConNombre = if Value.Is(FX_INSUMO_CC, type function) then Table.AddColumn(TeoRepartido, "__InsCC", each FX_INSUMO_CC([Centro de Costos], [Ins]), type text)
                   else error "No se pudo validar el mapa de equivalencias por centro de costos.",
    TeoResuelto = Table.RenameColumns(Table.RemoveColumns(TeoConNombre, {"Ins"}), {{"__InsCC", "Ins"}}),
    TeoNormal = Table.SelectRows(Table.TransformColumns(TeoResuelto, {
        {"Material", FnTextoODef, type text}, {"Tipologia", FnTextoODef, type text},
        {"Centro de Costos", each Text.Upper(FnTextoODef(_)), type text},
        {"TipoElemento", FnTextoODef, type text}, {"Piso", FnTextoODef, type text},
        {"Ins", each try Text.Upper(FnTextoODef(FnRemoveAccentMarks(_))) otherwise "", type text}}), each [Ins] <> ""),
    ClasifSalidas = FnMapaModa(
        Table.SelectRows(ClasifCompleto, each FnTieneTexto([#"#SALIDA"]) and FnTieneTexto([Clasificador]) and [Clasificador] <> "zOtros"),
        each FnClaveMaestro([Ins]), "Clasificador"),
    TeoFilas = Table.RenameColumns(Table.AddColumn(Table.AddColumn(Table.AddColumn(TeoNormal,
        "Tipo", each "TEORICO", type text),
        "# CC - Comparativo", each "SIN CUADRO COMPARATIVO", type text),
        "Clasificador", each
            let k = FnClaveMaestro([Ins]), m = FnBuscar(ClasifMaestro, k), s = FnBuscar(ClasifSalidas, k)
            in if m <> null then m
               else if [Material] <> "MAMPOSTERIA" then Text.Proper([Material])
               else if s <> null then s else "Mamposteria", type text),
        {{"CantTotal", "Cant. teorica"}}),
    ConTeorico = Table.Combine({ConAlmacen, TeoFilas}),

    FnOCDin = (v as any) as text => Text.Select(FnTextoODef(v), {"0".."9"}),
    FnModaTxt = (l as list) as nullable text =>
        let x = List.Select(List.Transform(l, FnTextoODef), each _ <> "") in if List.IsEmpty(x) then null else List.Mode(x),
    FnNITDin = (v as any) as text => let d = FnOCDin(v) in if Text.Length(d) = 10 and List.Contains({"8", "9"}, Text.Start(d, 1)) then Text.Start(d, 9) else d,
    FnFacturaDin = (v as any) as text => Text.TrimStart(Text.Select(Text.Upper(FnTextoODef(v)), {"A".."Z", "0".."9"}), "0"),
    FnSumaDin = (l as list) as number => List.Sum(List.Transform(l, each let n = FnNumONull(_) in if n = null then 0 else n)),
    ConOCd = Table.Buffer(Table.AddColumn(ConTeorico, "__OC", each FnOCDin([#"# OC / Contrato"]), type text)),
    FilasOC = Table.SelectRows(ConOCd, each [__OC] <> ""),
    DimOC = Table.Group(Table.SelectRows(FilasOC, each [Tipo] = "CONTRATO" or [Tipo] = "COMPRAS"), {"__OC"}, {
        {"Tipo OC", each if List.Contains([Tipo], "CONTRATO") then "Contrato" else "Orden de compra", type text},
        {"CC OC", each FnModaTxt([Centro de Costos]), type text},
        {"Contratista OC", each let c = FnModaTxt([Nombre Contratista]) in if c = null then "(sin contratista)" else c, type text}}),
    ClasifDin = Table.Group(Table.SelectRows(FilasOC, each [Tipo] <> "PROVISIONES"), {"__OC"}, {{"Clasificador OC", each FnModaTxt([Clasificador]), type text}}),
    DimOCCompleta = Table.Buffer(Table.TransformColumns(
        Table.ExpandTableColumn(Table.NestedJoin(DimOC, {"__OC"}, ClasifDin, {"__OC"}, "C", JoinKind.LeftOuter), "C", {"Clasificador OC"}),
        {{"Clasificador OC", each if _ = null then "(sin clasificador)" else _, type text}})),
    OCsValidas = List.Buffer(DimOCCompleta[__OC]),

    CCProyecto = FnModaTxt(ConTeorico[Centro de Costos]),
    MapaCCOC = Record.FromList(DimOCCompleta[CC OC], DimOCCompleta[__OC]),
    FnCCOC = (oc as nullable text) as nullable text => let c = if oc = null then null else Record.FieldOrDefault(MapaCCOC, oc, null) in if c = null then CCProyecto else c,

    FactVacia = #table(type table [OC = text, NIT = text, Factura = text, Valor = number, FacturaClave = text], {}),
    FactTodas = Table.Buffer(try SINCO_FACTURAS otherwise FactVacia),
    FactSinco = Table.SelectRows(FactTodas, each [OC] <> "" and List.Contains(OCsValidas, [OC])),
    MapaFactOC = let
            t = Table.SelectRows(Table.AddColumn(FactTodas, "__K", each FnNITDin([NIT]) & "|" & [FacturaClave], type text), each [FacturaClave] <> ""),
            g = Table.Group(t, {"__K"}, {{"OC", each List.Mode([OC]), type text}})
        in Record.FromList(g[OC], g[__K]),
    PorFactura = Table.Group(FactSinco, {"OC", "FacturaClave"}, {
        {"Factura", each List.First(List.Select([Factura], each _ <> "" and _ <> null), "(sin factura)"), type text},
        {"Invertido factura", each List.Sum([Valor]), type number}}),
    TextoFactura = Record.FromList(PorFactura[Factura], List.Transform(Table.ToRecords(PorFactura), each [OC] & "|" & [FacturaClave])),

    InvertidoOC = Table.Group(FilasOC, {"__OC"}, {{"__Inv", each FnSumaDin([VT Cortes]), type number}}),
    FacturadoOC = Table.Group(PorFactura, {"OC"}, {{"__Fact", each List.Sum([Invertido factura]), type number}}),
    SaldoDin = Table.SelectRows(Table.AddColumn(
        Table.ExpandTableColumn(Table.NestedJoin(
            Table.SelectRows(InvertidoOC, each List.Contains(OCsValidas, [__OC])), {"__OC"}, FacturadoOC, {"OC"}, "F", JoinKind.LeftOuter), "F", {"__Fact"}),
        "__Dif", each [__Inv] - (if [__Fact] = null then 0 else [__Fact]), type number), each Number.Abs([__Dif]) >= 1),
    FilasFacturaBase = Table.Combine({
        Table.SelectColumns(Table.SelectRows(PorFactura, each [FacturaClave] <> ""), {"OC", "Factura", "Invertido factura"}),
        Table.Group(Table.Combine({
            Table.SelectColumns(Table.TransformColumns(Table.SelectRows(PorFactura, each [FacturaClave] = ""), {{"Factura", each "(sin factura)", type text}}), {"OC", "Factura", "Invertido factura"}),
            Table.FromRecords(List.Transform(Table.ToRecords(SaldoDin), (r) => [OC = r[__OC], Factura = "(sin factura)", #"Invertido factura" = r[__Dif]]),
                type table [OC = text, Factura = text, #"Invertido factura" = number])}),
            {"OC", "Factura"}, {{"Invertido factura", each List.Sum([Invertido factura]), type number}})}),
    FilasFactura = Table.AddColumn(Table.AddColumn(Table.AddColumn(Table.AddColumn(FilasFacturaBase,
        "Tipo", each "FACTURA", type text),
        "# OC / Contrato", each [OC], type text),
        "Centro de Costos", each FnCCOC([OC]), type text),
        "# CC - Comparativo", each "SIN CUADRO COMPARATIVO", type text),
    ConFacturas = Table.Combine({ConOCd, Table.AddColumn(FilasFactura, "__OC", each [OC], type text)}),

    ConUbicacion = Table.AddColumn(ConFacturas, "__U", each
        if [Tipo] = "FACTURA" then [O = [OC], F = [Factura]]
        else if [Tipo] = "PROVISIONES" then
            let
                clave = FnFacturaDin([No_Factura]),
                ocS = if clave = "" then null else Record.FieldOrDefault(MapaFactOC, FnNITDin([NIT]) & "|" & clave, null),
                oc = if List.Contains(OCsValidas, [__OC]) then [__OC] else if ocS <> null and List.Contains(OCsValidas, ocS) then ocS else null
            in [O = oc, F = if clave = "" then "(sin factura)"
                    else Record.FieldOrDefault(TextoFactura, (if oc = null then "" else oc) & "|" & clave, FnTextoODef([No_Factura]))]
        else [O = if List.Contains(OCsValidas, [__OC]) then [__OC] else null, F = "(sin factura)"]),
    Ubicadas = Table.Buffer(Table.AddColumn(Table.AddColumn(
        Table.RemoveColumns(ConUbicacion, {"OC", "Factura"}, MissingField.Ignore),
        "OC", each [__U][O], type text), "Factura", each if [__U][O] = null then null else [__U][F], type text)),
    NoProvPorOC = Table.Group(Table.SelectRows(Ubicadas, each [OC] <> null and FnTieneTexto([No_Prov])), {"OC"}, {
        {"No_Prov OC", each Text.Combine(List.Sort(List.Distinct(List.Transform([No_Prov], each Text.From(try Int64.From(Number.From(_)) otherwise _))), each try Number.From(_) otherwise 0), " | "), type text}}),
    ConDimOC = Table.ExpandTableColumn(Table.NestedJoin(Ubicadas, {"OC"}, DimOCCompleta, {"__OC"}, "D", JoinKind.LeftOuter),
        "D", {"Tipo OC", "CC OC", "Contratista OC", "Clasificador OC"}),
    ConNoProv = Table.ExpandTableColumn(Table.NestedJoin(ConDimOC, {"OC"}, NoProvPorOC, {"OC"}, "N", JoinKind.LeftOuter), "N", {"No_Prov OC"}),

    InsSinDesperdicio = {"T_TRANSPORTE CONCRETO (M3)"},
    ConSal = Table.AddColumn(ConNoProv, "__Sal", each [Tipo] = "COMPRAS" and FnTieneTexto([#"#SALIDA"]), type logical),
    ConDesp = Table.AddColumn(Table.AddColumn(ConSal,
        "Clasif desperdicio", each if ([Tipo] = "TEORICO" or [__Sal]) and not List.Contains(InsSinDesperdicio, FnTextoODef([Ins])) then [Clasificador] else "(no aplica)", type text),
        "Contratista almacen", each if [Tipo] = "TEORICO" then "(TEORICO)" else if [__Sal] then [Nombre Contratista] else "(no aplica)", type text),
    ConSinProv = Table.TransformColumns(ConDesp, {{"Tipo OC", each if _ = null then "(no aplica)" else _, type text}}),
    ConProvTxt = Table.AddColumn(Table.RenameColumns(ConSinProv, {{"No_Prov OC", "__NP"}}), "No_Prov OC", each
        if [OC] = null then null else if [__NP] = null then "(sin provision)" else [__NP], type text),

    ConSeg = Table.AddColumn(Table.AddColumn(ConProvTxt,
        "No_Prov seg", each if [Tipo] = "PROVISIONES" and [OC] <> null and FnTieneTexto([No_Prov]) then Text.From(try Int64.From(Number.From([No_Prov])) otherwise [No_Prov])
            else if [No_Prov OC] = "(sin provision)" then "(sin provision)" else "(otros registros)", type text),
        "Fila real", each if [Tipo] = "PROVISIONES" and [OC] <> null and FnTieneTexto([No_Prov]) then null else 1, Int64.Type),
    ValoresOC = Table.Group(Table.SelectRows(ConSeg, each [OC] <> null), {"OC"}, {
        {"Aseg prov", each List.Sum([VT Asegurada]), type number},
        {"Comp prov", each List.Sum([VR total aprobacion]), type number},
        {"Inv prov", each List.Sum([Invertido factura]), type number},
        {"Desc prov", each List.Sum([Valor descuento]), type number}}),
    ParesOCProv = Table.Distinct(Table.SelectColumns(Table.SelectRows(ConSeg, each [Fila real] = null),
        {"OC", "No_Prov seg", "Tipo OC", "CC OC", "Contratista OC", "Clasificador OC"})),
    FilasProvOC = Table.ExpandTableColumn(Table.NestedJoin(ParesOCProv, {"OC"}, ValoresOC, {"OC"}, "V", JoinKind.LeftOuter), "V", {"Aseg prov", "Comp prov", "Inv prov", "Desc prov"}),

    InvPorFactura = Table.Group(Table.SelectRows(ConSeg, each [Tipo] = "FACTURA" and [Factura] <> "(sin factura)"), {"OC", "Factura"}, {{"Inv prov", each List.Sum([Invertido factura]), type number}}),

    SumaFactOC = Table.Buffer(Table.Group(InvPorFactura, {"OC"}, {{"__S", each List.Sum([Inv prov]), type number}})),
    ValoresOCB = Table.Buffer(ValoresOC),
    MapaSumaFact = Record.FromList(SumaFactOC[__S], SumaFactOC[OC]),
    MapaAseg = Record.FromList(ValoresOCB[Aseg prov], ValoresOCB[OC]),
    MapaComp = Record.FromList(ValoresOCB[Comp prov], ValoresOCB[OC]),
    FnFactorOC = (mapa as record, oc as text) as number =>
        let v = Record.FieldOrDefault(mapa, oc, null), s = Record.FieldOrDefault(MapaSumaFact, oc, null)
        in if v = null or s = null or v <= 0 or s <= 0 then 0 else List.Min({1, v / s}),
    FnNz = (n as nullable number) as number => if n = null then 0 else n,
    FilasProvFact0 = Table.ExpandTableColumn(Table.NestedJoin(ParesOCProv, {"OC"}, InvPorFactura, {"OC"}, "V", JoinKind.Inner), "V", {"Factura", "Inv prov"}),
    FilasProvFact = Table.AddColumn(Table.AddColumn(FilasProvFact0,
        "Aseg prov", each FnNz([Inv prov]) * FnFactorOC(MapaAseg, [OC]), type number),
        "Comp prov", each FnNz([Inv prov]) * FnFactorOC(MapaComp, [OC]), type number),
    RepartidoProv = Table.Group(FilasProvFact, {"OC", "No_Prov seg"}, {
        {"__Rep", each List.Sum([Inv prov]), type number}, {"__RepA", each List.Sum([Aseg prov]), type number}, {"__RepC", each List.Sum([Comp prov]), type number}}),
    FnResta = (tabla as table, col as text, rep as text) as table =>
        Table.ReplaceValue(tabla, each Record.Field(_, col), each FnNz(Record.Field(_, col)) - FnNz(Record.Field(_, rep)), Replacer.ReplaceValue, {col}),
    FilasProvResto = Table.RemoveColumns(FnResta(FnResta(FnResta(
        Table.ExpandTableColumn(Table.NestedJoin(FilasProvOC, {"OC", "No_Prov seg"}, RepartidoProv, {"OC", "No_Prov seg"}, "R", JoinKind.LeftOuter), "R", {"__Rep", "__RepA", "__RepC"}),
        "Inv prov", "__Rep"), "Aseg prov", "__RepA"), "Comp prov", "__RepC"), {"__Rep", "__RepA", "__RepC"}),

    FnClavePar = (oc as text, prov as text) as text => oc & "|" & prov,
    PrimeraParT = Table.Buffer(Table.Group(FilasProvFact, {"OC", "No_Prov seg"}, {{"__F", each List.Min([Factura]), type text}})),
    MapaPrimeraPar = Record.FromList(PrimeraParT[__F], List.Transform(Table.ToRecords(PrimeraParT), each FnClavePar([OC], [#"No_Prov seg"]))),
    FnRestoCero = (r as record) as logical => Number.Abs(FnNz(r[Aseg prov])) < 1 and Number.Abs(FnNz(r[Comp prov])) < 1 and Number.Abs(FnNz(r[Inv prov])) < 1 and Number.Abs(FnNz(r[Desc prov])) < 1,
    FilasProvRestoV = Table.SelectRows(FilasProvResto, each not (FnRestoCero(_) and Record.HasFields(MapaPrimeraPar, FnClavePar([OC], [#"No_Prov seg"])))),
    FilasProvBase = Table.Combine({Table.AddColumn(FilasProvRestoV, "Factura", each "(sin factura)", type text), FilasProvFact}),

    TeoClaves = Table.Group(Table.SelectRows(ConSeg, each [Tipo] = "TEORICO" and [Clasif desperdicio] <> "(no aplica)"),
        {"Centro de Costos", "Clasif desperdicio", "Ins", "Tipologia", "TipoElemento"}, {{"Teorico fijo", each List.Sum([Cant. teorica]), type number}}),
    ContratistasSal = List.Buffer(List.RemoveNulls(List.Distinct(Table.SelectRows(ConSeg, each [__Sal] = true and [Clasif desperdicio] <> "(no aplica)")[Contratista almacen]))),
    FilasTeoBase = Table.AddColumn(Table.ExpandListColumn(Table.AddColumn(TeoClaves, "Contratista almacen", each ContratistasSal), "Contratista almacen"),
        "Clasificador", each [Clasif desperdicio], type text),
    FnApoyo = (tabla as table, valores as record) as table =>
        List.Accumulate(Record.FieldNames(valores), tabla, (st, c) => Table.AddColumn(st, c, each Record.Field(valores, c))),
    FilasProv = Table.AddColumn(FnApoyo(Table.AddColumn(FilasProvBase, "Centro de Costos", each FnCCOC([OC]), type text), [Tipo = "FACTURA", #"# CC - Comparativo" = "SIN CUADRO COMPARATIVO",
        #"Clasif desperdicio" = "(no aplica)", #"Contratista almacen" = "(no aplica)"]), "# OC / Contrato", each [OC], type text),
    FilasTeo = FnApoyo(FilasTeoBase, [Tipo = "FACTURA", #"# CC - Comparativo" = "SIN CUADRO COMPARATIVO", #"Tipo OC" = "(no aplica)", #"No_Prov seg" = "(otros registros)", #"Fila real" = 1]),

    FnEsFactura = (r as record) as logical => r[Tipo] = "FACTURA" and r[OC] <> null and r[Factura] <> "(sin factura)",
    FnReparte = (tabla as table, col as text, mapa as record) as table =>
        Table.ReplaceValue(tabla, each Record.Field(_, col), each if FnEsFactura(_) then FnNz([Invertido factura]) * FnFactorOC(mapa, [OC]) else Record.Field(_, col), Replacer.ReplaceValue, {col}),
    ConSegRep = FnReparte(FnReparte(ConSeg, "VT Asegurada", MapaAseg), "VR total aprobacion", MapaComp),
    RestoBase = Table.Distinct(Table.RemoveColumns(Table.SelectRows(ConSeg, FnEsFactura), {"VT Asegurada", "VR total aprobacion", "Invertido factura", "Factura"}), {"OC"}),
    FnMontoResto = (oc as text, mapa as record) as number => - FnNz(Record.FieldOrDefault(MapaSumaFact, oc, null)) * FnFactorOC(mapa, oc),
    FilasResto = Table.SelectRows(Table.AddColumn(Table.AddColumn(Table.AddColumn(RestoBase,
        "Factura", each "(sin factura)", type text),
        "VT Asegurada", each FnMontoResto([OC], MapaAseg), type number),
        "VR total aprobacion", each FnMontoResto([OC], MapaComp), type number),
        each [VT Asegurada] <> 0 or [VR total aprobacion] <> 0),

    ConRealBase = Table.Buffer(Table.Combine({ConSegRep, FilasResto})),
    SinFactBase = Table.Buffer(Table.SelectRows(ConRealBase, each [OC] <> null and [Factura] = "(sin factura)")),
    FnCeroLista = (l as list) as logical => Number.Abs(FnSumaDin(l)) < 1,
    CeroOC = Table.SelectRows(Table.Group(SinFactBase, {"OC"}, {{"__Cero", each FnCeroLista([VT Asegurada]) and FnCeroLista([VR total aprobacion]) and FnCeroLista([Invertido factura]) and FnCeroLista([VR_Bruto_con_desc]) and FnCeroLista([Valor descuento]), type logical}}), each [__Cero]),
    PrimeraOCT = Table.Buffer(Table.Group(InvPorFactura, {"OC"}, {{"__F", each List.Min([Factura]), type text}})),
    MapaPrimeraOC = Record.FromList(PrimeraOCT[__F], PrimeraOCT[OC]),
    MapaCeroOC = Record.FromList(List.Repeat({true}, Table.RowCount(CeroOC)), CeroOC[OC]),

    BrutoPar = Table.Group(Table.SelectRows(SinFactBase, each [Fila real] = null), {"OC", "No_Prov seg"}, {{"__B", each FnSumaDin([VR_Bruto_con_desc]), type number}}),
    MapaBrutoPar = Record.FromList(BrutoPar[__B], List.Transform(Table.ToRecords(BrutoPar), each FnClavePar([OC], [#"No_Prov seg"]))),
    RestosCero = Table.SelectRows(FilasProvResto, each FnRestoCero(_) and Number.Abs(Record.FieldOrDefault(MapaBrutoPar, FnClavePar([OC], [#"No_Prov seg"]), 0)) < 1),
    MapaParCero = Record.FromList(List.Repeat({true}, Table.RowCount(RestosCero)), List.Transform(Table.ToRecords(RestosCero), each FnClavePar([OC], [#"No_Prov seg"]))),
    FnFacturaFinal = (r as record) as nullable text =>
        if r[OC] = null or r[Factura] <> "(sin factura)" then r[Factura]
        else let
            ocCero = Record.HasFields(MapaCeroOC, r[OC]) and Record.HasFields(MapaPrimeraOC, r[OC]),
            esProv = r[Fila real] = null,
            par = if esProv then FnClavePar(r[OC], r[#"No_Prov seg"]) else "",
            enPar = if esProv then Record.FieldOrDefault(MapaPrimeraPar, par, null) else null
        in if esProv and enPar <> null and (ocCero or Record.HasFields(MapaParCero, par)) then enPar
            else if ocCero then Record.Field(MapaPrimeraOC, r[OC]) else r[Factura],
    ConFacturaFinal = Table.AddColumn(Table.RenameColumns(ConRealBase, {{"Factura", "__FacturaBase"}}), "Factura", each FnFacturaFinal(_ & [Factura = [__FacturaBase]]), type text),
    ConApoyo = Table.Combine({ConFacturaFinal, FilasProv, FilasTeo}),

    ConCantidadClasificada = Table.AddColumn(ConApoyo, "Cantidad SINCO clasificado", each
        if [Tipo] = "CC" and List.Contains({"Contrato", "Orden de compra"}, [Tipo OC]) then ToNumber0([#"Cant. aprobacion"]) else 0, type number),
    ConValorClasificado = Table.AddColumn(ConCantidadClasificada, "VT SINCO clasificado", each
        if [Tipo] = "CC" and List.Contains({"Contrato", "Orden de compra"}, [Tipo OC]) then ToNumber0([#"VR total aprobacion"]) else 0, type number),
    FnClaveAseg = (r as record, nivel as number) as nullable text =>
        let oc = FnOCKeyValida(r[#"# OC / Contrato"]),
            cc = FnClaveMaestro(r[Centro de Costos]),
            ins = FnClaveMaestro(r[Ins]),
            act = FnParseCodAct(r[Codigo act], r[Actividad])
        in if oc = null or cc = "" or (nivel > 0 and ins = "") or (nivel > 1 and act = "") then null
           else cc & "|" & oc & (if nivel > 0 then "|" & ins else "") & (if nivel > 1 then "|" & act else ""),
    PuenteAseg = Table.Buffer(Table.SelectRows(FilasPuente, each
        not List.Contains({"SIN CUADRO COMPARATIVO", "POR ADJUDICAR"}, FnClaveComp(_)))),
    CompAsegExacto = FnMapaUnico(PuenteAseg, each FnClaveAseg(_, 2), "# CC - Comparativo"),
    CompAsegInsumo = FnMapaUnico(PuenteAseg, each FnClaveAseg(_, 1), "# CC - Comparativo"),
    CompAsegOC = FnMapaUnico(PuenteAseg, each FnClaveAseg(_, 0), "# CC - Comparativo"),
    ConComparativoAseg0 = Table.AddColumn(ConValorClasificado, "__ComparativoAseg", each
        if [Tipo] = "CONTRATO" or ([Tipo] = "COMPRAS" and ToNumber0([VT Asegurada]) <> 0) then
            let a = FnBuscar(CompAsegExacto, FnClaveAseg(_, 2)),
                b = if a <> null then a else FnBuscar(CompAsegInsumo, FnClaveAseg(_, 1)),
                c = if b <> null then b else FnBuscar(CompAsegOC, FnClaveAseg(_, 0))
            in if c = null then "SIN CUADRO COMPARATIVO" else c
        else [#"# CC - Comparativo"], type text),
    ConComparativoAseg = Table.RenameColumns(
        Table.RemoveColumns(ConComparativoAseg0, {"# CC - Comparativo"}), {{"__ComparativoAseg", "# CC - Comparativo"}}),
    ConCantidadAsegSINCO = Table.AddColumn(ConComparativoAseg, "Cantidad asegurada SINCO", each
        if [Tipo] = "CONTRATO" or [Tipo] = "COMPRAS" then ToNumber0([Cantidad asegurada]) else 0, type number),
    ConValorAsegSINCO = Table.AddColumn(ConCantidadAsegSINCO, "VT asegurado SINCO", each
        if [Tipo] = "CONTRATO" or [Tipo] = "COMPRAS" then ToNumber0([VT Asegurada]) else 0, type number),
    FinalRecortada = Table.SelectColumns(ConValorAsegSINCO, ColumnasFinalesArboleda & {"OC", "Tipo OC", "CC OC", "Contratista OC", "Clasificador OC", "Factura", "Invertido factura", "No_Prov seg", "Fila real", "Aseg prov", "Comp prov", "Inv prov", "Desc prov", "Teorico fijo", "Clasif desperdicio", "Contratista almacen", "VT Adicionales",
        "Cant. entradas", "Cant. salidas", "Cant. teorica", "Material", "Tipologia", "TipoElemento", "Piso", "Codigo ins", "Cantidad SINCO clasificado", "VT SINCO clasificado", "Cantidad asegurada SINCO", "VT asegurado SINCO"}, MissingField.UseNull),

    TablaMaestraFinal = Table.Buffer(FinalRecortada)
in
    TablaMaestraFinal
