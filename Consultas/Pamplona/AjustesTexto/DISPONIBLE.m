let
    Limpiar = F_Globales[FnLimpiarTabla],
    Texto = F_Globales[FnTrimText],
    CodigoInsumo = F_Globales[FnCodigoInsumo],
    ActividadVisible = F_Globales[FnActividadVisible],
    SepararSubcap = F_Globales[FnSepararSubcapDeNombre],
    AjustarSubcap = F_Globales[FnAplicarOverrideSubcap],
    ClaveTexto = (v as any) as text =>
        if v = null then "" else Text.Upper(F_Globales[FnRemoveAccentsSymbols](Texto(v))),
    CodigoActividad = (act as nullable text) as text =>
        let prefijo = ClaveTexto(if act = null then null else Text.BeforeDelimiter(act, "-"))
        in if Text.Select(prefijo, {"0".."9", "."}) = prefijo then prefijo else "",
    Datos = (r as record) as record =>
        let
            act = ActividadVisible(r[Actividad]),
            sub = AjustarSubcap(r[Subcapitulo]),
            sep = if ClaveTexto(sub) = "" then SepararSubcap(act) else null,
            subFinal = if sep = null then sub else sep[Subcap],
            codAct = CodigoActividad(act),
            codIns = CodigoInsumo(r[#"Codigo ins"])
        in [
            CCKey = ClaveTexto(r[Centro de Costos]),
            CapKey = ClaveTexto(r[Capitulo]),
            SubKey = ClaveTexto(subFinal),
            ActKey = if codAct = "" then ClaveTexto(act) else codAct,
            InsKey = ClaveTexto(r[Ins]),
            CodIns = codIns,
            SubVisible = subFinal
        ],
    ConDatos = (t as table) as table =>
        let
            columnas = {"CCKey", "CapKey", "SubKey", "ActKey", "InsKey", "CodIns", "SubVisible"},
            listo = Limpiar(t)
        in Table.ExpandRecordColumn(Table.AddColumn(listo, "__Datos", each Datos(_)), "__Datos", columnas),

    BasePpto = ConDatos(Table.SelectColumns(PPTO_BD,
        {"Centro de Costos", "Codigo act", "Codigo ins", "Capitulo", "Actividad", "Subcapitulo",
         "Ins", "Cantidad Presupuesto", "VT Presupuesto", "V/U Presupuesto"})),
    ClaveContexto = {"CCKey", "CapKey", "SubKey", "ActKey"},
    PptoIdentificado = Table.AddColumn(BasePpto, "__IdInsumo", each
        if [CodIns] <> null then "C:" & Text.From([CodIns], "en-US") else "N:" & [InsKey], type text),
    PptoAgrupado = Table.Buffer(Table.AddIndexColumn(Table.Group(PptoIdentificado, ClaveContexto & {"__IdInsumo"}, {
        {"Renglon", each _{0}, type record},
        {"ValorPPTO", each List.Sum([VT Presupuesto]), type number},
        {"CantidadPPTO", each List.Sum([Cantidad Presupuesto]), type number},
        {"UnitarioPPTO", each List.First(List.RemoveNulls([#"V/U Presupuesto"]), null), type nullable number}
    }), "__Bloque", 1, 1, Int64.Type)),
    Indice = Table.Buffer(Table.ExpandRecordColumn(Table.SelectColumns(PptoAgrupado,
        ClaveContexto & {"Renglon", "__Bloque"}), "Renglon", {"CodIns", "InsKey"})),

    BaseDescarga = ConDatos(Table.SelectRows(Table.SelectColumns(DESCARGAS,
        {"Centro de Costos", "Codigo ins", "Capitulo", "Actividad", "Subcapitulo",
         "Ins", "# CC - Comparativo", "Cantidad ppto (CC)", "Valor Total ppto (CC)", "V/U ppto (CC)"}),
        each [#"# CC - Comparativo"] <> null and [#"Valor Total ppto (CC)"] <> null)),
    // Cruce directo por codigo y contexto. Solo las descargas que no tienen
    // coincidencia exacta consultan las alternativas del mismo CC y actividad.
    IndiceCodigo = Table.SelectRows(Indice, each [CodIns] <> null),
    ClaveCodigo = ClaveContexto & {"CodIns"},
    DescExacta = Table.Buffer(Table.ExpandTableColumn(Table.NestedJoin(
        BaseDescarga, ClaveCodigo, IndiceCodigo, ClaveCodigo, "__Exacto", JoinKind.LeftOuter),
        "__Exacto", {"__Bloque"})),
    SinBloque = Table.RemoveColumns(Table.SelectRows(DescExacta, each [__Bloque] = null), {"__Bloque"}),
    Alternativas = Table.Group(Indice, {"CCKey", "ActKey"}, {{"__Lista", each List.Buffer(Table.ToRecords(_)), type list}}),
    ConCandidatos = Table.ExpandTableColumn(Table.NestedJoin(SinBloque, {"CCKey", "ActKey"},
        Alternativas, {"CCKey", "ActKey"}, "__Candidatos", JoinKind.LeftOuter), "__Candidatos", {"__Lista"}),
    Unico = (l as list) as nullable number => if List.Count(l) = 1 then l{0}[__Bloque] else null,
    BloqueDescarga = (d as record) as nullable number =>
        let
            candidatos = if d[__Lista] = null then {} else d[__Lista],
            contexto = List.Select(candidatos, each
                ([CapKey] = d[CapKey] or [CapKey] = "" or d[CapKey] = "") and
                ([SubKey] = d[SubKey] or [SubKey] = "" or d[SubKey] = "")),
            conCodigo = if d[CodIns] = null then {} else List.Select(contexto, each [CodIns] = d[CodIns]),
            porNombre = if d[InsKey] = "" then null else Unico(List.Select(contexto, each
                ([CodIns] = null or d[CodIns] = null) and [InsKey] = d[InsKey]))
        in if not List.IsEmpty(conCodigo) then Unico(conCodigo) else porNombre,
    DescResuelta = Table.RemoveColumns(Table.AddColumn(ConCandidatos, "__Bloque", each BloqueDescarga(_), Int64.Type), {"__Lista"}),
    DescConBloque = Table.Combine({Table.SelectRows(DescExacta, each [__Bloque] <> null), DescResuelta}),
    DescConCantidad = Table.AddColumn(DescConBloque, "Cantidad_Calc", each
        let c = [#"Cantidad ppto (CC)"], u = [#"V/U ppto (CC)"]
        in if c <> null then c else if u <> null and u <> 0 then [#"Valor Total ppto (CC)"] / u else null, type nullable number),
    DescConPpto = Table.ExpandTableColumn(Table.NestedJoin(DescConCantidad, {"__Bloque"}, PptoAgrupado,
        {"__Bloque"}, "__Ppto", JoinKind.LeftOuter), "__Ppto", {"Renglon"}, {"__Oficial"}),
    DescLista = Table.Buffer(Table.ExpandRecordColumn(Table.AddColumn(DescConPpto, "__Resultado", each
        let o = [__Oficial]
        in [
            ActividadFinal = if o = null then ActividadVisible([Actividad]) else ActividadVisible(o[Actividad]),
            SubcapFinal = if o = null then [SubVisible] else o[Subcapitulo],
            CapituloFinal = if o = null then [Capitulo] else o[Capitulo],
            CodigoAct = if o = null then CodigoActividad([Actividad]) else o[Codigo act]
        ]), "__Resultado", {"ActividadFinal", "SubcapFinal", "CapituloFinal", "CodigoAct"})),

    TotalesDescarga = Table.Group(Table.SelectRows(DescLista, each [__Bloque] <> null), {"__Bloque"}, {
        {"ValorDescargado", each List.Sum([#"Valor Total ppto (CC)"]), type number},
        {"CantidadDescargada", each List.Sum([Cantidad_Calc]), type nullable number}
    }),
    Saldos = Table.ExpandTableColumn(Table.NestedJoin(PptoAgrupado, {"__Bloque"}, TotalesDescarga,
        {"__Bloque"}, "__Desc", JoinKind.LeftOuter), "__Desc", {"ValorDescargado", "CantidadDescargada"}),
    Columnas = {"Centro de Costos", "Codigo act", "Capitulo", "Actividad", "Subcapitulo", "Codigo ins",
        "Ins", "# CC - Comparativo", "Tipo", "Cantidad_Calc", "V/U ppto (CC)", "Valor Total ppto (CC)"},
    FilasSaldo = Table.SelectRows(Table.AddColumn(Saldos, "Valor Total ppto (CC)", each
        [ValorPPTO] - (if [ValorDescargado] = null then 0 else [ValorDescargado]), type number),
        each Number.Round([#"Valor Total ppto (CC)"], 2) <> 0),
    SaldoCantidad = Table.AddColumn(FilasSaldo, "Cantidad_Calc", each
        if [CantidadPPTO] = null then null else [CantidadPPTO] - (if [CantidadDescargada] = null then 0 else [CantidadDescargada]), type nullable number),
    SaldoPpto = Table.ExpandRecordColumn(SaldoCantidad, "Renglon",
        {"Centro de Costos", "Codigo act", "Capitulo", "Actividad", "Subcapitulo", "Codigo ins", "Ins"}),
    PorAdjudicar = Table.AddColumn(Table.AddColumn(Table.RenameColumns(SaldoPpto, {{"UnitarioPPTO", "V/U ppto (CC)"}}),
        "# CC - Comparativo", each "Por adjudicar", type text), "Tipo", each "Por adjudicar", type text),
    DescColumnas = Table.RemoveColumns(DescLista, {"Actividad", "Capitulo", "Subcapitulo"}),
    Adjudicado = Table.AddColumn(Table.RenameColumns(DescColumnas,
        {{"ActividadFinal", "Actividad"}, {"CapituloFinal", "Capitulo"}, {"SubcapFinal", "Subcapitulo"}, {"CodigoAct", "Codigo act"}}),
        "Tipo", each "Adjudicado", type text),
    Resultado = Table.Combine({Table.SelectColumns(Adjudicado, Columnas), Table.SelectColumns(PorAdjudicar, Columnas)})
in
    Resultado
