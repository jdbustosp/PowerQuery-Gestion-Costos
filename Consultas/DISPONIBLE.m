let
    // ============================================================
    // 1. FUNCIONES DE LIMPIEZA EXTREMA
    // ============================================================
    // FnCleanText: Quita espacios fantasma, convierte a mayúsculas y vuelve los vacíos ("") en null reales.
    FnCleanText = (t as any) as nullable text => if t = null then null else let txt = Text.Trim(Text.From(t)) in if txt = "" then null else Text.Upper(txt),
    
    FnRemoveAccentsSymbols = (t as nullable text) as nullable text => let cleanT = FnCleanText(t) in if cleanT = null then null else let replacements = {{"Á","A"},{"É","E"},{"Í","I"},{"Ó","O"},{"Ú","U"},{"º",""},{"°",""},{"¨",""}}, result = List.Accumulate(replacements, cleanT, (state, current) => Text.Replace(state, current{0}, current{1})) in result,

    // ============================================================
    // 2. LECTURA DE BASES
    // ============================================================
    FuentePPTO  = Excel.CurrentWorkbook(){[Name="PPTO_BD"]}[Content],
    FuenteDetCC = Excel.CurrentWorkbook(){[Name="COMPARATIVOS"]}[Content],

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
    Bloques_Filled = Table.TransformColumns(Bloques_ExpandedAdj, {{"ValorTotal_PPTO_Bloque", each _ ?? 0, type number}, {"Unitario_PPTO_Bloque", each _ ?? 0, type number}, {"ValorAdjudicado_Bloque", each _ ?? 0, type number}, {"CantAdjudicada_Bloque", each _ ?? 0, type number}}),
    
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
