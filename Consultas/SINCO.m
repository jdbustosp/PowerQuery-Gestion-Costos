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