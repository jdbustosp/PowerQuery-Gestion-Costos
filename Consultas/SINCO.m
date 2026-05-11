let
    // 🔥 CONEXIÓN DIRECTA EN CASCADA
    Source = BD, 
    TablaComparativo = COMPARATIVOS,
    FnRemoveAccentsSymbols = F_Globales[FnRemoveAccentsSymbols],

    ListaOC_Excluir = List.Distinct(List.RemoveNulls(List.Transform(try TablaComparativo[#"# OC / Contrato"] otherwise {}, each if _ = null or Text.Trim(Text.From(_)) = "" then null else Text.Trim(Text.From(_))))),
    // Lookup O(1) en vez de List.Contains O(n) por fila
    SetOC = Record.FromList(List.Repeat({true}, List.Count(ListaOC_Excluir)), ListaOC_Excluir),

    FiltroExclusion = Table.SelectRows(Source, each
        (Text.Upper(if [Tipo] = null then "" else [Tipo]) <> "PPTO") and
        (let ocText = if [#"# OC / Contrato"] = null then "" else Text.Trim(Text.From([#"# OC / Contrato"])) in ocText = "" or not Record.HasFields(SetOC, {ocText}))
    ),

    // try/otherwise false: si alguna fila de BD trajo un Error en VT Asegurada (en vez de null o 0)
    // el filtro sin proteccion falla silenciosamente y devuelve 0 filas
    FiltroCeros = Table.SelectRows(FiltroExclusion, each try ([VT Asegurada] <> 0 and [VT Asegurada] <> null) otherwise false),

    // 🚀 Limpiar acentos en Nombre Contratista y Descripcion contrato
    LimpiezaTextos = Table.TransformColumns(FiltroCeros, {
        {"Nombre Contratista", each FnRemoveAccentsSymbols(if _ = null then null else Text.Trim(Text.From(_))), type text},
        {"Descripcion contrato", each FnRemoveAccentsSymbols(if _ = null then null else Text.Trim(Text.From(_))), type text}
    }, null, MissingField.Ignore),

    // Orden final de columnas (sin Cantidad_Calc, V/U ppto (CC), Valor Total ppto (CC))
    ColumnasFinales = Table.SelectColumns(LimpiezaTextos, 
        {"Centro de Costos", "Subcapitulo", "Capitulo", "Actividad", "Codigo ins", "Ins", 
         "# OC / Contrato", "Nombre Contratista", "Cantidad asegurada", "V/U asegurada", 
         "VT Asegurada", "Descripcion contrato", "Tipo"}, MissingField.Ignore),

    ResultadoFinal = Table.Buffer(ColumnasFinales)
in
    ResultadoFinal
