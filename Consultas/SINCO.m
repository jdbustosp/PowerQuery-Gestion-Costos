let
    // 🔥 CONEXIÓN DIRECTA EN CASCADA
    Source = BD, 
    TablaComparativo = COMPARATIVOS,
    FnRemoveAccentsSymbols = F_Globales[FnRemoveAccentsSymbols],

    ListaOC_Excluir = List.Distinct(List.RemoveNulls(List.Transform(try TablaComparativo[#"# OC / Contrato"] otherwise {}, each if _ = null or Text.Trim(Text.From(_)) = "" then null else Text.Trim(Text.From(_))))),

    FiltroExclusion = Table.SelectRows(Source, each 
        (Text.Upper(if [Tipo] = null then "" else [Tipo]) <> "PPTO") and 
        (let ocText = if [#"# OC / Contrato"] = null then "" else Text.Trim(Text.From([#"# OC / Contrato"])) in ocText = "" or not List.Contains(ListaOC_Excluir, ocText))
    ),

    // Filtro de ceros: solo VT Asegurada (ya no hay columnas PPTO)
    FiltroCeros = Table.SelectRows(FiltroExclusion, each [VT Asegurada] <> 0 and [VT Asegurada] <> null),

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

    ResultadoFinal = ColumnasFinales
in
    ResultadoFinal
