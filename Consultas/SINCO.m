let
    // 🔥 CONEXIÓN DIRECTA EN CASCADA
    Source = BD, 
    TablaComparativo = COMPARATIVOS,

    ListaOC_Excluir = List.Distinct(List.RemoveNulls(List.Transform(try TablaComparativo[#"# OC / Contrato"] otherwise {}, each if _ = null or Text.Trim(Text.From(_)) = "" then null else Text.Trim(Text.From(_))))),

    FiltroExclusion = Table.SelectRows(Source, each 
        (Text.Upper(if [Tipo] = null then "" else [Tipo]) <> "PPTO") and 
        (let ocText = if [#"# OC / Contrato"] = null then "" else Text.Trim(Text.From([#"# OC / Contrato"])) in ocText = "" or not List.Contains(ListaOC_Excluir, ocText))
    ),

    FiltroCeros = Table.SelectRows(FiltroExclusion, each ([VT Asegurada] <> 0 and [VT Asegurada] <> null) or ([#"Valor Total ppto (CC)"] <> 0 and [#"Valor Total ppto (CC)"] <> null)),

    ColumnasFinales = Table.SelectColumns(FiltroCeros, {"Codigo ins", "Ins", "Actividad", "Capitulo", "Subcapitulo", "Centro de Costos", "# OC / Contrato", "Nombre Contratista", "Cantidad asegurada", "V/U asegurada", "VT Asegurada", "Cantidad_Calc", "V/U ppto (CC)", "Valor Total ppto (CC)", "Descripcion contrato", "Tipo"}, MissingField.Ignore),

    // 🚀 Table.Sort eliminado: es una de las operaciones más costosas en PQ
    // y no aporta valor si el destino es Power BI (el orden se controla en DAX/Visual)
    ResultadoFinal = ColumnasFinales
in
    ResultadoFinal
