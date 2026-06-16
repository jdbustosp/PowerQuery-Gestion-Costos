let
    // ============================================================
    // FUNCIONES AUXILIARES GLOBALES
    // ============================================================
    FxToNumberFlex = F_Globales[FxToNumberFlex],
    FnCleanText = F_Globales[FnCleanText],

    // ============================================================
    // FUENTE LOCAL: tabla DESCARGA del libro actual
    // ============================================================
    TablaDescargaLocal = try Excel.CurrentWorkbook(){[Name="DESCARGA"]}[Content] otherwise null,
    TablaDescargas =
        if TablaDescargaLocal = null then
            error Error.Record(
                "DESCARGAS",
                "No se encontró la tabla local DESCARGA en este libro.",
                [TablaRequerida = "DESCARGA"]
            )
        else
            TablaDescargaLocal,

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
