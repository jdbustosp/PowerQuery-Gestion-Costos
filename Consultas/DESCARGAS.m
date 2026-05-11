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
