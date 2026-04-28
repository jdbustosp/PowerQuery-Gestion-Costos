let
    // ============================================================
    // FUNCIONES AUXILIARES GLOBALES (Desde F_Globales)
    // ============================================================
    FnFormatCodigoAct = F_Globales[FnFormatCodigoAct],
    FxToNumberFlex = F_Globales[FxToNumberFlex],

    // ============================================================
    // PROCESAMIENTO DE LA TABLA MANUAL
    // ============================================================
    // 🔥 ACELERADOR: Origen directo desde Excel
    Origen = Excel.CurrentWorkbook(){[Name="Det_CC"]}[Content],

    // 1. FILTRO DE FILAS VACÍAS
    FilasValidas = Table.SelectRows(Origen, each [Ins] <> null or [Actividad] <> null),

    // 2. EXTRAER Y ESTANDARIZAR CÓDIGO DE ACTIVIDAD
    AgregadoCodAct = Table.AddColumn(
        FilasValidas, 
        "Codigo act", 
        each 
            let
                txt = Text.Trim(Text.From(if [Actividad] = null then "" else [Actividad])),
                cod = Text.BeforeDelimiter(txt, "-", 0)
            in 
                if txt = "" then null else FnFormatCodigoAct(cod), 
        type text
    ),

    // 3. ETIQUETA DE TIPO
    AgregadoTipo = Table.AddColumn(AgregadoCodAct, "Tipo", each "CC", type text),

    // 4. LIMPIEZA Y TIPOS DE DATOS ROBUSTOS
    TextosLimpios = Table.TransformColumns(AgregadoTipo, {
        {"Centro de Costos", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Ins", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Actividad", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Capitulo", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Subcapitulo", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"# OC / Contrato", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Nombre Contratista", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"# CC - Comparativo", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"# CC", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Comparativo", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        {"Clasificador", each if _ = null then null else Text.Trim(Text.From(_)), type text},
        
        {"Cant. aprobacion", each FxToNumberFlex(_), type number},
        {"V/U aprobacion", each FxToNumberFlex(_), type number},
        {"VR total aprobacion", each FxToNumberFlex(_), type number},
        {"Cantidad ppto (CC)", each FxToNumberFlex(_), type number},
        {"V/U ppto (CC)", each FxToNumberFlex(_), type number},
        {"Valor Total ppto (CC)", each FxToNumberFlex(_), type number}
    }, null, MissingField.Ignore),

    TiposFinales = try Table.TransformColumnTypes(TextosLimpios, {{"Codigo ins", Int64.Type}}) otherwise TextosLimpios,

    // 5. SALIDA (Sin doble buffer)
    TablaFinal = TiposFinales
in
    TablaFinal
