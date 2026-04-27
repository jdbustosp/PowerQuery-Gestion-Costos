let
    // ============================================================
    // FUNCIONES AUXILIARES GLOBALES (Para hablar el mismo idioma)
    // ============================================================
    FnFormatCodigoAct = (raw as any) as nullable text =>
        let
            txtRaw  = if raw = null then null else Text.Trim(Text.From(raw)),
            result =
                if txtRaw = null or txtRaw = "" then null
                else
                    let
                        txtNorm = Text.Replace(Text.Replace(txtRaw, ",", "."), " ", ""),
                        hasDot  = Text.Contains(txtNorm, ".")
                    in
                        if hasDot then txtNorm
                        else
                            let
                                digits = Text.Select(txtNorm, {"0".."9"}),
                                len    = Text.Length(digits)
                            in
                                if len <= 3 then null
                                else Text.Range(digits, 0, len - 3) & "." & Text.Range(digits, len - 3, 3)
        in result,

    FxToNumberFlex = (value as any) as nullable number =>
        let
            v = value,
            numeroDirecto = if Value.Is(v, type number) then Number.From(v) else null
        in
            if numeroDirecto <> null then numeroDirecto else
            let t0 = if v=null then "" else Text.From(v), t = Text.Trim(Text.Replace(Text.Replace(t0, "#(00A0)", ""), " ", "")) in
            if t="" then null else
            let tryUS = try Number.FromText(t, "en-US"), valUS = if tryUS[HasError] then null else tryUS[Value] in
            if valUS<>null then valUS else
            let tryES = try Number.FromText(t, "es-ES"), valES = if tryES[HasError] then null else tryES[Value] in valES,

    // ============================================================
    // PROCESAMIENTO DE LA TABLA MANUAL
    // ============================================================
    // 🔥 ACELERADOR: Cargamos la tabla a la RAM desde el inicio
    Origen = Table.Buffer(Excel.CurrentWorkbook(){[Name="Det_CC"]}[Content]),

    // 1. FILTRO DE FILAS VACÍAS
    FilasValidas = Table.SelectRows(Origen, each [Ins] <> null or [Actividad] <> null),

    // 2. EXTRAER Y ESTANDARIZAR CÓDIGO DE ACTIVIDAD
    AgregadoCodAct = Table.AddColumn(
        FilasValidas, 
        "Codigo act", 
        each 
            let
                txt = Text.Trim(Text.From([Actividad] ?? "")),
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

    // 5. BUFFER QUIRÚRGICO
    TablaEnMemoria = Table.Buffer(TiposFinales)
in
    TablaEnMemoria
