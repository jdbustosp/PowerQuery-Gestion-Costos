let
    // Query de diagnostico: pegala como una consulta nueva en Power Query y cargala como tabla.
    // Te muestra cuantas filas trae cada query del modelo y si tiene errores.
    // Util para aislar el query roto sin tener que abrir cada uno por separado.

    Medir = (nombre as text, fn as function) =>
        let
            t0 = DateTime.LocalNow(),
            res = try fn() otherwise null,
            t1 = DateTime.LocalNow(),
            segundos = Duration.TotalSeconds(t1 - t0),
            filas = if res = null then -1 else try Table.RowCount(res) otherwise -1,
            errores = if res = null then -1 else try
                let conErr = Table.SelectRowsWithErrors(res) in Table.RowCount(conErr)
                otherwise -1
        in
            [Query = nombre, Filas = filas, Errores = errores, Segundos = Number.Round(segundos, 1)],

    Filas = {
        Medir("SP_CarpetasCC",          () => SP_CarpetasCC),
        Medir("SP_Archivos_Proyecto",   () => SP_Archivos_Proyecto),
        Medir("SP_Seguimiento_Parsed",  () => SP_Seguimiento_Parsed),
        Medir("ITEMSINSUMOS",           () => ITEMSINSUMOS),
        Medir("PPTO_BD",                () => PPTO_BD),
        Medir("DESCARGAS",              () => DESCARGAS),
        Medir("CONTRATOS",              () => CONTRATOS),
        Medir("COMPRAS",                () => COMPRAS),
        Medir("DESCUENTOS",             () => DESCUENTOS),
        Medir("APROBACIONES_SP",        () => APROBACIONES_SP),
        Medir("PROVISIONES_SP",         () => PROVISIONES_SP),
        Medir("COMPARATIVOS",           () => COMPARATIVOS),
        Medir("DISPONIBLE",             () => DISPONIBLE),
        Medir("BD",                     () => BD),
        Medir("SINCO",                  () => SINCO)
    },

    Resultado = Table.FromRecords(Filas)
in
    Resultado
