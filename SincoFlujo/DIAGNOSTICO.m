let
    // Query de diagnostico: pegala como una consulta nueva en Power Query y cargala como tabla.
    // Te muestra cuantas filas trae cada query del modelo, si tiene errores y cuanto tarda.
    // Util para aislar el query roto o lento sin tener que abrir cada uno por separado.

    Medir = (nombre as text, fn as function) =>
        let
            t0 = DateTime.LocalNow(),
            // El "if t0 = null" fuerza a evaluar t0 ANTES de la consulta; Table.Buffer
            // materializa el resultado una sola vez (RowCount + errores no re-evaluan).
            res = if t0 = null then null else (try Table.Buffer(fn()) otherwise null),
            filas = if res = null then -1 else try Table.RowCount(res) otherwise -1,
            errores = if res = null then -1 else try Table.RowCount(Table.SelectRowsWithErrors(res)) otherwise -1,
            // La dependencia artificial en filas/errores fuerza a que t1 se evalue DESPUES
            // del trabajo; sin esto la evaluacion perezosa marca 0 segundos.
            t1 = DateTime.LocalNow() + #duration(0, 0, 0, (filas - filas) + (errores - errores)),
            segundos = Duration.TotalSeconds(t1 - t0)
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
        Medir("BD",                     () => BD)
    },

    Resultado = Table.FromRecords(Filas)
in
    Resultado
