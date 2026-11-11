// ORA_PPTO_V2_BD: presupuesto V2 (vigente) por insumo, con el MISMO esquema que PPTO_BD, desde la hoja
// "Actualizacion_ppto" del archivo de descargas del proyecto (SharePoint, ORA_DESCARGAS_LIBRO).
// Esa hoja la mantiene el usuario (V2 = V1 + cambios aprobados); sobre ella se hacen las descargas.
// Columnas de la hoja: Proyecto, Centro de Costos, Subcapitulo, Capitulo, Actividad, Codigo ins, Ins,
// Cantidad ppto (CC), V/U ppto (CC), Valor Total ppto (CC). Codigo act = "CBS | PAQUETE" (como ORA_PPTO_BD).
let
    Arbol = ORA_ARBOL,
    FnNum = ORA_FUENTES[FnNum],
    Hoja = "Actualizacion_ppto",
    Nav = Table.SelectRows(ORA_DESCARGAS_LIBRO, each [Item] = Hoja and [Kind] = "Sheet"),
    Crudo = if Table.RowCount(Nav) = 0
        then error Error.Record("Falta presupuesto V2", "El archivo de descargas no tiene la hoja " & Hoja)
        else Table.PromoteHeaders(Nav{0}[Data], [PromoteAllScalars = true]),
    Req = {"Subcapitulo", "Capitulo", "Actividad", "Codigo ins", "Ins", "Cantidad ppto (CC)", "V/U ppto (CC)", "Valor Total ppto (CC)"},
    Faltan = List.Difference(Req, Table.ColumnNames(Crudo)),
    Validado = if List.IsEmpty(Faltan) then Crudo
        else error Error.Record("Columnas faltantes en " & Hoja, Text.Combine(Faltan, ", ")),
    Txt = (v as any) as nullable text => let t = if v = null then "" else Text.Trim(Text.From(v)) in if t = "" then null else t,
    Filas = Table.SelectRows(Validado, each Txt([Actividad]) <> null and Txt([Ins]) <> null),
    FnCbs = (a as nullable text) as nullable text => if a = null or not Text.Contains(a, " - ") then a else Text.BeforeDelimiter(a, " - "),
    Res = Table.FromRecords(List.Transform(Table.ToRecords(Filas), (r) =>
        let
            act = Txt(r[Actividad]), sub = Txt(r[Subcapitulo]),
            cant = FnNum(r[#"Cantidad ppto (CC)"]), vt = FnNum(r[#"Valor Total ppto (CC)"]), vu = FnNum(r[#"V/U ppto (CC)"])
        in [
            #"Centro de Costos" = ORA_FUENTES[CentroCostos],
            #"Codigo ins" = Txt(r[Codigo ins]),
            Ins = Txt(r[Ins]),
            #"Codigo act" = Arbol[FnCodAct](FnCbs(act), sub),
            Actividad = act,
            Capitulo = Txt(r[Capitulo]),
            Subcapitulo = sub,
            #"Cantidad Presupuesto" = cant,
            #"VT Presupuesto" = vt,
            #"V/U Presupuesto" = if vu <> null then vu else if cant <> null and cant <> 0 and vt <> null then vt / cant else null,
            Tipo = "PPTO V2"]),
        type table [#"Centro de Costos" = text, #"Codigo ins" = text, Ins = text, #"Codigo act" = text, Actividad = text,
            Capitulo = text, Subcapitulo = text, #"Cantidad Presupuesto" = number, #"VT Presupuesto" = number,
            #"V/U Presupuesto" = number, Tipo = text])
in
    Table.Buffer(Res)
