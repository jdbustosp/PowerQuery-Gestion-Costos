// ORA_CONTRATOS: contratos Oracle (CONTRATOS.xls, CO-) con el MISMO esquema que CONTRATOS de SINCO.
//   VT Contratado = Valor Orden (lo contratado); VT Cortes = Valor Recepcion (actas, invertido).
let
    Div = (a as nullable number, b as nullable number) as nullable number => if a = null or b = null or b = 0 then null else a / b,
    Ct = Table.SelectRows(ORA_MOV, each [Fuente] = "CONTRATOS"),
    Filas = Table.FromRecords(List.Transform(Table.ToRecords(Ct), (r) => [
        #"Centro de Costos" = r[Centro de Costos], #"Codigo ins" = r[Codigo ins], Ins = r[Ins],
        #"Codigo act" = r[Codigo act], Actividad = r[Actividad], Capitulo = r[Capitulo], Subcapitulo = r[Subcapitulo],
        #"# OC / Contrato" = r[Orden], #"Nombre Contratista" = r[Razon Social], #"Descripcion contrato" = r[Titulo],
        #"Cantidad Contratado" = r[Cantidad], #"VT Contratado" = r[Valor Orden],
        #"Cantidad Cortes" = r[Cantidad Recepcion], #"VT Cortes" = r[Valor Recepcion], Tipo = "Contrato"]),
        type table [#"Centro de Costos" = text, #"Codigo ins" = text, Ins = text, #"Codigo act" = text, Actividad = text,
            Capitulo = text, Subcapitulo = text, #"# OC / Contrato" = text, #"Nombre Contratista" = text,
            #"Descripcion contrato" = text, #"Cantidad Contratado" = number, #"VT Contratado" = number,
            #"Cantidad Cortes" = number, #"VT Cortes" = number, Tipo = text])
in
    Table.Buffer(Filas)
