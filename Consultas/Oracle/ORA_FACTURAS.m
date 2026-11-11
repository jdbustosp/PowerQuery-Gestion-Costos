// ORA_FACTURAS: reemplazo de SINCO_FACTURAS para libros Oracle, desde FACTURACION.xls (una fila por
// recepcion facturada de una orden OC- o contrato CO-). Mismo esquema que SINCO_FACTURAS:
//   OC        = Orden con prefijo (prefijos OC- y CO-): igual a la llave Oracle de la BD (FnOCOra)
//   Documento = Recepcion (ACP-..); Factura = Numero (numero del proveedor, el que traen las provisiones)
//   Valor     = Valor Recepción (Valor Factura a veces viene danado); Fecha = Fecha Factura_3 (ISO)
//   FacturaClave = letras y digitos en mayuscula sin ceros iniciales (igual a FnFacturaDin de la BD)
// Se excluyen facturas canceladas o rechazadas (segundo Estatus, Estado ERP CANCELLED) y las
// recepciones repetidas (misma Recepcion + Numero).
let
    Fuentes = ORA_FUENTES,
    FnNum = Fuentes[FnNum],
    T = Fuentes[FACTURACION],
    FnCol = (nombre as text) as list =>
        if Table.HasColumns(T, nombre) then Table.Column(T, nombre)
        else error Error.Record("Columna no encontrada", "FACTURACION.xls: falta la columna '" & nombre & "'",
            "Columnas: " & Text.Combine(Table.ColumnNames(T), " | ")),
    FnTxt = (v as any) as text => let t = try Text.Trim(Text.From(v)) otherwise "" in if t = null then "" else t,
    FnOC = (v as any) as text => Text.Select(Text.Upper(FnTxt(v)), {"A".."Z", "0".."9", "-"}),
    FnClave = (v as any) as text => Text.TrimStart(Text.Select(Text.Upper(FnTxt(v)), {"A".."Z", "0".."9"}), "0"),
    FnFecha = (v as any) as nullable date =>
        if v = null then null
        else if Value.Is(v, type date) then v
        else if Value.Is(v, type datetime) then Date.From(v)
        else if Value.Is(v, type number) then try Date.From(v) otherwise null
        else let t = FnTxt(v) in if t = "" then null
            else try Date.From(DateTime.FromText(Text.Start(t, 19), [Format = "yyyy-MM-ddTHH:mm:ss"]))
            otherwise try Date.FromText(t, [Culture = "es-CO"]) otherwise null,

    Base = Table.FromColumns({
        List.Transform(FnCol("Recepcion"), FnTxt),
        List.Transform(FnCol("Orden"), FnOC),
        List.Transform(FnCol("Nit"), FnTxt),
        List.Transform(FnCol("Numero"), FnTxt),
        // "Fecha Factura" (columna 11) viene con una fecha fija del informe; la real es la segunda (ISO).
        // Recepcion sin factura todavia (Numero vacio): Fecha Factura_3 vacia -> null (no la fecha fija)
        List.Transform(if Table.HasColumns(T, "Fecha Factura_3") then FnCol("Fecha Factura_3") else FnCol("Fecha Factura"), FnFecha),
        List.Transform(FnCol("Valor Recepción"), FnNum),
        List.Transform(FnCol("Estatus_1"), (v) => Text.Upper(FnTxt(v))),
        List.Transform(FnCol("Estado ERP"), (v) => Text.Upper(FnTxt(v)))
    }, {"Documento", "OC", "NIT", "Factura", "Fecha", "Valor", "__Est", "__ERP"}),
    Validas = Table.SelectRows(Base, each [Documento] <> "" and [Documento] <> "RECEPCION" and [OC] <> ""
        and not List.Contains({"CANCELADO", "RECHAZADO"}, [__Est]) and [__ERP] <> "CANCELLED"),
    SinRepetir = Table.Distinct(Validas, {"Documento", "Factura"}),
    ConClave = Table.AddColumn(SinRepetir, "FacturaClave", each FnClave([Factura]), type text),
    ConFuente = Table.AddColumn(ConClave, "Fuente", each "ORACLE", type text),
    Resultado = Table.SelectColumns(ConFuente, {"Fuente", "OC", "Documento", "NIT", "Factura", "Fecha", "Valor", "FacturaClave"})
in
    Table.Buffer(Value.ReplaceType(Resultado,
        type table [Fuente = text, OC = text, Documento = text, NIT = text, Factura = text, Fecha = date, Valor = number, FacturaClave = text]))
