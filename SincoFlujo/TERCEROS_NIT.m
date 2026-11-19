// TERCEROS_NIT: NIT del tercero por NOMBRE cuando los informes del centro (pagos / facturacion) no lo traen.
// Se llama desde CONTRATOS_SINCO, ORDENES_COMPRA y SALIDAS_SINCO con FxGitHub("TERCEROS_NIT")
// (no necesita consulta en el libro). Lee la carpeta PROVEEDORES de "0. Automatizacion sinco":
//   1. Terceros_NIT_MANUAL.xlsx (manda sobre todo): hoja NOMBRES (NOMBRE | NIT) para nombres con dos NIT
//      o que no estan en las fuentes; hoja SALIDAS (CENTRO | SALIDA ORIGEN | NIT | NOMBRE) para salidas
//      cuyo contratista en el origen es "-" o vacio.
//   2. El InformeDeProveedores_*.xlsx mas reciente (informe de proveedores de nuestro SINCO: Nit | Nombre).
//   3. Todas las Plantilla_Terceros_NUEVOS_*.xlsx (hoja MIGRACION: NIT en la col. 2, razon social en la 5
//      o nombres y apellidos en la 6..9).
// Nombre normalizado = mayusculas, solo A-Z y 0-9. Si un nombre tiene mas de un NIT en 2 y 3, no se
// devuelve (va a la hoja NOMBRES del manual). Si la carpeta o un archivo no se puede leer, devuelve vacio.
// Devuelve [PorNombre = (nombre) => NIT o null, SalidaSin = (salida) => [NIT, Nombre] o null].
let
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    Carpeta = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. Automatizacion sinco/PROVEEDORES",
    Headers = [Accept="application/json;odata=nometadata"],
    FnEncode = F_Globales[FnEncode],

    // Fuente: Name | TimeLastModified | Bin (funcion sin argumentos que baja el archivo)
    Lista = try Table.FromRecords(Json.Document(Web.Contents(SiteUrl, [
        RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(Carpeta) & "')/Files",
        Query = [#"$select" = "Name,ServerRelativeUrl,TimeLastModified"],
        Headers = Headers,
        Timeout = #duration(0, 0, 5, 0)
    ]))[value]) otherwise #table({"Name","ServerRelativeUrl","TimeLastModified"}, {}),
    Fuente = Table.AddColumn(Lista, "Bin", (r) => () => try Binary.Buffer(Web.Contents(SiteUrl, [
        RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(r[ServerRelativeUrl]) & "')/$value",
        Headers = Headers,
        Timeout = #duration(0, 0, 5, 0)
    ])) otherwise null),

    Txt = (v as any) as text =>
        if v = null then "" else if v is number then Number.ToText(Number.Round(v), "0")
        else Text.Trim(Text.Replace(Text.From(v), "#(00A0)", " ")),
    Norm = (v as any) as text => Text.Select(Text.Upper(Txt(v)), {"A".."Z","0".."9"}),
    NitDe = (v as any) as nullable text =>
        let t0 = Txt(v), t = if Text.EndsWith(t0, ".0") then Text.Start(t0, Text.Length(t0) - 2) else t0
        in if Text.Length(t) >= 6 and Text.Length(t) <= 11 and Text.Select(t, {"0".."9"}) = t then t else null,
    Hoja = (bin as nullable binary, nombre as text) as list =>
        if bin = null then {} else
        let
            libro = try Excel.Workbook(bin, false, true) otherwise null,
            h = if libro = null then null else Table.SelectRows(libro, each [Kind] = "Sheet" and Text.Upper([Item]) = nombre)
        in if h = null or Table.RowCount(h) = 0 then {} else List.Skip(Table.ToRows(h{0}[Data]), 1),
    Col = (fila as list, i as number) as any => if List.Count(fila) > i then fila{i} else null,
    Archivos = (prefijo as text) as table =>
        Table.Sort(Table.SelectRows(Fuente, each Text.StartsWith(Text.Upper([Name]), Text.Upper(prefijo))
            and Text.EndsWith(Text.Lower([Name]), ".xlsx")), {{"TimeLastModified", Order.Descending}}),

    // 2. Informe de proveedores de SINCO (el mas reciente)
    Informes = Archivos("InformeDeProveedores"),
    DeInforme = if Table.RowCount(Informes) = 0 then {} else
        List.Transform(Hoja(Informes{0}[Bin](), "INFORMEDEPROVEEDORES"), each {Norm(Col(_, 2)), NitDe(Col(_, 1))}),
    // 3. Plantillas de terceros nuevos
    DePlantillas = List.Combine(List.Transform(Table.ToRecords(Archivos("Plantilla_Terceros_NUEVOS")), (a) =>
        List.Transform(Hoja(a[Bin](), "MIGRACION"), (f) =>
            let rs = Txt(Col(f, 4)),
                nm = if rs <> "" then rs else Text.Combine(List.Select(List.Transform({5..8}, each Txt(Col(f, _))), each _ <> ""), " ")
            in {Norm(nm), NitDe(Col(f, 1))}))),
    Pares = Table.Distinct(Table.SelectRows(#table(type table [Clave = text, NIT = nullable text], DeInforme & DePlantillas),
        each [Clave] <> "" and [NIT] <> null)),
    Unicos = Table.SelectRows(Table.Group(Pares, {"Clave"}, {{"NITs", each [NIT], type list}}), each List.Count([NITs]) = 1),

    // 1. Manual
    Manual = Archivos("Terceros_NIT_MANUAL"),
    BinManual = if Table.RowCount(Manual) = 0 then null else Manual{0}[Bin](),
    ManNombres = Table.ToRows(Table.Distinct(#table({"C","N"}, List.Select(List.Transform(Hoja(BinManual, "NOMBRES"),
        each {Norm(Col(_, 0)), NitDe(Col(_, 1))}), each _{0} <> "" and _{1} <> null)), {"C"})),
    ManSalidas = Table.ToRows(Table.Distinct(#table({"S","F"}, List.Transform(List.Select(Hoja(BinManual, "SALIDAS"),
        each Txt(Col(_, 1)) <> "" and NitDe(Col(_, 2)) <> null), each {Txt(Col(_, 1)), _})), {"S"})),

    MapaAuto = Record.FromList(List.Transform(Unicos[NITs], each _{0}), List.Transform(Unicos[Clave], each "k" & _)),
    MapaManual = Record.FromList(List.Transform(ManNombres, each _{1}), List.Transform(ManNombres, each "k" & _{0})),
    MapaSalidas = Record.FromList(List.Transform(ManSalidas, each [NIT = NitDe(Col(_{1}, 2)), Nombre = Txt(Col(_{1}, 3))]),
        List.Transform(ManSalidas, each "k" & _{0})),
    Mapa = Record.Combine({MapaAuto, MapaManual}),

    Resultado = [
        PorNombre = (nombre as any) as nullable text =>
            let k = Norm(nombre) in if k = "" then null else Record.FieldOrDefault(Mapa, "k" & k, null),
        SalidaSin = (salida as any) as nullable record => Record.FieldOrDefault(MapaSalidas, "k" & Txt(salida), null),
        Cuenta = [Automaticos = Record.FieldCount(MapaAuto), Manuales = Record.FieldCount(MapaManual), Salidas = Record.FieldCount(MapaSalidas)]
    ]
in
    Resultado
