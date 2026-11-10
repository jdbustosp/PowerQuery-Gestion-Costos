let
    Version = "PAMPLONA-SEGUIMIENTO-20261110-R2",
    CodigoParser = Binary.Buffer(Web.Contents("https://raw.githubusercontent.com", [
        RelativePath="jdbustosp/PowerQuery-Gestion-Costos/refs/heads/main/Consultas/F_Globales.m",
        Headers=[Accept="text/plain"], Timeout=#duration(0,0,1,0)])),
    FirmaParser = Binary.ToText(Binary.Compress(CodigoParser, Compression.Deflate), BinaryEncoding.Base64),
    ColumnasCache = {"Centro de Costos", "Firma", "Parte", "Partes", "Datos"},
    Prev0 = try Excel.CurrentWorkbook(){[Name="CACHE_SEGUIMIENTO"]}[Content] otherwise #table(ColumnasCache, {}),
    Prev = Table.Buffer(Table.SelectColumns(Prev0, ColumnasCache, MissingField.UseNull)),
    Forzar = try Logical.From(Excel.CurrentWorkbook(){[Name="FORZAR_RELECTURA_REPORTES"]}[Content]{0}[Column1]) otherwise false,
    Archivos = Table.Buffer(Table.SelectRows(SP_Archivos_Proyecto, each
        Text.Contains([Name], "SEGUIMIENTO POR ITEMS", Comparer.OrdinalIgnoreCase) or
        Text.Contains([Name], "ANALISIS DE PRECIOS UNITARIOS", Comparer.OrdinalIgnoreCase))),
    LeerCentro = (SP_Archivos_Proyecto as table) as table => (
        let

            FxProcesarCentroCosto = F_Globales[FxProcesarCentroCosto],
            FnReadSPBinary = F_Globales[FnReadSPBinary],
            SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",

            ConCentroCosto = SP_Archivos_Proyecto,

            PickLatestBinary = (t as table, containsText as text) as nullable binary =>
                let
                    candidatos = Table.Sort(
                        Table.SelectRows(t, each Text.Contains([Name], containsText, Comparer.OrdinalIgnoreCase)),
                        {{"TimeLastModified", Order.Descending}, {"Name", Order.Ascending}}
                    ),
                    path = if Table.RowCount(candidatos) = 0 then null else candidatos{0}[ServerRelativeUrl]
                in
                    if path = null then null else FnReadSPBinary(SiteUrl, path),

            Agrupado = Table.Group(ConCentroCosto, {"Centro de Costos"}, {{"Binarios", each
                let
                    binPres = PickLatestBinary(_, "ANALISIS DE PRECIOS UNITARIOS"),
                    binSeg = PickLatestBinary(_, "SEGUIMIENTO POR ITEMS")
                in
                    if binPres <> null and binSeg <> null then [Bin_P = binPres, Bin_S = binSeg] else null
            }}),
            CentrosCompletos = Table.SelectRows(Agrupado, each [Binarios] <> null),
            TablaConDatos = Table.AddColumn(CentrosCompletos, "Datos", each FxProcesarCentroCosto([Binarios][Bin_S], [Binarios][Bin_P])),
            Expandido = Table.ExpandTableColumn(TablaConDatos, "Datos", {"Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo", "Cantidad Presupuesto", "VT Presupuesto", "Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido"}),
            ColumnasUtiles = Table.SelectColumns(Expandido, {"Centro de Costos", "Codigo ins", "Ins", "Codigo act", "Actividad", "Capitulo", "Subcapitulo", "Cantidad Presupuesto", "VT Presupuesto", "Cantidad Proyectado", "VT Proyectado", "Cantidad Consumido", "VT Consumido"}),
            TiposFinales = Table.TransformColumnTypes(ColumnasUtiles,{{"Centro de Costos", type text}, {"Codigo ins", Int64.Type}, {"Cantidad Presupuesto", type number}, {"VT Presupuesto", Currency.Type}, {"Cantidad Proyectado", type number}, {"VT Proyectado", Currency.Type}, {"Cantidad Consumido", type number}, {"VT Consumido", Currency.Type}}),

            Resultado = Table.Buffer(TiposFinales)
        in
            Resultado
    ),
    Firma = (t as table) as text => Version & "|" & FirmaParser & "|" & Text.Combine(List.Transform(Table.ToRecords(
        Table.Sort(t, {{"Name", Order.Ascending}, {"TimeLastModified", Order.Descending}})), each
        Text.From([Name]) & "|" & Text.From([ServerRelativeUrl]) & "|" &
        DateTimeZone.ToText(DateTimeZone.From([TimeLastModified]), "o", "en-US") & "|" & Text.From([Length])), "#(lf)"),
    Empaquetar = (cc as text, t as table) as table =>
        let firma = Firma(t),
            previo = Table.Sort(Table.SelectRows(Prev, each [Centro de Costos] = cc and [Firma] = firma), {{"Parte", Order.Ascending}}),
            partes = if Table.IsEmpty(previo) then 0 else try Number.From(previo{0}[Partes]) otherwise 0,
            secuencia = try List.Transform(previo[Parte], Number.From) otherwise {},
            completo = partes > 0 and Table.RowCount(previo) = partes and secuencia = List.Numbers(1, partes),
            valido = if not completo or Forzar then false else
                try (let j=Json.Document(Text.Combine(previo[Datos], "")) in Record.HasFields(j, {"Columnas", "Filas"})) otherwise false,
            datos = if valido then null else Table.Buffer(LeerCentro(t)),
            texto = if valido then "" else Text.FromBinary(Json.FromValue([Columnas=Table.ColumnNames(datos), Filas=Table.ToRows(datos)]), TextEncoding.Utf8),
            n = if valido then 0 else Number.RoundUp(Text.Length(texto)/24000),
            filas = if valido then previo else Table.FromRecords(List.Transform(List.Numbers(0,n), each
                [#"Centro de Costos"=cc, Firma=firma, Parte=_+1, Partes=n, Datos=Text.Middle(texto,_*24000,24000)]), ColumnasCache)
        in filas,
    Grupos = Table.Group(Archivos, {"Centro de Costos"}, {{"DatosCC", each _, type table}}),
    Faltantes = List.Difference(SP_CarpetasCC[Centro de Costos], Grupos[Centro de Costos]),
    Incompletos = Table.SelectRows(Grupos, each
        not List.AnyTrue(List.Transform([DatosCC][Name], each Text.Contains(_, "SEGUIMIENTO POR ITEMS", Comparer.OrdinalIgnoreCase))) or
        not List.AnyTrue(List.Transform([DatosCC][Name], each Text.Contains(_, "ANALISIS DE PRECIOS UNITARIOS", Comparer.OrdinalIgnoreCase)))),
    Revisados = if not List.IsEmpty(Faltantes) or not Table.IsEmpty(Incompletos) then
        error "No se pudo verificar Seguimiento/APU de todos los centros. Se conserva la carga anterior; revise disponibilidad de los reportes."
        else Grupos,
    ConCache = Table.AddColumn(Revisados, "Cache", each Empaquetar([Centro de Costos], [DatosCC])),
    Final = Table.TransformColumnTypes(Table.Combine(ConCache[Cache]), {
        {"Centro de Costos",type text},{"Firma",type text},{"Parte",Int64.Type},{"Partes",Int64.Type},{"Datos",type text}})
in
    Final
