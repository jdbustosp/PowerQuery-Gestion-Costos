let
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    ParamProyecto = Text.Trim(ProyectoActual),
    FechaVersion = try Text.Trim(Text.From(FechaVersionSINCO)) otherwise "",
    FnEncode = F_Globales[FnEncode],
    FnDecodeHtml = F_Globales[FnDecodeHtml],
    FnReadSPBinary = F_Globales[FnReadSPBinary],
    FxToNumberFlex = F_Globales[FxToNumberFlex],
    Columnas = F_Globales[FnBuildColumnas](15),
    BasePath = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. Reportes EDT - Control costos interno/" & ParamProyecto,

    FnText = (v as any) as text => try Text.Trim(Text.From(if v = null then "" else v)) otherwise "",
    FnDigits = (v as any) as nullable text => let d = Text.Select(FnText(v), {"0".."9"}) in if d = "" then null else d,

    Centros = try List.Distinct(SP_CarpetasCC[Centro de Costos]) otherwise List.Distinct(Table.Column(COMPRAS, "Centro de Costos")),
    FnFilesPrev = (cc as text) as table =>
        let
            path = BasePath & "/" & cc & "/Versiones previas/" & FechaVersion,
            raw = try Json.Document(Web.Contents(SiteUrl, [
                RelativePath = "/_api/web/GetFolderByServerRelativeUrl('" & FnEncode(path) & "')/Files",
                Query = [#"$select" = "Name,ServerRelativeUrl,TimeLastModified,Length"],
                Headers = [Accept="application/json;odata=nometadata"],
                Timeout = #duration(0,0,5,0)
            ])) otherwise null,
            tbl = if raw <> null and Record.HasFields(raw, "value") then Table.FromRecords(raw[value]) else #table({"Name","ServerRelativeUrl","TimeLastModified","Length"}, {}),
            add = Table.AddColumn(tbl, "Centro de Costos", each cc, type text)
        in
            add,

    ArchivosPrevios = Table.Buffer(if FechaVersion = "" or List.Count(Centros)=0 then #table({"Name","ServerRelativeUrl","Centro de Costos"}, {}) else Table.Combine(List.Transform(Centros, each FnFilesPrev(_)))),
    FnPick = (cc as text, containsText as text) as nullable binary =>
        let
            rows = Table.Sort(Table.SelectRows(ArchivosPrevios, each [Centro de Costos] = cc and Text.Contains([Name], containsText, Comparer.OrdinalIgnoreCase)), {{"TimeLastModified", Order.Descending}, {"Name", Order.Ascending}}),
            path = if Table.RowCount(rows)=0 then null else rows{0}[ServerRelativeUrl]
        in if path = null then null else FnReadSPBinary(SiteUrl, path),

    FnTable = (bin as binary) as table =>
        try Excel.Workbook(Binary.Buffer(bin), null, true){0}[Data]
        otherwise Html.Table(FnDecodeHtml(bin), Columnas, [RowSelector="tr"]),
    FnRename = (tbl as table) as table => Table.RenameColumns(tbl, List.Zip({Table.ColumnNames(tbl), List.Transform({1..List.Count(Table.ColumnNames(tbl))}, each "Columna" & Text.From(_))})),
    Actual = Table.SelectRows(COMPRAS, each [#"#ENTRADA"] <> null and FnDigits([#"#ENTRADA"]) <> null),
    ActualKey = Table.AddColumn(Actual, "__Key", each [Centro de Costos] & "|" & FnDigits([#"#ENTRADA"]), type text),

    FnKeysPrevCC = (cc as text) as table =>
        let
            bin = let b1 = FnPick(cc, "INFORME ENTRADAS DE ALMACEN") in if b1 <> null then b1 else FnPick(cc, "INFORME ENTRADAS DE ALMACÉN"),
            tbl0 = if bin = null then #table({"Columna1","Columna2","Columna4"}, {}) else FnRename(FnTable(bin)),
            rows = Table.SelectRows(tbl0, each FnText([Columna1]) <> "" and FnDigits([Columna2]) <> null and FnDigits([Columna4]) <> null and (try Date.From([Columna1]) otherwise null) <> null),
            keys = Table.AddColumn(rows, "__Key", each cc & "|" & FnDigits([Columna2]), type text),
            out = Table.SelectColumns(keys, {"__Key"})
        in out,
    PrevKeys = Table.Distinct(if List.Count(Centros)=0 then #table({"__Key"}, {}) else Table.Combine(List.Transform(Centros, each FnKeysPrevCC(_)))),
    JoinPrev = Table.NestedJoin(ActualKey, {"__Key"}, PrevKeys, {"__Key"}, "Prev", JoinKind.LeftAnti),
    Resultado = Table.RemoveColumns(JoinPrev, {"__Key"}, MissingField.Ignore)
in
    Resultado
