// PROYECCION_INSUMOS (hoja PROYECCION INSUMOS del flujo SINCO)
// Pares (actividad, insumo) de SEGUIMIENTO POR ITEMS cuya proyeccion difiere
// del presupuesto (cantidad o valor unitario, margen 0.001).
// Columnas fijas (subflujo PROYECCION INSUMOS):
//   A ACTIVIDAD (codigo)  B INSUMO (descripcion)
//   C CANTIDAD PROYECTADA (cantidad proyectada TOTAL de la constructora, no la diferencia)
//   D VALOR UNITARIO (proyectado)  E ESTADO (lo escribe el flujo)
// La clave del estado incluye cantidad y valor: si la constructora cambia la
// proyeccion, la fila vuelve a quedar pendiente (el flujo fija el valor, no suma).
// Contratado mayor que lo proyectado (tabla CONTRATOS_SINCO del libro, cantidad =
// Cantidad grupo x Cantidad detalle): la cantidad sube a la contratada (redondeada
// hacia arriba a 2 decimales) y el valor unitario baja para conservar el valor total
// proyectado; SINCO no deja contratar mas cantidad que la proyectada. Si el valor
// contratado supera el proyectado (cualquier monto), el total sube al contratado.
let
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    Headers = [Accept="application/json;odata=nometadata"],
    FnEncode = F_Globales[FnEncode],
    ArchivosProyecto = Table.Buffer(SP_Archivos_Flujo),

    Limpiar = (v as any) as text =>
        let t = if v = null then "" else Text.From(v)
        in Text.Trim(Text.Replace(Text.Replace(Text.Replace(t, "#(lf)", " "), "#(cr)", " "), "#(00A0)", " ")),

    FxNum = (v as any) as number =>
        let n = try Number.FromText(Text.Replace(Limpiar(v), ",", ""), "en-US") otherwise null
        in if n = null then 0 else n,

    EsCodigoActividad = (c0 as text) as logical =>
        let
            partes = Text.Split(c0, "."),
            dosPartes = List.Count(partes) = 2,
            p1Num = if dosPartes then (try Number.FromText(partes{0}) otherwise null) else null,
            p2Num = if dosPartes then (try Number.FromText(partes{1}) otherwise null) else null
        in
            dosPartes and p1Num <> null and p2Num <> null and Text.Length(partes{0}) > 0 and Text.Length(partes{1}) > 0,

    EsCodigoInsumo = (c0 as text) as logical =>
        c0 <> "" and Text.Select(c0, {"0".."9"}) = c0,

    // Clave de insumo sin tildes, espacios ni signos (el reporte y CONTRATOS_SINCO difieren en tildes)
    NormTxt = (v as any) as text =>
        let
            t = Text.Upper(Limpiar(v)),
            r = List.Accumulate({{"Á","A"},{"É","E"},{"Í","I"},{"Ó","O"},{"Ú","U"},{"Ü","U"},{"Ñ","N"}}, t, (s, p) => Text.Replace(s, p{0}, p{1}))
        in Text.Select(r, {"A".."Z", "0".."9"}),
    NumX = (v as any) as number =>
        if v is number then v else let n = try Number.From(v) otherwise null in if n = null then 0 else n,

    // Cantidad y valor contratados por (actividad, insumo), de la tabla CONTRATOS_SINCO ya cargada en el libro
    Contratado =
        let
            T = try Excel.CurrentWorkbook(){[Name = "CONTRATOS_SINCO"]}[Content] otherwise null,
            Ok = T <> null and List.ContainsAll(Table.ColumnNames(T), {"Grupo (actividad)", "Insumo", "Cantidad grupo", "Cantidad detalle (vacia = 1)", "Valor"}),
            Filas = if not Ok then {} else Table.TransformRows(T, (r) =>
                let
                    e = NumX(r[Cantidad grupo]),
                    g0 = NumX(r[#"Cantidad detalle (vacia = 1)"]),
                    g = if g0 = 0 then 1 else g0
                in [K = Limpiar(r[#"Grupo (actividad)"]) & "|" & NormTxt(r[Insumo]), CantC = e * g, ValC = e * NumX(r[Valor])]),
            Tab = Table.FromRecords(Filas, type table [K = text, CantC = number, ValC = number])
        in
            Table.Buffer(Table.Group(Tab, {"K"}, {{"CantC", each List.Sum([CantC]), type number}, {"ValC", each List.Sum([ValC]), type number}})),

    Filas = Table.SelectRows(ArchivosProyecto, each Text.Contains([Name], "SEGUIMIENTO POR ITEMS", Comparer.OrdinalIgnoreCase)),
    Ordenadas = Table.Sort(Filas, {{"Centro de Costos", Order.Ascending}, {"TimeLastModified", Order.Descending}}),
    Vacio = #table({"ACTIVIDAD","INSUMO","CANTIDAD PROYECTADA","VALOR UNITARIO"}, {}),

    Nuevo =
        if Table.RowCount(Ordenadas) = 0 then Vacio
        else
            let
                Binario = Binary.Buffer(Web.Contents(SiteUrl, [
                    RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(Ordenadas{0}[ServerRelativeUrl]) & "')/$value",
                    Headers = Headers,
                    Timeout = #duration(0, 0, 5, 0)
                ])),
                // Columnas 1-9: Cod, Desc, Tipo, UM, Presupuestado (Cant=5, VrU=6, Tot=7), Proyectado (Cant=8, VrU=9)
                Cols = List.Transform({1..9}, each {"Columna" & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}),
                Tabla = Html.Table(Text.FromBinary(Binario, 65001), Cols, [RowSelector="tr"]),

                Clasificado = Table.AddColumn(Tabla, "Clase", (r as record) =>
                    let
                        c0 = Limpiar(Record.Field(r, "Columna1")),
                        c2 = Limpiar(Record.Field(r, "Columna3")),
                        c3 = Limpiar(Record.Field(r, "Columna4"))
                    in
                        if c0 = "" or c0 = "Cod" or Text.StartsWith(c0, "SubCapitulo") then "Otro"
                        else if c2 <> "" and EsCodigoInsumo(c0) then "Insumo"  // sin UM tambien es insumo
                        else if EsCodigoActividad(c0) and Text.EndsWith(c0, ".000") then "Capitulo"
                        else if EsCodigoActividad(c0) then "Actividad"
                        else "Otro",
                    type text),
                ConActCod = Table.AddColumn(Clasificado, "ActCod", each if [Clase] = "Actividad" then Limpiar([Columna1]) else if [Clase] = "Capitulo" then "" else null, type text),
                FD = Table.FillDown(ConActCod, {"ActCod"}),

                Insumos = Table.SelectRows(FD, each [Clase] = "Insumo" and [ActCod] <> null and [ActCod] <> ""),
                Datos = Table.FromRecords(Table.TransformRows(Insumos, (r) => [
                    ACTIVIDAD = r[ActCod],
                    INSUMO = Text.Trim(Text.Replace(Limpiar(r[Columna2]), "160;", "")),
                    PresCant = FxNum(r[Columna5]),
                    PresVrU = FxNum(r[Columna6]),
                    ProyCant = FxNum(r[Columna8]),
                    ProyVrU = FxNum(r[Columna9])
                ])),
                ConK = Table.AddColumn(Datos, "K", each [ACTIVIDAD] & "|" & NormTxt([INSUMO]), type text),
                ConC = Table.ExpandTableColumn(Table.NestedJoin(ConK, {"K"}, Contratado, {"K"}, "__C", JoinKind.LeftOuter), "__C", {"CantC", "ValC"}, {"CantC", "ValC"}),
                ConNueva = Table.AddColumn(ConC, "NuevaCant", each
                    let c = if [CantC] = null then 0 else Number.RoundUp(Number.Round([CantC] * 100, 6)) / 100
                    in if c > [ProyCant] + 0.000001 then c else [ProyCant], type number),
                // Valor contratado mayor que el proyectado (cualquier monto): el total sube al contratado
                // (centavos hacia arriba). Lo contratado nunca puede quedar por encima de lo proyectado.
                ConTotal = Table.AddColumn(ConNueva, "NuevoTotal", each
                    let t = [ProyCant] * [ProyVrU], v = if [ValC] = null then 0 else Number.RoundUp(Number.Round([ValC] * 100, 6)) / 100
                    in if v > t + 0.005 then v else t, type number),
                Difieren = Table.SelectRows(ConTotal, each Number.Abs([ProyCant] - [PresCant]) > 0.001 or Number.Abs([ProyVrU] - [PresVrU]) > 0.001
                    or [NuevaCant] <> [ProyCant] or [NuevoTotal] <> [ProyCant] * [ProyVrU]),
                Salida = Table.FromRecords(Table.TransformRows(Difieren, (r) => [
                    ACTIVIDAD = r[ACTIVIDAD],
                    INSUMO = r[INSUMO],
                    #"CANTIDAD PROYECTADA" = Number.Round(r[NuevaCant], 6),
                    #"VALOR UNITARIO" =
                        if r[NuevaCant] = 0 then r[ProyVrU]
                        else if r[NuevoTotal] <> r[ProyCant] * r[ProyVrU] then Number.RoundUp(Number.Round(r[NuevoTotal] / r[NuevaCant] * 1000000, 3)) / 1000000
                        else if r[NuevaCant] = r[ProyCant] then Number.Round(r[ProyVrU], 6)
                        else Number.Round(r[ProyCant] * r[ProyVrU] / r[NuevaCant], 6)
                ])),
                Dedup = Table.Distinct(Salida, {"ACTIVIDAD","INSUMO"})
            in
                if Table.RowCount(Difieren) = 0 then Vacio
                else Table.Sort(Dedup, {{"ACTIVIDAD", Order.Ascending}, {"INSUMO", Order.Ascending}}),

    Final = FxConservarEstado(Nuevo, "PROYECCION_INSUMOS", {"ACTIVIDAD", "INSUMO", "CANTIDAD PROYECTADA", "VALOR UNITARIO"}, {"ESTADO"},
        {"ACTIVIDAD","INSUMO","CANTIDAD PROYECTADA","VALOR UNITARIO","ESTADO"}, false)
in
    Final
