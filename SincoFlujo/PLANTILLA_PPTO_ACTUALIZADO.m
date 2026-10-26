// PLANTILLA_PPTO_ACTUALIZADO
// MISMA plantilla de importacion de presupuesto (21 columnas, identica a
// PLANTILLA_PPTO_SINCO / _ORACLE) armada desde el presupuesto ACTUALIZADO del
// proyecto: "<ProyectoActual>.xlsx" de "0. Descargas pptos - Control costos
// interno", hoja Actualizacion_ppto (respaldo PPTO_TODOS_PROYECTOS). Se usa
// cuando el presupuesto a subir es el actualizado y no el del ERP del
// constructor (VERSALLES, 2026-10-25). El libro lo activa con su consulta
// PLANTILLA_PPTO_SINCO = FxGitHub("PLANTILLA_PPTO_ACTUALIZADO").
//
// Columnas del archivo: Proyecto, Centro de Costos, Subcapitulo (sector:
// TORRE 1, URBANISMO...), Capitulo ("01-00-000 - NOMBRE"), Actividad
// ("01-01-003 - NOMBRE (UM)"), Codigo ins, Ins ("COD - NOMBRE (UM)"),
// Cantidad ppto (CC), V/U ppto (CC), Valor Total ppto (CC).
//
// Reglas (aprobadas por Juan 2026-10-25):
//   - Los sectores se suman: una actividad por codigo (como el EDT de SINCO).
//   - Arbol CD > capitulo (NN-00-000) > subcapitulo (NN-NN-000) > actividad;
//     nombres de CONTROL.xlsx (Actual del centro) si existe, si no el del
//     archivo o el del capitulo.
//   - CANTIDAD de la actividad = cantidad del insumo con la MISMA UM de la
//     actividad: el "Sub." (subcontrato) si hay; si no, la cantidad que mas se
//     repite (empate: la menor). Sin insumo de esa UM: 1.
//   - Insumo: VrUnitSinIVA = suma V.Total / suma Cantidad; Cant APU = suma
//     Cantidad / CANTIDAD de la actividad (10 decimales: con 6 los insumos
//     muy pequenos descuadran el total de la actividad).
//   - Insumo sin UM en el archivo ("()"): toma la UM de la actividad.
//   - Tipo Insumo y Agrupacion: maestro INSUMOS_ORACLE.csv, como _ORACLE.
let
    ParamProyecto = Text.Trim(ProyectoActual),
    SiteUrl = "https://colsubsidio365.sharepoint.com/sites/MiGerenciaViv",
    RutaArchivo = "/sites/MiGerenciaViv/Departamento Tecnico/COORDINACION DE PRESUPUESTOS/0. Descargas pptos - Control costos interno/" & ParamProyecto & ".xlsx",
    Headers = [Accept="application/json;odata=nometadata"],
    FnEncode = F_Globales[FnEncode],

    Limpiar = (v as any) as text =>
        let t = if v = null then "" else Text.From(v)
        in Text.Trim(Text.Replace(Text.Replace(Text.Replace(t, "#(lf)", " "), "#(cr)", " "), "#(00A0)", " ")),
    FnNum = (v as any) as number =>
        if v = null then 0
        else if Value.Is(v, type number) then v
        else try Number.FromText(Text.From(v), "es-CO") otherwise 0,

    // "TEXTO (UM)" -> [Txt = "TEXTO", UM = "UM"]
    FnSepararUM = (t as text) as record =>
        let
            p = Text.PositionOf(t, "(", Occurrence.Last),
            fin = Text.End(t, 1) = ")"
        in
            if fin and p > 0
            then [Txt = Text.Trim(Text.Start(t, p)), UM = Text.Upper(Text.Trim(Text.Middle(t, p + 1, Text.Length(t) - p - 2)))]
            else [Txt = t, UM = null],
    // "01-01-003 - NOMBRE" -> [Cod = "01-01-003", Nom = "NOMBRE"]
    FnSepararCodigo = (t as text) as record =>
        let p = Text.PositionOf(t, " - ")
        in if p > 0 then [Cod = Text.Trim(Text.Start(t, p)), Nom = Text.Trim(Text.Middle(t, p + 3))] else [Cod = t, Nom = t],

    // =========================================================
    // 1. ARCHIVO DEL PRESUPUESTO ACTUALIZADO
    // =========================================================
    BinArchivo = try Binary.Buffer(Web.Contents(SiteUrl, [
            RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(RutaArchivo) & "')/$value",
            Headers = Headers,
            Timeout = #duration(0, 0, 5, 0)
        ])) otherwise error Error.Record("PLANTILLA_PPTO_ACTUALIZADO",
            "No se encontró el presupuesto actualizado " & ParamProyecto & ".xlsx en 0. Descargas pptos - Control costos interno",
            [Ruta = RutaArchivo]),
    Wb = Excel.Workbook(BinArchivo, null, true),
    FnHoja = (nombre as text) as nullable table =>
        let f = Table.SelectRows(Wb, each [Kind] = "Sheet" and Text.Upper([Name]) = Text.Upper(nombre))
        in if Table.RowCount(f) = 0 then null else f{0}[Data],
    HojaBase = FnHoja("Actualizacion_ppto") ?? FnHoja("PPTO_TODOS_PROYECTOS"),
    Hoja = if HojaBase = null
        then error Error.Record("PLANTILLA_PPTO_ACTUALIZADO", "El archivo no tiene la hoja Actualizacion_ppto ni PPTO_TODOS_PROYECTOS", [Ruta = RutaArchivo])
        else Table.PromoteHeaders(HojaBase, [PromoteAllScalars = true]),

    Filas = Table.SelectRows(Hoja, each
        Text.Upper(Limpiar([Proyecto])) = Text.Upper(ParamProyecto) and Limpiar([Actividad]) <> "" and Limpiar([Codigo ins]) <> ""),
    FilasValidas = if Table.RowCount(Filas) = 0
        then error Error.Record("PLANTILLA_PPTO_ACTUALIZADO", "El archivo no tiene filas del proyecto " & ParamProyecto, [Ruta = RutaArchivo])
        else Filas,
    Norm = Table.FromRecords(Table.TransformRows(FilasValidas, (r) =>
        let
            cap = FnSepararCodigo(Limpiar(r[Capitulo])),
            act = FnSepararCodigo(Limpiar(r[Actividad])),
            actUM = FnSepararUM(act[Nom]),
            ins = FnSepararUM(Limpiar(r[Ins]))
        in [
            CodCap = cap[Cod], NomCap = cap[Nom],
            CodAct = act[Cod], NomAct = actUM[Txt], UMAct = actUM[UM],
            Articulo = Limpiar(r[Codigo ins]), DescIns = ins[Txt], UMIns = ins[UM],
            Cantidad = FnNum(r[#"Cantidad ppto (CC)"]),
            VrTotal = FnNum(r[#"Valor Total ppto (CC)"])
        ])),

    // insumos: suma de sectores por (actividad, articulo)
    Insumos = Table.Buffer(Table.Group(Norm, {"CodAct", "Articulo"}, {
        {"DescIns", each List.First([DescIns]), type text},
        {"UMIns", each List.First([UMIns]), type nullable text},
        {"UMAct", each List.First([UMAct]), type nullable text},
        {"Cantidad", each List.Sum([Cantidad]), type number},
        {"VrTotal", each List.Sum([VrTotal]), type number}
    })),

    // =========================================================
    // 2. ACTIVIDADES Y SU CANTIDAD
    // =========================================================
    FnCantidadActividad = (t as table) as number =>
        let
            misma = Table.SelectRows(t, each [UMIns] <> null and [UMIns] = [UMAct] and [Cantidad] <> 0),
            subs = Table.SelectRows(misma, each Text.StartsWith(Text.Upper(Text.Trim([DescIns])), Text.Upper([Articulo]) & " - SUB.")),
            cants = List.Transform(Table.Column(misma, "Cantidad"), each Number.Round(_, 6)),
            grupos = List.Transform(List.Distinct(cants), (c) => [C = c, N = List.Count(List.Select(cants, each _ = c))]),
            moda = List.First(List.Sort(grupos, {{each [N], Order.Descending}, {each [C], Order.Ascending}}))[C]
        in
            if Table.RowCount(subs) > 0 then Table.Max(subs, "VrTotal")[Cantidad]
            else if Table.RowCount(misma) > 0 then moda
            else 1,
    Actividades = Table.Buffer(Table.Group(Norm, {"CodAct"}, {
        {"NomAct", each List.First([NomAct]), type text},
        {"UMAct", each List.First([UMAct]), type nullable text},
        {"CodCap", each List.First([CodCap]), type text},
        {"VrTotalAct", each List.Sum([VrTotal]), type number}
    })),
    ConCantidad = Table.Buffer(Table.AddColumn(Actividades, "CantAct", each
        FnCantidadActividad(Table.SelectRows(Insumos, (i) => i[CodAct] = [CodAct])), type number)),
    DictCantAct = Record.FromList(Table.Column(ConCantidad, "CantAct"), Table.Column(ConCantidad, "CodAct")),

    // =========================================================
    // 3. ARBOL: capitulo = 1er segmento + ceros; subcapitulo = 2 primeros
    // segmentos + ceros. Nombres: CONTROL.xlsx, si no el archivo.
    // =========================================================
    FnEnCeros = (cod as text, desde as number) as text =>
        let segs = Text.Split(cod, "-")
        in Text.Combine(List.Transform({0..List.Count(segs) - 1}, (i) => if i >= desde then Text.Repeat("0", Text.Length(segs{i})) else segs{i}), "-"),
    DictControl =
        try
            let
                Arch = Table.SelectRows(Table.Buffer(SP_Archivos_Flujo), each Text.Contains([Name], "CONTROL", Comparer.OrdinalIgnoreCase)),
                Ord = Table.Sort(Arch, {{"TimeLastModified", Order.Descending}}),
                Bin = Web.Contents(SiteUrl, [
                    RelativePath = "/_api/web/GetFileByServerRelativeUrl('" & FnEncode(Ord{0}[ServerRelativeUrl]) & "')/$value",
                    Headers = Headers, Timeout = #duration(0, 0, 5, 0)]),
                Data = Excel.Workbook(Bin, [InferSheetDimensions = true, DelayTypes = true]){0}[Data],
                Cols = Table.ColumnNames(Data),
                Pares = List.Distinct(List.Select(
                    List.Zip({List.Transform(Table.Column(Data, Cols{0}), Limpiar), List.Transform(Table.Column(Data, Cols{1}), Limpiar)}),
                    each _{0} <> "" and _{1} <> ""), each _{0})
            in
                Record.FromList(List.Transform(Pares, each _{1}), List.Transform(Pares, each _{0}))
        otherwise [],
    DictCapArchivo = Record.FromList(Table.Column(Table.Distinct(Norm, {"CodCap"}), "NomCap"), Table.Column(Table.Distinct(Norm, {"CodCap"}), "CodCap")),

    Nodos = Table.Buffer(Table.Distinct(Table.FromRecords(List.Combine(Table.TransformRows(ConCantidad, (r) =>
        let cap = FnEnCeros(r[CodAct], 1), sub = FnEnCeros(r[CodAct], 2)
        in {
            [Cod = cap, Padre = "CD", Nom = Record.FieldOrDefault(DictControl, cap, Record.FieldOrDefault(DictCapArchivo, cap, cap))],
            [Cod = sub, Padre = cap, Nom = Record.FieldOrDefault(DictControl, sub, Record.FieldOrDefault(DictCapArchivo, cap, sub))]
        }))), {"Cod"})),

    // =========================================================
    // 4. TIPO INSUMO y AGRUPACION (igual que PLANTILLA_PPTO_ORACLE)
    // =========================================================
    UrlMaestroIns = "https://raw.githubusercontent.com/jdbustosp/SINCO-SINCO/main/Datos/INSUMOS_ORACLE.csv?cb=" & Number.ToText(Number.From(DateTime.LocalNow())),
    CsvMaestro = try Table.Skip(Csv.Document(Web.Contents(UrlMaestroIns), [Delimiter = ";", Columns = 4, Encoding = 65001, QuoteStyle = QuoteStyle.None]), 1)
                 otherwise #table({"Column1", "Column2", "Column3", "Column4"}, {}),
    MaestroDist = Table.Distinct(CsvMaestro, {"Column1"}),
    DictTipoIns = Record.FromList(
        List.Transform(Table.Column(MaestroDist, "Column2"), Text.Trim),
        List.Transform(Table.Column(MaestroDist, "Column1"), Text.Trim)),
    DictGrupoIns = Record.FromList(
        List.Transform(Table.Column(MaestroDist, "Column4"), Text.Trim),
        List.Transform(Table.Column(MaestroDist, "Column1"), Text.Trim)),
    TiposValidos = {"E", "I", "O", "M", "N", "Y", "P", "S", "T"},
    FnTipoInsumo = (articulo as text) as text =>
        let
            delMaestro = Record.FieldOrDefault(DictTipoIns, articulo, null),
            p = Text.Upper(Text.Start(articulo, 1)),
            heuristica =
                if p = "L" then "N"
                else if p = "M" then "M"
                else if p = "P" then "M"
                else if p = "S" then "O"
                else if p = "Z" then "S"
                else if p = "H" then "S"
                else "Y",
            t = if delMaestro <> null and delMaestro <> "" then Text.Upper(delMaestro) else heuristica
        in
            if List.Contains(TiposValidos, t) then t else "Y",
    FnGrupoInsumo = (articulo as text) as text =>
        let g = Record.FieldOrDefault(DictGrupoIns, articulo, null)
        in if g = null or g = "" then "OTROS" else g,

    // =========================================================
    // 5. FILAS (claves de orden: codigo nodo/actividad, tipo, articulo)
    // =========================================================
    FilaVacia = [
        Código = null, Descripción = null, Padre = null, UM = null, CANTIDAD = null, SUBCAPITULO = null,
        #"ID PROYECTO" = null, VERSION = null, #"ID APU" = null, #"Cant APU" = null,
        Rend = null, IVA = null, VrUnitSinIVA = null, #"Tipo Insumo" = null,
        Agrupacion = null, #"COD CLIENTE" = null, #"Precio Cliente" = null, Clase = null,
        Tipo = null, #"Vr Unitario" = null, #"Vr Total" = null],
    FilasNodo = Table.TransformRows(Nodos, (r) => FilaVacia & [
        Código = r[Cod], Descripción = r[Nom], Padre = r[Padre], Tipo = "Capítulo",
        __K1 = r[Cod], __K3 = 0, __K4 = ""]),
    FilasActividad = Table.TransformRows(ConCantidad, (r) => FilaVacia & [
        Código = r[CodAct], Descripción = r[NomAct], Padre = FnEnCeros(r[CodAct], 2),
        UM = r[UMAct] ?? "GL", CANTIDAD = Number.Round(r[CantAct], 6), Tipo = "Actividad",
        #"Vr Unitario" = Number.Round(r[VrTotalAct] / r[CantAct], 6), #"Vr Total" = r[VrTotalAct],
        __K1 = r[CodAct], __K3 = 1, __K4 = ""]),
    FilasInsumo = Table.TransformRows(Insumos, (r) =>
        let cantAct = Record.Field(DictCantAct, r[CodAct])
        in FilaVacia & [
            Descripción = r[DescIns], Padre = r[CodAct], UM = if r[UMIns] = null or r[UMIns] = "" then r[UMAct] ?? "UN" else r[UMIns],
            #"Cant APU" = Number.Round(r[Cantidad] / cantAct, 10),
            Rend = 1, IVA = 0,
            VrUnitSinIVA = if r[Cantidad] <> 0 then Number.Round(r[VrTotal] / r[Cantidad], 6) else r[VrTotal],
            #"Tipo Insumo" = FnTipoInsumo(r[Articulo]), Agrupacion = FnGrupoInsumo(r[Articulo]),
            Tipo = "Insumo", #"Vr Unitario" = Number.Round(r[VrTotal] / cantAct, 6),
            __K1 = r[CodAct], __K3 = 2, __K4 = r[Articulo]]),

    Ordenada = Table.Sort(Table.FromRecords(FilasNodo & FilasActividad & FilasInsumo),
        {{"__K1", Order.Ascending}, {"__K3", Order.Ascending}, {"__K4", Order.Ascending}}),
    SinClaves = Table.RemoveColumns(Ordenada, {"__K1", "__K3", "__K4"}),

    FilaCD = Table.FromRecords({FilaVacia & [Código = "CD", Descripción = "COSTOS DIRECTOS", Tipo = "CD"]}),
    TablaFinal = Table.Combine({FilaCD, SinClaves}),

    TiposDefinidos = Table.TransformColumnTypes(TablaFinal,{
        {"Código", type text}, {"Descripción", type text}, {"Padre", type text}, {"UM", type text},
        {"CANTIDAD", type number}, {"SUBCAPITULO", type text}, {"ID PROYECTO", type text},
        {"VERSION", type text}, {"ID APU", type text}, {"Cant APU", type number},
        {"Rend", type number}, {"IVA", type number}, {"VrUnitSinIVA", type number},
        {"Tipo Insumo", type text}, {"Agrupacion", type text}, {"COD CLIENTE", type text},
        {"Precio Cliente", type number}, {"Clase", type text},
        {"Tipo", type text}, {"Vr Unitario", type number}, {"Vr Total", type number}
    })
in
    TiposDefinidos
