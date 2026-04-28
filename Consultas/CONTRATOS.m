let
    // ============================================================
    // 1. FUNCIONES AUXILIARES GLOBALES
    // ============================================================
    FnFormatCodigoAct = F_Globales[FnFormatCodigoAct],
    FxToNumberFlex = F_Globales[FxToNumberFlex],
    FnClaveLimpia = F_Globales[FnClaveLimpia],

    Columnas_HTML = List.Transform({1..40}, each {"Columna" & Text.From(_), "td:nth-child(" & Text.From(_) & "), th:nth-child(" & Text.From(_) & ")"}),

    // ============================================================
    // 2. FUNCIÓN MÁGICA: PROCESAR CORTES
    // ============================================================
    FxProcesarCortes = (BinarioCortes as binary) =>
        let
            // Uso de Windows-1252 para mitigar símbolos raros desde el origen
            TextoHTML = Text.FromBinary(Binary.Buffer(BinarioCortes), 1252), 
            Source = Html.Table(TextoHTML, Columnas_HTML, [RowSelector="tr"]), 
            AddFilaTexto = Table.AddColumn(Source, "FilaTexto", each let vals = Record.FieldValues(_), soloTexto = List.Transform(List.Select(vals, each _ <> null and _ <> ""), Text.From) in Text.Trim(Text.Combine(soloTexto, " ")), type text),
            AddOC = Table.AddColumn(AddFilaTexto, "# OC / Contrato", each let txt = [FilaTexto] in if txt <> null and Text.Contains(Text.Upper(txt), "CONTRATO NO") then let after = Text.TrimStart(Text.Replace(Text.Range(txt, Text.PositionOf(Text.Upper(txt), "CONTRATO NO") + 11), "#(00A0)", " "), {".", ":", " "}), first = Text.BeforeDelimiter(after, " "), num = Text.Select(if first = "" then after else first, {"0".."9"}) in if num = "" then null else num else null, type text),
            AddDesc = Table.AddColumn(AddOC, "Descripcion contrato", each let txt = [FilaTexto] in if txt <> null and Text.Contains(Text.Upper(txt), "CONTRATO NO") then let after = Text.TrimStart(Text.Range(txt, Text.PositionOf(Text.Upper(txt), "CONTRATO NO") + 11), {".", ":", " "}), idx = Text.PositionOfAny(after, {"A".."Z","a".."z"}), desc = if idx = -1 then null else Text.Range(after, idx), lim = if desc = null then null else if Text.Contains(Text.Upper(desc), "CONTRATISTA") then Text.BeforeDelimiter(Text.Upper(desc), "CONTRATISTA") else desc in if lim = null then null else Text.Trim(lim) else null, type text),
            AddNombre = Table.AddColumn(AddDesc, "Nombre Contratista", each let txt = [FilaTexto] in if txt <> null and Text.Contains(Text.Upper(txt), "CONTRATISTA") then Text.Trim(Text.TrimStart(Text.AfterDelimiter(Text.Upper(txt), "CONTRATISTA"), {":"," ","-"})) else null, type text),
            FillDown1 = Table.FillDown(AddNombre, {"# OC / Contrato","Descripcion contrato","Nombre Contratista"}),
            AddCodAct = Table.AddColumn(FillDown1, "CodigoAct", each let c = [Columna1], t = if c = null then null else Text.Trim(Text.From(c)) in if t <> null and t <> "" and (try Number.From(Text.Replace(t, ".", "")) otherwise null) <> null then FnFormatCodigoAct(t) else null, type text),
            AddActFuente = Table.AddColumn(AddCodAct, "ActividadFuente", each if [CodigoAct] <> null then [Columna2] else null, type text),
            FillDown2 = Table.FillDown(AddActFuente, {"CodigoAct", "ActividadFuente"}),
            
            AddCantC = Table.AddColumn(FillDown2, "Cantidades contrato", each FxToNumberFlex([Columna4]), type number),
            AddVTC = Table.AddColumn(AddCantC, "VT contrato", each FxToNumberFlex([Columna5]), type number),
            AddCantCortes = Table.AddColumn(AddVTC, "Cantidad Cortes", each FxToNumberFlex([Columna10]), type number),
            AddNums = Table.AddColumn(AddCantCortes, "VT Cortes", each FxToNumberFlex([Columna11]), type number),
            
            Filtered = Table.SelectRows(AddNums, each 
                [Columna2] <> null and 
                [CodigoAct] <> null and 
                ([Columna1] = null or Text.Trim(Text.From([Columna1])) = "") and 
                not Text.Contains(Text.Upper(Text.From([Columna1] ?? "")), "TOTAL") and
                not Text.Contains(Text.Upper(Text.From([Columna2] ?? "")), "TOTAL")
            ),
            
            // Creamos la clave robusta para el cruce en SINCO
            AddClave = Table.AddColumn(Filtered, "InsClave_Cruce", each FnClaveLimpia([Columna2]), type text)
        in AddClave, 

    // ============================================================
    // CONEXIÓN A SHAREPOINT (LECTURA DESDE CONSULTA COMPARTIDA)
    // ============================================================
    ArchivosProyecto = Table.SelectRows(SP_Archivos_Proyecto, each 
        Text.Contains([Name], "ESTADO DE CONTRATOS", Comparer.OrdinalIgnoreCase)
    ),
    ConCentroCosto = ArchivosProyecto,
    
    TablaConDatos = Table.AddColumn(ConCentroCosto, "Datos", each FxProcesarCortes([Content])),
    SoloDatos = Table.SelectColumns(TablaConDatos, {"Centro de Costos", "Datos"}),
    
    Expandido = Table.ExpandTableColumn(SoloDatos, "Datos", {"# OC / Contrato", "Descripcion contrato", "Nombre Contratista", "CodigoAct", "ActividadFuente", "Cantidades contrato", "VT contrato", "Cantidad Cortes", "VT Cortes", "Columna2", "InsClave_Cruce"}),
    
    Expandido_Clean = Table.TransformColumns(Expandido, {
        {"Centro de Costos", each if _ = null then null else Text.Upper(Text.Trim(Text.From(_))), type text},
        {"CodigoAct", each FnFormatCodigoAct(_), type text}
    }, null, MissingField.Ignore),

    Expandido_Unico = Table.Buffer(Table.Distinct(Expandido_Clean)),

    // ============================================================
    // 4. LECTURA DE LA CONSULTA MAESTRA (EL DICCIONARIO OFICIAL)
    // ============================================================
    ITEMS_Raw = ITEMSINSUMOS,
    
    ITEMS_Clean = Table.Buffer(Table.TransformColumns(ITEMS_Raw, {
        {"Centro de Costos", each if _ = null then null else Text.Upper(Text.Trim(Text.From(_))), type text},
        {"Codigo act", each FnFormatCodigoAct(_), type text}
    }, null, MissingField.Ignore)),

    ITEMS_Jerarquia = Table.Buffer(Table.Group(ITEMS_Clean, {"Centro de Costos", "Codigo act"}, {
        {"Ref.Act", each List.First(List.RemoveNulls([Actividad])), type text},
        {"Ref.Cap", each List.First(List.RemoveNulls([Capitulo])), type text}, 
        {"Ref.Sub", each List.First(List.RemoveNulls([Subcapitulo])), type text}
    })),

    // 🔥 PREPARAMOS TBITEMS: Creamos la misma llave y extraemos su nombre oficial perfecto (con unidades) y su código.
    ITEMS_Insumos_Base = Table.AddColumn(ITEMS_Clean, "InsClave_Cruce", each FnClaveLimpia([Ins]), type text),
    ITEMS_Insumos_Dist = Table.Buffer(Table.Group(ITEMS_Insumos_Base, {"Centro de Costos", "Codigo act", "InsClave_Cruce"}, {
        {"Ref.InsOficial", each List.First([Ins]), type text},
        {"Ref.CodIns", each List.First([Codigo ins]), type any}
    })),

    // ============================================================
    // 5. CRUCES FINALES Y REEMPLAZO DE NOMBRES
    // ============================================================
    MergedJerarquia = Table.NestedJoin(Expandido_Unico, {"Centro de Costos", "CodigoAct"}, ITEMS_Jerarquia, {"Centro de Costos", "Codigo act"}, "JER", JoinKind.LeftOuter),
    ExpandedJerarquia = Table.ExpandTableColumn(MergedJerarquia, "JER", {"Ref.Act", "Ref.Cap", "Ref.Sub"}, {"Ref.Act", "Ref.Cap", "Ref.Sub"}),
    
    MergedInsumos = Table.NestedJoin(ExpandedJerarquia, {"Centro de Costos", "CodigoAct", "InsClave_Cruce"}, ITEMS_Insumos_Dist, {"Centro de Costos", "Codigo act", "InsClave_Cruce"}, "INS", JoinKind.LeftOuter),
    ExpandedInsumos = Table.ExpandTableColumn(MergedInsumos, "INS", {"Ref.CodIns", "Ref.InsOficial"}, {"Ref.CodIns", "Ref.InsOficial"}),

    AddFinalCols = Table.AddColumn(ExpandedInsumos, "FinalCols", each [
        // Si cruzó, toma el nombre OFICIAL de TbItems (Ej: "X_PROVISIONAL DE ENERGIA (GL)"). Si no, deja el de SINCO.
        I = if [Ref.InsOficial] <> null then [Ref.InsOficial] else (if [Columna2] = null or Text.Trim([Columna2]) = "" then "SIN DESCRIPCION" else Text.Trim([Columna2])),
        
        A_Original = let a0 = Text.Trim(Text.From([ActividadFuente] ?? "")) in if a0 = "" then null else if [CodigoAct] <> null and not Text.StartsWith(a0, Text.From([CodigoAct])) then Text.From([CodigoAct]) & " - " & a0 else a0,
        A = [Ref.Act] ?? A_Original
    ]),
    ExpandedFinal = Table.ExpandRecordColumn(AddFinalCols, "FinalCols", {"I", "A"}, {"Ins", "Actividad"}),
    
    AddCodIns = Table.AddColumn(ExpandedFinal, "Codigo ins_Final", each [Ref.CodIns]),
    
    Selected = Table.SelectColumns(Table.AddColumn(AddCodIns, "Tipo", each "Contrato"), {"Centro de Costos", "Codigo ins_Final", "Ins", "CodigoAct", "Actividad", "Ref.Cap", "Ref.Sub", "# OC / Contrato", "Nombre Contratista", "Descripcion contrato", "Cantidades contrato", "VT contrato", "Cantidad Cortes", "VT Cortes", "Tipo"}),
    Renamed = Table.RenameColumns(Selected, {{"Codigo ins_Final", "Codigo ins"}, {"CodigoAct", "Codigo act"}, {"Ref.Cap", "Capitulo"}, {"Ref.Sub", "Subcapitulo"}, {"Cantidades contrato", "Cantidad Contratado"}, {"VT contrato", "VT Contratado"}}),
    
    Typed = Table.TransformColumnTypes(Renamed, {{"Centro de Costos", type text}, {"Codigo ins", Int64.Type}, {"Cantidad Contratado", type number}, {"VT Contratado", type number}, {"Cantidad Cortes", type number}, {"VT Cortes", type number}}),
    
    FilteredZeros = Table.SelectRows(Typed, each ([VT Contratado] <> 0 and [VT Contratado] <> null)),
    
    TablaFinal = FilteredZeros
in
    TablaFinal
