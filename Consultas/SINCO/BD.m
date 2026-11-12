let
    Base = Table.Buffer(BD_Origen),
    // Los libros de formato anterior conservan sus columnas de salida.
    EsLegado = not Table.HasColumns(Base, "VT asegurado SINCO"),
    ColValor = if EsLegado then "VT Asegurada" else "VT asegurado SINCO",
    ColCantidad = if EsLegado then "Cantidad asegurada" else "Cantidad asegurada SINCO",
    ClaveTexto = (v as any) as text =>
        let t = F_Globales[FnNormalizeSpaces](F_Globales[FnRemoveAccentMarks](v))
        in if t = null then "" else Text.Upper(t),
    EsReal = (comp as text) as logical =>
        not List.Contains({"", "SIN CUADRO COMPARATIVO", "POR ADJUDICAR", "-", "0"}, comp),
    ClaveComparativo = (r as record) as nullable text =>
        let
            cc = ClaveTexto(r[Centro de Costos]),
            prefijo = Text.BeforeDelimiter(cc & "-", "-"),
            comp = ClaveTexto(r[#"# CC - Comparativo"]),
            cola = if prefijo <> "" and Text.StartsWith(comp, prefijo & "-") then Text.AfterDelimiter(comp, "-") else "",
            consecutivo = Text.BeforeDelimiter(cola & "-", "-"),
            incluyeCC = Text.Contains(cola, "-") and consecutivo <> "" and Text.Select(consecutivo, {"0".."9"}) = consecutivo,
            local = if incluyeCC then cola else comp
        in if cc = "" or not EsReal(local) then null
           else cc & "|" & F_Globales[FnNormalizeComparativo](local),
    ConInsumo = (clave as nullable text, r as record) as nullable text =>
        let ins = ClaveTexto(r[Ins])
        in if clave = null or ins = "" then null else clave & "|" & ins,
    CategoriaValida = (v as any) as logical =>
        not List.Contains({"", "ZOTROS", "SIN CLASIFICADOR", "(SIN CLASIFICADOR)", "NO APLICA", "-"}, ClaveTexto(v)),
    IndiceUnico = (t as table, clave as text, valor as text) as record =>
        let
            grupos = Table.Group(Table.SelectRows(t, each Record.Field(_, clave) <> null), {clave}, {
                {"Valores", each List.Distinct(List.Transform(Table.Column(_, valor), Text.Trim), Comparer.OrdinalIgnoreCase), type list}
            }),
            unicos = Table.SelectRows(grupos, each List.Count([Valores]) = 1)
        in Record.FromList(List.Transform(unicos[Valores], each _{0}), Table.Column(unicos, clave)),
    Buscar = (indice as record, clave as nullable text) as nullable text =>
        if clave = null then null else Record.FieldOrDefault(indice, clave, null),
    AprobadosConClave = Table.AddColumn(
        Table.SelectRows(Base, each [Tipo] = "CC CONSOLIDADO" and CategoriaValida([Clasificador])),
        "__Comp", each ClaveComparativo(_), type nullable text),
    Aprobados = Table.Buffer(Table.SelectRows(AprobadosConClave, each [__Comp] <> null)),
    AprobadosPorComparativo = IndiceUnico(Aprobados, "__Comp", "Clasificador"),
    AprobadosPorInsumo = IndiceUnico(
        Table.AddColumn(Aprobados, "__Ins", each ConInsumo([__Comp], _), type nullable text), "__Ins", "Clasificador"),
    Manual = try Excel.CurrentWorkbook(){[Name = "MaestroInsumos"]}[Content]
             otherwise #table(type table [Ins = nullable text, Clasificador = nullable text], {}),
    ManualValido = Table.SelectRows(Manual, each ClaveTexto([Ins]) <> "" and CategoriaValida([Clasificador])),
    ManualPorInsumo = IndiceUnico(
        Table.AddColumn(ManualValido, "__Ins", each ClaveTexto([Ins]), type text), "__Ins", "Clasificador"),
    TieneValor = (v as any) as logical =>
        let n = try Number.From(v) otherwise null in n <> null and n <> 0,
    EsAsegurado = (r as record) as logical =>
        List.Contains({"COMPRAS", "CONTRATO"}, r[Tipo]) and
        (TieneValor(Record.Field(r, ColValor)) or TieneValor(Record.Field(r, ColCantidad))),
    // La base anterior al formato de 67 columnas no ubicaba todos los
    // asegurados en el comparativo. Vincular solo por coincidencia univoca.
    ClaveOC = (r as record) as nullable text =>
        let
            cc = ClaveTexto(r[Centro de Costos]),
            oc = Text.Select(ClaveTexto(r[#"# OC / Contrato"]), {"0".."9"})
        in if cc = "" or oc = "" or Text.Remove(oc, {"0"}) = "" then null else cc & "|" & oc,
    Puente = Table.Buffer(Table.AddColumn(
        Table.SelectRows(Base, each [Tipo] = "CC" and ClaveComparativo(_) <> null),
        "__OC", each ClaveOC(_), type nullable text)),
    ComparativoPorOC = IndiceUnico(Puente, "__OC", "# CC - Comparativo"),
    ComparativoPorOCIns = IndiceUnico(
        Table.AddColumn(Puente, "__OCIns", each ConInsumo([__OC], _), type nullable text),
        "__OCIns", "# CC - Comparativo"),
    ComparativoResuelto = (r as record) as nullable text =>
        if not EsAsegurado(r) or ClaveComparativo(r) <> null then r[#"# CC - Comparativo"]
        else
            let
                oc = ClaveOC(r),
                porIns = Buscar(ComparativoPorOCIns, ConInsumo(oc, r)),
                comp = if porIns <> null then porIns else Buscar(ComparativoPorOC, oc)
            in if comp = null then r[#"# CC - Comparativo"] else comp,
    ConComparativo = if EsLegado then Table.ReplaceValue(Base, each [#"# CC - Comparativo"],
        each ComparativoResuelto(_), Replacer.ReplaceValue, {"# CC - Comparativo"}) else Base,
    Clasificador = (r as record) as nullable text =>
        if not List.Contains({"CC", "ADJUDICADO"}, r[Tipo]) and not EsAsegurado(r) then r[Clasificador]
        else
            let
                comp = ClaveComparativo(r),
                ins = Buscar(AprobadosPorInsumo, ConInsumo(comp, r)),
                aprobado = if ins <> null then ins else Buscar(AprobadosPorComparativo, comp),
                manual = if aprobado <> null then null else Buscar(ManualPorInsumo, ClaveTexto(r[Ins]))
            in if aprobado <> null then aprobado else if manual <> null then manual else "zOtros",
    Resultado = Table.ReplaceValue(ConComparativo, each [Clasificador], each Clasificador(_), Replacer.ReplaceValue, {"Clasificador"})
in
    Resultado
