let
    Base = Table.Buffer(FxFormatoPamplona("BD")),
    ClaveTexto = (v as any) as text =>
        let t = F_Globales[FnNormalizeSpaces](F_Globales[FnRemoveAccentMarks](v))
        in if t = null then "" else Text.Upper(t),
    ClaveComparativo = (r as record) as nullable text =>
        let
            cc = ClaveTexto(r[Centro de Costos]),
            prefijo = Text.BeforeDelimiter(cc & "-", "-"),
            comp = ClaveTexto(r[#"# CC - Comparativo"]),
            cola = if prefijo <> "" and Text.StartsWith(comp, prefijo & "-") then Text.AfterDelimiter(comp, "-") else "",
            consecutivo = Text.BeforeDelimiter(cola & "-", "-"),
            incluyeCC = Text.Contains(cola, "-") and consecutivo <> "" and Text.Select(consecutivo, {"0".."9"}) = consecutivo,
            local = if incluyeCC then cola else comp,
            normal = F_Globales[FnNormalizeComparativo](local)
        in if cc = "" or List.Contains({"", "SIN CUADRO COMPARATIVO", "POR ADJUDICAR", "-", "0"}, local) then null
           else cc & "|" & normal,
    ClaveInsumo = (r as record, comp as text) as nullable text =>
        let ins = ClaveTexto(r[Ins]) in if ins = "" then null else comp & "|" & ins,
    CategoriaValida = (v as any) as logical =>
        not List.Contains({"", "ZOTROS", "SIN CLASIFICADOR", "(SIN CLASIFICADOR)", "NO APLICA", "-"}, ClaveTexto(v)),
    AprobadosConClave = Table.AddColumn(
        Table.SelectRows(Base, each [Tipo] = "CC CONSOLIDADO" and CategoriaValida([Clasificador])),
        "__Comp", each ClaveComparativo(_), type nullable text),
    Aprobados = Table.Buffer(Table.SelectRows(AprobadosConClave, each [__Comp] <> null)),
    IndiceUnico = (t as table, c as text) as record =>
        let
            grupos = Table.Group(Table.SelectRows(t, each Record.Field(_, c) <> null), {c}, {
                {"Clases", each List.Distinct(List.Transform([Clasificador], Text.Trim), Comparer.OrdinalIgnoreCase), type list}
            }),
            unicos = Table.SelectRows(grupos, each List.Count([Clases]) = 1)
        in Record.FromList(List.Transform(unicos[Clases], each _{0}), Table.Column(unicos, c)),
    PorComparativo = IndiceUnico(Aprobados, "__Comp"),
    PorInsumo = IndiceUnico(Table.AddColumn(Aprobados, "__Ins", each ClaveInsumo(_, [__Comp]), type nullable text), "__Ins"),
    Buscar = (indice as record, clave as nullable text) as nullable text =>
        if clave = null then null else Record.FieldOrDefault(indice, clave, null),
    // El asegurado vinculado hereda la categoria aprobada de su comparativo.
    // Las reglas manuales siguen en la base para los usos por insumo.
    Categoria = (r as record) as nullable text =>
        if not List.Contains({"COMPRAS", "CONTRATO"}, r[Tipo]) or
            (List.Contains({0, null}, r[#"VT asegurado SINCO"]) and List.Contains({0, null}, r[#"Cantidad asegurada SINCO"])) then r[Clasificador]
        else
            let
                comp = ClaveComparativo(r),
                ins = if comp = null then null else Buscar(PorInsumo, ClaveInsumo(r, comp)),
                aprobada = if ins <> null then ins else Buscar(PorComparativo, comp)
            in if comp = null then r[Clasificador] else if aprobada <> null then aprobada else "zOtros",
    Resultado = Table.ReplaceValue(Base, each [Clasificador], each Categoria(_), Replacer.ReplaceValue, {"Clasificador"})
in
    Resultado
