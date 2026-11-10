

let
    Norm = (v as any) as text => let t = try Text.Upper(Text.Trim(F_Globales[FnRemoveAccentMarks](Text.From(v)))) otherwise null
        in if t = null then "" else Text.Combine(List.Select(Text.Split(t, " "), each _ <> ""), " "),
    Centro = (v as any) as text => Text.Trim(Text.BeforeDelimiter(Norm(v) & "-", "-")),
    Tabla = try Excel.CurrentWorkbook(){[Name="EQUIVALENCIAS_INSUMOS"]}[Content]
        otherwise #table({"Centro de Costos", "Ins memoria", "Codigo ins SINCO"}, {}),
    Faltantes = List.Difference({"Centro de Costos", "Ins memoria", "Codigo ins SINCO"}, Table.ColumnNames(Tabla)),
    Datos = if not List.IsEmpty(Faltantes) then error "Revise los encabezados de EQUIVALENCIAS_INSUMOS: Centro de Costos, Ins memoria y Codigo ins SINCO."
            else Table.Buffer(Table.SelectRows(Tabla, each Centro([Centro de Costos]) <> "" and Norm([Ins memoria]) <> "")),
    Catalogo = Table.Buffer(Table.Distinct(Table.SelectColumns(ITEMSINSUMOS, {"Centro de Costos", "Codigo ins", "Ins"}))),
    Claves = Table.AddColumn(Datos, "__K", each Centro([Centro de Costos]) & "|" & Norm([Ins memoria]), type text),
    Grupos = Table.Group(Claves, {"__K"}, {
        {"__CC", each Centro(List.First([Centro de Costos])), type text},
        {"__Codigos", each List.Distinct(List.Transform([Codigo ins SINCO], Norm)), type list}}),
    Revisadas = Table.AddColumn(Grupos, "__Nombre", each
        if List.Count([__Codigos]) <> 1 or [__Codigos]{0} = "" then null
        else let cc=[__CC], codigo=[__Codigos]{0}, nombres=List.Distinct(Table.SelectRows(Catalogo,
             each Centro([Centro de Costos]) = cc and Norm([Codigo ins]) = codigo)[Ins])
             in if List.Count(nombres) = 1 then Text.From(nombres{0}) else null, type nullable text),
    Invalidas = Table.SelectRows(Revisadas, each [__Nombre] = null),
    Mapa = Record.FromList(Revisadas[__Nombre], Revisadas[__K]),
    Resolver = (cc as any, ins as any) as text => Record.FieldOrDefault(Mapa, Centro(cc) & "|" & Norm(ins), Text.From(ins))
in
    if Table.RowCount(Invalidas) > 0 then error "Equivalencia ambigua, sin codigo o sin un unico insumo oficial en su CC. Revise EQUIVALENCIAS_INSUMOS."
    else Resolver
