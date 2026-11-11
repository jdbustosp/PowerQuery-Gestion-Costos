// Agrega a BD_BASE el presupuesto V2 por insumo y los niveles de la estructura CBS.
let
    Arbol = ORA_ARBOL,
    // Forzar DISPONIBLE antes de BD_BASE: su try/otherwise central no debe ocultar una descarga ilegible.
    Base = if Table.RowCount(Table.Buffer(DISPONIBLE)) >= 0 then BD_BASE
        else error "No se pudo evaluar el presupuesto adjudicado y por adjudicar",
    // V2 es una medida independiente por CBS. La base y sus medidas iniciales se conservan.
    BaseConV2 = Table.AddColumn(Base, "VT Presupuesto v2", each null, type nullable number),
    FilasV2 = Table.SelectColumns(ORA_PPTO_V2, Table.ColumnNames(BaseConV2), MissingField.UseNull),
    Conjunto = Table.Combine({BaseConV2, FilasV2}),
    // Completar etiquetas solo en libros que habiliten la opción. El cruce de costos ya se evaluó.
    FormatearInsumos = (tabla as table) as table =>
        let
            Norm = (v as any) as text => if v = null then "" else Text.Upper(Text.Trim(Text.From(v))),
            EsUnidad = (v as text) as logical => v <> "" and not List.Contains({"-", "N/A", "S/U", "SIN UM"}, v),
            DeAsegurado = Table.TransformColumns(
                Table.SelectColumns(ORA_ASEGURADO, {"Articulo", "Registro", "UM"}),
                {{"Articulo", Norm, type text}, {"Registro", Norm, type text}, {"UM", Norm, type text}}),
            DePpto = Table.TransformColumns(
                Table.SelectColumns(Arbol[Insumos], {"Articulo", "UM"}),
                {{"Articulo", Norm, type text}, {"UM", Norm, type text}}),
            Validar = (t as table) as table => Table.SelectRows(t, each [Articulo] <> "" and EsUnidad([UM])),
            PorRegistro = Table.AddColumn(Validar(DeAsegurado), "Clave", each [Registro] & "|" & [Articulo], type text),
            MapaUnico = (t as table, clave as text) as record =>
                let
                    g = Table.Group(t, {clave}, {{"Unidades", each List.Distinct([UM]), type list}}),
                    u = Table.SelectRows(g, each List.Count([Unidades]) = 1)
                in Record.FromList(List.Transform(u[Unidades], each _{0}), Table.Column(u, clave)),
            MapaRegistro = MapaUnico(PorRegistro, "Clave"),
            MapaArticulo = MapaUnico(Table.Combine({Validar(DeAsegurado), Validar(DePpto)}), "Articulo"),
            Complementos = Record.FieldOrDefault(ORA_CONFIG, "UnidadesComplementarias", []),
            Unidades = List.Buffer(List.Distinct(List.Combine({DeAsegurado[UM], DePpto[UM], Record.FieldValues(Complementos), {"SIN UM"}}))),
            ConNombre = Table.AddColumn(tabla, "__InsUM", each
                let
                    nombre = if [Ins] = null then "" else Text.Trim(Text.From([Ins])),
                    codigo = Norm([Codigo ins]),
                    segmentos = Text.Split(codigo, "-"),
                    esArticulo = List.Count(segmentos) = 2 and Text.Length(segmentos{1}) > 0
                        and Text.Select(segmentos{1}, {"0".."9"}) = segmentos{1},
                    sufijo = if Text.EndsWith(nombre, ")") then
                        (try Norm(Text.BeforeDelimiter(Text.AfterDelimiter(nombre, "(", {0, RelativePosition.FromEnd}), ")")) otherwise "")
                        else "",
                    unidad = Record.FieldOrDefault(MapaRegistro, Norm([#"# OC / Contrato"]) & "|" & codigo, null)
                        ?? Record.FieldOrDefault(MapaArticulo, codigo, null)
                        ?? Record.FieldOrDefault(Complementos, codigo, "SIN UM")
                in if nombre = "" or not esArticulo or (sufijo <> "" and List.Contains(Unidades, sufijo)) then [Ins]
                    else nombre & " (" & unidad & ")", type nullable text),
            Final = Table.ReorderColumns(
                Table.RenameColumns(Table.RemoveColumns(ConNombre, {"Ins"}), {{"__InsUM", "Ins"}}),
                Table.ColumnNames(tabla))
        in Final,
    Etiquetas = if Record.FieldOrDefault(ORA_CONFIG, "MostrarUnidadInsumo", false) then FormatearInsumos(Conjunto) else Conjunto,
    FnCod = (a as any) as nullable text =>
        let t = try Text.Trim(Text.From(a)) otherwise null
        in if t = null or not Text.Contains(t, " - ") then null else Text.BeforeDelimiter(t, " - "),
    ConNiveles = Table.AddColumn(Etiquetas, "__N", each Arbol[FnNiveles](FnCod([Actividad]))),
    Resultado = Table.ExpandRecordColumn(ConNiveles, "__N", {"Nivel 1", "Nivel 2", "Nivel 3", "Nivel 4"})
in
    Table.Buffer(Resultado)
