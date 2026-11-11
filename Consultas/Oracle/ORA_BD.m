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
    FnCod = (a as any) as nullable text =>
        let t = try Text.Trim(Text.From(a)) otherwise null
        in if t = null or not Text.Contains(t, " - ") then null else Text.BeforeDelimiter(t, " - "),
    ConNiveles = Table.AddColumn(Conjunto, "__N", each Arbol[FnNiveles](FnCod([Actividad]))),
    Resultado = Table.ExpandRecordColumn(ConNiveles, "__N", {"Nivel 1", "Nivel 2", "Nivel 3", "Nivel 4"})
in
    Table.Buffer(Resultado)
