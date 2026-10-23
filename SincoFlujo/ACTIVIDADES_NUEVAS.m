// ACTIVIDADES_NUEVAS (hoja ACTIVIDADES NUEVAS del flujo SINCO)
// Actividades que estan en SEGUIMIENTO POR ITEMS pero no en el catalogo APU.
// Columnas fijas (las lee el subflujo ACTIVIDADES): A CODIGO, B ACTIVIDAD,
// C UM, D CAPITULO, E SUBCAPITULO; F ESTADO lo escribe el flujo y se conserva
// entre refrescos (clave CODIGO).
let
    ActividadesSeg = SEGUIMIENTO_ACTIVIDADES,
    CodigosAPU = Table.Distinct(Table.TransformColumns(APU_CODIGOS, {"Codigo", each Text.Trim(Text.Upper(Text.From(_))), type text})),

    ConClave = Table.AddColumn(ActividadesSeg, "ClaveCodigo", each Text.Trim(Text.Upper(Text.From([Codigo]))), type text),
    Faltantes = Table.NestedJoin(ConClave, {"ClaveCodigo"}, CodigosAPU, {"Codigo"}, "Match", JoinKind.LeftAnti),
    Ordenado = Table.Sort(Faltantes, {{"Codigo", Order.Ascending}}),

    Nuevo = Table.RenameColumns(
        Table.SelectColumns(Ordenado, {"Codigo", "Descripcion", "UM", "Capitulo", "Subcapitulo"}),
        {{"Codigo", "CODIGO"}, {"Descripcion", "ACTIVIDAD"}, {"Capitulo", "CAPITULO"}, {"Subcapitulo", "SUBCAPITULO"}}
    ),

    Salida = FxConservarEstado(Nuevo, "ACTIVIDADES_NUEVAS", {"CODIGO"}, {"ESTADO"},
        {"CODIGO", "ACTIVIDAD", "UM", "CAPITULO", "SUBCAPITULO", "ESTADO"})
in
    Salida
