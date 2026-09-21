let
    // ============================================================
    // COMPARATIVOS_POR_SUBIR: comparativos aprobados en SharePoint
    // (APROBACIONES_SP) que todavia no estan en la tabla manual Det_CC.
    // Una sola columna con el comparativo (numero-nombre). El cruce es por el
    // nombre normalizado (espacios, guiones, numero a 3 digitos, sin tildes,
    // mayusculas). No se listan aprobaciones con total <= 0 (anulaciones).
    // ============================================================
    FnNormComp = F_Globales[FnNormalizeComparativo],
    FnSinTildes = F_Globales[FnRemoveAccentsSymbols],
    FnClave = (t as any) as text =>
        let n = FnNormComp(t) in if n = null then "" else Text.Upper(FnSinTildes(n)),
    FnNum = (v as any) as number => let n = try Number.From(v) otherwise null in if n = null then 0 else n,

    Manual0 = try COMPARATIVOS otherwise #table({"# CC - Comparativo"}, {}),
    ClavesManual = List.Buffer(List.Distinct(List.Transform(Manual0[#"# CC - Comparativo"], FnClave))),

    Aprob0 = try APROBACIONES_SP otherwise #table({"# CC - Comparativo", "VT CC cons"}, {}),
    Aprob1 = Table.AddColumn(Aprob0, "__Clave", each FnClave([#"# CC - Comparativo"]), type text),
    Aprob = Table.Group(Table.SelectRows(Aprob1, each [__Clave] <> ""), {"__Clave"}, {
        {"Comparativo por subir", each List.First(List.RemoveNulls([#"# CC - Comparativo"])), type text},
        {"__Aprobado", each List.Sum(List.Transform([VT CC cons], FnNum)), type number}
    }),
    Pendientes = Table.SelectRows(Aprob, each [__Aprobado] > 0 and not List.Contains(ClavesManual, [__Clave])),
    Final = Table.Sort(Table.SelectColumns(Pendientes, {"Comparativo por subir"}), {{"Comparativo por subir", Order.Ascending}})
in
    Final
