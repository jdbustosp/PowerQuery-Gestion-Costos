let
    // ============================================================
    // COMPARATIVOS_POR_SUBIR: lo que falta cuadrar en la tabla manual Det_CC.
    //  - Aprobado en SharePoint (APROBACIONES_SP) que no esta en Det_CC, o que
    //    esta con otro valor, o sin # OC / Contrato.
    //  - Comparativos de las DESCARGAS que no aparecen con el mismo nombre en
    //    Det_CC: mientras no coincidan, la Proyeccion Colsubsidio no sabe que ya
    //    se contrataron y siguen proyectando lo descargado.
    // La clave es el nombre del comparativo normalizado (espacios, guiones,
    // numero a 3 digitos, sin tildes, mayusculas).
    // ============================================================
    FnNormComp = F_Globales[FnNormalizeComparativo],
    FnSinTildes = F_Globales[FnRemoveAccentsSymbols],
    FnClave = (t as any) as text =>
        let n = FnNormComp(t) in if n = null then "" else Text.Upper(FnSinTildes(n)),
    FnPrefijo = (clave as text) as text =>
        if Text.Contains(clave, "-") then Text.BeforeDelimiter(clave, "-") else "",
    FnNum = (v as any) as number => let n = try Number.From(v) otherwise null in if n = null then 0 else n,
    FnTxt = (v as any) as text => let t = try Text.Trim(Text.From(v)) otherwise "" in if t = null then "" else t,
    Tolerancia = 10000,

    // ---------- Det_CC (tabla manual) ----------
    Manual0 = try COMPARATIVOS otherwise #table({"# CC - Comparativo", "VR total aprobacion", "# OC / Contrato"}, {}),
    Manual1 = Table.AddColumn(Manual0, "__Clave", each FnClave([#"# CC - Comparativo"]), type text),
    Manual = Table.Buffer(Table.Group(Table.SelectRows(Manual1, each [__Clave] <> ""), {"__Clave"}, {
        {"Nombre en Det_CC", each List.First(List.RemoveNulls([#"# CC - Comparativo"])), type text},
        {"En Det_CC", each List.Sum(List.Transform([VR total aprobacion], FnNum)), type number},
        {"OC en Det_CC", each Text.Combine(List.Sort(List.Distinct(List.Select(List.Transform([#"# OC / Contrato"], FnTxt), each _ <> ""))), ", "), type text}
    })),
    NombresManualPorPrefijo = Table.Buffer(Table.Group(
        Table.AddColumn(Manual, "__Pref", each FnPrefijo([__Clave]), type text),
        {"__Pref"}, {{"Posible nombre en Det_CC", each Text.Combine([Nombre en Det_CC], " / "), type text}})),

    // ---------- Aprobado en SharePoint ----------
    Aprob0 = try APROBACIONES_SP otherwise #table({"# CC - Comparativo", "VT CC cons", "Centro de Costos", "Nombre Contratista"}, {}),
    Aprob1 = Table.AddColumn(Aprob0, "__Clave", each FnClave([#"# CC - Comparativo"]), type text),
    Aprob = Table.Group(Table.SelectRows(Aprob1, each [__Clave] <> ""), {"__Clave"}, {
        {"Comparativo", each List.First(List.RemoveNulls([#"# CC - Comparativo"])), type text},
        {"Centro de Costos", each List.First(List.RemoveNulls([Centro de Costos])), type text},
        {"Contratista", each List.First(List.RemoveNulls([Nombre Contratista])), type text},
        {"Aprobado SharePoint", each List.Sum(List.Transform([VT CC cons], FnNum)), type number}
    }),
    AprobCruce = Table.ExpandTableColumn(
        Table.NestedJoin(Aprob, {"__Clave"}, Manual, {"__Clave"}, "__M", JoinKind.LeftOuter),
        "__M", {"En Det_CC", "OC en Det_CC"}),
    AprobEstado = Table.AddColumn(AprobCruce, "Estado", each
        if [En Det_CC] = null then "Aprobado, falta subirlo a Det_CC"
        else if Number.Abs(FnNum([Aprobado SharePoint]) - FnNum([En Det_CC])) > Tolerancia then "Valor distinto al aprobado"
        else if FnTxt([OC en Det_CC]) = "" then "Falta # OC / Contrato en Det_CC"
        else null, type text),
    AprobPendientes = Table.SelectRows(AprobEstado, each [Estado] <> null),

    // ---------- Descargas sin nombre igual en Det_CC ----------
    Desc0 = try DESCARGAS otherwise #table({"# CC - Comparativo", "Valor Total ppto (CC)", "Centro de Costos"}, {}),
    Desc1 = Table.AddColumn(Desc0, "__Clave", each FnClave([#"# CC - Comparativo"]), type text),
    Desc = Table.Group(Table.SelectRows(Desc1, each [__Clave] <> ""), {"__Clave"}, {
        {"Comparativo", each List.First(List.RemoveNulls([#"# CC - Comparativo"])), type text},
        {"Centro de Costos", each List.First(List.RemoveNulls([Centro de Costos])), type text},
        {"Descargado", each List.Sum(List.Transform([#"Valor Total ppto (CC)"], FnNum)), type number}
    }),
    DescSinPareja = Table.NestedJoin(Desc, {"__Clave"}, Manual, {"__Clave"}, "__M", JoinKind.LeftAnti),
    DescEstado = Table.AddColumn(Table.RemoveColumns(DescSinPareja, {"__M"}), "Estado",
        each "Descarga con nombre distinto o sin subir a Det_CC", type text),

    // ---------- Union + sugerencia por numero de comparativo ----------
    Union = Table.Combine({AprobPendientes, DescEstado}),
    ConPref = Table.AddColumn(Union, "__Pref", each FnPrefijo([__Clave]), type text),
    ConSugerencia = Table.ExpandTableColumn(
        Table.NestedJoin(ConPref, {"__Pref"}, NombresManualPorPrefijo, {"__Pref"}, "__S", JoinKind.LeftOuter),
        "__S", {"Posible nombre en Det_CC"}),
    ConSugerencia2 = Table.ReplaceValue(ConSugerencia, each [#"Posible nombre en Det_CC"], each if [En Det_CC] <> null then null else [#"Posible nombre en Det_CC"], Replacer.ReplaceValue, {"Posible nombre en Det_CC"}),
    // Solo dos columnas: en la hoja Comparativos quedan libres R y S (Det_CC empieza en U).
    FnPesos = (v as any) as text => "$ " & Number.ToText(Number.Round(FnNum(v)), "N0", "es-CO"),
    ConTexto = Table.AddColumn(ConSugerencia2, "Qu#(00E9) falta", each
        if [Estado] = "Aprobado, falta subirlo a Det_CC" then
            "Subir a Det_CC (aprobado " & FnPesos([Aprobado SharePoint]) & ")"
            & (if [#"Posible nombre en Det_CC"] <> null then ". En Det_CC hay con el mismo n#(00FA)mero: " & [#"Posible nombre en Det_CC"] else "")
        else if [Estado] = "Valor distinto al aprobado" then
            "Revisar valor: aprobado " & FnPesos([Aprobado SharePoint]) & " vs Det_CC " & FnPesos([En Det_CC])
        else if [Estado] = "Falta # OC / Contrato en Det_CC" then
            "Falta el # OC / Contrato en Det_CC"
        else
            "El nombre en descargas no coincide con Det_CC (descargado " & FnPesos([Descargado]) & ")"
            & (if [#"Posible nombre en Det_CC"] <> null then ". Cambiarlo a: " & [#"Posible nombre en Det_CC"] else ". No est#(00E1) en Det_CC"),
        type text),
    Final = Table.Sort(
        Table.SelectColumns(ConTexto, {"Comparativo", "Qu#(00E9) falta"}),
        {{"Comparativo", Order.Ascending}})
in
    Final
