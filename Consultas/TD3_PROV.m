let
    // TD3_PROV (2026-10-05): provisiones (bruto con descuentos) por OC, factura y numero de
    // provision, para la dinamica de contratos/OC. Se ubican en la OC que dice la provision;
    // si esa OC no existe en SINCO pero NIT + factura si existen, en la OC de SINCO.
    // Las demas (OC invalida y factura que no esta en SINCO) quedan fuera de la dinamica.
    FnTexto = (v as any) as text => let t = try Text.Trim(Text.From(v)) otherwise null in if t = null then "" else t,
    FnDigitos = (v as any) as text => Text.Select(FnTexto(v), {"0".."9"}),
    FnNIT = (v as any) as text => let d = FnDigitos(v) in if Text.Length(d) = 10 and List.Contains({"8", "9"}, Text.Start(d, 1)) then Text.Start(d, 9) else d,
    FnFactura = (v as any) as text => Text.TrimStart(Text.Select(Text.Upper(FnTexto(v)), {"A".."Z", "0".."9"}), "0"),
    FnNum = (v as any) as number => let n = try Number.From(v) otherwise null in if n = null then 0 else n,
    OCsValidas = List.Buffer(TD3_OC[OC]),
    // NIT + factura -> OC segun SINCO
    MapaSinco = let
            t = Table.SelectRows(Table.AddColumn(SINCO_FACTURAS, "__K", each FnNIT([NIT]) & "|" & [FacturaClave], type text), each [FacturaClave] <> ""),
            g = Table.Group(t, {"__K"}, {{"OC", each List.Mode([OC]), type text}})
        in Record.FromList(g[OC], g[__K]),
    Prov = Table.SelectColumns(PROVISIONES_SP, {"# OC / Contrato", "NIT", "No_Factura", "No_Prov", "VR_Bruto_con_desc"}, MissingField.UseNull),
    ConCampos = Table.AddColumn(Prov, "R", each
        let
            ocProv = FnDigitos([#"# OC / Contrato"]),
            clave = FnFactura([No_Factura]),
            ocSinco = if clave = "" then null else Record.FieldOrDefault(MapaSinco, FnNIT([NIT]) & "|" & clave, null),
            oc = if List.Contains(OCsValidas, ocProv) then ocProv else if ocSinco <> null and List.Contains(OCsValidas, ocSinco) then ocSinco else null
        in [OC = oc, FacturaClave = clave, Factura = if clave = "" then "(sin factura)" else FnTexto([No_Factura])]),
    Expandido = Table.ExpandRecordColumn(ConCampos, "R", {"OC", "FacturaClave", "Factura"}),
    Ubicadas = Table.SelectRows(Expandido, each [OC] <> null),
    Final = Table.FromRecords(List.Transform(Table.ToRecords(Ubicadas), (r) => [
            FactKey = r[OC] & "|" & (if r[FacturaClave] = "" then "SIN" else r[FacturaClave]),
            OC = r[OC], Factura = r[Factura],
            No_Prov = try Int64.From(Number.From(r[No_Prov])) otherwise null,
            Bruto = FnNum(r[VR_Bruto_con_desc])]),
        type table [FactKey = text, OC = text, Factura = text, No_Prov = Int64.Type, Bruto = number])
in
    Final
