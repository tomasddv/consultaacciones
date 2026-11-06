from __future__ import annotations

import json
from pathlib import Path

import pandas as pd


BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
OLD_GROUPS = Path(r"C:\Users\triesgo\Documents\Codex\2026-08-06\https-distribuidoradelvalle-chesserp-com-ar889-ventas\work\erp_groups_full_2026-10-22-all-actions.json")
CLIENTS = DATA / "reporte_clientes_drive_2026-10-26.xlsx"
ERP_EXPORT = Path(r"C:\Users\triesgo\Downloads\Acciones Comerciales (17).xlsx")
PREVIOUS_ERP_EXPORT = Path(r"C:\Users\triesgo\Downloads\Acciones Comerciales (16).xlsx")


NEW_ACTIONS = [
    dict(promo="328365", accion_id="3558", segmento="MAYORISTA", tipo="1890", alcance="LATA 1890 MAYORISTA", nombre="6% DCTO EN LATA 1890 DESDE 100 BULTOS", descuento_regla="Desde 100 bultos: 6%", canal="MAY", calibre="1890 LATA", grupo="3120", grupo_nombre="MAYORISTA PAIS", estado_grupo="Confirmado ERP", tiers=[{"min": 100.0, "pct": 6.0}], familia="MAYORISTA 1890 LATA"),
    dict(promo="331085", accion_id="3569", segmento="SABORIZADAS", tipo="DROP", alcance="DROP BAJO SABORIZADAS", nombre="DROP BAJO SABORIZADAS SUR", descuento_regla="Desde 2 bultos: 25%; desde 4: 50%", canal="SUR", calibre="H2OH", grupo="3070", grupo_nombre="DROP BAJO MS SUR", estado_grupo="Confirmado ERP", tiers=[{"min": 2.0, "pct": 25.0}, {"min": 4.0, "pct": 50.0}], familia="SABORIZADAS"),
    dict(promo="331087", accion_id="3570", segmento="SABORIZADAS", tipo="DROP", alcance="DROP ALTO SABORIZADAS", nombre="DROP ALTO SABORIZADAS SUR", descuento_regla="Desde 5 bultos: 25%; desde 10: 50%", canal="SUR", calibre="H2OH", grupo="3071", grupo_nombre="DROP ALTO MS SUR", estado_grupo="Confirmado ERP", tiers=[{"min": 5.0, "pct": 25.0}, {"min": 10.0, "pct": 50.0}], familia="SABORIZADAS"),
    dict(promo="330980", accion_id="3571", segmento="MAYORISTA", tipo="DROP", alcance="DROP SABORIZADAS MAYORISTA", nombre="DROP SABORIZADAS MAY NACIONAL", descuento_regla="Desde 30 bultos: 25%; desde 60: 50%", canal="MAY", calibre="H2OH", grupo="2958", grupo_nombre="MAY UNG TT PAIS", estado_grupo="Confirmado ERP", tiers=[{"min": 30.0, "pct": 25.0}, {"min": 60.0, "pct": 50.0}], familia="SABORIZADAS MAYORISTA"),
    dict(promo="335569", accion_id="3596", segmento="PREMIUM", tipo="JAGER", alcance="ESC JAGER TOTAL PAIS", nombre="ESC JAGER", descuento_regla="Desde 0,01 bultos: 12%; desde 2: 17%", canal="TODOS", calibre="JAGER", grupo="3170", grupo_nombre="TOTAL PAIS VF", estado_grupo="Confirmado ERP", tiers=[{"min": 0.01, "pct": 12.0}, {"min": 2.0, "pct": 17.0}], familia="JAGER"),
    dict(promo="335570", accion_id="3597", segmento="PREMIUM", tipo="JAGER", alcance="CLIENTES ESPECIALES JAGER", nombre="CTES ESP JAGER", descuento_regla="Desde 0,01 bultos: 0,01%", canal="TODOS", calibre="JAGER", grupo="3169", grupo_nombre="GC-100335570", estado_grupo="Confirmado ERP", tiers=[{"min": 0.01, "pct": 0.01}], familia="JAGER"),
    dict(promo="335977", accion_id="3598", segmento="RED BULL", tipo="ESCALA", alcance="RED BULL 15% AS", nombre="RED BULL 15% AS", descuento_regla="Desde 0,16 bultos: 15%; 0,5: 11%; 1: 9%; 2,5: 7%", canal="AS", calibre="RED BULL", grupo="3174", grupo_nombre="RED BULL 15% AS", estado_grupo="Confirmado ERP", tiers=[{"min": 0.16, "pct": 15.0}, {"min": 0.5, "pct": 11.0}, {"min": 1.0, "pct": 9.0}, {"min": 2.5, "pct": 7.0}], familia="RED BULL"),
    dict(promo="335973", accion_id="3599", segmento="RED BULL", tipo="ESCALA", alcance="RED BULL 20% AS", nombre="RED BULL 20% AS", descuento_regla="Desde 0,16 bultos: 20%; 0,5: 16%; 1: 14%; 2,5: 12%", canal="AS", calibre="RED BULL", grupo="3175", grupo_nombre="RED BULL 20% AS", estado_grupo="Confirmado ERP", tiers=[{"min": 0.16, "pct": 20.0}, {"min": 0.5, "pct": 16.0}, {"min": 1.0, "pct": 14.0}, {"min": 2.5, "pct": 12.0}], familia="RED BULL"),
    dict(promo="335962", accion_id="3600", segmento="RED BULL", tipo="ESCALA", alcance="RED BULL 20% KT", nombre="RED BULL 20% KT", descuento_regla="Desde 0,16 bultos: 16%; 0,66: 14%; 1: 12%", canal="K+T", calibre="RED BULL", grupo="3173", grupo_nombre="RED BULL 20% KT", estado_grupo="Confirmado ERP", tiers=[{"min": 0.16, "pct": 16.0}, {"min": 0.66, "pct": 14.0}, {"min": 1.0, "pct": 12.0}], familia="RED BULL"),
    dict(promo="335945", accion_id="3601", segmento="RED BULL", tipo="VOLUMEN", alcance="VOLUMEN RB KT", nombre="VOLUMEN RB KT", descuento_regla="Desde 10 bultos: 30%", canal="K+T", calibre="RED BULL", grupo="3176", grupo_nombre="RED BULL VOLUMEN KT", estado_grupo="Confirmado ERP", tiers=[{"min": 10.0, "pct": 30.0}], familia="RED BULL"),
    dict(promo="336337", accion_id="3603", segmento="GASEOSAS", tipo="ADICIONAL", alcance="PEPSI Y 7UP AS", nombre="DCTO ADIC 5% PEPSI, 7UP 2L-1.5L AS", descuento_regla="Desde 15 bultos: 5%", canal="AS/MAY", calibre="PEPSI/7UP", grupo="2927", grupo_nombre="AS/MAY UNG TT PAIS", estado_grupo="Confirmado ERP", tiers=[{"min": 15.0, "pct": 5.0}], familia="GASEOSAS PEPSI / 7UP"),
    dict(promo="336336", accion_id="3604", segmento="GASEOSAS", tipo="ADICIONAL", alcance="PEPSI Y 7UP KT", nombre="DCTO ADIC 5% PEPSI, 7UP 2L-1.5L", descuento_regla="Desde 8 bultos: 5%", canal="K+T", calibre="PEPSI/7UP", grupo="2928", grupo_nombre="K+T UNG TT PAIS", estado_grupo="Confirmado ERP", tiers=[{"min": 8.0, "pct": 5.0}], familia="GASEOSAS PEPSI / 7UP"),
    dict(promo="336298", accion_id="3605", segmento="RED BULL", tipo="ESCALA", alcance="RED BULL 20% OTROS", nombre="RED BULL 20% OTROS", descuento_regla="Desde 0,16 bultos: 20%", canal="OTROS", calibre="RED BULL", grupo="3172", grupo_nombre="RED BULL 20% OTROS", estado_grupo="Confirmado ERP", tiers=[{"min": 0.16, "pct": 20.0}], familia="RED BULL"),
    dict(promo="336295", accion_id="3606", segmento="RED BULL", tipo="VOLUMEN", alcance="VOLUMEN RB AS/MAY", nombre="VOLUMEN RB AS/MAY", descuento_regla="Desde 50 bultos: 30%", canal="AS/MAY", calibre="RED BULL", grupo="3178", grupo_nombre="RED BULL VOLUMEN AS/MAY", estado_grupo="Confirmado ERP", tiers=[{"min": 50.0, "pct": 30.0}], familia="RED BULL"),
    dict(promo="336292", accion_id="3607", segmento="RED BULL", tipo="DROP", alcance="DROP RB AS", nombre="DROP RB AS", descuento_regla="Desde 0,5 bultos: 4%; 1: 6%; 2,5: 8%; 50: 0,01%", canal="AS", calibre="RED BULL", grupo="2929", grupo_nombre="AS UNG TT PAIS", estado_grupo="Confirmado ERP", tiers=[{"min": 0.5, "pct": 4.0}, {"min": 1.0, "pct": 6.0}, {"min": 2.5, "pct": 8.0}, {"min": 50.0, "pct": 0.01}], familia="RED BULL"),
    dict(promo="336286", accion_id="3608", segmento="RED BULL", tipo="DROP", alcance="DROP RB KT", nombre="DROP RB KT", descuento_regla="Desde 0,16 bultos: 4%; 0,66: 6%; 1: 8%; 9,99: 0,01%", canal="K+T", calibre="RED BULL", grupo="3177", grupo_nombre="UNG K+T (SIN BASE)", estado_grupo="Confirmado ERP", tiers=[{"min": 0.16, "pct": 4.0}, {"min": 0.66, "pct": 6.0}, {"min": 1.0, "pct": 8.0}, {"min": 9.99, "pct": 0.01}], familia="RED BULL"),
    dict(promo="244406", accion_id="3609", segmento="ISIDRA", tipo="BOTELLA", alcance="BOTELLA ISIDRA", nombre="BOTELLA ISIDRA", descuento_regla="Desde 1 bulto: 60%", canal="MH/MAY/REF", calibre="ISIDRA", grupo="3030", grupo_nombre="MH+MAY+REF ADY", estado_grupo="Confirmado ERP", tiers=[{"min": 1.0, "pct": 60.0}], familia="ISIDRA"),
    dict(promo="336396", accion_id="3610", segmento="0.0", tipo="0.0", alcance="DINAMICA 0.0 KT PATAGONIA", nombre="DINAMICA 0.0 KT PATAGONIA", descuento_regla="Desde 2 bultos: 5%; 5: 10%; 10: 15%", canal="K+T", calibre="STELLA 0.0", grupo="3019", grupo_nombre="K+T TOTAL PAIS", estado_grupo="Confirmado ERP", tiers=[{"min": 2.0, "pct": 5.0}, {"min": 5.0, "pct": 10.0}, {"min": 10.0, "pct": 15.0}], familia="CERVEZA 0.0"),
    dict(promo="336403", accion_id="3611", segmento="0.0", tipo="0.0", alcance="DINAMICA 0.0 AS PATAGONIA", nombre="DINAMICA 0.0 AS PATAGONIA", descuento_regla="Desde 5 bultos: 5%; 10: 10%; 15: 15%", canal="AS", calibre="STELLA 0.0", grupo="3079", grupo_nombre="AUTOSERVICIOS PAIS", estado_grupo="Confirmado ERP", tiers=[{"min": 5.0, "pct": 5.0}, {"min": 10.0, "pct": 10.0}, {"min": 15.0, "pct": 15.0}], familia="CERVEZA 0.0"),
    dict(promo="336472", accion_id="3612", segmento="VALUE", tipo="LATON", alcance="DCTO 7% LATON VALUE 1890 AS+MAY", nombre="DCTO 7% LATON VALUE 1890. AS+MAY", descuento_regla="Desde 96 bultos: 7%", canal="AS/MAY", calibre="1890 LATON", grupo="3179", grupo_nombre="AS + MAY SUR", estado_grupo="Confirmado ERP", tiers=[{"min": 96.0, "pct": 7.0}], familia="VALUE LATON 710"),
    dict(promo="336474", accion_id="3613", segmento="VALUE", tipo="LATON", alcance="DCTO 7% LATON VALUE 1890 KT", nombre="DCTO 7% LATON VALUE 1890. KT", descuento_regla="Desde 48 bultos: 7%", canal="K+T", calibre="1890 LATON", grupo="3019", grupo_nombre="K+T TOTAL PAIS", estado_grupo="Confirmado ERP", tiers=[{"min": 48.0, "pct": 7.0}], familia="VALUE LATON 710"),
]

EXCLUDED = [
    ("3565", "", "CMQ PEPSI BLACK 1500CC 19-10", "Sin beneficio ni promo de compañía"),
    ("3575", "330712", "ACCION BIG PIZZA - BR AO BUD", "Beneficios vacíos en ERP"),
    ("3576", "330718", "ACCION BIG PIZZA TOTAL ARG", "Beneficios vacíos en ERP"),
    ("3579", "327795", "11% 269 PALETA PATAGONIA K+T", "Beneficios vacíos en ERP"),
    ("3580", "320366", "11% 269 PALETA PATAGONIA AS", "Beneficios vacíos en ERP"),
    ("3583", "319057", "ESCALONADA QUILMES Y BRAHMA 269 CC MAYORISTAS", "Escalas y beneficios vacíos en ERP"),
]


def action_row(action: dict) -> dict:
    return {
        **{key: action[key] for key in ["promo", "accion_id", "segmento", "tipo", "alcance", "nombre", "descuento_regla", "canal", "calibre", "grupo", "grupo_nombre", "estado_grupo"]},
        "vigencia_excel": "06/11/2026",
        "tiers": json.dumps(action["tiers"], ensure_ascii=False),
        "familia": action["familia"],
    }


def main() -> None:
    audit = pd.read_csv(DATA / "action_audit.csv", dtype=str).fillna("")
    current = pd.read_excel(ERP_EXPORT)
    active_ids = set(current["Identificador"].astype(str))
    previous = pd.read_excel(PREVIOUS_ERP_EXPORT)
    previous_ids = set(previous["Identificador"].astype(str))
    audit = audit[audit["accion_id"].isin(active_ids)]
    new_ids = {action["accion_id"] for action in NEW_ACTIONS} | {item[0] for item in EXCLUDED}
    audit = audit[~audit["accion_id"].isin(new_ids)]
    audit = pd.concat([audit, pd.DataFrame([action_row(action) for action in NEW_ACTIONS])], ignore_index=True)
    excluded_rows = []
    for action_id, promo, name, reason in EXCLUDED:
        excluded_rows.append({
            "promo": promo, "accion_id": action_id, "segmento": "EXCLUIDA", "tipo": "EXCLUIDA",
            "alcance": reason, "nombre": name, "descuento_regla": "", "canal": "", "calibre": "",
            "grupo": "", "grupo_nombre": "", "estado_grupo": f"Excluida: {reason}",
            "vigencia_excel": "06/11/2026", "tiers": "[]", "familia": "EXCLUIDA",
        })
    excluded_rows = [row for row in excluded_rows if row["accion_id"] in active_ids]
    audit = pd.concat([audit, pd.DataFrame(excluded_rows)], ignore_index=True)
    audit.to_csv(DATA / "action_audit.csv", index=False, encoding="utf-8-sig")

    groups = {entry["gid"]: [str(client["codigo"]) for client in entry["clientes"]] for entry in json.loads(OLD_GROUPS.read_text(encoding="utf-8"))}
    capture = json.loads((DATA / "november_groups_capture.json").read_text(encoding="utf-8"))
    # Chess may briefly keep the previous group's single visible row while a
    # hash-route changes. Ignore those known partial captures and retain the
    # last complete snapshot instead.
    single_client_groups = {"2931", "3156", "3160", "3161"}
    capture = {
        group: clients
        for group, clients in capture.items()
        if group and (clients != ["1506"] or group in single_client_groups)
    }
    groups.update(capture)
    groups["3020"] = groups["3079"]
    groups["3169"] = [
        "100536", "100578", "100646", "101029", "101063", "101305", "109", "1734",
        "2240", "2412", "2671", "2745", "2962", "2985", "3113", "3212", "3218",
        "3293", "3333", "3335", "3511", "3526", "3717", "3853", "3884", "3968",
        "3970", "4030", "4146", "4217", "4270", "4341", "436", "4663", "4671",
        "4809", "4941", "4987", "5022", "5062", "5113", "5116", "5260", "5292",
        "5409", "5471", "906",
    ]

    master = pd.read_excel(CLIENTS, dtype={"Cliente": str}).fillna("")
    master["Cliente"] = master["Cliente"].astype(str).str.replace(r"\.0$", "", regex=True)
    master = master.drop_duplicates("Cliente").set_index("Cliente")
    groups["3170"] = master.index.astype(str).tolist()
    master_subcanal = master["Subcanal"].astype(str).str.replace(r"\.0$", "", regex=True)
    groups["2927"] = master.index[master_subcanal.isin("200 4 931".split())].astype(str).tolist()
    groups["2928"] = master.index[master_subcanal.isin("1 100 101 102 103 104 3 400 5 704 705 706 814 815 828 829 830 831 832 853 854 855 856 857 858 914 922 923 924 925 926 927 928 929 930 934 936".split())].astype(str).tolist()
    groups["2929"] = master.index[master_subcanal.eq("200")].astype(str).tolist()
    groups["3030"] = master.index[master_subcanal.isin("1 100 101 102 103 104 2 200 3 4 400 5 500 501 502 550 551 600 601 602 650 651 652 653 654 655 701 703 704 705 706 760 761 762 763 764 765 766 767 768 769 770 771 772".split())].astype(str).tolist()
    groups["3172"] = ["5109"]
    groups["3173"] = "100038 100156 100480 100902 101158 1962 2192 2407 3057 3156 3546 3583 3742 3941 3963 3982 4077 4267 4578 4606 4608 4629 4691 4745 4771 4889 4935 5016".split()
    groups["3174"] = ["3992"]
    groups["3175"] = "100550 100572 1506 2393 2532 2683 2764 2822 3015 3050 3206 3323 3400 3404 3440 3744 3785 3863 3965 4551 4975 5192 891 916 921".split()
    groups["3176"] = "1 10 100005 100009 100022 100034 100036 100044 100045 100056 100066 100072 100074 100123 100126 100129 100147 100148 100159 100162 100175 100176 100177 100181 100201 100216".split()
    groups["3178"] = "100330 100571 101255 1258 1830 1846 2164 2252 2501 2531 2674 2676 2733 2965 2968 3041 3134 3182 3221 3223 3224 3457 3631 4260 4337 4371 4788 5104 5232 5369 5411 5450 927".split()
    groups["3177"] = master.index[master_subcanal.isin("1 100 101 102 103 104 3 400 5 704 705 706 914 922 923 924 925 926 927 928 929 930".split())].astype(str).tolist()
    groups["3179"] = master.index[master_subcanal.isin("200 4 931 936".split())].astype(str).tolist()

    records = []
    confirmed = audit[audit["estado_grupo"].str.startswith("Confirmado")].copy()
    for _, action in confirmed.iterrows():
        group_id = str(action["grupo"]).replace(".0", "")
        if group_id not in groups:
            continue
        tiers = json.loads(action["tiers"] or "[]")
        if not tiers:
            continue
        best = max(tiers, key=lambda tier: float(tier["pct"]))
        detail = {
            "promo": action["promo"], "accion_id": action["accion_id"], "nombre": action["nombre"],
            "grupo": group_id, "grupo_nombre": action["grupo_nombre"], "estado_grupo": action["estado_grupo"],
            "descuento_regla": action["descuento_regla"],
        }
        for code in groups[group_id]:
            records.append({"codigo": str(code), "familia": action["familia"], "pct": float(best["pct"]), "desde": float(best["min"]), "detail": detail})

    rows = []
    frame = pd.DataFrame(records)
    for (code, family), block in frame.groupby(["codigo", "familia"], sort=False):
        info = master.loc[code] if code in master.index else pd.Series(dtype=object)
        if family == "CORE RGB / 473" or family == "CORONA EXCEPCION":
            pct = block["pct"].max()
            threshold = block.loc[block["pct"].idxmax(), "desde"]
        else:
            pct = block["pct"].sum()
            threshold = block["desde"].max()
        route_code = str(info.get("Fuerza de venta 1 Ruta de venta", "")).replace(".0", "")
        route_name = str(info.get("Fuerza de venta 1 Descripcion ruta de venta", ""))
        rows.append({
            "codigo": code,
            "fantasia": info.get("Nombre de fantasia", ""),
            "razon_social": info.get("Razon social", ""),
            "promotor": info.get("Fuerza de venta 1 Descripcion personal comercial", ""),
            "ruta": " - ".join(value for value in [route_code, route_name] if value),
            "subcanal": info.get("Descripcion subcanal", ""),
            "familia": family,
            "porcentaje_maximo": pct,
            "desde_bultos": threshold,
            "acciones": len(block),
            "detalle": json.dumps(block["detail"].tolist(), ensure_ascii=False),
            "tiene_inferidos": False,
        })
    output = pd.DataFrame(rows).sort_values(["codigo", "familia"])
    output.to_csv(DATA / "client_discounts.csv", index=False, encoding="utf-8-sig")

    meta = {
        "source_actions": str(ERP_EXPORT),
        "source_clients": "Google Drive/reporte de clientes.xlsx (26/10/2026)",
        "source_groups": "Chess ERP, relevamiento 06/11/2026",
        "actions_total": int(len(audit)),
        "actions_confirmed": int(audit["estado_grupo"].str.startswith("Confirmado").sum()),
        "actions_excluded": len(EXCLUDED),
        "actions_pending": 0,
        "clients_with_discounts": int(output["codigo"].nunique()),
        "rows": int(len(output)),
        "erp_actions_observed": int(len(current)),
        "erp_actions_added_since_previous": sorted(active_ids - previous_ids),
        "erp_actions_removed_since_previous": sorted(previous_ids - active_ids),
        "new_actions_excluded_pending_scope_confirmation": [item[0] for item in EXCLUDED],
    }
    (DATA / "snapshot_meta.json").write_text(json.dumps(meta, indent=2, ensure_ascii=False), encoding="utf-8")


if __name__ == "__main__":
    main()
