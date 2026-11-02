from __future__ import annotations

import json
from pathlib import Path

import pandas as pd


BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
OLD_GROUPS = Path(r"C:\Users\triesgo\Documents\Codex\2026-08-06\https-distribuidoradelvalle-chesserp-com-ar889-ventas\work\erp_groups_full_2026-10-22-all-actions.json")
CLIENTS = DATA / "reporte_clientes_drive_2026-10-26.xlsx"
ERP_EXPORT = Path(r"C:\Users\triesgo\Downloads\Acciones Comerciales (15).xlsx")


NEW_ACTIONS = [
    dict(promo="328365", accion_id="3558", segmento="MAYORISTA", tipo="1890", alcance="LATA 1890 MAYORISTA", nombre="6% DCTO EN LATA 1890 DESDE 100 BULTOS", descuento_regla="Desde 100 bultos: 6%", canal="MAY", calibre="1890 LATA", grupo="3120", grupo_nombre="MAYORISTA PAIS", estado_grupo="Confirmado ERP", tiers=[{"min": 100.0, "pct": 6.0}], familia="MAYORISTA 1890 LATA"),
    dict(promo="330535", accion_id="3572", segmento="CORE", tipo="CORE", alcance="EXCEPCION TOPE CORE", nombre="SUR- CORE K+T (TIER 3) EXCEPCION TOPE", descuento_regla="Desde 6 bultos: 4%; desde 10: 5%", canal="K+T", calibre="CORE", grupo="3154", grupo_nombre="EXCEPCION TOPE CORE SUR OCTUBRE KT", estado_grupo="Confirmado ERP", tiers=[{"min": 6.0, "pct": 4.0}, {"min": 10.0, "pct": 5.0}], familia="CORE RGB / 473"),
    dict(promo="331085", accion_id="3569", segmento="SABORIZADAS", tipo="DROP", alcance="DROP BAJO SABORIZADAS", nombre="DROP BAJO SABORIZADAS SUR", descuento_regla="Desde 2 bultos: 25%; desde 4: 50%", canal="SUR", calibre="H2OH", grupo="3070", grupo_nombre="DROP BAJO MS SUR", estado_grupo="Confirmado ERP", tiers=[{"min": 2.0, "pct": 25.0}, {"min": 4.0, "pct": 50.0}], familia="SABORIZADAS"),
    dict(promo="331087", accion_id="3570", segmento="SABORIZADAS", tipo="DROP", alcance="DROP ALTO SABORIZADAS", nombre="DROP ALTO SABORIZADAS SUR", descuento_regla="Desde 5 bultos: 25%; desde 10: 50%", canal="SUR", calibre="H2OH", grupo="3071", grupo_nombre="DROP ALTO MS SUR", estado_grupo="Confirmado ERP", tiers=[{"min": 5.0, "pct": 25.0}, {"min": 10.0, "pct": 50.0}], familia="SABORIZADAS"),
    dict(promo="330980", accion_id="3571", segmento="MAYORISTA", tipo="DROP", alcance="DROP SABORIZADAS MAYORISTA", nombre="DROP SABORIZADAS MAY NACIONAL", descuento_regla="Desde 30 bultos: 25%; desde 60: 50%", canal="MAY", calibre="H2OH", grupo="2958", grupo_nombre="MAY UNG TT PAIS", estado_grupo="Confirmado ERP", tiers=[{"min": 30.0, "pct": 25.0}, {"min": 60.0, "pct": 50.0}], familia="SABORIZADAS MAYORISTA"),
    dict(promo="331990", accion_id="3586", segmento="PREMIUM", tipo="EXCEPCION", alcance="CORONA 710 AS", nombre="EXCEPCION TOPE ESCALON EXTRASUR CORONA 710 AS", descuento_regla="Desde 40 bultos: 3%", canal="AS", calibre="CORONA 710", grupo="3160", grupo_nombre="GC-100331990", estado_grupo="Confirmado ERP", tiers=[{"min": 40.0, "pct": 3.0}], familia="CORONA EXCEPCION"),
    dict(promo="329242", accion_id="3587", segmento="PREMIUM", tipo="EXCEPCION", alcance="CORONA 710 Y 473 AS", nombre="SUR CORONA 710 Y 473 AS EXCEPCION TOPE", descuento_regla="Desde 3 bultos: 10%; desde 20: 12%", canal="AS", calibre="CORONA 710/473", grupo="3161", grupo_nombre="EXCEPCION TOPE CORONA OCTUBRE SUR", estado_grupo="Confirmado ERP", tiers=[{"min": 3.0, "pct": 10.0}, {"min": 20.0, "pct": 12.0}], familia="CORONA EXCEPCION"),
    dict(promo="307202", accion_id="3593", segmento="ENERGIZANTES", tipo="RED BULL", alcance="RED BULL TODOS", nombre="RED BULL 15%", descuento_regla="Desde 0,16 bultos: 15%", canal="TODOS", calibre="RED BULL", grupo="3156", grupo_nombre="RED BULL 15%", estado_grupo="Confirmado ERP", tiers=[{"min": 0.16, "pct": 15.0}], familia="RED BULL"),
    dict(promo="335362", accion_id="3594", segmento="0.0", tipo="0.0", alcance="DINAMICA 0.0 AS", nombre="DINAMICA 0.0 AS", descuento_regla="Desde 5 bultos: 5%; desde 10: 10%", canal="AS", calibre="0.0", grupo="3020", grupo_nombre="AS TOTAL PAIS CVZA", estado_grupo="Confirmado ERP", tiers=[{"min": 5.0, "pct": 5.0}, {"min": 10.0, "pct": 10.0}], familia="CERVEZA 0.0"),
    dict(promo="335363", accion_id="3595", segmento="0.0", tipo="0.0", alcance="DINAMICA 0.0 KT", nombre="DINAMICA 0.0 KT", descuento_regla="Desde 2 bultos: 5%; desde 5: 10%", canal="K+T", calibre="0.0", grupo="3019", grupo_nombre="K+T TOTAL PAIS", estado_grupo="Confirmado ERP", tiers=[{"min": 2.0, "pct": 5.0}, {"min": 5.0, "pct": 10.0}], familia="CERVEZA 0.0"),
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
        "vigencia_excel": "02/11/2026",
        "tiers": json.dumps(action["tiers"], ensure_ascii=False),
        "familia": action["familia"],
    }


def main() -> None:
    audit = pd.read_csv(DATA / "action_audit.csv", dtype=str).fillna("")
    current = pd.read_excel(ERP_EXPORT)
    active_ids = set(current["Identificador"].astype(str))
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
            "vigencia_excel": "02/11/2026", "tiers": "[]", "familia": "EXCLUIDA",
        })
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

    master = pd.read_excel(CLIENTS, dtype={"Cliente": str}).fillna("")
    master["Cliente"] = master["Cliente"].astype(str).str.replace(r"\.0$", "", regex=True)
    master = master.drop_duplicates("Cliente").set_index("Cliente")

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
        "source_groups": "Chess ERP, relevamiento 02/11/2026",
        "actions_total": int(len(audit)),
        "actions_confirmed": int(audit["estado_grupo"].str.startswith("Confirmado").sum()),
        "actions_excluded": len(EXCLUDED),
        "actions_pending": 0,
        "clients_with_discounts": int(output["codigo"].nunique()),
        "rows": int(len(output)),
        "erp_actions_observed": int(len(current)),
        "erp_actions_added_since_previous": sorted(new_ids),
        "erp_actions_removed_since_previous": ["3441", "3488", "3499"],
        "new_actions_excluded_pending_scope_confirmation": [item[0] for item in EXCLUDED],
    }
    (DATA / "snapshot_meta.json").write_text(json.dumps(meta, indent=2, ensure_ascii=False), encoding="utf-8")


if __name__ == "__main__":
    main()
