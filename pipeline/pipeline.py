"""MIS dashboard pipeline.

Reads every workbook in input/, takes the Sales Dump, Prod Dump, Samples and Secondary Sales sheets,
normalises brands/channels, reconciles Sales vs Prod, masks values and writes:
  docs/data.js                      masked data for the hosted dashboard (GitHub Pages)
  output/dashboard_standalone.html  same dashboard with data inlined (single file to share)
  output/dashboard_actual.html      same, with ACTUAL (unmasked) values - local only, never published
  output/data_health.txt            reconciliation report
  data/*.csv                        unmasked normalised rows (local only)

Usage:  python pipeline/pipeline.py
"""
import json
import re
import sys
from datetime import datetime, date
from pathlib import Path

import openpyxl
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config as C  # noqa: E402

SALES_MEASURES = ["gmv", "discount", "dealer_margin", "rto", "net_gmv", "final_sales"]
PROD_MEASURES = ["net_gmv", "final_sales", "tax", "units", "rm", "pm", "shrink", "shipper", "jw"]
SAMPLE_MEASURES = ["units", "cogs"]
SEC_MEASURES = ["gmv", "units", "margin", "tax", "rm", "pm", "shrink", "shipper", "jw"]
COGS_SOURCE = {}  # (file, sheet) -> COGS columns actually read, reported in data_health.txt
MONTH_NAMES = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}


# ---------------------------------------------------------------- extract

def to_month(v):
    """Return 'YYYY-MM' from a date, a yyyymm number, or text like 'Jul-26'."""
    if v is None or v == "" or (not isinstance(v, str) and pd.isna(v)):
        return None
    if isinstance(v, (datetime, date)):
        return f"{v.year}-{v.month:02d}"
    if isinstance(v, (int, float)):
        n = int(v)
        if 190001 <= n <= 299912:
            return f"{n // 100}-{n % 100:02d}"
        return None
    s = str(v).strip()
    m = re.match(r"^(\d{4})[-/]?(\d{2})", s)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    m = re.match(r"^([A-Za-z]{3})[A-Za-z]*[-' ]+(\d{2,4})$", s)
    if m and m.group(1).lower() in MONTH_NAMES:
        y = int(m.group(2))
        y = y + 2000 if y < 100 else y
        return f"{y}-{MONTH_NAMES[m.group(1).lower()]:02d}"
    return None


def read_sheet(wb, sheet, header_row, fields):
    ws = wb[sheet]
    rows = ws.iter_rows(min_row=header_row, values_only=True)
    header = [str(h).strip() if h is not None else "" for h in next(rows)]
    idx, used = {}, {}
    for field, spec in fields.items():
        for name, occ in (spec if isinstance(spec, list) else [spec]):
            hits = [i for i, h in enumerate(header) if h == name]
            if len(hits) > occ:
                idx[field], used[field] = hits[occ], name
                break
        else:
            raise SystemExit(f"[{sheet}] none of the columns {spec} found in header row {header_row}")
    names, cols = list(idx), list(idx.values())
    data = [[r[c] if c < len(r) else None for c in cols] for r in rows]
    return pd.DataFrame(data, columns=names), used


def extract_file(path):
    print(f"Reading {path.name} ...", flush=True)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = {}
    for sheet, (hdr, fields) in C.SHEETS.items():
        if sheet not in wb.sheetnames:
            raise SystemExit(f"{path.name}: sheet '{sheet}' missing")
        df, used = read_sheet(wb, sheet, hdr, fields)
        cogs_cols = [used[f] for f in ("rm", "pm", "shrink", "shipper", "jw", "cogs") if f in used]
        COGS_SOURCE[(path.name, sheet)] = ", ".join(cogs_cols)
        if "month_fallback" in df:
            df["month"] = df["month"].where(df["month"].notna(), df["month_fallback"])
            df = df.drop(columns="month_fallback")
        df["month"] = df["month"].map(to_month)
        df["source_file"] = path.name
        out[sheet] = df
        print(f"  {sheet}: {len(df):,} rows", flush=True)
    wb.close()
    return out


def load_inputs():
    # Sorted by name (MIS_YYYY_MM...), not modified time: a git checkout gives every file the
    # same mtime, so name order is the only rule that behaves the same locally and on GitHub.
    files = sorted(p for p in C.INPUT_DIR.glob("*.xlsx") if not p.name.startswith("~$"))
    if not files:
        raise SystemExit(f"No .xlsx files in {C.INPUT_DIR}")
    extracted, owner = [], {}
    for f in files:  # later name takes over any month it shares with an earlier file
        e = extract_file(f)
        extracted.append((f, e))
        for df in e.values():
            for m in df["month"].dropna().unique():
                owner[m] = f.name
    merged = {}
    for sheet in C.SHEETS:
        parts = [e[sheet][e[sheet]["month"].map(owner) == f.name] for f, e in extracted]
        merged[sheet] = pd.concat(parts, ignore_index=True)
    return merged, owner


# ---------------------------------------------------------------- normalise

def norm_text(v):
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()


CANON_CHANNEL = {k.lower(): k for k in C.CHANNEL_GROUPS}
CANON_CHANNEL.update(C.CHANNEL_ALIASES)


def norm_brand(v):
    return C.BRANDS.get(norm_text(v).lower(), norm_text(v) or "Unknown")


def norm_channel(v):
    s = norm_text(v)
    return CANON_CHANNEL.get(s.lower(), s or "Unknown")


def normalise(merged):
    unmapped = {}
    out = {}
    for sheet, df in merged.items():
        df = df.copy()
        dropped = int(df["month"].isna().sum())
        df = df[df["month"].notna()]
        measures = {"Sales Dump": SALES_MEASURES, "Prod Dump": PROD_MEASURES,
                    "Samples": SAMPLE_MEASURES, "Secondary Sales": SEC_MEASURES}[sheet]
        for m in measures:
            df[m] = pd.to_numeric(df[m], errors="coerce").fillna(0.0)
        if "dump_type" in df:
            df["dump_type"] = df["dump_type"].map(norm_text).replace("", "Internal")
        else:
            df["dump_type"] = "Internal"
        df["brand_raw"], df["channel_raw"] = df["brand"], df["channel"]
        df["brand"] = df["brand"].map(norm_brand)
        df["channel"] = df["channel"].map(norm_channel)
        df["channel_group"] = df["channel"].map(C.CHANNEL_GROUPS).fillna("Other")
        bad_b = sorted(set(df.loc[~df["brand"].isin(C.BRANDS.values()), "brand"]))
        bad_c = sorted(set(df.loc[~df["channel"].isin(C.CHANNEL_GROUPS), "channel"]))
        unmapped[sheet] = {"dropped_rows_without_month": dropped, "unmapped_brands": bad_b, "unmapped_channels": bad_c}
        if sheet == "Prod Dump":
            df["sku"] = df["sku"].map(norm_text)
            df["sku_name"] = df["sku_name"].map(norm_text)
            for col in ["category", "sub_category", "sku_type"]:
                df[col] = df[col].map(norm_text)
                missing = df[col].str.lower().isin(C.UNMAPPED_CATEGORY)
                # adjustment/return rows often lack category: borrow it from the same SKU elsewhere
                known = df[~missing & (df["sku"] != "")].groupby("sku")[col].agg(lambda x: x.value_counts().index[0])
                df.loc[missing, col] = df.loc[missing, "sku"].map(known)
                df[col] = df[col].fillna("Unmapped")
            df["sub_category"] = df["sub_category"].str.strip().str.title()
            df["cogs"] = df[["rm", "pm", "shrink", "shipper", "jw"]].sum(axis=1)
        if sheet == "Secondary Sales":
            df["sku"] = df["sku"].map(norm_text)
            df["sku_name"] = df["sku_name"].map(norm_text)
            df["sub_category"] = df["sub_category"].map(norm_text).str.title().replace("", "Unmapped")
            df["cogs"] = df[["rm", "pm", "shrink", "shipper", "jw"]].sum(axis=1)
        if sheet == "Samples":
            df["sample_type"] = df["sample_type"].map(norm_text).replace("", "Samples")
        out[sheet] = df
    return out, unmapped


# ---------------------------------------------------------------- validate

def validate(d, unmapped, owner):
    sd, pdd, sm = d["Sales Dump"], d["Prod Dump"], d["Samples"]
    checks = []

    def add(name, ok, detail):
        checks.append({"check": name, "ok": bool(ok), "detail": detail})

    e1 = (sd["gmv"] - sd["discount"] - sd["dealer_margin"] - sd["rto"] - sd["final_sales"]).abs().max()
    e2 = (sd["gmv"] - sd["rto"] - sd["net_gmv"]).abs().max()
    add("Sales Dump: Final Sales = GMV - Discount - Dealer Margin - RTO", e1 <= 1, f"max row diff Rs {e1:,.2f}")
    add("Sales Dump: Actual GMV = GMV - RTO", e2 <= 1, f"max row diff Rs {e2:,.2f}")

    keys = ["month", "dump_type", "brand", "channel"]
    a = sd.groupby(keys)[["net_gmv", "final_sales"]].sum()
    b = pdd.groupby(keys)[["net_gmv", "final_sales"]].sum()
    j = a.join(b, how="outer", lsuffix="_sales", rsuffix="_prod").fillna(0)
    j["d_gmv"] = j["net_gmv_sales"] - j["net_gmv_prod"]
    j["d_sales"] = j["final_sales_sales"] - j["final_sales_prod"]
    bad = j[(j["d_gmv"].abs() > C.RECON_TOLERANCE) | (j["d_sales"].abs() > C.RECON_TOLERANCE)]
    add("Sales Dump (GMV - RTO) = Prod Dump Updated GMV, per month x dump type x brand x channel",
        bad["d_gmv"].abs().le(C.RECON_TOLERANCE).all(),
        f"{len(j)} cells checked, {int((j['d_gmv'].abs() > C.RECON_TOLERANCE).sum())} off by more than Rs {C.RECON_TOLERANCE:g}")
    add("Sales Dump Final Sales = Prod Dump Updated Sales, per month x dump type x brand x channel",
        bad["d_sales"].abs().le(C.RECON_TOLERANCE).all(),
        f"{len(j)} cells checked, {int((j['d_sales'].abs() > C.RECON_TOLERANCE).sum())} off by more than Rs {C.RECON_TOLERANCE:g}")

    months = sorted(owner)
    missing = {s: [m for m in months if m not in set(df["month"])] for s, df in d.items()}
    add("Every month present in all four dumps", not any(missing.values()),
        "; ".join(f"{s} missing {v}" for s, v in missing.items() if v) or f"months: {', '.join(months)}")

    for s, u in unmapped.items():
        add(f"{s}: brands and channels mapped", not u["unmapped_brands"] and not u["unmapped_channels"],
            f"unmapped brands {u['unmapped_brands'] or 'none'}; unmapped channels {u['unmapped_channels'] or 'none'}; "
            f"rows dropped for no month: {u['dropped_rows_without_month']}")

    nocogs = pdd[(pdd["units"] > 0) & (pdd["cogs"] == 0) & (pdd["net_gmv"] > 0)]
    share = nocogs["net_gmv"].sum() / max(pdd.loc[pdd["net_gmv"] > 0, "net_gmv"].sum(), 1)
    add("Prod Dump: rows with units and GMV but no COGS (warning only)", True,
        f"{len(nocogs):,} rows, {share:.2%} of positive Net GMV")
    unm = pdd.loc[pdd["category"] == "Unmapped", "net_gmv"].abs().sum() / max(pdd["net_gmv"].abs().sum(), 1)
    add("Prod Dump: Net GMV with no category (warning only)", True,
        f"{unm:.2%} of absolute Net GMV ({int((pdd['category'] == 'Unmapped').sum()):,} rows)")
    for (f, sheet), cols in sorted(COGS_SOURCE.items()):
        if cols:
            add(f"COGS columns used: {sheet} in {f} (info)", True, cols)

    hard_fail = [c for c in checks if not c["ok"] and "warning" not in c["check"]]
    return checks, j.reset_index(), hard_fail


# ---------------------------------------------------------------- aggregate + mask

class Dict:
    """Dictionary-encodes a dimension so the JSON stays small."""

    def __init__(self):
        self.values, self.index = [], {}

    def __call__(self, v):
        if v not in self.index:
            self.index[v] = len(self.values)
            self.values.append(v)
        return self.index[v]


def cube(df, dims, measures, enc, factor, digits=0):
    g = df.groupby(dims, dropna=False)[measures].sum().reset_index()
    g = g[(g[measures].abs() > 0.005).any(axis=1)]
    rows = []
    for rec in g.itertuples(index=False):
        r = [enc[dim](getattr(rec, dim)) for dim in dims]
        r += [round(float(getattr(rec, m)) * factor, digits) for m in measures]
        rows.append(r)
    return {"dims": dims, "measures": measures, "rows": rows}


def build_payload(d, checks, recon, factor):
    """factor = C.MASK_FACTOR for anything shared; 1.0 for the local actual-values copy."""
    sd, pdd, sm, sec = d["Sales Dump"], d["Prod Dump"], d["Samples"], d["Secondary Sales"]
    enc = {k: Dict() for k in ["month", "dump_type", "brand", "channel", "category", "sub_category", "sku", "sample_type", "channel_group"]}

    keys = ["month", "dump_type", "brand", "channel"]
    s = sd.groupby(keys)[SALES_MEASURES].sum()
    p = pdd.groupby(keys)[["tax", "units", "rm", "pm", "shrink", "shipper", "jw"]].sum()
    pnl = s.join(p, how="outer").fillna(0).reset_index()
    pnl_cube = cube(pnl, keys, SALES_MEASURES + ["tax", "units", "rm", "pm", "shrink", "shipper", "jw"], enc, factor)

    cat_cube = cube(pdd, keys + ["category", "sub_category"],
                    ["net_gmv", "final_sales", "tax", "units", "rm", "pm", "shrink", "shipper", "jw"], enc, factor)

    named = pd.concat([pdd[["sku", "sku_name"]], sec[["sku", "sku_name"]]])
    names = named[named["sku_name"] != ""].groupby("sku")["sku_name"].agg(lambda x: x.value_counts().index[0])
    sku_cube = cube(pdd, ["month", "dump_type", "brand", "channel_group", "sku"],
                    ["net_gmv", "final_sales", "tax", "units", "rm", "pm", "shrink", "shipper", "jw"], enc, factor)
    smp_cube = cube(sm, ["month", "brand", "channel", "sample_type"], ["units", "cogs"], enc, factor)
    sec_cube = cube(sec, ["month", "brand", "channel", "sku"], SEC_MEASURES, enc, factor)

    # name and category for every SKU code seen in Prod Dump or Secondary Sales
    top = lambda x: x.value_counts().index[0]
    cats = pdd.groupby("sku")["category"].agg(top)
    subs = pd.concat([pdd.groupby("sku")["sub_category"].agg(top), sec.groupby("sku")["sub_category"].agg(top)])
    subs = subs[~subs.index.duplicated()]
    sku_info = [[names.get(k, ""), cats.get(k, "Unmapped"), subs.get(k, "Unmapped")] for k in enc["sku"].values]

    channel_group = [C.CHANNEL_GROUPS.get(c, "Other") for c in enc["channel"].values]
    months = sorted(set(sd["month"]) | set(pdd["month"]) | set(sm["month"]) | set(sec["month"]))

    health = {
        "checks": checks,
        "recon_cells": int(len(recon)),
        "recon_max_abs_diff_gmv": round(float(recon["d_gmv"].abs().max()) * factor, 2),
        "recon_max_abs_diff_sales": round(float(recon["d_sales"].abs().max()) * factor, 2),
        "rows": {"sales_dump": int(len(sd)), "prod_dump": int(len(pdd)), "samples": int(len(sm)), "secondary": int(len(sec))},
    }
    return {
        "meta": {
            "generated": datetime.now().strftime("%d %b %Y %H:%M"),
            "built_at": datetime.now().isoformat(timespec="minutes"),
            "months": months,
            "headline_dump_types": C.HEADLINE_DUMP_TYPES,
            "masked": factor != 1.0,
        },
        "dicts": {k: v.values for k, v in enc.items()},
        "channel_group_of": channel_group,
        "sku_info": sku_info,
        "pnl": pnl_cube,
        "cat": cat_cube,
        "sku": sku_cube,
        "samples": smp_cube,
        "secondary": sec_cube,
        "health": health,
    }


# ---------------------------------------------------------------- write

def write_outputs(d, payload, actual, checks, recon, owner):
    C.DATA_DIR.mkdir(exist_ok=True)
    C.OUTPUT_DIR.mkdir(exist_ok=True)
    for sheet, df in d.items():
        df.to_csv(C.DATA_DIR / (sheet.lower().replace(" ", "_") + ".csv"), index=False)
    recon.to_csv(C.DATA_DIR / "recon_sales_vs_prod.csv", index=False)

    js = "window.MIS_DATA = " + json.dumps(payload, separators=(",", ":"), default=str) + ";\n"
    (C.DOCS_DIR / "data.js").write_text(js, encoding="utf-8")

    # single-file copies: masked one to share, actual one for you only (output/ is git-ignored)
    page = (C.DOCS_DIR / "index.html").read_text(encoding="utf-8")
    tag = '<script src="data.js"></script>'
    inline = lambda p: page.replace(tag, "<script>window.MIS_DATA = " + json.dumps(p, separators=(",", ":"), default=str).replace("</", "<\\/") + ";</script>")
    if tag in page:
        (C.OUTPUT_DIR / "dashboard_standalone.html").write_text(inline(payload), encoding="utf-8")
        (C.OUTPUT_DIR / "dashboard_actual.html").write_text(inline(actual), encoding="utf-8")

    lines = [f"MIS pipeline run {payload['meta']['generated']}", "",
             "Months and source file:"] + [f"  {m}: {f}" for m, f in sorted(owner.items())] + [""]
    for c in checks:
        lines.append(f"[{'PASS' if c['ok'] else 'FAIL'}] {c['check']}\n        {c['detail']}")
    (C.OUTPUT_DIR / "data_health.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("\n".join(lines))
    print(f"\nWrote docs/data.js ({len(js) / 1e6:.2f} MB), output/dashboard_standalone.html, output/data_health.txt")
    print("Actual (unmasked) values: output/dashboard_actual.html - local only, do not share")


def main():
    merged, owner = load_inputs()
    d, unmapped = normalise(merged)
    checks, recon, hard_fail = validate(d, unmapped, owner)
    payload = build_payload(d, checks, recon, C.MASK_FACTOR)
    actual = build_payload(d, checks, recon, 1.0)
    write_outputs(d, payload, actual, checks, recon, owner)
    if hard_fail:
        print(f"\n{len(hard_fail)} check(s) FAILED - review output/data_health.txt before publishing.")
        sys.exit(1)


if __name__ == "__main__":
    main()
