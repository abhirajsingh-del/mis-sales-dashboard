"""Writes a masked copy of an MIS workbook with formulas intact: every rupee and unit figure × MASK_FACTOR.

Works on the .xlsx XML directly, so every formula stays exactly as written and only the stored
numbers change: input values in amount/unit columns, and the cached results of formulas in those
columns (the formulas are linear, so the scaled result is what Excel recalculates to).
Rates, %, month codes, per-unit costs, MRPs and IDs are kept as they are.

Only the dumps the dashboard uses plus their lookup sheets are kept. Summary and pivot sheets and
every pivot cache are removed - pivot caches hold copies of the real figures.

Usage:  python pipeline/mask_workbook.py [source.xlsx] [dest.xlsx]
        defaults: the last workbook in input/  ->  output/<name>_masked.xlsx
Keep the result out of input/ - its name sorts after the original and would replace it.
"""
import codecs
import html
import posixpath
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config as C  # noqa: E402

# sheet -> (header row, default for numeric columns, exceptions)
#   "scale" sheets: every numeric column is scaled except the listed ones (a missed column is
#   over-scaled, never leaked). "keep" sheets: nothing is scaled except the listed ones.
# Exceptions are header names, (header, occurrence) where a name repeats, or a test function.
PER_UNIT = lambda h: h.startswith("p.u") or h.endswith("p.u")
SHEETS = {
    "Sales Dump": (3, "scale", {"Month", "Actual Month", "Margin %",
                                "RTO (Website)"}),  # holds two stray 202609 month codes, no amounts
    "Prod Dump": (4, "scale", {"Channel / PO Number", "Range", "Month-MIS", "Actual Month", "MRP", "Actual MRP",
                               "Tax_Rate (Report)", "Tax Rate Accounts", "Tax remove?", "Tax rate Diff",
                               "COGS % on GMV", "Discount %", "RK Discount %", ("Check", 0), PER_UNIT}),
    "Samples": (2, "scale", {"yearmonth1", "yearmonth2", "Final Month", PER_UNIT}),
    "Secondary Sales": (5, "scale", {"yearmonth", "Margin %", "Tax rate", "COGS % on GMV",
                                     ("Check", 0), ("Check", 2), PER_UNIT}),
    "COGS P.u": (3, "keep", set()),
    "updated COGS p.u": (2, "keep", set()),
    "Tax and Sku Change": (1, "keep", {"Sales total", "Sales 2 m"}),
    "All Product Mastersheet": (1, "keep", set()),
    "COMBO SKU CODE": (1, "keep", set()),
}

REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ROW = re.compile(r"<row\b[^>]*?(?:/>|>.*?</row>)", re.S)
ROW_NUM = re.compile(r'<row\b[^>]*?\br="(\d+)"')
CELL = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
T_ATTR = re.compile(r'\bt="(\w+)"')
V_TAG = re.compile(r"<v>([^<]*)</v>")
CHUNK = 8 << 20


def column_plan(header, default, exceptions):
    """True for each column that gets scaled."""
    seen, plan = Counter(), []
    for h in header:
        name = h.strip()
        occ = seen[name]
        seen[name] += 1
        hit = name in exceptions or (name, occ) in exceptions or any(callable(e) and e(name) for e in exceptions)
        plan.append(hit if default == "keep" else not hit)
    return plan


def col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n - 1


def shared_strings(z):
    xml = z.read("xl/sharedStrings.xml").decode("utf-8")
    out = []
    for m in re.finditer(r"<si(?:\s[^>]*)?(?:/>|>(.*?)</si>)", xml, re.S):
        out.append(html.unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", m.group(1) or "", re.S))))
    return out


def cell_text(attrs, inner, sst):
    t = T_ATTR.search(attrs)
    t = t.group(1) if t else "n"
    if t == "inlineStr":
        return html.unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", inner or "", re.S)))
    v = V_TAG.search(inner or "")
    if not v:
        return ""
    return sst[int(v.group(1))] if t == "s" else html.unescape(v.group(1))


class SheetMasker:
    """Rewrites one worksheet's rows; formulas pass through untouched."""

    def __init__(self, name, hr, default, exceptions, sst, factor, removed_ref):
        self.name, self.hr, self.default, self.exceptions = name, hr, default, exceptions
        self.sst, self.factor, self.removed_ref = sst, factor, removed_ref
        self.plan = None
        self.stats = Counter()

    def header(self, row_xml):
        cells = {col_index(m.group(1)): cell_text(m.group(3), m.group(4), self.sst) for m in CELL.finditer(row_xml)}
        names = [cells.get(i, "") for i in range(max(cells) + 1 if cells else 0)]
        flags = column_plan(names, self.default, self.exceptions)
        self.plan = {i for i, f in enumerate(flags) if f}
        self.names = names

    def row(self, m):
        xml = m.group(0)
        rn = ROW_NUM.search(xml)
        r = int(rn.group(1)) if rn else 0
        self.stats["formulas"] += xml.count("<f")
        for hit in self.removed_ref.findall(xml) if self.removed_ref else []:
            self.stats["refs to removed sheets"] += 1
            self.stats.setdefault("examples", set()).add(hit) if len(self.stats.get("examples", ())) < 5 else None
        if r == self.hr:
            self.header(xml)
            return xml
        if r < self.hr:
            return CELL.sub(self.above, xml)
        if self.plan is None:
            return xml
        return CELL.sub(self.data, xml)

    def above(self, m):
        """Rows above the header: keep labels and formulas, drop every stored number."""
        attrs, inner = m.group(3), m.group(4)
        if inner is None:
            return m.group(0)
        t = T_ATTR.search(attrs)
        t = t.group(1) if t else "n"
        if "<f" in inner:
            self.stats["above: formula values dropped"] += 1
            return f'<c r="{m.group(1)}{m.group(2)}"{T_ATTR.sub("", attrs)}>{V_TAG.sub("", inner)}</c>'
        if t in ("n", "e", "b"):
            self.stats["above: numbers dropped"] += 1
            return f'<c r="{m.group(1)}{m.group(2)}"{T_ATTR.sub("", attrs)}/>'
        return m.group(0)

    def data(self, m):
        attrs, inner = m.group(3), m.group(4)
        if inner is None or col_index(m.group(1)) not in self.plan:
            return m.group(0)
        t = T_ATTR.search(attrs)
        if t and t.group(1) != "n":
            return m.group(0)
        v = V_TAG.search(inner)
        if not v or not v.group(1):
            return m.group(0)
        new = repr(float(v.group(1)) * self.factor)
        self.stats["formula results scaled" if "<f" in inner else "values scaled"] += 1
        return f'<c r="{m.group(1)}{m.group(2)}"{attrs}>{inner[:v.start(1)]}{new}{inner[v.end(1):]}</c>'

    def stream(self, src, dst):
        dec = codecs.getincrementaldecoder("utf-8")()
        buf, phase = "", "pre"
        while True:
            chunk = src.read(CHUNK)
            final = not chunk
            buf += dec.decode(chunk, final)
            if phase == "pre":
                i = buf.find("<sheetData>")
                if i < 0:
                    if "<sheetData/>" in buf or final:
                        dst.write(buf.encode("utf-8")); buf, phase = "", "post"
                    if final:
                        break
                    continue
                dst.write(buf[:i + 11].encode("utf-8")); buf, phase = buf[i + 11:], "rows"
            if phase == "rows":
                end = buf.find("</sheetData>")
                upto = end if end >= 0 else buf.rfind("</row>") + 6 if "</row>" in buf else 0
                if upto:
                    dst.write(ROW.sub(self.row, buf[:upto]).encode("utf-8"))
                    buf = buf[upto:]
                if end >= 0:
                    phase = "post"
            if phase == "post":
                if self.removed_ref and self.removed_ref.search(buf):
                    self.stats["refs to removed sheets (validations/formatting)"] += len(self.removed_ref.findall(buf))
                dst.write(buf.encode("utf-8")); buf = ""
            if final:
                break


def main():
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else sorted(p for p in C.INPUT_DIR.glob("*.xlsx") if not p.name.startswith("~$"))[-1]
    dst = Path(sys.argv[2]) if len(sys.argv) > 2 else C.OUTPUT_DIR / f"{src.stem}_masked.xlsx"
    if dst.resolve().parent == C.INPUT_DIR.resolve():
        raise SystemExit("Write the masked copy outside input/ - the pipeline would read it as the newest file.")
    dst.parent.mkdir(exist_ok=True)
    print(f"Masking {src.name} x {C.MASK_FACTOR} -> {dst} (formulas kept)")

    z = zipfile.ZipFile(src)
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rel_target = {m.group(1): m.group(2) for m in re.finditer(r'<Relationship Id="([^"]+)"[^>]*?Target="([^"]+)"', rels)}
    rel_target.update({m.group(2): m.group(1) for m in re.finditer(r'<Relationship [^>]*?Target="([^"]+)"[^>]*?Id="([^"]+)"', rels)})
    sheets = [(m.group(0), html.unescape(m.group(1)), m.group(2)) for m in re.finditer(r'<sheet [^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"[^>]*/>', wb)]
    missing = set(SHEETS) - {n for _, n, _ in sheets}
    if missing:
        raise SystemExit(f"Sheets not found: {missing}")
    kept = [(tag, n, rid) for tag, n, rid in sheets if n in SHEETS]
    removed = [(tag, n, rid) for tag, n, rid in sheets if n not in SHEETS]
    part = lambda rid: "xl/" + rel_target[rid]

    # parts to drop: removed sheets and what only they use, plus every pivot table / cache
    drop = set()
    for _, n, rid in removed:
        p = part(rid)
        drop.add(p)
        rp = posixpath.join(posixpath.dirname(p), "_rels", posixpath.basename(p) + ".rels")
        if rp in z.namelist():
            drop.add(rp)
            for t in re.findall(r'Target="([^"]+)"[^>]*?(?<!External")/?>', z.read(rp).decode("utf-8")):
                if not t.startswith("http"):
                    q = posixpath.normpath(posixpath.join(posixpath.dirname(p), t))
                    drop.add(q)
                    qr = posixpath.join(posixpath.dirname(q), "_rels", posixpath.basename(q) + ".rels")
                    if qr in z.namelist():
                        drop.add(qr)
    drop |= {n for n in z.namelist() if n.startswith(("xl/pivotCache/", "xl/pivotTables/"))}
    kept_parts = {part(rid) for _, _, rid in kept}
    drop -= kept_parts

    # workbook.xml: sheets, pivot caches, sheet-local defined names (re-indexed)
    old_idx = {n: i for i, (_, n, _) in enumerate(sheets)}
    new_idx = {old_idx[n]: i for i, (_, n, _) in enumerate(kept)}
    for tag, _, _ in removed:
        wb = wb.replace(tag, "")
    wb = re.sub(r"<pivotCaches>.*?</pivotCaches>", "", wb, flags=re.S)

    def fix_name(m):
        loc = re.search(r'localSheetId="(\d+)"', m.group(0))
        if loc and int(loc.group(1)) not in new_idx:
            return ""
        if any(f"'{n}'!" in m.group(0) or f"{n}!" in m.group(0) for _, n, _ in removed):
            return ""
        return re.sub(r'localSheetId="(\d+)"', lambda x: f'localSheetId="{new_idx[int(x.group(1))]}"', m.group(0))
    wb = re.sub(r"<definedName\b[^>]*>.*?</definedName>", fix_name, wb, flags=re.S)
    wb = wb.replace("<definedNames></definedNames>", "")
    wb = re.sub(r'(<workbookView\b[^>]*?)\s(?:activeTab|firstSheet)="\d+"', r"\1", wb)

    drop_rids = {rid for _, _, rid in removed} | {m.group(1) for m in re.finditer(r'<Relationship Id="([^"]+)"[^>]*pivotCacheDefinition"', rels)} \
        | {m.group(2) for m in re.finditer(r'<Relationship [^>]*Type="[^"]*pivotCacheDefinition"[^>]*Id="([^"]+)"', rels)}
    rels = re.sub(r"<Relationship\b[^>]*/>", lambda m: "" if (re.search(r'Id="([^"]+)"', m.group(0)).group(1) in drop_rids
                                                                   or "pivotCacheDefinition" in m.group(0)) else m.group(0), rels)
    ct = z.read("[Content_Types].xml").decode("utf-8")
    ct = re.sub(r'<Override [^>]*PartName="/([^"]+)"[^>]*/>', lambda m: "" if m.group(1) in drop else m.group(0), ct)

    removed_names = [n for _, n, _ in removed]
    removed_ref = re.compile("|".join(re.escape(f"'{n}'!") for n in removed_names) + "|" +
                             "|".join(rf"(?<![\w'.]){re.escape(n)}!" for n in removed_names if re.fullmatch(r"\w+", n)))
    sst = shared_strings(z)
    by_part = {part(rid): n for _, n, rid in kept}

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED, allowZip64=True) as out:
        for info in z.infolist():
            n = info.filename
            if n in drop:
                continue
            if n == "xl/workbook.xml":
                out.writestr(info, wb)
            elif n == "xl/_rels/workbook.xml.rels":
                out.writestr(info, rels)
            elif n == "[Content_Types].xml":
                out.writestr(info, ct)
            elif n in by_part:
                name = by_part[n]
                hr, default, exc = SHEETS[name]
                sm = SheetMasker(name, hr, default, exc, sst, C.MASK_FACTOR, removed_ref)
                zi = zipfile.ZipInfo(n, info.date_time)
                zi.compress_type = zipfile.ZIP_DEFLATED
                with z.open(info) as s, out.open(zi, "w", force_zip64=True) as d:
                    sm.stream(s, d)
                scaled = [h for i, h in enumerate(sm.names) if i in sm.plan and h.strip()]
                print(f"\n[{name}] header row {hr}")
                print("  " + ", ".join(f"{k}: {v:,}" for k, v in sm.stats.items() if k != "examples"))
                print(f"  scaled columns: {', '.join(scaled) if scaled else '(none)'}")
                if sm.stats.get("examples"):
                    print(f"  formulas still pointing at removed sheets, e.g. {sorted(sm.stats['examples'])}")
            else:
                out.writestr(info, z.read(n))
    print(f"\nRemoved sheets: {', '.join(removed_names)}")
    print(f"Dropped {len(drop)} parts (pivot tables, pivot caches, removed sheets and their drawings)")
    print(f"Saved {dst}")


if __name__ == "__main__":
    main()
