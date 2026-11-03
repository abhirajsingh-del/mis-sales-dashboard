"""Builds template/MIS_Template_Next_Month.xlsx from the header rows of an existing MIS workbook.

Same sheets, same header rows and column names, with the Updated / Revised per-unit COGS
columns removed: from Oct-26 onward only the original (finalised) per-unit COGS is given.

Usage:  python pipeline/make_template.py [source.xlsx]
"""
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config as C  # noqa: E402

# Columns dropped per sheet: exact names, or everything from a given column onward.
DROP = {
    "Prod Dump": {"names": {"U comment", "U RM p.u", "U PM p.u", "U Shrink p.u", "U Shipper p.u", "U JW p.u",
                            "U Total p.u", "U check p.u", "U RM", "U PM", "U Shrink", "U Shipper", "U JW",
                            "U Total", "U PM Total", "U check"}},
    "Samples": {"names": {"Updated p.u", "Updated COGS", "Diff Of Value"}},
    "Secondary Sales": {"from": "p.u RM- Revised"},
}

INSTRUCTIONS = [
    ("How to use this template", None),
    ("1. Fill each dump sheet from its header row down (Sales Dump row 4, Prod Dump row 5, Samples row 3, Secondary Sales row 6).", None),
    ("2. Keep the header names exactly as they are. Extra columns are fine; renamed or missing columns stop the pipeline.", None),
    ("3. Values or formulas both work. If you use formulas, save the file in Excel / Google Sheets so the results are stored.", None),
    ("4. Save in input/ as MIS_YYYY_MM.xlsx (e.g. MIS_2026_10.xlsx). Files are read in name order; if two files hold the same month, the later name wins.", None),
    ("5. Run:  python pipeline/pipeline.py   then check output/data_health.txt (all checks should say PASS).", None),
    ("", None),
    ("Columns the dashboard reads", None),
    ("Sales Dump", "Dump Type, Type, Platform, Type (2nd = B2B/B2C), Month, GMV, Discount Value, Dealer Margin, RTO/Cancellation, Final Sales, Actual GMV, MIS Brand, MIS Channel"),
    ("Prod Dump", "Dump Type, Listing_SKU, Item SKU Code, Item Type Name, Type, Category, Sub_category, Channel (2nd), Updated GMV, Updated Sales, Units, Month-MIS (yyyymm), MIS Brand, MIS Channel, MIS Tax Amount, RM, PM, Shrink, Shipper, JW"),
    ("Samples", "yearmonth1, Final Month, Channel, Split_SKu, sum(Units), Type, MIS Brand, MIS Channel, Value"),
    ("Secondary Sales", "Split SKU Code, SKU Name, Channel (3rd), Updated GMV, Units, yearmonth, Sub category, MIS Brand, MIS Channel, Margin, Tax value, RM, PM, Shipper, Shrink, JW"),
    ("", None),
    ("Rules the pipeline checks", None),
    ("Sales Dump", "Final Sales = GMV - Discount Value - Dealer Margin - RTO/Cancellation;  Actual GMV = GMV - RTO/Cancellation"),
    ("Sales vs Prod", "Per month x Dump Type x MIS Brand x MIS Channel: Sales Dump Actual GMV = Prod Dump Updated GMV and Final Sales = Updated Sales (within Rs 100)"),
    ("Prod Dump COGS", "RM / PM / Shipper / Shrink / JW = Units x per-unit cost (single finalised per-unit sheet, COGS P.u)"),
    ("Prod Dump tax", "MIS Tax Amount = Updated Sales / (100 + GST %) x GST %"),
    ("Samples", "Value = sum(Units) x p.u"),
    ("Secondary Sales", "Margin = Updated GMV x Margin %;  Tax value = (Updated GMV - Margin) / (100 + Tax rate) x Tax rate"),
    ("Month", "Any of: a date in the month, yyyymm (202610), or text like Oct-26"),
    ("MIS Brand", "mCaffeine / Hyphen / DND (any capitalisation)"),
    ("Dump Type", "Internal, Reco or Others"),
]


def main():
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else sorted(p for p in C.INPUT_DIR.glob("*.xlsx") if not p.name.startswith("~$"))[-1]
    wb = openpyxl.load_workbook(src, read_only=True)
    out = openpyxl.Workbook()
    ins = out.active
    ins.title = "Instructions"
    bold = Font(bold=True)
    for i, (a, b) in enumerate(INSTRUCTIONS, 1):
        ins.cell(i, 1, a)
        if b:
            ins.cell(i, 2, b)
            ins.cell(i, 1).font = bold
        elif a and not a[0].isdigit():
            ins.cell(i, 1).font = Font(bold=True, size=12)
    ins.column_dimensions["A"].width = 22
    ins.column_dimensions["B"].width = 140

    head_fill = PatternFill("solid", fgColor="DDEBF7")
    sheets = list(C.SHEETS) + ["COGS P.u"]
    header_rows = {**{s: h for s, (h, _) in C.SHEETS.items()}, "COGS P.u": 3}
    for sheet in sheets:
        hr = header_rows[sheet]
        header = next(wb[sheet].iter_rows(min_row=hr, max_row=hr, values_only=True))
        header = list(header)
        while header and header[-1] is None:
            header.pop()
        rule = DROP.get(sheet, {})
        names = [str(h).strip() if h else "" for h in header]
        if "from" in rule and rule["from"] in names:  # absent when the source is already a template-format file
            header = header[:names.index(rule["from"])]
            while header and (header[-1] is None or str(header[-1]).strip() == ""):
                header.pop()
        header = [h for h in header if str(h).strip() not in rule.get("names", set())]
        ws = out.create_sheet(sheet)
        ws.cell(1, 1, f"Header is row {hr}. Paste data from row {hr + 1}. Rows above the header are ignored.")
        ws.cell(1, 1).font = Font(italic=True, color="666666")
        for j, h in enumerate(header, 1):
            c = ws.cell(hr, j, h)
            c.font, c.fill = bold, head_fill
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.column_dimensions[c.column_letter].width = max(10, min(28, len(str(h or "")) + 3))
        ws.freeze_panes = ws.cell(hr + 1, 1)
        print(f"{sheet}: {len(header)} columns")

    dest = C.ROOT / "template" / "MIS_Template_Next_Month.xlsx"
    dest.parent.mkdir(exist_ok=True)
    out.save(dest)
    print(f"Saved {dest}")


if __name__ == "__main__":
    main()
