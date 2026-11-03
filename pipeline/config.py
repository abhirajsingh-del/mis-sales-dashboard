"""Settings for the MIS dashboard pipeline. Edit here, not in pipeline.py."""
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
INPUT_DIR = ROOT / "input"     # monthly MIS workbooks (.xlsx) - never published
DATA_DIR = ROOT / "data"       # unmasked normalised CSVs - local only, git-ignored
DOCS_DIR = ROOT / "docs"       # GitHub Pages site - masked data only
OUTPUT_DIR = ROOT / "output"   # health report + standalone single-file dashboard

# Every rupee and unit value is multiplied by this before it leaves the pipeline.
# Ratios (%) are unaffected because numerator and denominator are scaled together.
MASK_FACTOR = 0.75

# Dump types counted in the headline numbers (matches the MIS Sales Pivot).
# "Others" (B2B adjustment rows) can be switched on in the dashboard.
HEADLINE_DUMP_TYPES = ["Internal", "Reco"]

# Reconciliation tolerance in rupees per month x brand x channel cell
# (source sheets carry small rounding gaps, e.g. Rs 67 on a Rs 1.9 Cr cell).
RECON_TOLERANCE = 100.0

# Sheet name -> (header row, {field: (header text, occurrence)}).
# Occurrence is 0 for the first column with that header, 1 for the second, etc.
# A list of specs means "first one found": JAS-26 carries finalised Updated/Revised COGS
# columns; later months carry only the original per-unit COGS, so they fall back to it.
SHEETS = {
    "Sales Dump": (3, {
        "dump_type": ("Dump Type", 0),
        "platform_type": ("Type", 0),
        "platform": ("Platform", 0),
        "b2x": ("Type", 1),
        "month": ("Month", 0),
        "gmv": ("GMV", 0),
        "discount": ("Discount Value", 0),
        "dealer_margin": ("Dealer Margin", 0),
        "rto": ("RTO/Cancellation", 0),
        "final_sales": ("Final Sales", 0),
        "net_gmv": ("Actual GMV", 0),
        "brand": ("MIS Brand", 0),
        "channel": ("MIS Channel", 0),
    }),
    "Prod Dump": (4, {
        "dump_type": ("Dump Type", 0),
        "listing_sku": ("Listing_SKU", 0),
        "sku": ("Item SKU Code", 0),
        "sku_name": ("Item Type Name", 0),
        "sku_type": ("Type", 0),
        "category": ("Category", 0),
        "sub_category": ("Sub_category", 0),
        "platform": ("Channel", 1),
        "net_gmv": ("Updated GMV", 0),
        "final_sales": ("Updated Sales", 0),
        "units": ("Units", 0),
        "month": ("Month-MIS", 0),
        "brand": ("MIS Brand", 0),
        "channel": ("MIS Channel", 0),
        "tax": ("MIS Tax Amount", 0),
        "rm": [("U RM", 0), ("RM", 0)],
        "pm": [("U PM", 0), ("PM", 0)],
        "shrink": [("U Shrink", 0), ("Shrink", 0)],
        "shipper": [("U Shipper", 0), ("Shipper", 0)],
        "jw": [("U JW", 0), ("JW", 0)],
    }),
    "Samples": (2, {
        "month": ("Final Month", 0),
        "month_fallback": ("yearmonth1", 0),
        "platform": ("Channel", 0),
        "sku": ("Split_SKu", 0),
        "units": ("sum(Units)", 0),
        "sample_type": ("Type", 0),
        "brand": ("MIS Brand", 0),
        "channel": ("MIS Channel", 0),
        "cogs": [("Updated COGS", 0), ("Value", 0)],
    }),
    # Sell-through reported by B2B partners (Nykaa, Blinkit, Zepto ...). Shown on its own
    # page and never added to primary totals, so B2B sales are not counted twice.
    "Secondary Sales": (5, {
        "sku": ("Split SKU Code", 0),
        "sku_name": ("SKU Name", 0),
        "platform": ("Channel", 2),
        "gmv": ("Updated GMV", 0),
        "units": ("Units", 0),
        "month": ("yearmonth", 0),
        "brand": ("MIS Brand", 0),
        "channel": ("MIS Channel", 0),
        "sub_category": ("Sub category", 0),
        "margin": ("Margin", 0),
        "tax": ("Tax value", 0),
        "rm": [("RM-Revised", 0), ("RM", 0)],
        "pm": [("PM-Revised", 0), ("PM", 0)],
        "shipper": [("Shipper-Revised", 0), ("Shipper", 0)],
        "shrink": [("Shrink-Revised", 0), ("Shrink", 0)],
        "jw": [("JW-Revised", 0), ("JW", 0)],
    }),
}

BRANDS = {"hyphen": "Hyphen", "mcaffeine": "mCaffeine", "dnd": "DND"}

# Canonical channel -> channel group. Matching is case-insensitive.
CHANNEL_GROUPS = {
    "Website": "D2C",
    "Amazon": "Marketplace", "Flipkart": "Marketplace", "Myntra": "Marketplace",
    "Nykaa": "Marketplace", "Purplle": "Marketplace", "Meesho": "Marketplace",
    "Snapdeal": "Marketplace", "Tira": "Marketplace", "Tata 1MG": "Marketplace",
    "HYPD": "Marketplace", "First Cry B2C": "Marketplace", "Smytten B2C": "Marketplace",
    "Slikk": "Marketplace", "First Club": "Marketplace",
    "Blinkit": "Quick Commerce", "Zepto": "Quick Commerce", "Swiggy Insta": "Quick Commerce",
    "Flipkart Qcom": "Quick Commerce", "Big Basket": "Quick Commerce",
    "Amazon B2B": "B2B / Modern Trade", "First Cry B2B": "B2B / Modern Trade",
    "Dmart": "B2B / Modern Trade", "Pop Tech": "B2B / Modern Trade",
    "Gluzo": "B2B / Modern Trade", "Smytten B2B": "B2B / Modern Trade",
    "Noble Chemist": "B2B / Modern Trade",
    "Institutional": "Institutional",
}

# Extra spellings seen in the dumps -> canonical channel.
CHANNEL_ALIASES = {
    "dnc-blinkit": "Blinkit",
    "dnc-myntra or": "Myntra",
    "swiggy": "Swiggy Insta",
    "bigbasket": "Big Basket",
    # Secondary Sales spellings
    "firstcry b2b": "First Cry B2B",
    "rk world": "Amazon B2B",
    "zepto-mini": "Zepto",
    "swiggy insta-mini": "Swiggy Insta",
    "blinkit freebies": "Blinkit",
}

# Category values treated as missing.
UNMAPPED_CATEGORY = {"", "none", "#n/a", "no_sku_detail", "no_sku_detail(hyp_shopify_in)", "nan"}
