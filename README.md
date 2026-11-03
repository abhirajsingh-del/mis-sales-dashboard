# MIS Sales Dashboard

A monthly GMV / Sales / COGS dashboard for mCaffeine, Hyphen and DND. It covers channel, brand, category and SKU cuts, and it is built from the MIS workbook's **Sales Dump, Prod Dump, Samples and Secondary Sales** sheets.

All rupee and unit values are **masked** (× 0.75, set in `pipeline/config.py`) before they reach the dashboard. Every % shown is the actual figure.

```
input/                      monthly MIS workbooks (raw, unmasked)
pipeline/config.py          mask factor, sheet/column map, brand + channel mapping, channel groups
pipeline/pipeline.py        extract → normalise → reconcile → aggregate → mask → write
pipeline/make_template.py   rebuilds template/MIS_Template_Next_Month.xlsx
template/                   empty workbook for the next month (single per-unit COGS)
docs/index.html             the dashboard (GitHub Pages serves docs/)
docs/data.js                masked data the dashboard reads (written by the pipeline)
output/data_health.txt      reconciliation report for the last run       (local only)
output/dashboard_standalone.html  single-file copy to email / share, masked (local only)
output/dashboard_actual.html      single-file copy with ACTUAL values     (local only - never share)
data/*.csv                  unmasked normalised rows                       (local only)
```

## Every month

1. Fill `template/MIS_Template_Next_Month.xlsx`, or keep using your usual MIS workbook. The sheet names and header names must stay the same. COGS comes from the original per-unit columns (`RM`, `PM`, `Shrink`, `Shipper`, `JW`; Samples `Value`). The JAS-26 file's Updated / Revised columns are used only where they exist.
2. Save it in `input/` as **`MIS_YYYY_MM.xlsx`**, e.g. `MIS_2026_10.xlsx`. Files are read in name order, and if two files contain the same month, the later name wins. So to correct a month, either re-save that file or add a newer-named file.
3. Build:
   ```bash
   pip install -r requirements.txt
   ```
   ```bash
   python pipeline/pipeline.py
   ```
   It prints the health checks. Every line should say **PASS** (they are also in `output/data_health.txt` and on the dashboard's Data Health tab). If a check fails, the run exits with an error so a bad month isn't published by accident.
4. To see the real numbers yourself, open `output/dashboard_actual.html`. Every run writes it from the same data, unmasked. It sits in the git-ignored `output/` folder, so it is never committed or hosted.
5. Publish. Choose one of the two modes below.

### Hosting on GitHub Pages

One-time setup: create the repo, push this folder, then go to **Settings → Pages → Source: GitHub Actions**. The workflow in `.github/workflows/build-dashboard.yml` publishes `docs/` on every push.

| Mode | What goes to GitHub | Monthly step |
|---|---|---|
| **A. Build locally (default)** | Only masked `docs/` and the code. `input/*.xlsx` is git-ignored. | Run the pipeline on your laptop, then commit and push `docs/data.js`. |
| **B. Build on GitHub** | Raw workbooks too. Remove `input/*.xlsx` from `.gitignore`. | Commit the new `input/MIS_YYYY_MM.xlsx` and push. The Action runs the pipeline, commits `docs/data.js` and redeploys. |

Notes:
- **A GitHub Pages site is public on the internet**, even when the repo is private, unless your org is on GitHub Enterprise Cloud with private Pages. The masked numbers are still real ratios, so check this with your team before turning Pages on.
- Mode B keeps unmasked workbooks in the repo. GitHub rejects files over 100 MB; the JAS file is about 91 MB.

## What the numbers mean

| Metric | Definition | Source |
|---|---|---|
| Gross GMV | MRP value | Sales Dump `GMV` |
| Net GMV | GMV − RTO/Cancellation | Sales Dump `Actual GMV` = Prod Dump `Updated GMV` |
| Final Sales | GMV − Discount − Dealer Margin − RTO (incl. GST) | Sales Dump `Final Sales` = Prod Dump `Updated Sales` |
| Tax | GST inside Final Sales | Prod Dump `MIS Tax Amount` |
| Net Sales | Final Sales − Tax | |
| COGS | RM + PM + Shrink + Shipper + Job Work | Prod Dump |
| Samples cost | Samples, Compensation, Marketing and ATL units × per-unit cost | Samples |
| Gross Margin | Net Sales − COGS − Samples cost | |
| Secondary | Partner sell-through. Net = GMV − Margin − Tax | Secondary Sales. Shown on its own, never added to primary totals |

**Dump type:**
- **Internal** rows are the main data.
- **Reco** rows are discount reconciliation lines.
- **Others** rows are B2B adjustments.

The headline view is Internal + Reco, which matches the MIS Sales Pivot. Switch "Others" on in the Dump type filter to include the adjustments.

**Checks the pipeline runs:**
- the Sales Dump identities hold;
- Sales Dump `GMV − RTO` = Prod Dump `Updated GMV`, and Final Sales = `Updated Sales`, for every month × dump type × brand × channel (within ₹100);
- every brand and channel maps to a known value;
- every month is present in all four dumps.

It also reports, without failing, the COGS columns used per file and the rows with no COGS or no category.

## Changing things

- New channel spelling: add it to `CHANNEL_ALIASES` in `pipeline/config.py`. A new channel goes in `CHANNEL_GROUPS`.
- New column positions are fine, because columns are found by header name. A renamed header needs a matching edit in `SHEETS` in `config.py`.
- Preview locally:
  ```bash
  python -m http.server 8000 --directory docs
  ```
  Or just open `output/dashboard_standalone.html`.
