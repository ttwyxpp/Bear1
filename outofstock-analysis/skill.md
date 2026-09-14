---
name: outofstock-analysis
description: Generate zero-stock (out-of-stock) analysis HTML page and Monthly_Inventory Excel for any Skyline sister company using bi-monthly inventory CSV snapshots.
---

# Out-of-Stock Analysis Skill

Generates two outputs per company:
1. `[Company]_ZeroStock_Analysis.html` — standalone zero-stock analysis page (charts, heatmap, per-SKU matrix)
2. `[Company]_Monthly_Inventory.xlsx` — all orderable SKUs across 15 bi-monthly snapshots (Jan 1 – Aug 1)

---

## Source Files Required

| File | Description |
|------|-------------|
| `[Company]/[Company] X-X.csv` | 15 inventory snapshots (1-1, 1-15, 2-1 … 8-1) |
| `[Company]/SKU-[Company].xlsx` | Orderable SKU list (col 1, from row 2) |

All files for Artisan H1 2026 are at:
`C:\Users\Jonathan\Desktop\Mid year\Artisan\`

---

## SKU Filtering Rules

**Excluded color prefixes** (not tracked):
`SAO, SIB, SB, LG, RW, MI, MB, MG, CG, OB, OS`

**Excluded patterns**: `-MINI` suffix, `-SD` suffix

**Tracked colors** (for per-color breakdown):
`SW, GR, NB, SWO, SDW, SA, IB, TC, SAG, DDW, DSG, AG, UBX, HG, HW, PWB`

---

## CSV Format

Each snapshot CSV has 5 columns: `Item, Description, Inv. Value, % of Inv. Value, On Hand`
- Header row is at line index 5 (0-based); data starts at index 6
- Zero-stock items have **empty** "On Hand" field → parsed as `$t[$sku] = 0`
- Negative quantities appear as `-\d+` → marked as `'ERROR'`
- Values may be quote-wrapped — use `.Trim('"')`

---

## JS Data Constants (ZeroStock HTML)

```javascript
const DATES = ['Jan 1','Jan 15','Feb 1','Feb 15','Mar 1','Mar 15','Apr 1','Apr 15','May 1','May 15','Jun 1','Jun 15','Jul 1','Jul 15','Aug 1'];

// Count of SKUs at qty=0 per snapshot (15 values)
const ZERO_COUNT = [31, 22, 32, ...];

// Duration buckets: how many SKUs had zero-stock events
const ZERO_DUR = {
  '1 snapshot (2 wks)': 69,
  '2-4 snapshots (1-2 mo)': 44,
  '5+ snapshots (2.5+ mo)': 14,
  'Still at 0 (Aug 1)': 9
};

// Per tracked-color, count of snapshots with >=1 zero for that color's SKUs
const ZERO_BY_COLOR = {
  'SW': [2, 1, 3, ...],  // 15 values
  ...
};

// SKUs still at zero at Aug 1 snapshot
const STILL_ZERO = ['SW-DS-W1530', ...];

// SKUs at zero for 5+ snapshots: [sku, count]
const CHRONIC_ZERO = [['SW-DS-W1530', 7], ...];

// Per-SKU zero-stock boolean matrix: [sku, [15 bits (0 or 1)]]
// Sorted by color prefix then SKU name, only includes SKUs with any zero event
const ZERO_MATRIX = [
  ['AG-DS-3VDB12', [0,0,1,1,0,0,0,0,0,0,0,0,0,0,0]],
  ...
];
```

---

## HTML Page Sections

| Section | Chart/Table | Key Data |
|---------|-------------|----------|
| Overview KPIs | 4 cards | total zero events, still at 0, peak date, chronic count |
| Zero-Stock Trend | Line chart | `ZERO_COUNT` over 15 dates |
| Duration Breakdown | Bar chart | `ZERO_DUR` |
| Color Totals | Bar chart | sum of `ZERO_BY_COLOR` per color |
| Color x Date Heatmap | Table | `ZERO_BY_COLOR` — yellow=1-2, orange=3-5, red=6+ |
| Per-SKU Matrix | Wide table | `ZERO_MATRIX` — red=zero, gray=has stock, grouped by color |
| Still at 0 | Table | `STILL_ZERO` |
| Chronic Zero-Stock | Table | `CHRONIC_ZERO` |

---

## Master Build Script

`C:\Users\Jonathan\Desktop\Mid year\Artisan\build_all_companies.ps1`

Processes all 5 companies in one run. Generates ZeroStock HTML for all; builds Monthly_Inventory.xlsx for Milestone, Oasis, Skyline, Spring Forest (Artisan has its own rebuild_full.ps1).

### Critical PS 5.1 Bug: HashSet from Variable

**WRONG** (returns null in PS 5.1):
```powershell
$TSET = [System.Collections.Generic.HashSet[string]]$TORD
```

**CORRECT** (must use literal array):
```powershell
$TSET = [System.Collections.Generic.HashSet[string]]@('SW','GR','NB','SWO','SDW','SA','IB','TC','SAG','DDW','DSG','AG','UBX','HG','HW','PWB')
```

A null HashSet causes `InvokeMethodOnNull` on `$TSET.Contains($pfx)`. With `$ErrorActionPreference = 'Continue'`, the error is silently swallowed and the per-color counts stay 0.

### Critical PS 5.1 Bug: Encoding in Here-Strings

PS 5.1 reads `.ps1` files as Windows system codepage (CP1252), not UTF-8. The Write tool saves UTF-8 without BOM. Chinese characters in here-strings will be garbled in the output HTML.

**Fix**: Replace all Chinese characters in here-strings with ASCII or HTML entity equivalents. Do NOT put Chinese text directly inside PowerShell here-strings in scripts saved by the Write tool.

---

## Per-Company Data Directories

All under `C:\Users\Jonathan\Desktop\Mid year\Artisan\`:

| Company | Dir | SKU file | CSV pattern |
|---------|-----|----------|-------------|
| Artisan | `Artisan\` | `SKU- Artisan.xlsx` | `Artisan X-X.csv` |
| Milestone | `Milestone\` | `SKU-Milestone.xlsx` | `Milestone X-X.csv` |
| Oasis | `Oasis\` | `SKU-Oasis.xlsx` | `Oasis X-X.csv` |
| Skyline | `Skyline\` | `SKU-Skyline.xlsx` | `Skyline X-X.csv` |
| Spring Forest | `Spring Forest\` | `SKU-Spring Forest.xlsx` | `Spring Forest X-X.csv` |

---

## Supporting Scripts

| Script | Purpose |
|--------|---------|
| `build_all_companies.ps1` | Master: all 5 companies, ZeroStock HTML + Excel |
| `compute_zero.ps1` | Artisan zero stats only (diagnostic / data extraction) |
| `compute_overvol.ps1` | Artisan overstock volume analysis (qty>30 × vol_m3) |
| `rebuild_full.ps1` | Artisan Monthly_Inventory.xlsx rebuild |

---

## Artisan H1 2026 Computed Stats (Reference)

- Total orderable SKUs: 3,236 (excl. MINI, SD, 11 excluded colors)
- SKUs with any zero event: 127
- Still at 0 on Aug 1: 9
- Chronic (5+ snapshots): 14

`ZERO_COUNT = [31,22,32,29,20,29,34,17,17,9,11,17,9,13,9]`

---

## Overstock Volume Analysis

Separate section in `Artisan_H1_2026_Inventory_Report.html` (main report, not ZeroStock page).

- Definition: SKU qty > 30 at a snapshot
- Volume per SKU: `qty × fob-specs.json[sku].vol_m3`
- FOB lookup key: take part before `/` for slash-variant SKUs; strip `-FL` suffix
- Container size: 65 m³ (`const CBM = 65`)
- Script: `compute_overvol.ps1`

`OVER_VOL_M3 = [4242.4,4353.15,4380.15,4951.59,5405.94,5186.7,4730.86,4785,4785,4386.17,3965.43,4032.19,3624.88,3536.04,3975.2]`
