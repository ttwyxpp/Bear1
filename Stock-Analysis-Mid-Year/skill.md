---
name: Stock-Analysis-Mid-Year
description: Generate a mid-year inventory analysis HTML report + PDF for any Skyline sister-company warehouse. Covers stock trends, 库销比 (stock-to-sales ratio), PO pipeline, and PO volume analysis with actual FOB spec volumes.
---

# Mid-Year Inventory Analysis Report Skill

Generate a self-contained HTML report (+ PDF export) for any warehouse's H1 inventory data. The AZ report at `C:\Users\Jonathan\Desktop\2026 inventory report\AZ_H1_2026_Inventory_Report.html` is the canonical reference — clone and adapt it for each new company.

---

## Required Source Files

| File | Description |
|------|-------------|
| `[CO] mid-year data.xlsx` | Main inventory Excel; one sheet per PO plus Sheet1 with color snapshot data |
| `fob-specs.json` | FOB volumes/weights per SKU — at `C:\Users\Jonathan\.claude\skills\Skyline-SKU\sku-data\fob-specs.json` |

The mid-year Excel must already exist with the structure described below. Place it alongside the output HTML.

---

## Excel Source File Structure

### Sheet1 — Color Snapshot Data

Each **row = one color**, each **column = one snapshot date**. Extract:
- `stock_qty` per color per date → feeds `RAW` (cabinet equivalents)
- `30_day_sales` per date → feeds `SALES`
- `po_qty` per color per date → feeds `PO_DATA`
- `sea_qty` per color per date (on-water subset of PO) → feeds `SEA_DATA`
- `ratio` per color per date (stock ÷ monthly_sales × 30) → feeds `RATIO`

Dates for AZ H1 2026: Jan 6, Jan 21, Mar 3, Apr 8, May 8, Jun 6, Jul 16 (7 snapshots).
Adjust dates to match the actual report period. `null` = color not yet tracked at that date.

### PO Sheets — One Per Open Purchase Order

Sheet name format: `PO#NAC260303-AZ` (company suffix varies). Columns:
| Col | Field |
|-----|-------|
| 1 | Item (SKU) |
| 3 | Quantity (total ordered) |
| 4 | Received |
| 6 | Quantity on Shipments (includes received + in-transit) |

**On-water qty** = `max(0, QtyOnShipments − Received)`

---

## JavaScript Data Structures (in HTML)

```javascript
const DATES = ['Jan 6', 'Jan 21', 'Mar 3', 'Apr 8', 'May 8', 'Jun 6', 'Jul 16'];

// Stock per color per date [cabinets]; null = not tracked yet
const RAW = { 'SW': [16.81, 19.64, ...], 'GR': [...], ... };
const TOTAL = [97.79, 98.34, 123.57, ...];   // warehouse total per date

// 30-day sales per snapshot date [cabinets]
const SALES = [8.90, 8.90, 14.00, ...];

// 库销比 per color per date
const RATIO = { 'SW': [10.75, 12.57, ...], ... };
const TOTAL_RATIO = [10.99, 11.05, 8.83, ...];

// PO on-order and at-sea per color per date [cabinets]
const PO_DATA  = { 'SW': [7.39, 3.03, ...], ... };
const SEA_DATA = { 'SW': [5.56, 1.20, ...], ... };
const TOTAL_PO  = [46.70, 40.83, ...];
const TOTAL_SEA = [25.68, 26.27, ...];

// PO Volume Analysis (from PO sheets × fob-specs.json)
const CBM = 65;   // m³ per container

// Per-PO totals (Completion chart)
const PO_COMP = [
  { label:'260303', tot:15.22, water:2.54, rcvd:11.95, pend:0.73 },
  ...
];

// Color totals for bar chart (old/frameless only, sorted ascending)
const PO_VOL2 = [ ['UBX',16.23,1.68], ['SW',13.59,3.99], ... ];

// Per-PO per-color detail (ALL colors)
const PO_TABLE = [
  ['UBX', [9.16,0.59], [0.31,0.31], ...],  // [tot,water] per PO
  ...
];
```

---

## Color Groups

```javascript
const OLD_KEYS   = ['SW','GR','NB','SWO','SDW'];
const NEW_KEYS   = ['SA','IB','TC','SAG','DDW','DSG','AG','UBX','RW','SB','LG','SIB','SAO'];
const FRMLS_KEYS = ['HG','HW','PWB','MI','MB','MG','CG','OB','OS'];

// 2026 new colors — excluded from "Current PO vs On Water" bar chart
const NEW_COLORS_2026 = new Set(['AG','DDW','DSG','IB','LG','RW','SA','SAG','SAO','SB','SIB','TC']);
```

---

## Report Sections

| Section | Chart/Table | Data Used |
|---------|-------------|-----------|
| Overview KPIs | 4 cards | `TOTAL[0]`, `TOTAL[5]`, `SALES` |
| Total Stock Trend | Line chart | `TOTAL`, `DATES` |
| Stock by Color Category | Stacked bar | `oldData`, `newData`, `frmlsData` |
| Individual Color Trends | Line (top 12) | `RAW`, top 12 by Jun stock |
| Stock vs. 30-Day Sales | Dual-axis line | `TOTAL`, `SALES` |
| 库销比 KPIs | 4 cards | `TOTAL_RATIO` at key dates |
| 库销比 Overall Trend | Line chart | `TOTAL_RATIO` |
| 库销比 by Color (Jul 16) | Horizontal bar | `RATIO[color][6]` |
| 库销比 Trend — Key Colors | Multi-line | `RATIO` for 10 core colors |
| 库销比 Detail Table | Table | `RATIO`, all colors |
| PO KPIs | 4 cards | `TOTAL_PO`, `TOTAL_SEA`, `SALES` |
| Total Pipeline Breakdown | Grouped bar | `TOTAL`, `TOTAL_PO`, `TOTAL_SEA` |
| PO Trend vs. Sales | Line chart | `TOTAL_PO`, `SALES` |
| Open PO by Color (Jul 16) | Horizontal bar | `PO_DATA[color][6]`, excludes `NEW_COLORS_2026` |
| PO Detail Table | Table | `PO_DATA`, `SEA_DATA` |
| PO Volume KPIs | 4 cards | `PO_COMP` totals |
| Current PO vs On Water — Volume by Color | Stacked horizontal bar | `PO_VOL2` (old/frameless only) |
| PO Completion Status — by Order | Stacked horizontal bar | `PO_COMP` |
| PO Volume Detail Table | Wide table (Color × PO) | `PO_TABLE` (all colors) |
| Color Stock Summary | Table | `RAW`, `PO_DATA` |

---

## PO Volume Extraction Script (PowerShell)

Reads all PO sheets from the mid-year Excel, looks up volumes in `fob-specs.json`, and outputs the three JS arrays (`PO_COMP`, `PO_VOL2`, `PO_TABLE`).

```powershell
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open('PATH\TO\MID-YEAR.xlsx')
$fob = (Get-Content 'C:\Users\Jonathan\.claude\skills\Skyline-SKU\sku-data\fob-specs.json' -Raw) | ConvertFrom-Json

$CBM = 65.0
# Update PO_NAMES and PO_LABELS to match actual PO sheet names in the file
$PO_NAMES  = @('PO#NAC260303-AZ','PO#NAC260403-AZ','PO#NAC260508-AZ','PO#NAC260606-AZ','PO#NAC260626-AZ')
$PO_LABELS = @('260303','260403','260508','260606','260626')

$INDIVIDUAL = [System.Collections.Generic.HashSet[string]]@(
  'AG','CG','CW','DDW','DSG','GR','HG','HW','IB','LG','MB','MG','MI',
  'NB','OB','OS','PWB','RW','SA','SAG','SAO','SB','SDW','SE','SIB','SW','SWO','TC','UBX'
)

$colorData = @{}; $poTotals = @{}
foreach ($po in $PO_LABELS) { $poTotals[$po] = @{tot=0.0;rcvd=0.0;water=0.0} }

for ($pi=0; $pi -lt $PO_NAMES.Count; $pi++) {
  $ws = $wb.Sheets.Item($PO_NAMES[$pi])
  $po = $PO_LABELS[$pi]
  $lastRow = $ws.UsedRange.Rows.Count

  for ($r=2; $r -le $lastRow; $r++) {
    $sku = $ws.Cells.Item($r,1).Value2
    if ($null -eq $sku -or $sku -eq '') { continue }
    $sku = $sku.Trim()

    $qv=$ws.Cells.Item($r,3).Value2; if ($null -eq $qv){$qv=0}
    $rv=$ws.Cells.Item($r,4).Value2; if ($null -eq $rv){$rv=0}
    $sv=$ws.Cells.Item($r,6).Value2; if ($null -eq $sv){$sv=0}
    $qty=[double]$qv; $rcvd=[double]$rv; $ship=[double]$sv

    # FOB lookup with slash-variant and -FL normalization
    $lookupKey = if ($sku -match '/') { $sku.Split('/')[0] } else { $sku }
    $lookupKey = $lookupKey -replace '-FL$',''
    $fobEntry = $fob.$lookupKey
    if ($null -eq $fobEntry) { continue }
    $vol = $fobEntry.vol_m3

    # Color grouping
    $color = $sku.Split('-')[0]
    if ($sku -match '^UBX-') { $color = 'UBX' }
    if ($sku -match '^PWB-') { $color = 'PWB' }
    if (-not $INDIVIDUAL.Contains($color)) { $color = 'Other' }

    if (-not $colorData.ContainsKey($color)) {
      $arr = @(); for ($i=0;$i -lt 5;$i++) { $arr += ,@{tot=0.0;rcvd=0.0;water=0.0} }
      $colorData[$color] = $arr
    }
    $tv = $qty * $vol
    $wv = [Math]::Max(0, $ship - $rcvd) * $vol
    $colorData[$color][$pi].tot   += $tv
    $colorData[$color][$pi].water += $wv
    $colorData[$color][$pi].rcvd  += $rcvd * $vol
    $poTotals[$po].tot   += $tv
    $poTotals[$po].water += $wv
    $poTotals[$po].rcvd  += $rcvd * $vol
  }
}
$wb.Close($false); $xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null

function R2($v) { [Math]::Round($v,2) }

# Output PO_COMP
Write-Host "const PO_COMP = ["
foreach ($po in $PO_LABELS) {
  $t=$poTotals[$po].tot/$CBM; $w=$poTotals[$po].water/$CBM
  $rc=$poTotals[$po].rcvd/$CBM; $pend=$t-$rc-$w
  Write-Host "  { label:'$po', tot:$(R2 $t), water:$(R2 $w), rcvd:$(R2 $rc), pend:$(R2 $pend) },"
}
Write-Host "];"

# Output PO_TABLE (row order: largest total first)
Write-Host "`nconst PO_TABLE = ["
$sortOrder = @('UBX','SW','SWO','TC','SA','SDW','GR','AG','IB','DDW','NB','SAG','DSG',
               'LG','SB','SIB','SAO','RW','PWB','HG','MB','MG','MI','OB','OS','SE','CG','CW','HW','Other')
foreach ($c in $sortOrder) {
  if (-not $colorData.ContainsKey($c)) { continue }
  $cells = for ($i=0;$i -lt 5;$i++) {
    "[$(R2($colorData[$c][$i].tot/$CBM)),$(R2($colorData[$c][$i].water/$CBM))]"
  }
  Write-Host "  ['$c', $($cells -join ', ')],"
}
Write-Host "];"
```

**Key rules:**
- `??` null-coalescing is **not available** in PowerShell 5.1 — use `if ($null -eq $x){$x=0}`
- Slash-variant SKUs: `AG-DS-3DB12/3VDB12` → try `AG-DS-3DB12` as lookup key (first part before `/`)
- Strip `-FL` suffix before FOB lookup
- `on_water = max(0, QtyOnShipments − Received)` — QtyOnShipments includes already-received items

---

## Generating the Report for a New Company

### Step 1 — Prepare the source Excel

Ensure the mid-year Excel has:
- Sheet1 with color snapshot data across all dates
- One sheet per open PO, named `PO#NAC[YYMMDD]-[CO]` (e.g. `PO#NAC260303-FL`)

### Step 2 — Copy and adapt the HTML template

1. Copy `AZ_H1_2026_Inventory_Report.html` → `[CO]_H1_2026_Inventory_Report.html`
2. Update the report header: company name, date range, source file name
3. Replace all data arrays by extracting from the new company's Excel:
   - `RAW`, `TOTAL`, `SALES` — from Sheet1 stock/sales snapshots
   - `RATIO`, `TOTAL_RATIO` — from Sheet1 ratio columns
   - `PO_DATA`, `SEA_DATA`, `TOTAL_PO`, `TOTAL_SEA` — from Sheet1 PO columns
   - `PO_COMP`, `PO_VOL2`, `PO_TABLE` — from PO sheets via the extraction script above
4. Update PO sheet names in the extraction script and in the HTML table headers
5. Update `DATES` array to match actual snapshot dates

### Step 3 — Adjust chart Y-axis ranges

The AZ report has hardcoded Y-axis ranges for some charts:
- Total Stock Trend: `min: 80, max: 135` — adjust for the new company's scale
- 库销比 Trend: `min: 4, max: 14` — adjust to actual ratio range
- Dual-axis Sales: `y min/max`, `y1 min/max` — adjust for new volumes

### Step 4 — Generate PDF

```powershell
$co     = 'FL'   # change per company
$html   = "C:\Users\Jonathan\Desktop\2026 inventory report\${co}_H1_2026_Inventory_Report.html"
$pdf    = "C:\Users\Jonathan\Desktop\2026 inventory report\${co}_H1_2026_Inventory_Report.pdf"
$chrome = 'C:\Program Files\Google\Chrome\Application\chrome.exe'

& $chrome --headless=new --disable-gpu `
  --print-to-pdf="$pdf" `
  --print-to-pdf-no-header `
  --run-all-compositor-stages-before-draw `
  --virtual-time-budget=6000 `
  "file:///$($html -replace '\\','/')"
```

**Print CSS** is already in the HTML template:
- A4 landscape, 10mm margins
- `break-before: page` on every `.card` → one chart per page
- Charts expand to fill page height (~580–600px)

---

## PO Volume Analysis — Data Consistency Rules

The three PO Volume arrays must be computed from the **same extraction pass** to stay consistent:
- `PO_COMP[i].tot` must equal the column sum of `PO_TABLE` for that PO index (within rounding)
- `PO_VOL2` totals (summed across all POs per color) must match the color totals in `PO_TABLE`
- `PO_TABLE` must include **all colors** — not just old colors — or the totals will not match `PO_COMP`

If you notice PO Completion chart and PO Volume Detail Table totals diverging, it means they were computed from different passes. Re-run the extraction script to regenerate all three from the same data.

---

## 库销比 (Stock-to-Sales Ratio) Coloring

| Range | Color | Label |
|-------|-------|-------|
| < 5   | `#dc2626` red | Critical |
| 5–10  | `#ea580c` orange | Low |
| 10–30 | `#16a34a` green | Healthy |
| 30–60 | `#ca8a04` yellow | Elevated |
| > 60  | `#9d174d` pink | Overstock |

The 库销比 by Color bar chart excludes colors that were newly introduced and have unreliable early ratios. For AZ H1 2026 the exclude list was:
```javascript
const NEW_COLOR_EXCLUDE = new Set(['SIB','LG','MI','MB','SB','MG','CG','RW','SAO','OB','OS']);
```
Adjust this per report period and company.

---

## Libraries Used

```html
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
```

All charts use `responsive: true, maintainAspectRatio: false`. Register plugin at top of script:
```javascript
Chart.register(ChartDataLabels);
```

---

## Reference Files

| File | Path |
|------|------|
| AZ canonical report (HTML) | `C:\Users\Jonathan\Desktop\2026 inventory report\AZ_H1_2026_Inventory_Report.html` |
| AZ source Excel | `C:\Users\Jonathan\Desktop\2026 inventory report\AZ mid-year data.xlsx` |
| FOB specs JSON | `C:\Users\Jonathan\.claude\skills\Skyline-SKU\sku-data\fob-specs.json` |
| SKU data (per company) | `C:\Users\Jonathan\.claude\skills\Skyline-SKU\sku-data\[CO].json` |
