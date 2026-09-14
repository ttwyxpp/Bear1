---
name: midyear-combined-report
description: Generate combined Mid-Year Inventory Report HTML for each of 5 sister companies, merging selected charts from the Stock Report and ZeroStock Analysis into one self-contained page.
---

# Mid-Year Combined Inventory Report Skill

Produces `[Company]_Mid-Year_Inventory_Report.html` for each of the 5 sister companies by merging data/charts from two source reports.

---

## Output Files

All in `C:\Users\Jonathan\Desktop\Mid year\Artisan\`:

| File | Company |
|------|---------|
| `Artisan_Mid-Year_Inventory_Report.html` | Artisan Cabinetry LLC (HOU) |
| `Milestone_Mid-Year_Inventory_Report.html` | Milestone Cabinetry Inc. (FL) |
| `Oasis_Mid-Year_Inventory_Report.html` | Oasis Cabinetry LLC (AZ) |
| `Skyline_Mid-Year_Inventory_Report.html` | Skyline Cabinetry Inc. (TX) |
| `Spring Forest_Mid-Year_Inventory_Report.html` | Spring Forest Cabinetry (NC) |

---

## Source Files

| Source | Path | Purpose |
|--------|------|---------|
| Stock Reports | `C:\Users\Jonathan\Desktop\Mid year\Artisan\Stock Report\` | Charts 1–5 (inventory + ratio + pipeline) |
| ZeroStock HTMLs | `C:\Users\Jonathan\Desktop\Mid year\Artisan\` | Charts 6–9 (zero-stock + overstock) |

### Stock Report filenames → company mapping

| File | Company |
|------|---------|
| `HOU_Inventory_Report.html` | Artisan |
| `FL_Inventory_Report (1).html` | Milestone |
| `AZ_H1_2026_Inventory_Report (1).html` | Oasis |
| `TX_Inventory_Report.html` | Skyline |
| `NC_Inventory_Report.html` | Spring Forest |

---

## Report Sections (9 charts total)

### Section 1 — Inventory Analysis (from Stock Report)

| # | Chart | Canvas ID | Data Variables |
|---|-------|-----------|---------------|
| 1 | Total Stock Trend | `totalChart` | `TOTAL`, `INV_DATES` |
| 2 | Stock by Color Category | `catChart` | `RAW`, `OLD_KEYS`, `NEW_KEYS`, `FRMLS_KEYS` |
| 3 | 库销比 Overall Trend | `ratioTrendChart` | `TOTAL_RATIO`, `INV_DATES` |
| 4 | 库销比 Trend — Key Colors | `ratioColorChart` | `RC_KEYS`, `RC_DATA`, `INV_DATES` |
| 5 | Total Pipeline Breakdown | `pipelineChart` | `TOTAL`, `TOTAL_PO`, `TOTAL_SEA`, `INV_DATES` |

### Section 2 — Zero-Stock & Overstock Analysis (from ZeroStock HTML)

| # | Chart | Canvas ID | Data Variables |
|---|-------|-----------|---------------|
| 6 | Zero-Stock SKU Count Over Time | `zeroTrendChart` | `ZERO_COUNT`, `ZS_DATES` |
| 7 | Zero-Stock Duration Breakdown | `zeroDurChart` | `ZERO_DUR` |
| 8 | Overstock as % of Total Inventory | `overVolChart` | `OVER_VOL_M3`, `TOTAL_VOL_M3`, `ZS_DATES` |
| 9 | Overstock Volume Trend (m³) | `overChgChart` | `OVER_VOL_M3`, `ZS_DATES`, `CBM` |

**Variable name conflict**: Both source reports use `DATES`. In the combined report:
- Stock report dates → renamed to `INV_DATES`
- ZeroStock dates → renamed to `ZS_DATES`

All chart creation code that references these variables must use the renamed versions.

---

## Critical Bug: Missing Hidden Element IDs

The Stock Report JS references many DOM elements (KPI cards, detail tables) that are NOT shown in the combined report. These are placed in a `display:none` div. **If any referenced ID is missing, the JS throws a TypeError that stops all subsequent chart creation.**

The combined report's hidden div MUST include all of the following IDs:

```html
<!-- Hidden elements required by JS from source reports but not shown here -->
<div style="display:none">
  <canvas id="colorChart"></canvas>
  <canvas id="stockSalesChart"></canvas>
  <canvas id="ratioBarChart"></canvas>
  <canvas id="poSalesChart"></canvas>
  <canvas id="poColorChart"></canvas>
  <canvas id="poVolChart"></canvas>
  <canvas id="poCompChart"></canvas>
  <canvas id="zeroColorChart"></canvas>
  <div id="kpiRow"></div>
  <div id="ratioKpiRow"></div>
  <div id="poKpiRow"></div>
  <div id="poVolKpiRow"></div>
  <span id="colorChartNote"></span>
  <table id="heatTable"></table>
  <table id="matrixTable"></table>
  <table id="overVolTable"></table>
  <table><tbody id="szTbody"></tbody></table>
  <table><tbody id="chTbody"></tbody></table>
  <table><tbody id="ratioTable"></tbody></table>
  <table><tbody id="poTable"></tbody></table>
  <table><tbody id="summaryTable"></tbody></table>
  <table><tbody id="poVolTable"></tbody></table>
</div>
```

**Why this matters**: `buildRatioKpis()` runs before `ratioTrendChart` is created. If `ratioKpiRow` is missing, it throws and charts 3, 4, 5 never render. Same pattern applies to `poKpiRow` (before `pipelineChart`) and the table IDs.

---

## How to Regenerate

1. **Run `build_all_companies.ps1`** to produce fresh ZeroStock HTML files for all 5 companies. These contain the `ZERO_COUNT`, `ZERO_DUR`, `OVER_VOL_M3`, `TOTAL_VOL_M3` data.

2. **Extract JS data blocks** from each company's ZeroStock HTML (`[Company]_ZeroStock_Analysis.html`) and the matching Stock Report HTML.

3. **Build the combined HTML** with the 9-section structure above. Key rules:
   - Include both CDN scripts: Chart.js 4.4.0 + chartjs-plugin-datalabels 2.2.0
   - Register plugin: `Chart.register(ChartDataLabels);`
   - Rename `DATES` → `INV_DATES` throughout stock report JS, `DATES` → `ZS_DATES` throughout zero-stock JS
   - Include the complete hidden element div (all IDs listed above)
   - Section 1 header color: `#1a3c5e` (blue); Section 2 header: `#dc2626` (red)

---

## Overstock Chart Notes

- **Overstock as % (overVolChart)**: Stacked bar; total = 100%; orange = overstock%, blue = normal. Labels shown above bar when <12% (orange text), inside bar when ≥12% (white text). `layout.padding.top: 28`, `clip: false`. Legend at bottom.
- **Overstock Volume Trend (overChgChart)**: Line chart, Y-axis in containers (÷ 65 m³). Point color: orange for first point, red for increases, green for decreases.
- `CBM = 65` (m³ per container)
- FW/FE/FG color prefixes: excluded from zero-stock sections and overstock numerator, but INCLUDED in `TOTAL_VOL_M3` (denominator).

---

## CSS Classes Required

```css
.section-hdr { font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.2px;
  color: #1a3c5e; border-left: 4px solid #1a3c5e; padding-left: 12px; margin: 28px 0 16px; }
.section-hdr.red { color: #dc2626; border-left-color: #dc2626; }
.card { background: white; border-radius: 10px; padding: 20px 24px; margin-bottom: 20px;
  box-shadow: 0 1px 4px rgba(0,0,0,.08); }
.grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
.chart-h280 { position: relative; height: 280px; margin-top: 16px; }
.chart-h300 { position: relative; height: 300px; margin-top: 16px; }
.chart-h320 { position: relative; height: 320px; margin-top: 16px; }
.chart-h360 { position: relative; height: 360px; margin-top: 16px; }
.chart-h560 { position: relative; height: 560px; margin-top: 16px; }
```

Charts 7 & 8 are placed side-by-side in `.grid-2` (using `.chart-h280`). All others are full-width.
