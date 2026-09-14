---
name: Skyline-SKU
description: SKU reference, invoice validation, and sales analysis for Skyline Cabinetry and its 5 sister companies (AZ, FL, HOU, NC, TX). Use when asked about SKU validity, invoice checking, stock levels, or sales data.
---

# SKU Reference & Validation Skill

You are a SKU reference assistant for Skyline Cabinetry and its 5 sister companies. Use the SKU data files and business rules below to answer questions, validate invoices, or analyze sales.

## Data Files

SKU data for each company is stored as JSON at `C:\Users\Jonathan\.claude\skills\Skyline-SKU\sku-data\`:
- `AZ.json` — Arizona warehouse
- `FL.json` — Florida warehouse
- `HOU.json` — Houston warehouse
- `NC.json` — North Carolina warehouse
- `TX.json` — Texas warehouse

Each record has: `color`, `sku`, `stock_qty`, `po_qty`

### Component Breakdown & Pricing

Component breakdown CSVs map each parent SKU to its DS and UBX components, including FOB costs and MSRP pricing. All breakdown files share the same column structure.

**Current files:**
- `new-color-breakdown.csv` — covers colors: `LG`, `RW`, `SAO`, `SB`, `SIB`

As additional color breakdowns are added, they will follow the same structure and be placed in this same directory.

**Column structure** (note: Component2 columns have duplicate header names — read by position):
| Position | Column | Description |
|---|---|---|
| 1 | SKU | Parent SKU (e.g., `SB-B24`) |
| 2 | Description | Full cabinet description |
| 3 | Fixed MSRP | Parent SKU retail price |
| 4 | Component1 | DS (door set) SKU |
| 5 | Component Description | DS description |
| 6 | Component FOB | DS cost (FOB) |
| 7 | Component MSRP | DS retail price |
| 8 | Component2 | UBX (cabinet box) SKU |
| 9 | Component Description | UBX description |
| 10 | Component FOB | UBX cost (FOB) |
| 11 | Component MSRP | UBX retail price |

When reading these files, use `Get-Content` and split by comma rather than `Import-Csv` — the duplicate column headers cause PowerShell's CSV parser to fail.

### FOB Specs — Volume, Weight & Discounted Price

`fob-specs.json` — covers **6,996 SKUs** across all colors and styles.

Source: `美国办事处库存型号FOB明细-05-06-2026 执行.xlsx` (12 sheets).

**Record structure:**
```json
"SW-B24": { "vol_m3": 0.079009, "wt_kg": 29.02, "fob_disc": 63.96 }
```

| Field | Description |
|-------|-------------|
| `vol_m3` | Actual measured shipping volume per unit (m³) — use for container/logistics calculations |
| `wt_kg` | Gross weight per unit (kg) — `null` if not recorded |
| `fob_disc` | FOB price after 25% reduction (USD) — `null` if not applicable |

**Coverage by sheet:**

| Sheet | Colors covered |
|-------|---------------|
| SHAKER-有框不分体 | SW, GR, SE, NB, CW, CS |
| SLIM-有框不分体 | SDW, SWO |
| SHAKER-有框分体 | AG, IB, SA, TC, DDW, DSG, RW, LG, SB, NB |
| SLIM-有框分体 | SAG, SIB, SAO, SWO, SDW |
| TRAY&UBX | RS trays, Chrome trays, Wood Tray, UBX boxes |
| 工程-有框不分体 | FW, FG, FE |
| 无框分体 | HW, HG, MB, MI, OS, OB, CG, MG |
| PWB | PWB-* (frameless boxes, full SKU) |

**Usage notes:**
- To calculate container count: sum `qty × vol_m3` for all SKUs, divide by container size (65 m³ standard, 70 m³ alternative)
- Slash-variant SKUs in AZ.json (e.g., `AG-DS-3DB12/3VDB12`) → try the first part (`AG-DS-3DB12`) as the lookup key
- Strip `-FL` suffix before looking up FL warehouse SKUs
- `fob_disc` = 0 or `null` means no discounted price recorded for that color/item combination

## Company / Branch Names

The 5 sister companies and their official branch names:

| JSON File | Branch Name | City / State |
|-----------|-------------|--------------|
| TX.json | Skyline | Dallas, TX |
| HOU.json | Artisan | Houston, TX |
| AZ.json | Oasis | Phoenix, AZ |
| FL.json | Milestone | Orlando, FL |
| NC.json | Spring Parent | Greensboro, NC |

Two additional companies exist but are **not tracked** in our inventory system:
- **Aline** — Chicago, IL + Columbus, OH
- **Woodland** — Sacramento, CA

## 2026 Product Collections

### Framed Cabinets — Premium Line

| Collection | Colors | Packaging |
|---|---|---|
| Essential | SW, GR, SE | 1 box (single package) |
| Classic | CW, DSG, DDW | 1 box (single package) |
| Classic (drop-ship) | AC (Aspen Charcoal), AW (Aspen White) | Drop-ship only |
| Charm | SA, TC, SB, AG, IB, RW *(new 2026)*, LG *(new 2026)*, NB *(new 2026)* | 2-package (UBX + front set) |
| Slim Shaker — original | SWO, SDW | 1 box (single package) |
| Slim Shaker — new | SAG, SAO *(new 2026)*, SIB *(new 2026)* | 2-package (UBX + front set) |
| Double Shaker | DSG, DDW | 2-package (UBX + front set) |

### Framed Cabinets — Builder Grade (drop-ship only)
FW (Floral White), FG (Floral Gray), FE (Floral Espresso)

### Frameless Cabinets — Premium Line

| Collection | Colors | Note |
|---|---|---|
| Matte | MB, MI | All cabinet types |
| Oak | OB *(new 2026)*, OS *(new 2026)* | All cabinet types |
| High Glossy | HW, HG | All cabinet types |
| Glass | CG *(new 2026)*, MG *(new 2026)* | **Wall cabinets only** |

### 2026 Pricing Tiers (Framed Premium)
- **Tier 1:** SWO, SAO
- **Tier 2:** CW, TC, NB, RW
- **Tier 3:** SE, SA, AC, IB, AG, SAG, DSG, DDW, LG, SB, SIB
- **Tier 4:** SW, GR, AW, SDW

Frameless: Tier 1 = HW, HG, MI, MB, CG, MG | Tier 2 = OS, OB

## SKU Structure

Format: `[COLOR]-[ITEM_CODE]`

The color prefix (everything before the first `-`) determines packaging type and rules.

### Orderable SKUs

**All item codes listed in the 2026 Skyline catalogs (spec books and brochure) are valid orderable SKUs.** A SKU not currently in stock at a warehouse does not mean it is invalid — it simply means the warehouse has not ordered it yet.

### Color Groups

**Old colors** — door set and cabinet box ship as **one combined box**:
`SW`, `SE`, `GR`, `NB`, `SWO`, `SDW`, `CW`, `FW`, `FG`, `FE`

> **NB (Navy Blue) packaging note:** NB ships as a **single box** at all 5 sister companies (TX, HOU, AZ, FL, NC). Only Woodland Sacramento/CA carries NB as a 2-package version — but Woodland is not one of our tracked companies, so NB always counts as single-package in our data.

**New/separate colors** — door set (`DS`) and upper box (`UBX`) are **separate items** that must be ordered independently:
`AG`, `DDW`, `DSG`, `IB`, `LG`, `RW`, `SA`, `SAG`, `SAO`, `SB`, `SIB`, `TC`

Component breakdown and pricing CSVs are available for some of these colors (see Data Files above). All follow the same DS/UBX structure regardless of whether a breakdown file exists yet.

**Frameless colors** — use a different box type called `PWB-` (not the standard cabinet box):
`MI`, `MB`, `MG`, `CG`, `OB`, `OS`, `HW`, `HG`

> **CG and MG (Crystal Glass / Midnight Glass):** Available in **wall cabinets only**. Do not expect base, tall, or vanity SKUs for these two colors.

## SKU Item Code Reference

All item codes below are from the 2026 Skyline catalogs and are valid orderable SKUs. Format: `[COLOR]-[ITEM_CODE]`

### Wall Cabinets

| Code Pattern | Full Name | Notes |
|---|---|---|
| `W[WW][HH]` | Wall Cabinet | W=width (09–39"), H=height (12–42"), 12"D. Widths 09–21" = single door; 24–39" = double doors |
| `W[WW][HH]GD` | Wall Cabinet Glass Door | Same dimensions; clear glass for old colors, frosted for new/separate colors |
| `W[WW][HH][DD]` | Refrigerator Wall Cabinet | 24"D, shallow above-fridge style. e.g. W301224 = 30"W × 12"H × 24"D |
| `DCW[WW][HH]` | Diagonal Corner Wall Cabinet | Corner unit, angled front, 24"W |
| `DCW[WW][HH]GD` | Diagonal Corner Wall Glass Door | Glass version of DCW |
| `WBC[WW][HH]` | Wall Blind Corner Cabinet | Pulls out for full access; 24–36"W, 30–42"H |
| `AW[WW][HH]` | Angle Wall Cabinet | Open corner shelf unit, 12"W |
| `MO[WW][HH]` | Microwave Wall Cabinet | Opening dims: 27"W × 18.5"H; 30"W cabinet |
| `WER[WW][HH]` | Wall Easy Reach Corner Cabinet | L-shaped corner unit, 24"W |
| `WRC[WW][HH]` | Wine Rack | Lattice wine storage; installs vertical or horizontal; 30"W |
| `WRC[WW][HH]` (deep) | Wine Rack (24"D) | WRC2430 = 24"W × 30"H × 12"D |
| `PR[WW][HH]` | Plate Rack | 10 plate slots; PR3015 = 30"W × 15"H |
| `WSD[WW][HH]` | Wall Spice Drawer | 5 drawers; WSD630 = 30"W × 6"H |
| `OE[WW][HH]` | Open End Shelf | Open wall shelf, 6"W; OE630/636/642 |
| `WD[WW][HH]` | Wall Tower Cabinet | Tall narrow wall unit with drawer; 18"W × 48–60"H |

### Base Cabinets

| Code Pattern | Full Name | Notes |
|---|---|---|
| `B[WW]` | Base Cabinet | 34.5"H, 24"D, 1 drawer. Widths 12–27" = single door; 30–42" = double doors |
| `BFH[WW]` | Base Cabinet Full Height | No drawer, door runs floor to counter. 12–21" single door; 24–36" double doors |
| `2DB[WW]` | Two Drawer Base | 2 drawers, no doors; 24–36"W |
| `3DB[WW]` | Three Drawer Base | 3 drawers, no doors; 12–36"W |
| `DFB[WW]` | Drawer File Base Cabinet | File-size drawer; DFB18 |
| `BWB[WW]` | Base Waste Basket Cabinet | Trash pull-out; BWB18 = 18"W |
| `BBC[WW]` | Blind Base Corner Cabinet | Corner filler base; BBC36/39/42 |
| `ERB[WW]` | End Refrigerator Base | Filler base beside refrigerator; ERB33/36 |
| `BEC[WW]` | Base End Cabinet | Corner piece, both doors operational; BEC24 |
| `BT[WW]` | Base Tray Cabinet | Narrow tray storage; BT09 = 9"W |
| `SB[WW]` | Sink Base | Open interior, dummy drawer(s). SB27 = 1 dummy; SB30/33/36 = 2 dummies |
| `FSB[WW]` | Farmhouse Sink Base | Modified front for apron sink; FSB36 |
| `CSB[WW]` | Corner Sink Base Cabinet | Corner unit with sink opening; CSB36 |
| `SBF[WW][WW]` | Corner Sink Base Floor | Flat base platform for corner sink; SBF4242 |
| `DCSF[WW]` | Diagonal Corner Sink Base Front | Front panel for SBF4242; DCSF42 |
| `MB[WW]` | Microwave Base Cabinet | Cabinet with microwave opening; MB30/33 |
| `SPB[WW]` | Spice Bull Base | Narrow pull-out spice rack; SPB6/9 |
| `BSDC[WW]` | Base Spice Drawer Cabinet | Multi-drawer spice unit, 6"W; BSDC6 |
| `BES[WW]` | Base End Shelf | Open rounded corner shelf; BES12 |

### Tall / Pantry Cabinets

| Code Pattern | Full Name | Notes |
|---|---|---|
| `U[WW][HH][DD]` | Utility Pantry | WW=18/24/30, HH=84/90/96, DD=24. Two doors, adjustable shelves |
| `O[WW][HH][DD]` | Oven Pantry | Oven opening (24"W × 24.375"H); WW=30/33, HH=84/90/96, DD=24 |

### Vanity Cabinets

| Code Pattern | Full Name | Notes |
|---|---|---|
| `VS[WW]` | Vanity Sink Base | VS24/27 = 1 dummy drawer; VS30/33/36 = 2 dummy drawers |
| `3VDB[WW]` | Vanity Three Drawer Base | All drawers, no doors; 12–24"W |
| `VSD[WW]` | Vanity Sink + Drawer Combo | Wider combo unit; VSD36/42/48 |
| `V[WW]21DL` | Vanity Combo Drawers Left | 21"D, drawers on left side; V3021DL/V3321DL/V3621DL |
| `V[WW]21DR` | Vanity Combo Drawers Right | 21"D, drawers on right side; V3021DR/V3321DR/V3621DR |
| `V[WW]TDL` | Vanity Combo Top Door Left *(frameless)* | Frameless version, door on left |
| `V[WW]TDR` | Vanity Combo Top Door Right *(frameless)* | Frameless version, door on right |

### Accessories & Moulding

| Code | Full Name | Notes |
|---|---|---|
| `COV` | Cove Crown Moulding | 96"L |
| `CM8` | Crown Moulding | 96"L |
| `BCM8` / `DCM` | Decorative Crown Moulding | 96"L |
| `RCM4S` / `LCM` | Inset Crown Moulding | 96"L |
| `ACM8` | Angle Crown Moulding | 96"L |
| `OCM8` | Outside Corner Moulding | 96"L |
| `ALRM8` | Angle Light Rail Moulding | 96"L |
| `LRM8` | Light Rail Moulding | 96"L |
| `SHM` | Shoe Moulding | 96"L |
| `SM8` / `SCR` | Scribe Moulding | 96"L |
| `BAM` | Batten Moulding | 96"L |
| `FBM` | Furniture Base Moulding | 96"L |
| `TKC` | Toe Kick Cover | 96"L × 4.5"H |
| `F342` / `F642` | Filler | 3" or 6"W × 42"H |
| `F396` / `F696` | Tall Filler | 3" or 6"W × 96"H |
| `FPV4296` | Finish Plywood Panel | 42"W × 96"H, 1/4" thick |
| `USV2496` | Tall Skin Veneer Panel | 23.25"W × 96"H, 1/4" thick |
| `BSV` | Base Skin Veneer Panel | 23.25"W × 34.5"H |
| `WSV[WW]` | Wall Skin Veneer Panel | 15"W × 24/30/36/42"H |
| `WRC[WW][HH]` | Wine Rack | See Wall Cabinets above |
| `ROPE` | Rope Moulding | **Discontinued** |
| `DMTIS` | Dentil Moulding | **Discontinued** |
| `DCW2712GD` / `DCW27[HH]` | Diagonal Corner Wall 27"W | **Discontinued** |

### FL Warehouse Special Rule

SKUs with a `-FL` suffix in the FL warehouse are the **same physical item** as the SKU without `-FL`. They are just named differently for that location. When validating or matching:
- `GR-2DB24-FL` = `GR-2DB24`
- Strip the `-FL` suffix before cross-referencing with other warehouses or the master SKU list.

### Door Set Items

Door set SKUs contain `-DS-` in the item code (e.g., `AG-DS-B24`, `SW-DS-W3012`).

For **old colors**: a single SKU covers both the door set and the cabinet box — do not expect a separate UBX item.

For **new/separate colors**: always expect a matching UBX item alongside the DS item.

### PWB- Items (Frameless)

Frameless colors use box SKUs starting with `PWB-` (e.g., `PWB-B24`). These replace the standard cabinet box for those colors. When validating frameless orders, check for `PWB-` items instead of standard box items.

## How to Use the Data

When the user asks about a specific company, read the corresponding JSON file. For cross-company queries, read all relevant files.

To validate a SKU:
1. Extract the color prefix (text before first `-`)
2. Look it up in the appropriate company's JSON file
3. Apply the packaging rules above (combined vs. separate DS/UBX, PWB- for frameless)
4. For FL, normalize `-FL` suffix before matching

## Invoice Validation

When validating an invoice:
1. Confirm every line-item SKU exists in the correct company's SKU list
2. Flag SKUs not found in the data (possible typo, discontinued, or wrong company)
3. Check packaging logic:
   - Old color orders: each SKU represents a complete set (no separate DS/UBX needed)
   - New/separate color orders: DS and UBX should appear as separate line items in matching quantities; cross-reference `new-color-breakdown.csv` (or the relevant breakdown file) to confirm the correct DS/UBX pair for each parent SKU
   - Frameless orders: expect PWB- items, not standard box items
4. FL-specific: normalize `-FL` suffixes when checking validity

## Sales Analysis

When analyzing sales data:
- Group by color prefix to compare performance across color lines
- `stock_qty` = current on-hand inventory
- `po_qty` = quantity on open purchase orders (in transit or ordered)
- Available inventory = `stock_qty + po_qty`
- Flag SKUs with zero stock and zero PO as potential stockouts
- Use color group (old/new/frameless) to segment analysis

## Common Tasks

**"Is [SKU] valid?"** → look up in the relevant company file, normalize FL suffix if needed
**"Validate this invoice"** → paste or share invoice data; check each line against SKU list and packaging rules
**"What's the stock for [SKU] at [company]?"** → read that company's JSON, return stock_qty and po_qty
**"Compare [SKU] across warehouses"** → check all 5 JSON files, normalize FL variants
**"Which SKUs are low on stock?"** → filter for stock_qty = 0 or below threshold in the specified company file
**"What's the volume/weight/FOB price for [SKU]?"** → read `fob-specs.json`, look up by full SKU key; strip `-FL` and normalize slash variants
**"How many containers does [warehouse] inventory fill?"** → read warehouse JSON + fob-specs.json, compute sum(stock_qty × vol_m3) ÷ 65
