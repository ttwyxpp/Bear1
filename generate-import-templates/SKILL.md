---
name: generate-import-templates
description: Generate PO and Inbound import templates from a supplier invoice Excel file. Use when the user wants to process an invoice, create import templates, or run the import template generator.
---

# Import Template Generator Skill

You generate a PO CSV file and an Inbound CSV import template from a supplier invoice Excel file. Collect all required inputs from the user, then execute the logic directly via PowerShell.

## Scripts (reference only — do not run interactively)

- PowerShell: `C:\Users\Jonathan\Desktop\Scripts\Generate-ImportTemplates.ps1`
- Python: `C:\Users\Jonathan\Desktop\Scripts\generate_import_templates.py`

These scripts use `Read-Host` and cannot be run non-interactively. Instead, perform the steps below directly.

## Inputs to Collect

Ask the user for these before proceeding:

1. **Invoice file path** — full path to the `.xlsx` invoice file
2. **Blackwater?** — is this invoice from Blackwater? (Y/N). If yes, all prices × 1.05
3. **PO number** — e.g. `POTX20945` (for Inbound template only). Auto-prefix with `PO#` if the user doesn't include it.

**Actual Shipping Date** is read automatically from the invoice at row 9, col 7. Do not ask the user for it.

## Invoice Layout (Excel, 1-indexed rows and columns)

| Cell | Field |
|---|---|
| Row 8, Col 7 | Invoice NO. |
| Row 9, Col 7 | Ship Date (used as Actual Shipping Date for Inbound) |
| Row 13, Col 7 | Container No. |
| Rows 17+ | Line items |

**Line item parsing (rows 17 to end):**
- If col1 starts with `PO#` AND col2 contains `Item` → PO section header row; set current section PO to col1, skip
- If col1 starts with `TOTAL` → stop
- Otherwise, if col2 (SKU) is non-empty AND col3 (price) is non-empty AND col4 (qty) is numeric → data row

**Price parsing:** strip leading `$` and commas from col3, parse as float. If Blackwater: `round(price * 1.05, 2)`.

## Deduplication

After parsing all line items, deduplicate by SKU: if the same SKU appears more than once, sum their quantities into a single row. Use the price from the first occurrence (duplicate rows within the same invoice always have the same price).

```powershell
$deduped = @{}
foreach ($item in $items) {
    if ($deduped.ContainsKey($item.SKU)) {
        $deduped[$item.SKU].Qty += $item.Qty
    } else {
        $deduped[$item.SKU] = [PSCustomObject]@{ SKU=$item.SKU; Price=$item.Price; Qty=$item.Qty }
    }
}
$items = $deduped.Values
```

## SKU Validation

Validate all SKUs against TX warehouse:

```powershell
$txData = Get-Content 'C:\Users\Jonathan\.claude\skills\Skyline-SKU\sku-data\TX.json' -Raw | ConvertFrom-Json
$txSkus = $txData | ForEach-Object { $_.sku }
```

- Report any SKUs not found in `$txSkus`
- Proceed automatically (do not ask the user whether to continue)

## External ID Format

`yyyyMMddhmm` — current date + 12-hour clock hour (no leading zero) + 2-digit minutes

```powershell
$h12  = if ((Get-Date).Hour % 12 -eq 0) { 12 } else { (Get-Date).Hour % 12 }
$extId = (Get-Date -Format "yyyyMMdd") + $h12 + (Get-Date -Format "mm")
```

## Output Paths

```
Dir:     C:\Users\Jonathan\Desktop\PO Invoice Issue\
PO:      <container>_<safeName>.csv
Inbound: <container>_Inbound.csv
```

`<safeName>` = `invoiceNo` with characters `# / \ : * ? " < > |` replaced by `-`, then stripped of leading/trailing `-`.

```powershell
$safeName = ($invoiceNo -replace '[#/\\:*?"<>|]', '-').Trim('-')
$outDir   = "C:\Users\Jonathan\Desktop\PO Invoice Issue"
$poPath   = "$outDir\$container`_$safeName.csv"
$inPath   = "$outDir\$container`_Inbound.csv"
```

## PO Template — CSV

Output the PO as a CSV file. Note any invalid SKUs in the summary text instead of highlighting.

**Columns** (row 1 = header):

| A | B | C | D |
|---|---|---|---|
| `Item  ` | `qty` | `External ID` | `Item Rate` |

One row per deduplicated item.

```powershell
$csvLines = @('"Item  ","qty","External ID","Item Rate"')
foreach ($item in $items) {
    $csvLines += "$($item.SKU),$($item.Qty),$extId,$($item.Price)"
}
$csvLines | Out-File $poPath -Encoding utf8
```

Invalid SKUs: list them in the summary report (cannot highlight in CSV).

## Inbound Template CSV

Header row:
```
receiving Location,Actual Shipping Date,vessel number,Expected delivery Date,po,external id,Memo,Item,Quantity expected
```

One row per deduplicated item:
```
Main-Texas,<actualShipDate>,<container>,<expectedDelivery>,<inboundPO>,<extId>,<container>,<SKU>,<qty>
```

**Expected delivery Date** = Actual Shipping Date + 45 days. Parse the ship date from the invoice, add 45 days, and format as `M/d/yyyy`.

```powershell
$expectedDelivery = ([datetime]::Parse($shipDate)).AddDays(45).ToString("M/d/yyyy")
```

Write with `-Encoding utf8`.

## Execution Steps

1. Collect inputs (ask the user)
2. Open the invoice file via Excel COM (ReadOnly):
   ```powershell
   $excel = New-Object -ComObject Excel.Application
   $excel.Visible = $false
   $excel.DisplayAlerts = $false
   $wb = $excel.Workbooks.Open($InvoicePath, $null, $true)
   $ws = $wb.Sheets.Item(1)
   ```
3. Read `invoiceNo`, `shipDate` (row 9 col 7), and `container`
4. Loop rows 17 to `$ws.UsedRange.Rows.Count`, parse line items
5. Close invoice Excel and release COM objects
6. Deduplicate items by SKU (sum quantities)
7. Validate SKUs against TX JSON; report invalid ones
8. Generate external ID
9. Write PO CSV; note any invalid SKUs in the summary
10. Write Inbound CSV
11. Report summary: invoice no, container, item count (after dedup), any invalid SKUs, external ID, Blackwater flag, output paths

## Common Tasks

**"Generate templates from this invoice"** → collect inputs, run steps above, report output paths  
**"Run the import template generator"** → same as above  
**"Process invoice [path]"** → use provided path, still ask for remaining inputs
