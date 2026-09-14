---
name: FW-NPI
description: Generate FW breakdown Excel file (Cabinet kit, ADA Cabinet kit, Door kit, Door, No NB Match) from FW demand data, following the same structure as the NB breakdown document. Use when the user wants to create or update the FW product breakdown.
---

# FW NPI Breakdown Skill

Generate `FW breakdown.xlsx` saved to `c:\Users\Jonathan\Desktop\FOB\NB-DS\`.

The file mirrors the NB breakdown structure (`NB breakdown - Final.xlsx`) but for FW-prefix products. Prices and descriptions are left blank for the user to fill in.

## Source Files

| File | Purpose |
|------|---------|
| `c:\Users\Jonathan\Desktop\FOB\NB-DS\NB breakdown - Final.xlsx` | Template — defines which cabinet types get DS+UBX breakdown and how shared door sets are named |
| `c:\Users\Jonathan\Desktop\FOB\美国办事处库存型号FOB明细-05-06-2026 执行.xlsx` | Sheet `SKU-体积重量FOB` — full list of all FW-XXXXX SKUs (136 items) |

## Output File

`c:\Users\Jonathan\Desktop\FOB\NB-DS\FW breakdown.xlsx` — 5 sheets:

| Sheet | Contents |
|-------|---------|
| Cabinet kit | Standard FW items with NB counterparts + manually added items (89 total) |
| ADA Cabinet kit | ADA-prefixed FW items (29 total) |
| Door kit | Simplified kit SKU → actual shared component mapping (non-ADA + ADA) |
| Door | All unique door set component SKUs (non-ADA + ADA) |
| No NB Match | FW items with no NB counterpart and no breakdown (fillers, accessories, etc.) |

## Column Layout (Cabinet kit & ADA Cabinet kit)

Matches NB breakdown exactly:

| Col | Header | Value |
|-----|--------|-------|
| A | SKU | FW-XXXXX |
| B | Description | (blank) |
| C | Fixed MSRP | (blank) |
| D | Component1 | FW-DS-XXXXX (door set) |
| E | Component Description | "Door Set XXXXX" |
| F | Componenet FOB | (blank) |
| G | Component MSRP | (blank) |
| H | Component2 | UBX-XXXXX (cabinet box) |
| I | Component Description | "Cabinet Box for XXXXX" |
| J | Componenet FOB | (blank) |
| K | Component MSRP | (blank) |

## Matching Logic

### Step 1 — NB code extraction
Read all rows from NB `Cabinet kit` sheet. Strip `NB-` prefix and `-KIT` suffix to get base codes (e.g. `NB-2DB24-KIT` → `2DB24`). Store each code's DS component and UBX box.

### Step 2 — FW item list
Read `SKU-体积重量FOB` sheet of the FOB file. Collect all rows where col A starts with `FW-`. Strip `FW-` prefix to get codes.

### Step 3 — Non-ADA matching
For each NB base code, check if the same code exists in the FW set. If yes → include in **Cabinet kit** using the transformed component names.

**DS transformation (shared component simplification):**
- NB DS may be shared: `NB-DS-3DB12/3VDB12`, `NB-DS-BFH12/W1230`, etc.
- Split on `/`, keep only parts where `FW-[part]` exists in FW set, replace `NB-DS-` with `FW-DS-`
- Examples:
  - `NB-DS-3DB12/3VDB12` → `FW-DS-3DB12/3VDB12` (both 3DB12 and 3VDB12 exist in FW)
  - `NB-DS-BFH12/W1230` → `FW-DS-W1230` (BFH12 not in FW, only W1230 kept)
  - `NB-DS-W3012/W301224` → `FW-DS-W3012` (W301224 not in FW)
  - `NB-DS-W3612/W361224` → `FW-DS-W3612/W361224` (both exist in FW)
  - `NB-DS-SB27/VS27` → `FW-DS-VS27` (SB27 not in FW)
  - `NB-DS-SB30/VS30` → `FW-DS-SB30/VS30` (both exist in FW)

**UBX transformation (shared box simplification):**
- Same logic: split on `/`, keep only parts existing in FW set
- Examples:
  - `UBX-SB36/FSB36` → `UBX-SB36` (FSB36 not in FW)
  - `UBX-W2430/WBC2430` → `UBX-W2430` (WBC2430 not in FW)

### Step 4 — Manually added items
These FW items have no NB counterpart but the user confirmed they need a breakdown. Add them to **Cabinet kit** with simple `FW-DS-XXXXX` + `UBX-XXXXX` (no sharing):

```
VB09, VB12, VB15, VB18, W3912, W4230, W4236, W4242
```

### Step 5 — ADA matching
For each `FW-ADA-XXXXX` item, strip `FW-ADA-` to get `adaCode`:

**Special R-suffix pairs** (hard-coded shared DS):
| ADA code | DS component |
|----------|-------------|
| SB33R | FW-DS-ADA-SB33R/VS33R |
| VS33R | FW-DS-ADA-SB33R/VS33R |
| SB36R | FW-DS-ADA-SB36R/VS36R |
| VS36R | FW-DS-ADA-SB36R/VS36R |

**All other ADA items:** look up `adaCode` in NB map, then transform DS:
- Split NB DS on `/`, for each part check if `FW-ADA-[part]` exists in FW set
- Keep matching parts, prefix with `FW-DS-ADA-`
- Examples:
  - `NB-DS-3DB12/3VDB12` → `FW-DS-ADA-3DB12/3VDB12` (both ADA-3DB12 and ADA-3VDB12 exist)
  - `NB-DS-3DB21/3VDB21` → `FW-DS-ADA-3DB21` (ADA-3VDB21 not in FW)
  - `NB-DS-BT09/W0930` → `FW-DS-ADA-BT09` (ADA-W0930 not in FW)
  - `NB-DS-SB27/VS27` → `FW-DS-ADA-VS27` (ADA-SB27 not in FW)
  - `NB-DS-SB33/VS33` → `FW-DS-ADA-VS33` (ADA-SB33 not in FW; ADA-VS33 exists)

UBX for ADA items is always `UBX-ADA-[adaCode]` (no sharing).

### Step 6 — Door kit sheet
For each kit item (non-ADA + ADA), if the simplified SKU differs from the actual component:
- Non-ADA: `FW-DS-[code]` ≠ actual DS → add entry mapping simplified → actual
- ADA: `FW-DS-ADA-[adaCode]` ≠ actual DS → add entry mapping simplified → actual

### Step 7 — Door sheet
Collect all unique actual DS components from Cabinet kit + ADA Cabinet kit. Sort and list with description `"Door Set [code]"`.

### Step 8 — No NB Match sheet
All FW items not matched in any of the above steps.

## PowerShell Implementation

```powershell
# Close any open FW breakdown file
try{
    $xl=[System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    foreach($wbx in @($xl.Workbooks)){if($wbx.Name -like '*FW breakdown*'){$wbx.Close($false)}}
}catch{}

$excel=New-Object -ComObject Excel.Application
$excel.Visible=$false; $excel.DisplayAlerts=$false

# Read NB mapping
$wbNB=$excel.Workbooks.Open('c:\Users\Jonathan\Desktop\FOB\NB-DS\NB breakdown - Final.xlsx')
$wsNB=$wbNB.Worksheets.Item(1); $nbMap=@{}
for($r=2;$r -le $wsNB.UsedRange.Rows.Count;$r++){
    $sku=$wsNB.Cells.Item($r,1).Text;$ds=$wsNB.Cells.Item($r,4).Text;$ubx=$wsNB.Cells.Item($r,8).Text
    if($sku){$nbMap[($sku -replace '^NB-','' -replace '-KIT$','')]=@{DS=$ds;UBX=$ubx}}
};$wbNB.Close($false)

# Read FW items
$wbFOB=$excel.Workbooks.Open('c:\Users\Jonathan\Desktop\FOB\美国办事处库存型号FOB明细-05-06-2026 执行.xlsx')
$wsFOB=$wbFOB.Worksheets.Item('SKU-体积重量FOB')
$fwSet=[System.Collections.Generic.HashSet[string]]::new()
$fwAll=[System.Collections.Generic.List[string]]::new()
for($r=1;$r -le $wsFOB.UsedRange.Rows.Count;$r++){
    $v=$wsFOB.Cells.Item($r,1).Text
    if($v -like 'FW-*'){$fwSet.Add(($v -replace '^FW-',''))|Out-Null;$fwAll.Add($v)}
};$wbFOB.Close($false)

function Simplify-DS($nbDS,$fwSet){
    $inner=($nbDS -replace '^NB-DS-','');$parts=$inner -split '/'
    $keep=@($parts|Where-Object{$fwSet.Contains($_)})
    if($keep.Count -gt 0){return 'FW-DS-'+($keep -join '/')};return 'FW-DS-'+$inner
}
function ADA-DS($nbDS,$fwSet){
    $inner=($nbDS -replace '^NB-DS-','');$parts=$inner -split '/'
    $keep=@($parts|Where-Object{$fwSet.Contains("ADA-$_")})
    if($keep.Count -gt 0){return 'FW-DS-ADA-'+($keep -join '/')};return 'FW-DS-ADA-'+$inner
}

# Non-ADA NB-matched items
$kitItems=[System.Collections.Generic.List[object]]::new()
$matchedCodes=[System.Collections.Generic.HashSet[string]]::new()
foreach($code in ($nbMap.Keys|Sort-Object)){
    if($fwSet.Contains($code)){
        $matchedCodes.Add($code)|Out-Null
        $ds=$nbMap[$code].DS;$ubx=$nbMap[$code].UBX
        $fwDS=Simplify-DS $ds $fwSet
        if($ubx -like 'UBX-*/*'){
            $inner=$ubx.Substring(4);$parts=$inner -split '/';$keep=@($parts|Where-Object{$fwSet.Contains($_)})
            $fwUBX=if($keep.Count -gt 0){'UBX-'+($keep -join '/')}else{$ubx}
        }else{$fwUBX=$ubx}
        $kitItems.Add(@("FW-$code",$fwDS,$fwUBX))
    }
}
# Manually confirmed items
$manualItems=@('VB09','VB12','VB15','VB18','W3912','W4230','W4236','W4242')
$manualCodes=[System.Collections.Generic.HashSet[string]]::new()
foreach($code in $manualItems){$kitItems.Add(@("FW-$code","FW-DS-$code","UBX-$code"));$manualCodes.Add($code)|Out-Null}
$kitItems=[System.Collections.Generic.List[object]]($kitItems|Sort-Object{$_[0]})

# ADA items
$adaRPairs=@{
    'SB33R'='FW-DS-ADA-SB33R/VS33R';'VS33R'='FW-DS-ADA-SB33R/VS33R'
    'SB36R'='FW-DS-ADA-SB36R/VS36R';'VS36R'='FW-DS-ADA-SB36R/VS36R'
}
$adaItems=[System.Collections.Generic.List[object]]::new()
$adaMatchedSkus=[System.Collections.Generic.HashSet[string]]::new()
foreach($fwSku in ($fwAll|Where-Object{$_ -like 'FW-ADA-*'}|Sort-Object)){
    $adaCode=$fwSku -replace '^FW-ADA-',''
    if($adaRPairs.ContainsKey($adaCode)){
        $adaItems.Add(@($fwSku,$adaRPairs[$adaCode],"UBX-ADA-$adaCode"))
        $adaMatchedSkus.Add($fwSku)|Out-Null;continue
    }
    if($nbMap.ContainsKey($adaCode)){
        $fwDS=ADA-DS $nbMap[$adaCode].DS $fwSet
        $adaItems.Add(@($fwSku,$fwDS,"UBX-ADA-$adaCode"))
        $adaMatchedSkus.Add($fwSku)|Out-Null
    }
}

# Unmatched
$unmatched=$fwAll|Where-Object{
    $code=$_ -replace '^FW-',''
    (-not $matchedCodes.Contains($code)) -and (-not $manualCodes.Contains($code)) -and (-not $adaMatchedSkus.Contains($_))
}

# Door kit (non-ADA + ADA)
$dkRows=[System.Collections.Generic.List[object]]::new();$seen=@{}
foreach($it in $kitItems){
    $code=$it[0] -replace '^FW-','';$simKit="FW-DS-$code"
    if($simKit -ne $it[1] -and -not $seen.ContainsKey($simKit)){$seen[$simKit]=$true;$dkRows.Add(@($simKit,"Door Set $code",$it[1]))}
}
foreach($it in $adaItems){
    $adaCode=$it[0] -replace '^FW-ADA-','';$simKit="FW-DS-ADA-$adaCode"
    if($simKit -ne $it[1] -and -not $seen.ContainsKey($simKit)){$seen[$simKit]=$true;$dkRows.Add(@($simKit,"Door Set ADA-$adaCode",$it[1]))}
}

# Door (all unique DS)
$uDS=[System.Collections.Generic.SortedSet[string]]::new()
foreach($it in $kitItems){$uDS.Add($it[1])|Out-Null}
foreach($it in $adaItems){$uDS.Add($it[1])|Out-Null}

# Build workbook
$wb=$excel.Workbooks.Add()
while($wb.Worksheets.Count -gt 1){$wb.Worksheets.Item($wb.Worksheets.Count).Delete()}
$hdrs=@("SKU","Description","Fixed MSRP","Component1","Component Description","Componenet FOB","Component MSRP ","Component2","Component Description","Componenet FOB","Component MSRP ")

function Write-KitSheet($ws,$items){
    for($c=1;$c -le $hdrs.Count;$c++){$ws.Cells.Item(1,$c).Value2=$hdrs[$c-1]}
    $r=2;foreach($it in $items){
        $ws.Cells.Item($r,1).Value2=$it[0]
        $ws.Cells.Item($r,4).Value2=$it[1]
        $ws.Cells.Item($r,5).Value2="Door Set $($it[1] -replace '^FW-DS-','')"
        $ws.Cells.Item($r,8).Value2=$it[2]
        $ws.Cells.Item($r,9).Value2="Cabinet Box for $($it[2] -replace '^UBX-','')"
        $r++
    }
}

$ws1=$wb.Worksheets.Item(1);$ws1.Name="Cabinet kit";Write-KitSheet $ws1 $kitItems
$ws2=$wb.Worksheets.Add([System.Type]::Missing,$ws1);$ws2.Name="ADA Cabinet kit";Write-KitSheet $ws2 $adaItems

$ws3=$wb.Worksheets.Add([System.Type]::Missing,$ws2);$ws3.Name="Door kit"
$h2=@("Kit 1","Kit 1 Description","Kit 1 Actual FOB","Kit 1 AdIusted MSRP","compoent")
for($c=1;$c -le $h2.Count;$c++){$ws3.Cells.Item(1,$c).Value2=$h2[$c-1]}
$r=2;foreach($dk in $dkRows){$ws3.Cells.Item($r,1).Value2=$dk[0];$ws3.Cells.Item($r,2).Value2=$dk[1];$ws3.Cells.Item($r,5).Value2=$dk[2];$r++}

$ws4=$wb.Worksheets.Add([System.Type]::Missing,$ws3);$ws4.Name="Door"
$h3=@("Components","Component Description","Componenet Actual FOB","Component MSRP ")
for($c=1;$c -le $h3.Count;$c++){$ws4.Cells.Item(1,$c).Value2=$h3[$c-1]}
$r=2;foreach($ds in $uDS){$ws4.Cells.Item($r,1).Value2=$ds;$ws4.Cells.Item($r,2).Value2="Door Set $($ds -replace '^FW-DS-','')";$r++}

$ws5=$wb.Worksheets.Add([System.Type]::Missing,$ws4);$ws5.Name="No NB Match"
$ws5.Cells.Item(1,1).Value2="SKU";$ws5.Cells.Item(1,2).Value2="Description"
$r=2;foreach($sku in $unmatched){$ws5.Cells.Item($r,1).Value2=$sku;$r++}

$out="c:\Users\Jonathan\Desktop\FOB\NB-DS\FW breakdown.xlsx"
$wb.SaveAs($out,51);$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)|Out-Null
Write-Output "Done: $out"
```

## Notes

- All price/MSRP/description columns are left blank for manual entry.
- The **manually added items** list (`VB09, VB12, VB15, VB18, W3912, W4230, W4236, W4242`) may grow — update the `$manualItems` array if the user provides additional SKUs to include.
- The **No NB Match** sheet lists remaining FW items (fillers, accessories, etc.) that are intentionally excluded from breakdown.
- If the source FOB file path or NB breakdown path changes, update the file paths in the script.
