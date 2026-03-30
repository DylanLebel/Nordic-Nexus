# ==============================================================================
#  Create-TestOrder.ps1 - Nordic Minesteel Technologies
#  Utility to generate safe, redacted test assets for offline development.
# ==============================================================================

param(
    [string]$OrderNumber = "99999",
    [string]$JobPrefix   = "99999-10",
    [string]$OutPath     = "drawing-collector-v3/Test-Assets/Test-Order-$OrderNumber.txt"
)

$script:testParts = @(
    @{ Part = "4823-P27"; Rev = "B"; Desc = "TRIP ROLLER" },
    @{ Part = "4823-P30-L"; Rev = "ADD"; Desc = "LATCH ARM PLATE WELDMENT LEFT" }
)

# --- 1. Create a "Safe-for-Home" Mock OCR Text File ---
# Based on SO 25136 with company names redacted for privacy.
$content = @"
Nordic Minesteel Technologies Inc.
SALES ORDER ACKNOWLEDGEMENT

Customer: REDACTED CUSTOMER CORP
Sales Order: $OrderNumber
Date: 2021-01-15

Line  Part Number/Description                      Rev  Order Qty
----  -------------------------------------------  ---  ---------
1     $($testParts[0].Part)                           $($testParts[0].Rev)    2.00EA
      $($testParts[0].Desc)
      Rel Date: 2021-03-22  Quantity: 2.00
      Our Part: $($testParts[0].Part)

2     $($testParts[1].Part)                           $($testParts[1].Rev)    1.00EA
      $($testParts[1].Desc)
      Rel Date: 2021-03-22  Quantity: 1.00
      Our Part: $($testParts[1].Part)

Line Total: 2,348.20
Order Total: 2,348.20
"@

$content | Set-Content -Path $OutPath -Encoding UTF8
Write-Host "[+] Created Mock OCR Data: $OutPath" -ForegroundColor Green

# --- 2. Create a Mock Drawing Index ---
$indexPath = "drawing-collector-v3/Test-Assets/mock_index.csv"
$indexData = @()
foreach ($p in $testParts) {
    if ($p.Part -notmatch "FHCS") { # Only mock drawings for NMT parts
        $indexData += [PSCustomObject]@{
            PartNumber  = $p.Part
            Revision    = $p.Rev
            Description = $p.Desc
            FullPath    = "C:\MockVault\$($p.Part)_Rev$($p.Rev).pdf"
            FileName    = "$($p.Part)_Rev$($p.Rev).pdf"
            Extension   = ".pdf"
            LastWrite   = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        }
    }
}

$indexData | Export-Csv -Path $indexPath -NoTypeInformation -Encoding UTF8
Write-Host "[+] Created Mock Drawing Index: $indexPath" -ForegroundColor Green

# --- 3. Create Mock PDF Files ---
foreach ($item in $indexData) {
    $mockPdfPath = "drawing-collector-v3/Test-Assets/$($item.FileName)"
    "MOCK PDF CONTENT FOR $($item.PartNumber)" | Set-Content -Path $mockPdfPath
    Write-Host "    -> Created dummy file: $mockPdfPath"
}
