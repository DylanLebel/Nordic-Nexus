# ==============================================================================
#  Test-Extraction.ps1 - Nordic Minesteel Technologies
#  Diagnostic runner for Epicor Order Extraction
# ==============================================================================

$scriptDir = Split-Path $PSCommandPath -Parent
$extractor = Join-Path $scriptDir "Extract-EpicorOrderParts.ps1"

# --- Example 1: Order 26816 ---
$ocr1 = @"
Sales Order Acknowledgement 1 of 2
Sales Order: 26816
Sold To:
Freeport-McMoran/P.T. Freeport Indonesia
333 North Central Ave.
United States
Phoenix AZ 85004-2121
Ship To:
P.T. Freeport Indonesia
C/O DSV Air and Sea
Door 217-219
2200 Yukon Court
Milton ONTARIO L9E 1N5
Canada
Phone: 504-458-4386
Order Date:09-Jun-2025
Need By:
Terms:Net 15
21-Aug-2025
PO Number:
Sales Person:
Ship Via:Land
Philip Stairmand
5200950400 Shipping Terms:
Tax ID:13-2578-365
Ex Works
Quote #:12274
Color Code:BLACK
Lead time 8-9 weeks from receipt of PO
United States Dollar
Line Part Number/Description Rev Order Qty Unit Price Ext. Price
1 4823-P4-34 11 1.00EA 2,770.080 2,770.08
RUBBER
Rel Date Quantity
1 21-Aug-2025 1.00
Supplied w Hardware
2 FHCS-0.625-11x2.000-ZP F 1 6.00EA 0.000 0.00
FLAT HEAD 5/8-11 X 2 Lg - ZINC PLATED - FULL THRD
Rel Date Quantity
1 21-Aug-2025 6.00
Hardware for 4823-P4-34. Qty 6 required per unit
3 FW-0.625-SAE-ZP NA 6.00EA 0.000 0.00
FLAT WASHER 5/8 - SAE - ZINC PLATED
Rel Date Quantity
1 21-Aug-2025 6.00
Hardware for 4823-P4-34. Qty 6 required per unit
4 HN-0.625-11-G5-ZP 1 6.00EA 0.000 0.00
HEX NUT 5/8-11 GR 5 - ZINC PLATED
Rel Date Quantity
"@

# --- Example 2: Order 27111 ---
$ocr2 = @"
Sales Order Acknowledgement 1 of 1
Sales Order: 27111
Sold To:
Freeport-McMoran/P.T. Freeport Indonesia
333 North Central Ave.
Phoenix AZ 85004-2121 United States
Ship To:
P.T. Freeport Indonesia
C/O DSV Air and Sea
Door 217-219
2200 Yukon Court
Milton ONTARIO L9E 1N5
Canada
Phone: 504-458-4386
Order Date:12-Mar-2026
Need By:
Terms:Net 60
08-Jun-2026
PO Number:
Sales Person:
Ship Via:Land
Dylan McFee
5201015116 Shipping Terms:
Tax ID:13-2578-365
Ex Works
Quote #:12634
COLOUR CODE : RED
Lead time: 10-12 Weeks from receipt of revised Purchase Order
United States Dollar
Line Part Number/Description Rev Order Qty Unit Price Ext. Price
1 BRG03 1 10.00EA 237.750 2,377.50
TAPERED ROLLER BRG
Rel Need By Ship By Quantity
1 08-Jun-2026 01-Jun-2026 10.00
2 4823-P6 3 4.00EA 4,910.000 19,640.00
HINGE ASSEMBLY
Rel Need By Ship By Quantity
1 08-Jun-2026 01-Jun-2026 4.00
3 JOB CHARGES add 
oper
0.00EA 0.000 0.00
JOB/ ORDER ADDITIONAL CHARGES
"@

Write-Host "Running Extraction Tests..." -ForegroundColor Yellow
powershell.exe -File $extractor -OcrText $ocr1
powershell.exe -File $extractor -OcrText $ocr2
