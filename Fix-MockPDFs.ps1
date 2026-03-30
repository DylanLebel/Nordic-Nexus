# ==============================================================================
#  Fix-MockPDFs.ps1 - Nordic Minesteel Technologies
#  Generates VALID, viewable PDF files for testing.
# ==============================================================================

function New-ValidMockPDF {
    param([string]$Path, [string]$Content)
    
    # This is a 'Minimal PDF' header/structure that PDF readers will accept
    $pdfTemplate = @"
%PDF-1.1
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << >> >>
endobj
4 0 obj
<< /Length $( $Content.Length + 50 ) >>
stream
BT
/F1 24 Tf
100 700 Td
($Content) Tj
ET
endstream
endobj
xref
0 5
0000000000 65535 f
0000000009 00000 n
0000000058 00000 n
0000000115 00000 n
0000000212 00000 n
trailer
<< /Size 5 /Root 1 0 R >>
startxref
310
%%EOF
"@
    $pdfTemplate | Set-Content -Path $Path -Encoding Ascii
}

$assets = "drawing-collector-v3/Test-Assets"
Write-Host "[*] Upgrading mock files to valid PDFs..." -ForegroundColor Cyan

New-ValidMockPDF -Path "$assets/4823-P27_RevB.pdf" -Content "MOCK DRAWING: 4823-P27 Rev B"
New-ValidMockPDF -Path "$assets/4823-P30-L_RevADD.pdf" -Content "MOCK DRAWING: 4823-P30-L Rev ADD"

Write-Host "[+] Done! You can now open these in PDFgear." -ForegroundColor Green
