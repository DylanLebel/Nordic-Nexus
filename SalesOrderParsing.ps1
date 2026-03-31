function Invoke-SalesOrderParseLog {
    param(
        [scriptblock]$Logger,
        [string]$Message,
        [string]$Level = "INFO"
    )

    if ($null -eq $Logger) { return }
    try {
        & $Logger $Message $Level
    } catch { }
}

function Parse-SalesOrderText {
    param(
        [string]$Text,
        [scriptblock]$Logger = $null
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        Invoke-SalesOrderParseLog -Logger $Logger -Message "  No sales-order text supplied to parser" -Level "WARN"
        return @()
    }

    $lines = $Text -split "`n"
    $parts = [System.Collections.Generic.List[object]]::new()
    $orderNumber = "UNKNOWN"
    $inTable = $false
    $currentPart = $null
    $skipDescPrefixes = @('REL ', 'NEED BY', 'SHIP BY', 'QUANTITY', 'UNIT PRICE', 'EXT. PRICE', 'PAGE ', 'SALES ORDER ACKNOWLEDG', 'ORDERACK:', 'LINE TOTAL', 'TOTAL TAX', 'ORDER TOTAL', 'CANADIAN DOLLARS', 'EXT. ', 'LINE MISCELLANEOUS', 'ORDER MISCELLANEOUS')

    $normalizeQtyText = {
        param([string]$q)
        if ([string]::IsNullOrWhiteSpace($q)) { return "" }
        $x = ([string]$q).Trim()
        $x = $x -replace ',', '.'
        $x = $x -replace '\.0+$', ''
        return $x
    }

    $parsePdfRow = {
        param([string]$rowText)
        if ([string]::IsNullOrWhiteSpace($rowText)) { return $null }
        $t = ($rowText -replace '\s+$', '').Trim()
        $m = [regex]::Match(
            $t,
            '^(?<line>\d{1,3})\s+(?<part>[A-Z0-9][A-Z0-9._-]{3,})\s+(?<rev>[A-Z0-9]+)\b',
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
        )
        $lineNo = ""
        $partNum = ""
        $rev = ""
        $qty = ""
        $matchedWithoutRev = $false

        if ($m.Success) {
            $lineNo = [string]$m.Groups['line'].Value.Trim()
            $partNum = [string]$m.Groups['part'].Value.Trim().ToUpperInvariant()
            $rev = [string]$m.Groups['rev'].Value.Trim().ToUpperInvariant()
        } else {
            $m = [regex]::Match(
                $t,
                '^(?<line>\d{1,3})\s+(?<part>[A-Z0-9][A-Z0-9._-]{3,})(?<rest>.*)$',
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
            )
            if (-not $m.Success) { return $null }
            $lineNo = [string]$m.Groups['line'].Value.Trim()
            $partNum = [string]$m.Groups['part'].Value.Trim().ToUpperInvariant()
            $matchedWithoutRev = $true
        }

        if ($partNum -notmatch '^[A-Z0-9][A-Z0-9._-]{3,}$') { return $null }
        if ($partNum -match '(?i)^(JOB|TOTAL|SUB|TAX|NET|FREIGHT|CHARGES)') { return $null }
        if ($partNum -match '^\d{1,2}-[A-Z]{3}-\d{2,4}$') { return $null }
        if (-not $matchedWithoutRev -and $rev -notmatch '^[A-Z0-9]+$') { return $null }

        if (-not $matchedWithoutRev) {
            $revEndPos = $m.Groups['rev'].Index + $m.Groups['rev'].Length
            if ($revEndPos -lt $t.Length) {
                $afterRev = $t.Substring($revEndPos)
                if ($afterRev -match '^[.,]\d') {
                    $rev = ""
                }
            }
        }

        $qtyMatch = [regex]::Match($t, '(?<qty>\d+(?:[.,]\d+)?)EA\b', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($qtyMatch.Success) {
            $qty = & $normalizeQtyText $qtyMatch.Groups['qty'].Value
        }

        return [ordered]@{
            Line = $lineNo
            Part = $partNum
            Rev  = $rev
            Qty  = $qty
        }
    }

    $isLikelyDescriptionLine = {
        param([string]$text)
        if ([string]::IsNullOrWhiteSpace($text)) { return $false }
        $u = $text.Trim().ToUpperInvariant()
        if ($u -match '^\d{1,3}\s+') { return $false }
        foreach ($prefix in $skipDescPrefixes) {
            if ($u.StartsWith($prefix)) { return $false }
        }
        if ($u -match '(?i)SALES\s+ORDER\s+ACKNOWLEDG') { return $false }
        if ($u -match '\d+\s+OF\s+\d+') { return $false }
        if ($u -match '^\d+\s+\d{1,2}-[A-Z]{3}-\d{4}\b') { return $false }
        if ($u -match '^\d[\d\s,\.]+$') { return $false }
        if ($u -match '^[A-Z]') {
            return ($u -match '[A-Z]{3,}')
        }
        if ($u -match '\b\d+(?:[.,]\d+)?EA\b') { return $false }
        if ($u -match '\b\d{1,3}(?:,\d{3})+\.\d{2}\b') { return $false }
        return ($u -match '[A-Z]{3,}')
    }

    $isPriceNoiseLine = {
        param([string]$text)
        if ([string]::IsNullOrWhiteSpace($text)) { return $false }
        $u = $text.Trim().ToUpperInvariant()
        if ($u -match '^[A-Z]{3,}') { return $false }
        if ($u -match '\b\d+(?:[.,]\d+)?EA\b') { return $true }
        if ($u -match '\b\d{1,3}(?:,\d{3})*\.\d{2,3}\b') { return $true }
        if ($u -match '^\d[\d\s,\.EA]+$') { return $true }
        return $false
    }

    $startsNewNonPartRow = {
        param([string]$text)
        if ([string]::IsNullOrWhiteSpace($text)) { return $false }
        $m = [regex]::Match(
            $text.Trim(),
            '^(?<line>\d{1,3})\s+(?<token>[A-Z][A-Z0-9/_-]{1,})\b',
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
        )
        if (-not $m.Success) { return $false }
        $token = [string]$m.Groups['token'].Value.Trim().ToUpperInvariant()
        return ($token -match '^(JOB|TOTAL|SUB|TAX|NET|FREIGHT|CHARGES|LINE|ORDER)$')
    }

    $stripTrailingPriceNoise = {
        param([string]$text)
        if ([string]::IsNullOrWhiteSpace($text)) { return "" }
        $cleaned = [regex]::Replace($text, '\s+\d+(?:[.,]\d+)?EA\b.*$', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        $cleaned = [regex]::Replace($cleaned, '\s+\d{1,3}(?:,\d{3})*\.\d{2,3}\s*$', '')
        return $cleaned.Trim()
    }

    $tryParseQuantityLine = {
        param([string]$text)
        if ([string]::IsNullOrWhiteSpace($text)) { return "" }
        $m = [regex]::Match($text.Trim(), '^\d+\s+\d{1,2}-[A-Z]{3}-\d{4}\s+\d{1,2}-[A-Z]{3}-\d{4}\s+(?<qty>\d+(?:[.,]\d+)?)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($m.Success) { return (& $normalizeQtyText $m.Groups['qty'].Value) }
        return ""
    }

    $tryParseSplitRevLine = {
        param([string]$text)
        if ([string]::IsNullOrWhiteSpace($text)) { return "" }
        $t = $text.Trim()
        $m = [regex]::Match($t, '^(?<rev>[A-Z0-9]{1,4})\s+\d{1,3}(?:,\d{3})*\.\d{2,3}$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($m.Success) { return [string]$m.Groups['rev'].Value.Trim().ToUpperInvariant() }
        return ""
    }

    $flushCurrentPart = {
        if ($null -eq $currentPart) { return }
        $desc = [string]$currentPart.Description
        $desc = [regex]::Replace($desc, '\s+', ' ').Trim()
        $currentPart.Description = $desc
        [void]$parts.Add([PSCustomObject]@{
            Part         = [string]$currentPart.Part
            Order        = [string]$currentPart.Order
            Line         = [string]$currentPart.Line
            Rev          = [string]$currentPart.Rev
            Qty          = [string]$currentPart.Qty
            Description  = [string]$currentPart.Description
            InternalPart = [string]$currentPart.InternalPart
        })
    }

    foreach ($lineRaw in $lines) {
        $line = $lineRaw.Trim()
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        if ($line -match '(?i)Sales Order:\s*(\d+)') {
            $orderNumber = $matches[1]
        }

        if ($line -match '(?i)Line\s+Part\s+Number') {
            $inTable = $true
            continue
        }

        if (-not $inTable) { continue }

        $parsedRow = & $parsePdfRow $line
        if ($null -ne $parsedRow) {
            & $flushCurrentPart
            $currentPart = [ordered]@{
                Part         = [string]$parsedRow.Part
                Order        = $orderNumber
                Line         = [string]$parsedRow.Line
                Rev          = [string]$parsedRow.Rev
                Qty          = [string]$parsedRow.Qty
                Description  = ""
                InternalPart = ""
                AwaitQty     = $false
                AwaitSplitRev = [string]::IsNullOrWhiteSpace([string]$parsedRow.Rev)
            }
            continue
        }

        if ($null -eq $currentPart) { continue }

        if (& $startsNewNonPartRow $line) {
            & $flushCurrentPart
            $currentPart = $null
            continue
        }
        if ($line -match '^(?i)Rel\s+Need\s+By\s+Ship\s+By\s+Quantity\b') {
            $currentPart.AwaitQty = $true
            continue
        }
        if ([bool]$currentPart.AwaitQty) {
            $qtyFromLine = & $tryParseQuantityLine $line
            if (-not [string]::IsNullOrWhiteSpace($qtyFromLine)) {
                $currentPart.Qty = $qtyFromLine
                $currentPart.AwaitQty = $false
                continue
            }
        }
        if ([bool]$currentPart.AwaitSplitRev) {
            $revFromLine = & $tryParseSplitRevLine $line
            if (-not [string]::IsNullOrWhiteSpace($revFromLine)) {
                $currentPart.Rev = $revFromLine
                $currentPart.AwaitSplitRev = $false
                continue
            }
        }
        if ($line -match '(?i)(?:Our\s+Part|NMT\s+Part|Internal\s+Part)[:#\s]+\s*([A-Z0-9][A-Z0-9._-]{3,})') {
            $currentPart.InternalPart = $matches[1].Trim().ToUpperInvariant()
            continue
        }
        if (& $isPriceNoiseLine $line) {
            continue
        }
        if (& $isLikelyDescriptionLine $line) {
            $currentPart.AwaitSplitRev = $false
            $descText = & $stripTrailingPriceNoise $line
            if (-not [string]::IsNullOrWhiteSpace($descText)) {
                if ([string]::IsNullOrWhiteSpace($currentPart.Description)) {
                    $currentPart.Description = $descText
                } else {
                    $currentPart.Description += " " + $descText
                }
            }
            continue
        }
        if ($line -match '^\d{1,3}\s+\S+') {
            & $flushCurrentPart
            $currentPart = $null
            continue
        }
    }

    & $flushCurrentPart
    Invoke-SalesOrderParseLog -Logger $Logger -Message "  Extracted $($parts.Count) parts from Sales Order $orderNumber" -Level "INFO"
    return @($parts)
}
