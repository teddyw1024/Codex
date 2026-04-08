param(
    [Parameter(Position = 0)]
    [string]$Serial,
    [switch]$AsJson
)

$ErrorActionPreference = "Stop"

function Remove-HtmlTags {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ""
    }

    return (($Value -replace "<[^>]+>", " ") -replace "\s+", " ").Trim()
}

function Try-ParseDate {
    param(
        [string]$Value,
        [ref]$OutDate
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $false
    }

    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse($Value, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeLocal, [ref]$parsed)) {
        $OutDate.Value = $parsed.Date
        return $true
    }

    if ([datetime]::TryParse($Value, [ref]$parsed)) {
        $OutDate.Value = $parsed.Date
        return $true
    }

    return $false
}

function Invoke-TextGet {
    param(
        [string]$Url,
        [hashtable]$Headers
    )

    $params = @{
        Uri             = $Url
        Method          = "Get"
        TimeoutSec      = 30
        UseBasicParsing = $true
    }

    if ($Headers) {
        $params.Headers = $Headers
    }

    return (Invoke-WebRequest @params).Content
}

function Get-LenovoWarranty {
    param([string]$InputSerial)

    $headers = @{
        "User-Agent"      = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
        "Accept-Language" = "en-US,en;q=0.9"
        "Referer"         = "https://pcsupport.lenovo.com/us/en/warranty-lookup"
    }

    $productIds = @($InputSerial)
    if (-not $InputSerial.ToLowerInvariant().EndsWith(".ibase")) {
        $productIds += "$InputSerial.ibase"
    }
    $fallbackResult = $null

    foreach ($productId in ($productIds | Select-Object -Unique)) {
        $encodedProductId = [System.Uri]::EscapeDataString($productId)
        $url = "https://pcsupport.lenovo.com/us/en/api/v4/warranty/getwarrantyandrepair?productid=$encodedProductId"

        try {
            $response = Invoke-RestMethod -Method Get -Uri $url -Headers $headers -TimeoutSec 30
        } catch {
            continue
        }

        if (-not ($response.PSObject.Properties.Name -contains "Warranty")) {
            continue
        }

        $warranty = $response.Warranty
        if ($null -eq $warranty -or $warranty -is [string]) {
            continue
        }

        $serialFound = "$($warranty.Serial)"
        $baseProductId = "$($warranty.BaseProductId)"
        $productPath = "$($warranty.ProductId)"
        if ([string]::IsNullOrWhiteSpace($serialFound) -and [string]::IsNullOrWhiteSpace($baseProductId) -and [string]::IsNullOrWhiteSpace($productPath)) {
            continue
        }

        $dateCandidates = New-Object "System.Collections.Generic.List[datetime]"
        $startDateCandidates = New-Object "System.Collections.Generic.List[datetime]"
        $coverageGroups = @(
            "BaseWarranties",
            "UpmaWarranties",
            "ContractWarranties",
            "AodWarranties",
            "InstantWarranties",
            "SaeWarranties",
            "BaseUpmaWarranties"
        )

        foreach ($groupName in $coverageGroups) {
            $groupItems = $warranty.$groupName
            if ($null -eq $groupItems) {
                continue
            }

            foreach ($item in @($groupItems)) {
                $parsedEndDate = $null
                if (Try-ParseDate -Value "$($item.End)" -OutDate ([ref]$parsedEndDate)) {
                    [void]$dateCandidates.Add($parsedEndDate)
                }

                $parsedStartDate = $null
                if (Try-ParseDate -Value "$($item.Start)" -OutDate ([ref]$parsedStartDate)) {
                    [void]$startDateCandidates.Add($parsedStartDate)
                }
            }
        }

        if ($dateCandidates.Count -eq 0 -and $warranty.EntireWarrantyPeriod -and $warranty.EntireWarrantyPeriod.End) {
            try {
                $epochMs = [int64]$warranty.EntireWarrantyPeriod.End
                $fallbackDate = [DateTimeOffset]::FromUnixTimeMilliseconds($epochMs).DateTime.Date
                [void]$dateCandidates.Add($fallbackDate)
            } catch {
            }
        }

        if ($startDateCandidates.Count -eq 0 -and $warranty.EntireWarrantyPeriod -and $warranty.EntireWarrantyPeriod.Start) {
            try {
                $startEpochMs = [int64]$warranty.EntireWarrantyPeriod.Start
                $fallbackStartDate = [DateTimeOffset]::FromUnixTimeMilliseconds($startEpochMs).DateTime.Date
                [void]$startDateCandidates.Add($fallbackStartDate)
            } catch {
            }
        }

        $expirationDate = $null
        if ($dateCandidates.Count -gt 0) {
            $expirationDate = ($dateCandidates | Sort-Object | Select-Object -Last 1).ToString("yyyy-MM-dd")
        }

        $startDate = $null
        if ($startDateCandidates.Count -gt 0) {
            $startDate = ($startDateCandidates | Sort-Object | Select-Object -First 1).ToString("yyyy-MM-dd")
        }

        $status = "Unknown"
        if ($warranty.PSObject.Properties.Name -contains "RemainingDays") {
            try {
                $remainingDays = [int]$warranty.RemainingDays
                if ($remainingDays -gt 0) {
                    $status = "Active"
                } else {
                    $status = "Expired"
                }
            } catch {
            }
        }

        if ($status -eq "Unknown" -and $expirationDate) {
            $parsedExpiration = [datetime]::ParseExact($expirationDate, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)
            if ($parsedExpiration -ge (Get-Date).Date) {
                $status = "Active"
            } else {
                $status = "Expired"
            }
        }

        $basicWarrantyUrl = "https://pcsupport.lenovo.com/us/en/basicwarrantylookup?sn=$([System.Uri]::EscapeDataString($InputSerial))"

        $model = "$($warranty.ProductName)".Trim()
        if ([string]::IsNullOrWhiteSpace($model)) {
            $machineType = "$($warranty.MachineType)".Trim()
            $mtm = "$($warranty.MTM)".Trim()
            if (-not [string]::IsNullOrWhiteSpace($machineType) -and -not [string]::IsNullOrWhiteSpace($mtm)) {
                $model = "$machineType-$mtm"
            } elseif (-not [string]::IsNullOrWhiteSpace($machineType)) {
                $model = $machineType
            } elseif (-not [string]::IsNullOrWhiteSpace($mtm)) {
                $model = $mtm
            } else {
                $model = $null
            }
        }

        $specUrl = $null
        $specAvailable = "Unknown"
        $productPathSource = "$($warranty.FullProductId)"
        if ([string]::IsNullOrWhiteSpace($productPathSource)) {
            $productPathSource = "$($warranty.ProductId)"
        }

        if (-not [string]::IsNullOrWhiteSpace($productPathSource)) {
            $normalizedProductPath = ($productPathSource -replace "\\", "/").Trim()
            $normalizedProductPath = $normalizedProductPath.TrimStart("/").TrimEnd("/")
            if (-not [string]::IsNullOrWhiteSpace($normalizedProductPath)) {
                $specUrl = "https://pcsupport.lenovo.com/us/en/products/$normalizedProductPath"
                $specAvailable = "Yes"
            }
        }

        if ([string]::IsNullOrWhiteSpace($specUrl)) {
            $specUrl = $basicWarrantyUrl
        }

        $note = $null
        if (-not $expirationDate) {
            $note = "Warranty record was found, but no explicit expiration date was returned."
        }

        $candidateResult = [pscustomobject]@{
            Brand          = "Lenovo"
            Serial         = $InputSerial
            Model          = $model
            SpecAvailable  = $specAvailable
            SpecUrl        = $specUrl
            StartDate      = $startDate
            ExpirationDate = $expirationDate
            Status         = $status
            ResultUrl      = $basicWarrantyUrl
            Notes          = $note
        }

        if ($null -eq $fallbackResult) {
            $fallbackResult = $candidateResult
        }

        $hasRichProductData = -not [string]::IsNullOrWhiteSpace("$($warranty.ProductName)") -or
            -not [string]::IsNullOrWhiteSpace("$($warranty.ProductId)") -or
            -not [string]::IsNullOrWhiteSpace("$($warranty.FullProductId)")

        if ($hasRichProductData) {
            return $candidateResult
        }
    }

    return $fallbackResult
}

function Get-DellWarranty {
    param([string]$InputSerial)

    if ($InputSerial -notmatch "^[A-Za-z0-9]{7}$") {
        return $null
    }

    $headers = @{
        "Accept-Language" = "en-us"
        "Accept-Encoding" = "identity"
        "Content-Type"    = "application/x-www-form-urlencoded"
        "Origin"          = "https://support.dell.com"
        "User-Agent"      = "Mozilla/5.0"
    }

    if ($env:DELL_ABCK) {
        $headers["Cookie"] = "_abck=$($env:DELL_ABCK)"
    }

    $overviewUrl = "https://www.dell.com/support/home/en-us/product-support/servicetag/$InputSerial/overview"
    $specUrl = "https://www.dell.com/support/home/en-us/product-support/servicetag/$InputSerial/docs"

    try {
        $overviewHtml = Invoke-TextGet -Url $overviewUrl -Headers $headers
    } catch {
        return [pscustomobject]@{
            Brand          = "Dell"
            Serial         = $InputSerial
            Model          = $null
            SpecAvailable  = "Unknown"
            SpecUrl        = $specUrl
            StartDate      = $null
            ExpirationDate = $null
            Status         = "Unknown"
            ResultUrl      = $overviewUrl
            Notes          = "Dell page could not be fetched automatically. Open the result URL to confirm warranty details."
        }
    }

    if ($overviewHtml -match "(?i)access denied|you don.?t have permission") {
        return [pscustomobject]@{
            Brand          = "Dell"
            Serial         = $InputSerial
            Model          = $null
            SpecAvailable  = "Unknown"
            SpecUrl        = $specUrl
            StartDate      = $null
            ExpirationDate = $null
            Status         = "Unknown"
            ResultUrl      = $overviewUrl
            Notes          = "Dell blocked automated access from this network. Open the result URL to confirm warranty details."
        }
    }

    if ($overviewHtml -match "(?i)IsInvalidSelection=True|invalid service tag") {
        return $null
    }

    $model = $null
    $modelMatch = [regex]::Match($overviewHtml, "<h1[^>]*>(?<model>.*?)</h1>", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if ($modelMatch.Success) {
        $cleanModel = Remove-HtmlTags -Value $modelMatch.Groups["model"].Value
        if (-not [string]::IsNullOrWhiteSpace($cleanModel)) {
            $model = $cleanModel
        }
    }

    if ([string]::IsNullOrWhiteSpace($model)) {
        $titleMatch = [regex]::Match($overviewHtml, "<title[^>]*>(?<title>.*?)</title>", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline)
        if ($titleMatch.Success) {
            $titleText = Remove-HtmlTags -Value $titleMatch.Groups["title"].Value
            if (-not [string]::IsNullOrWhiteSpace($titleText) -and $titleText -notmatch "(?i)^access denied$") {
                $model = $titleText
            }
        }
    }

    $encryptedTag = $null
    $tagPatterns = @(
        "encryptedTag\s*=\s*['""](?<tag>[^'""]+)",
        """encryptedTag""\s*:\s*""(?<tag>[^""]+)""",
        """encryptedServiceTag""\s*:\s*""(?<tag>[^""]+)"""
    )

    foreach ($pattern in $tagPatterns) {
        $tagMatch = [regex]::Match($overviewHtml, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($tagMatch.Success) {
            $encryptedTag = $tagMatch.Groups["tag"].Value
            break
        }
    }

    if ([string]::IsNullOrWhiteSpace($encryptedTag)) {
        return [pscustomobject]@{
            Brand          = "Dell"
            Serial         = $InputSerial
            Model          = $model
            SpecAvailable  = "Unknown"
            SpecUrl        = $specUrl
            StartDate      = $null
            ExpirationDate = $null
            Status         = "Unknown"
            ResultUrl      = $overviewUrl
            Notes          = "Unable to locate Dell warranty payload automatically. Open the result URL to confirm details."
        }
    }

    $warrantyUrl = "https://www.dell.com/support/warranty/en-us/warrantydetails/servicetag/$([System.Uri]::EscapeDataString($encryptedTag))/mse/IPS?_=f"
    try {
        $warrantyHtml = Invoke-TextGet -Url $warrantyUrl -Headers $headers
    } catch {
        return [pscustomobject]@{
            Brand          = "Dell"
            Serial         = $InputSerial
            Model          = $model
            SpecAvailable  = "Yes"
            SpecUrl        = $specUrl
            StartDate      = $null
            ExpirationDate = $null
            Status         = "Unknown"
            ResultUrl      = $overviewUrl
            Notes          = "Dell warranty details could not be fetched automatically. Open the result URL to confirm details."
        }
    }

    $expirationDateCandidates = New-Object "System.Collections.Generic.List[datetime]"
    $startDateCandidates = New-Object "System.Collections.Generic.List[datetime]"
    $allDateCandidates = New-Object "System.Collections.Generic.List[datetime]"
    $options = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline

    $serviceRowMatches = [regex]::Matches($warrantyHtml, "<tr[^>]*>\s*<td[^>]*>.*?</td>\s*<td[^>]*>(?<start>.*?)</td>\s*<td[^>]*>(?<end>.*?)</td>", $options)
    foreach ($match in $serviceRowMatches) {
        $rowStartText = Remove-HtmlTags -Value $match.Groups["start"].Value
        $rowStartDate = $null
        if (Try-ParseDate -Value $rowStartText -OutDate ([ref]$rowStartDate)) {
            [void]$startDateCandidates.Add($rowStartDate)
        }

        $rowEndText = Remove-HtmlTags -Value $match.Groups["end"].Value
        $rowEndDate = $null
        if (Try-ParseDate -Value $rowEndText -OutDate ([ref]$rowEndDate)) {
            [void]$expirationDateCandidates.Add($rowEndDate)
        }
    }

    $expirationMatches = [regex]::Matches($warrantyHtml, "id\s*=\s*['""]expiration(?:Dt|Date)['""][^>]*>(?<value>.*?)</", $options)
    foreach ($match in $expirationMatches) {
        $textValue = Remove-HtmlTags -Value $match.Groups["value"].Value
        $parsedDate = $null
        if ($textValue -match "(?i)Expire[sd]\s+(?<date>.+)$" -and (Try-ParseDate -Value $Matches["date"] -OutDate ([ref]$parsedDate))) {
            [void]$expirationDateCandidates.Add($parsedDate)
        } elseif (Try-ParseDate -Value $textValue -OutDate ([ref]$parsedDate)) {
            [void]$expirationDateCandidates.Add($parsedDate)
        }
    }

    $startMatches = [regex]::Matches($warrantyHtml, "id\s*=\s*['""]start(?:Dt|Date)['""][^>]*>(?<value>.*?)</", $options)
    foreach ($match in $startMatches) {
        $textValue = Remove-HtmlTags -Value $match.Groups["value"].Value
        $parsedDate = $null
        if (Try-ParseDate -Value $textValue -OutDate ([ref]$parsedDate)) {
            [void]$startDateCandidates.Add($parsedDate)
        }
    }

    $longDateMatches = [regex]::Matches($warrantyHtml, "(?<date>(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2},\s+\d{4})", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    foreach ($match in $longDateMatches) {
        $parsedDate = $null
        if (Try-ParseDate -Value $match.Groups["date"].Value -OutDate ([ref]$parsedDate)) {
            [void]$allDateCandidates.Add($parsedDate)
        }
    }

    $shortDateMatches = [regex]::Matches($warrantyHtml, "(?<date>\b\d{1,2}/\d{1,2}/\d{2,4}\b)")
    foreach ($match in $shortDateMatches) {
        $parsedDate = $null
        if (Try-ParseDate -Value $match.Groups["date"].Value -OutDate ([ref]$parsedDate)) {
            [void]$allDateCandidates.Add($parsedDate)
        }
    }

    if ($expirationDateCandidates.Count -eq 0 -and $allDateCandidates.Count -gt 0) {
        [void]$expirationDateCandidates.Add(($allDateCandidates | Sort-Object | Select-Object -Last 1))
    }

    if ($startDateCandidates.Count -eq 0 -and $allDateCandidates.Count -gt 0) {
        [void]$startDateCandidates.Add(($allDateCandidates | Sort-Object | Select-Object -First 1))
    }

    $startDate = $null
    if ($startDateCandidates.Count -gt 0) {
        $startDate = ($startDateCandidates | Sort-Object | Select-Object -First 1).ToString("yyyy-MM-dd")
    }

    $expirationDate = $null
    if ($expirationDateCandidates.Count -gt 0) {
        $expirationDate = ($expirationDateCandidates | Sort-Object | Select-Object -Last 1).ToString("yyyy-MM-dd")
    }

    $status = "Unknown"
    $statusMatch = [regex]::Match($warrantyHtml, "warrantyExpiringLabel[^>]*>(?<label>[^<]+)<", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($statusMatch.Success) {
        $rawStatus = Remove-HtmlTags -Value $statusMatch.Groups["label"].Value
        if (-not [string]::IsNullOrWhiteSpace($rawStatus)) {
            $status = $rawStatus
        }
    }

    if ($status -eq "Unknown" -and $expirationDate) {
        $parsedExpiration = [datetime]::ParseExact($expirationDate, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)
        if ($parsedExpiration -ge (Get-Date).Date) {
            $status = "Active"
        } else {
            $status = "Expired"
        }
    }

    $note = $null
    if (-not $startDate -and -not $expirationDate) {
        $note = "Dell warranty page was found, but start and expiration dates could not be parsed automatically."
    } elseif (-not $startDate) {
        $note = "Dell warranty page was found, but start date could not be parsed automatically."
    } elseif (-not $expirationDate) {
        $note = "Dell warranty page was found, but expiration date could not be parsed automatically."
    }

    return [pscustomobject]@{
        Brand          = "Dell"
        Serial         = $InputSerial
        Model          = $model
        SpecAvailable  = "Yes"
        SpecUrl        = $specUrl
        StartDate      = $startDate
        ExpirationDate = $expirationDate
        Status         = $status
        ResultUrl      = $overviewUrl
        Notes          = $note
    }
}

if ([string]::IsNullOrWhiteSpace($Serial)) {
    $Serial = Read-Host "Enter Lenovo serial number or Dell service tag"
}

$Serial = ($Serial -replace "\s+", "").ToUpperInvariant()
if ([string]::IsNullOrWhiteSpace($Serial)) {
    Write-Error "No serial number/service tag was provided."
    exit 1
}

$result = Get-LenovoWarranty -InputSerial $Serial
if (-not $result) {
    $result = Get-DellWarranty -InputSerial $Serial
}

if (-not $result) {
    Write-Error "Unable to find a Lenovo or Dell warranty record for '$Serial'."
    exit 1
}

if ($AsJson) {
    $result | ConvertTo-Json -Depth 4
    exit 0
}

$expirationOutput = if ([string]::IsNullOrWhiteSpace("$($result.ExpirationDate)")) { "Unknown" } else { $result.ExpirationDate }
$startOutput = if ([string]::IsNullOrWhiteSpace("$($result.StartDate)")) { "Unknown" } else { $result.StartDate }
$modelOutput = if ([string]::IsNullOrWhiteSpace("$($result.Model)")) { "Unknown" } else { $result.Model }
$specAvailableOutput = if ([string]::IsNullOrWhiteSpace("$($result.SpecAvailable)")) { "Unknown" } else { $result.SpecAvailable }
$specUrlOutput = if ([string]::IsNullOrWhiteSpace("$($result.SpecUrl)")) { "Unknown" } else { $result.SpecUrl }
$noteOutput = if ([string]::IsNullOrWhiteSpace("$($result.Notes)")) { $null } else { $result.Notes }

Write-Output "Brand: $($result.Brand)"
Write-Output "Serial: $($result.Serial)"
Write-Output "Model: $modelOutput"
Write-Output "Spec Available: $specAvailableOutput"
Write-Output "Spec URL: $specUrlOutput"
Write-Output "Warranty Start: $startOutput"
Write-Output "Warranty Expiration: $expirationOutput"
Write-Output "Warranty Status: $($result.Status)"
Write-Output "Result URL: $($result.ResultUrl)"
if ($noteOutput) {
    Write-Output "Notes: $noteOutput"
}
