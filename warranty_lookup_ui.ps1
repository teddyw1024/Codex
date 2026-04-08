Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:LookupScriptPath = Join-Path $PSScriptRoot "warranty_lookup.ps1"
$script:LookupJob = $null

$form = New-Object System.Windows.Forms.Form
$form.Text = "Laptop Warranty Lookup"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(876, 700)
$form.MinimumSize = New-Object System.Drawing.Size(876, 700)
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 251)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Location = New-Object System.Drawing.Point(0, 0)
$headerPanel.Size = New-Object System.Drawing.Size(876, 78)
$headerPanel.BackColor = [System.Drawing.Color]::FromArgb(30, 70, 130)
$form.Controls.Add($headerPanel)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Warranty Checker"
$titleLabel.ForeColor = [System.Drawing.Color]::White
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 16)
$titleLabel.AutoSize = $true
$titleLabel.Location = New-Object System.Drawing.Point(22, 16)
$headerPanel.Controls.Add($titleLabel)

$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Text = "Enter a Lenovo serial or Dell service tag to check start/end dates and status."
$subtitleLabel.ForeColor = [System.Drawing.Color]::FromArgb(225, 236, 252)
$subtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$subtitleLabel.AutoSize = $true
$subtitleLabel.Location = New-Object System.Drawing.Point(24, 48)
$headerPanel.Controls.Add($subtitleLabel)

$inputPanel = New-Object System.Windows.Forms.Panel
$inputPanel.Location = New-Object System.Drawing.Point(20, 96)
$inputPanel.Size = New-Object System.Drawing.Size(820, 120)
$inputPanel.BackColor = [System.Drawing.Color]::White
$inputPanel.BorderStyle = "FixedSingle"
$form.Controls.Add($inputPanel)

$serialLabel = New-Object System.Windows.Forms.Label
$serialLabel.Text = "Serial Number / Service Tag"
$serialLabel.AutoSize = $true
$serialLabel.Location = New-Object System.Drawing.Point(16, 14)
$serialLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$inputPanel.Controls.Add($serialLabel)

$txtSerial = New-Object System.Windows.Forms.TextBox
$txtSerial.Location = New-Object System.Drawing.Point(20, 40)
$txtSerial.Size = New-Object System.Drawing.Size(540, 30)
$txtSerial.Font = New-Object System.Drawing.Font("Segoe UI", 11)
$txtSerial.BorderStyle = "FixedSingle"
$inputPanel.Controls.Add($txtSerial)

$btnCheck = New-Object System.Windows.Forms.Button
$btnCheck.Text = "Check Warranty"
$btnCheck.Location = New-Object System.Drawing.Point(578, 39)
$btnCheck.Size = New-Object System.Drawing.Size(130, 34)
$btnCheck.BackColor = [System.Drawing.Color]::FromArgb(42, 96, 170)
$btnCheck.ForeColor = [System.Drawing.Color]::White
$btnCheck.FlatStyle = "Flat"
$btnCheck.FlatAppearance.BorderSize = 0
$inputPanel.Controls.Add($btnCheck)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Clear"
$btnClear.Location = New-Object System.Drawing.Point(718, 39)
$btnClear.Size = New-Object System.Drawing.Size(66, 34)
$btnClear.BackColor = [System.Drawing.Color]::FromArgb(233, 237, 243)
$btnClear.ForeColor = [System.Drawing.Color]::FromArgb(45, 58, 79)
$btnClear.FlatStyle = "Flat"
$btnClear.FlatAppearance.BorderSize = 0
$inputPanel.Controls.Add($btnClear)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 84)
$progressBar.Size = New-Object System.Drawing.Size(540, 12)
$progressBar.Style = "Marquee"
$progressBar.MarqueeAnimationSpeed = 28
$progressBar.Visible = $false
$inputPanel.Controls.Add($progressBar)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Ready"
$statusLabel.AutoSize = $true
$statusLabel.Location = New-Object System.Drawing.Point(578, 82)
$statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(80, 94, 113)
$inputPanel.Controls.Add($statusLabel)

$resultPanel = New-Object System.Windows.Forms.Panel
$resultPanel.Location = New-Object System.Drawing.Point(20, 232)
$resultPanel.Size = New-Object System.Drawing.Size(820, 410)
$resultPanel.BackColor = [System.Drawing.Color]::White
$resultPanel.BorderStyle = "FixedSingle"
$form.Controls.Add($resultPanel)

$resultTitle = New-Object System.Windows.Forms.Label
$resultTitle.Text = "Result"
$resultTitle.AutoSize = $true
$resultTitle.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 11)
$resultTitle.Location = New-Object System.Drawing.Point(16, 12)
$resultPanel.Controls.Add($resultTitle)

$captionColor = [System.Drawing.Color]::FromArgb(90, 103, 125)
$valueColor = [System.Drawing.Color]::FromArgb(33, 45, 64)

$lblBrandCaption = New-Object System.Windows.Forms.Label
$lblBrandCaption.Text = "Brand:"
$lblBrandCaption.AutoSize = $true
$lblBrandCaption.ForeColor = $captionColor
$lblBrandCaption.Location = New-Object System.Drawing.Point(20, 50)
$resultPanel.Controls.Add($lblBrandCaption)

$lblBrandValue = New-Object System.Windows.Forms.Label
$lblBrandValue.Text = "-"
$lblBrandValue.AutoSize = $true
$lblBrandValue.ForeColor = $valueColor
$lblBrandValue.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$lblBrandValue.Location = New-Object System.Drawing.Point(150, 50)
$resultPanel.Controls.Add($lblBrandValue)

$lblSerialCaption = New-Object System.Windows.Forms.Label
$lblSerialCaption.Text = "Serial:"
$lblSerialCaption.AutoSize = $true
$lblSerialCaption.ForeColor = $captionColor
$lblSerialCaption.Location = New-Object System.Drawing.Point(20, 80)
$resultPanel.Controls.Add($lblSerialCaption)

$lblSerialValue = New-Object System.Windows.Forms.Label
$lblSerialValue.Text = "-"
$lblSerialValue.AutoSize = $true
$lblSerialValue.ForeColor = $valueColor
$lblSerialValue.Location = New-Object System.Drawing.Point(150, 80)
$resultPanel.Controls.Add($lblSerialValue)

$lblModelCaption = New-Object System.Windows.Forms.Label
$lblModelCaption.Text = "Model:"
$lblModelCaption.AutoSize = $true
$lblModelCaption.ForeColor = $captionColor
$lblModelCaption.Location = New-Object System.Drawing.Point(20, 110)
$resultPanel.Controls.Add($lblModelCaption)

$lblModelValue = New-Object System.Windows.Forms.Label
$lblModelValue.Text = "-"
$lblModelValue.AutoSize = $true
$lblModelValue.ForeColor = $valueColor
$lblModelValue.Location = New-Object System.Drawing.Point(150, 110)
$resultPanel.Controls.Add($lblModelValue)

$lblSpecAvailableCaption = New-Object System.Windows.Forms.Label
$lblSpecAvailableCaption.Text = "Spec Available:"
$lblSpecAvailableCaption.AutoSize = $true
$lblSpecAvailableCaption.ForeColor = $captionColor
$lblSpecAvailableCaption.Location = New-Object System.Drawing.Point(20, 140)
$resultPanel.Controls.Add($lblSpecAvailableCaption)

$lblSpecAvailableValue = New-Object System.Windows.Forms.Label
$lblSpecAvailableValue.Text = "-"
$lblSpecAvailableValue.AutoSize = $true
$lblSpecAvailableValue.ForeColor = $valueColor
$lblSpecAvailableValue.Location = New-Object System.Drawing.Point(150, 140)
$resultPanel.Controls.Add($lblSpecAvailableValue)

$lblSpecUrlCaption = New-Object System.Windows.Forms.Label
$lblSpecUrlCaption.Text = "Spec URL:"
$lblSpecUrlCaption.AutoSize = $true
$lblSpecUrlCaption.ForeColor = $captionColor
$lblSpecUrlCaption.Location = New-Object System.Drawing.Point(20, 170)
$resultPanel.Controls.Add($lblSpecUrlCaption)

$lnkSpecUrl = New-Object System.Windows.Forms.LinkLabel
$lnkSpecUrl.Text = "-"
$lnkSpecUrl.Location = New-Object System.Drawing.Point(150, 170)
$lnkSpecUrl.Size = New-Object System.Drawing.Size(646, 30)
$lnkSpecUrl.LinkColor = [System.Drawing.Color]::FromArgb(28, 95, 178)
$lnkSpecUrl.ActiveLinkColor = [System.Drawing.Color]::FromArgb(17, 65, 125)
$lnkSpecUrl.VisitedLinkColor = [System.Drawing.Color]::FromArgb(28, 95, 178)
$lnkSpecUrl.Enabled = $false
$resultPanel.Controls.Add($lnkSpecUrl)

$lblStartCaption = New-Object System.Windows.Forms.Label
$lblStartCaption.Text = "Warranty Start:"
$lblStartCaption.AutoSize = $true
$lblStartCaption.ForeColor = $captionColor
$lblStartCaption.Location = New-Object System.Drawing.Point(20, 205)
$resultPanel.Controls.Add($lblStartCaption)

$lblStartValue = New-Object System.Windows.Forms.Label
$lblStartValue.Text = "-"
$lblStartValue.AutoSize = $true
$lblStartValue.ForeColor = $valueColor
$lblStartValue.Location = New-Object System.Drawing.Point(150, 205)
$resultPanel.Controls.Add($lblStartValue)

$lblEndCaption = New-Object System.Windows.Forms.Label
$lblEndCaption.Text = "Warranty Expiration:"
$lblEndCaption.AutoSize = $true
$lblEndCaption.ForeColor = $captionColor
$lblEndCaption.Location = New-Object System.Drawing.Point(20, 235)
$resultPanel.Controls.Add($lblEndCaption)

$lblEndValue = New-Object System.Windows.Forms.Label
$lblEndValue.Text = "-"
$lblEndValue.AutoSize = $true
$lblEndValue.ForeColor = $valueColor
$lblEndValue.Location = New-Object System.Drawing.Point(150, 235)
$resultPanel.Controls.Add($lblEndValue)

$lblStatusCaption = New-Object System.Windows.Forms.Label
$lblStatusCaption.Text = "Warranty Status:"
$lblStatusCaption.AutoSize = $true
$lblStatusCaption.ForeColor = $captionColor
$lblStatusCaption.Location = New-Object System.Drawing.Point(20, 265)
$resultPanel.Controls.Add($lblStatusCaption)

$lblStatusValue = New-Object System.Windows.Forms.Label
$lblStatusValue.Text = "-"
$lblStatusValue.AutoSize = $true
$lblStatusValue.ForeColor = $valueColor
$lblStatusValue.Location = New-Object System.Drawing.Point(150, 265)
$resultPanel.Controls.Add($lblStatusValue)

$lblUrlCaption = New-Object System.Windows.Forms.Label
$lblUrlCaption.Text = "Result URL:"
$lblUrlCaption.AutoSize = $true
$lblUrlCaption.ForeColor = $captionColor
$lblUrlCaption.Location = New-Object System.Drawing.Point(20, 295)
$resultPanel.Controls.Add($lblUrlCaption)

$lnkResultUrl = New-Object System.Windows.Forms.LinkLabel
$lnkResultUrl.Text = "-"
$lnkResultUrl.Location = New-Object System.Drawing.Point(150, 295)
$lnkResultUrl.Size = New-Object System.Drawing.Size(646, 30)
$lnkResultUrl.LinkColor = [System.Drawing.Color]::FromArgb(28, 95, 178)
$lnkResultUrl.ActiveLinkColor = [System.Drawing.Color]::FromArgb(17, 65, 125)
$lnkResultUrl.VisitedLinkColor = [System.Drawing.Color]::FromArgb(28, 95, 178)
$lnkResultUrl.Enabled = $false
$resultPanel.Controls.Add($lnkResultUrl)

$lblNotesCaption = New-Object System.Windows.Forms.Label
$lblNotesCaption.Text = "Notes:"
$lblNotesCaption.AutoSize = $true
$lblNotesCaption.ForeColor = $captionColor
$lblNotesCaption.Location = New-Object System.Drawing.Point(20, 330)
$resultPanel.Controls.Add($lblNotesCaption)

$txtNotes = New-Object System.Windows.Forms.TextBox
$txtNotes.Location = New-Object System.Drawing.Point(150, 330)
$txtNotes.Size = New-Object System.Drawing.Size(646, 62)
$txtNotes.Multiline = $true
$txtNotes.ReadOnly = $true
$txtNotes.BorderStyle = "FixedSingle"
$txtNotes.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 253)
$txtNotes.Text = "-"
$resultPanel.Controls.Add($txtNotes)

function Set-ResultValue {
    param(
        [System.Windows.Forms.Control]$Control,
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        $Control.Text = "-"
    } else {
        $Control.Text = $Value
    }
}

function Set-IdleState {
    $progressBar.Visible = $false
    $btnCheck.Enabled = $true
    $txtSerial.Enabled = $true
    $txtSerial.Focus() | Out-Null
}

function Set-BusyState {
    $progressBar.Visible = $true
    $btnCheck.Enabled = $false
    $txtSerial.Enabled = $false
}

function Clear-ResultFields {
    Set-ResultValue -Control $lblBrandValue -Value ""
    Set-ResultValue -Control $lblSerialValue -Value ""
    Set-ResultValue -Control $lblModelValue -Value ""
    Set-ResultValue -Control $lblSpecAvailableValue -Value ""
    Set-ResultValue -Control $lblStartValue -Value ""
    Set-ResultValue -Control $lblEndValue -Value ""
    Set-ResultValue -Control $lblStatusValue -Value ""
    $lnkSpecUrl.Text = "-"
    $lnkSpecUrl.Links.Clear()
    $lnkSpecUrl.Enabled = $false
    $lnkResultUrl.Text = "-"
    $lnkResultUrl.Links.Clear()
    $lnkResultUrl.Enabled = $false
    $txtNotes.Text = "-"
}

function Show-ResultData {
    param($Data)

    Set-ResultValue -Control $lblBrandValue -Value ([string]$Data.Brand)
    Set-ResultValue -Control $lblSerialValue -Value ([string]$Data.Serial)
    Set-ResultValue -Control $lblModelValue -Value ([string]$Data.Model)
    Set-ResultValue -Control $lblSpecAvailableValue -Value ([string]$Data.SpecAvailable)
    Set-ResultValue -Control $lblStartValue -Value ([string]$Data.StartDate)
    Set-ResultValue -Control $lblEndValue -Value ([string]$Data.ExpirationDate)
    Set-ResultValue -Control $lblStatusValue -Value ([string]$Data.Status)

    $specUrl = [string]$Data.SpecUrl
    $lnkSpecUrl.Links.Clear()
    if ([string]::IsNullOrWhiteSpace($specUrl)) {
        $lnkSpecUrl.Text = "-"
        $lnkSpecUrl.Enabled = $false
    } else {
        $lnkSpecUrl.Text = $specUrl
        [void]$lnkSpecUrl.Links.Add(0, $specUrl.Length, $specUrl)
        $lnkSpecUrl.Enabled = $true
    }

    $url = [string]$Data.ResultUrl
    $lnkResultUrl.Links.Clear()
    if ([string]::IsNullOrWhiteSpace($url)) {
        $lnkResultUrl.Text = "-"
        $lnkResultUrl.Enabled = $false
    } else {
        $lnkResultUrl.Text = $url
        [void]$lnkResultUrl.Links.Add(0, $url.Length, $url)
        $lnkResultUrl.Enabled = $true
    }

    Set-ResultValue -Control $txtNotes -Value ([string]$Data.Notes)
}

function Start-Lookup {
    param([string]$SerialInput)

    if (-not (Test-Path -LiteralPath $script:LookupScriptPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Could not find warranty_lookup.ps1 in this folder.`n`nExpected: $($script:LookupScriptPath)",
            "Script Not Found",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    $serial = $SerialInput.Trim()
    if ([string]::IsNullOrWhiteSpace($serial)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a Lenovo serial number or Dell service tag.",
            "Missing Input",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    Clear-ResultFields
    $statusLabel.Text = "Checking warranty..."
    Set-BusyState

    $script:LookupJob = Start-Job -ScriptBlock {
        param($scriptPath, $lookupSerial)

        try {
            $arguments = @(
                "-NoProfile",
                "-ExecutionPolicy", "Bypass",
                "-File", $scriptPath,
                "-Serial", $lookupSerial,
                "-AsJson"
            )

            $raw = & powershell.exe @arguments 2>&1
            if ($LASTEXITCODE -ne 0) {
                return [pscustomobject]@{
                    Ok    = $false
                    Error = (($raw | ForEach-Object { "$_" }) -join [Environment]::NewLine).Trim()
                }
            }

            return [pscustomobject]@{
                Ok   = $true
                Json = (($raw | ForEach-Object { "$_" }) -join [Environment]::NewLine).Trim()
            }
        } catch {
            return [pscustomobject]@{
                Ok    = $false
                Error = $_.Exception.Message
            }
        }
    } -ArgumentList $script:LookupScriptPath, $serial
}

$pollTimer = New-Object System.Windows.Forms.Timer
$pollTimer.Interval = 250
$pollTimer.Add_Tick({
    if ($null -eq $script:LookupJob) {
        return
    }

    if ($script:LookupJob.State -eq "Running") {
        return
    }

    $jobState = $script:LookupJob.State
    $payload = $null
    try {
        $payload = Receive-Job -Job $script:LookupJob -ErrorAction SilentlyContinue
    } finally {
        Remove-Job -Job $script:LookupJob -Force -ErrorAction SilentlyContinue
        $script:LookupJob = $null
    }

    Set-IdleState

    if ($jobState -ne "Completed") {
        $statusLabel.Text = "Lookup failed."
        [System.Windows.Forms.MessageBox]::Show(
            "Warranty lookup did not complete successfully.",
            "Lookup Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    $result = if ($payload -is [System.Array]) { $payload | Select-Object -Last 1 } else { $payload }
    if ($null -eq $result -or -not $result.Ok) {
        $statusLabel.Text = "Lookup failed."
        $errorText = if ($null -eq $result) { "No result was returned." } else { [string]$result.Error }
        if ([string]::IsNullOrWhiteSpace($errorText)) {
            $errorText = "The lookup script reported an unknown error."
        }

        [System.Windows.Forms.MessageBox]::Show(
            $errorText,
            "Lookup Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    try {
        $data = $result.Json | ConvertFrom-Json -ErrorAction Stop
    } catch {
        $statusLabel.Text = "Lookup failed."
        [System.Windows.Forms.MessageBox]::Show(
            "Received an invalid response from warranty_lookup.ps1.`n`n$($_.Exception.Message)",
            "Invalid Response",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    Show-ResultData -Data $data
    $statusLabel.Text = "Lookup completed."
})
$pollTimer.Start()

$btnCheck.Add_Click({
    Start-Lookup -SerialInput $txtSerial.Text
})

$btnClear.Add_Click({
    $txtSerial.Text = ""
    Clear-ResultFields
    $statusLabel.Text = "Ready"
    $txtSerial.Focus() | Out-Null
})

$txtSerial.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.SuppressKeyPress = $true
        Start-Lookup -SerialInput $txtSerial.Text
    }
})

$lnkResultUrl.Add_LinkClicked({
    param($sender, $e)
    if ($null -ne $e.Link -and $e.Link.LinkData) {
        Start-Process -FilePath ([string]$e.Link.LinkData)
    }
})

$lnkSpecUrl.Add_LinkClicked({
    param($sender, $e)
    if ($null -ne $e.Link -and $e.Link.LinkData) {
        Start-Process -FilePath ([string]$e.Link.LinkData)
    }
})

$form.Add_FormClosing({
    if ($null -ne $script:LookupJob) {
        Stop-Job -Job $script:LookupJob -ErrorAction SilentlyContinue
        Remove-Job -Job $script:LookupJob -Force -ErrorAction SilentlyContinue
        $script:LookupJob = $null
    }
})

Clear-ResultFields
$txtSerial.Focus() | Out-Null
[void]$form.ShowDialog()
