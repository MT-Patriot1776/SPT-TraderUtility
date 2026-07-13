
# ========================================================================
# TarkovAssortBuilder_v13.ps1
# Always builds assort.json from an items helper, even if APIs return nothing.
# - GUI: browse input/output, pick currency (RUB/USD/EUR), discount/markup,
#   unmatched handling (Zero / Default / Leave), default value, minify.
# - Loyalty Level dropdown (1..4) applies same level to all items.
# - Ordered output: items (top), barter_scheme (middle), loyal_level_items (bottom).
# - Items: _comment, _id, _tpl, slotId, parentId, upd { UnlimitedCount, StackObjectsCount }.
# - barter_scheme entry: _comment first, then _tpl, then count.
# - Parser: accepts one ID per line (plain text) or JSON array of IDs; ignores helper levels.
# - ASCII only.
# ========================================================================

# --- Bootstrap: use native Windows PowerShell STA on stock Windows 11 ---
$IsScriptHost = [bool]$PSCommandPath
if ($IsScriptHost -and ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA')) {
    $exe = "$env:Windir\System32\WindowsPowerShell\v1.0\powershell.exe"
    $args = "-NoProfile -ExecutionPolicy Bypass -STA -File `"$PSCommandPath`""
    Start-Process -FilePath $exe -ArgumentList $args -WorkingDirectory (Split-Path $PSCommandPath)
    exit
}

# --- WinForms assemblies ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Helpers: logging, minify, default out path ---
function Update-Log { param([System.Windows.Forms.TextBox]$tb, [string]$line)
    if ($tb) { $tb.AppendText($line + [Environment]::NewLine); $tb.SelectionStart = $tb.Text.Length; $tb.ScrollToCaret() }
}
function Minify-Json { param([Parameter(Mandatory=$true)][string]$JsonText)
    $sb = New-Object System.Text.StringBuilder
    $inString = $false; $escape = $false
    foreach ($ch in $JsonText.ToCharArray()) {
        if ($inString) {
            $null = $sb.Append($ch)
            if ($escape) { $escape = $false }
            elseif ($ch -eq '\') { $escape = $true }
            elseif ($ch -eq '"') { $inString = $false }
        } else {
            switch ($ch) {
                '"' { $inString = $true; $null = $sb.Append($ch) }
                ' ' {}
                "`t" {}
                "`r" {}
                "`n" {}
                default { $null = $sb.Append($ch) }
            }
        }
    }
    $sb.ToString()
}
function Format-JsonIndent { param([Parameter(Mandatory=$true)][string]$JsonText)
    $json = Minify-Json $JsonText
    $sb = New-Object System.Text.StringBuilder
    $indent = 0
    $inString = $false
    $escape = $false

    for ($i = 0; $i -lt $json.Length; $i++) {
        $ch = $json[$i]

        if ($inString) {
            $null = $sb.Append($ch)
            if ($escape) { $escape = $false }
            elseif ($ch -eq '\') { $escape = $true }
            elseif ($ch -eq '"') { $inString = $false }
            continue
        }

        switch ($ch) {
            '"' {
                $inString = $true
                $null = $sb.Append($ch)
            }
            '{' {
                $null = $sb.Append($ch)
                $indent++
                $null = $sb.Append([Environment]::NewLine)
                $null = $sb.Append(' ' * ($indent * 2))
            }
            '[' {
                $null = $sb.Append($ch)
                $indent++
                $null = $sb.Append([Environment]::NewLine)
                $null = $sb.Append(' ' * ($indent * 2))
            }
            '}' {
                $indent--
                $null = $sb.Append([Environment]::NewLine)
                $null = $sb.Append(' ' * ($indent * 2))
                $null = $sb.Append($ch)
            }
            ']' {
                $indent--
                $null = $sb.Append([Environment]::NewLine)
                $null = $sb.Append(' ' * ($indent * 2))
                $null = $sb.Append($ch)
            }
            ',' {
                $null = $sb.Append($ch)
                $null = $sb.Append([Environment]::NewLine)
                $null = $sb.Append(' ' * ($indent * 2))
            }
            ':' {
                $null = $sb.Append(': ')
            }
            default {
                $null = $sb.Append($ch)
            }
        }
    }

    return $sb.ToString()
}
function Get-DefaultOutPath { param([string]$InPath)
    $dir = [System.IO.Path]::GetDirectoryName($InPath)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($InPath)
    [System.IO.Path]::Combine($dir, "${base}_assort.json")
}

# --- Currency normalization ---
function Normalize-Currency {
    param([string]$cur)
    if (-not $cur) { return 'USD' }
    $c = $cur.Trim()
    if ($c.Length -eq 1) {
        $code = [int][char]$c
        switch ($code) {
            36   { return 'USD' }  # '$'
            8364 { return 'EUR' }  # '€'
            8381 { return 'RUB' }  # '₽'
            default { return $c.ToUpper() }
        }
    }
    return $c.ToUpper()
}

# --- Currency tpl IDs ---
$CurrencyTplMap = @{
    'RUB' = '5449016a4bdc2d6f028b456f'
    'USD' = '5696686a4bdc2da3298b456a'
    'EUR' = '569668774bdc2da2298b4568'
}

# --- GUI ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Tarkov Assort Builder"
$form.Size = New-Object System.Drawing.Size(780, 690)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$lblIn = New-Object System.Windows.Forms.Label
$lblIn.Text = "Items helper file:"
$lblIn.Location = New-Object System.Drawing.Point(15, 20)
$lblIn.Size = New-Object System.Drawing.Size(120, 20)
$txtIn = New-Object System.Windows.Forms.TextBox
$txtIn.Location = New-Object System.Drawing.Point(140, 18)
$txtIn.Size = New-Object System.Drawing.Size(500, 22)
$btnBrowseIn = New-Object System.Windows.Forms.Button
$btnBrowseIn.Text = "Browse..."
$btnBrowseIn.Location = New-Object System.Drawing.Point(650, 16)
$btnBrowseIn.Size = New-Object System.Drawing.Size(100, 26)

$lblOut = New-Object System.Windows.Forms.Label
$lblOut.Text = "Output assort.json:"
$lblOut.Location = New-Object System.Drawing.Point(15, 60)
$lblOut.Size = New-Object System.Drawing.Size(120, 20)
$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = New-Object System.Drawing.Point(140, 58)
$txtOut.Size = New-Object System.Drawing.Size(500, 22)
$btnBrowseOut = New-Object System.Windows.Forms.Button
$btnBrowseOut.Text = "Save As..."
$btnBrowseOut.Location = New-Object System.Drawing.Point(650, 56)
$btnBrowseOut.Size = New-Object System.Drawing.Size(100, 26)

$lblCur = New-Object System.Windows.Forms.Label
$lblCur.Text = "Currency:"
$lblCur.Location = New-Object System.Drawing.Point(15, 100)
$lblCur.Size = New-Object System.Drawing.Size(120, 20)
$cmbCur = New-Object System.Windows.Forms.ComboBox
$cmbCur.Location = New-Object System.Drawing.Point(140, 98)
$cmbCur.Size = New-Object System.Drawing.Size(140, 22)
$cmbCur.DropDownStyle = 'DropDownList'
[void]$cmbCur.Items.AddRange(@('RUB','USD','EUR'))
$cmbCur.SelectedItem = 'RUB'

$lblLeft  = New-Object System.Windows.Forms.Label; $lblLeft.Text  = "Discount"; $lblLeft.Location  = New-Object System.Drawing.Point(140, 135); $lblLeft.Size  = New-Object System.Drawing.Size(100, 16)
$lblRight = New-Object System.Windows.Forms.Label; $lblRight.Text = "Markup";  $lblRight.Location = New-Object System.Drawing.Point(540, 135); $lblRight.Size = New-Object System.Drawing.Size(100, 16); $lblRight.TextAlign = 'MiddleRight'
$lblDisc  = New-Object System.Windows.Forms.Label; $lblDisc.Text  = "0%";       $lblDisc.Location  = New-Object System.Drawing.Point(15, 170); $lblDisc.Size  = New-Object System.Drawing.Size(120, 20)
$trkDisc = New-Object System.Windows.Forms.TrackBar
$trkDisc.Location = New-Object System.Drawing.Point(140, 155)
$trkDisc.Size = New-Object System.Drawing.Size(500, 45)
$trkDisc.Minimum = -100; $trkDisc.Maximum = 500; $trkDisc.TickFrequency = 10; $trkDisc.Value = 0
$trkDisc.Add_Scroll({
    $v = [int]$trkDisc.Value
    if ($v -lt 0)      { $lblDisc.Text = "Discount: {0}%" -f ([math]::Abs($v)); $lblDisc.ForeColor = [System.Drawing.Color]::SteelBlue }
    elseif ($v -eq 0)  { $lblDisc.Text = "0%"; $lblDisc.ForeColor = [System.Drawing.Color]::Black }
    else               { $lblDisc.Text = "Markup: {0}%"  -f $v; $lblDisc.ForeColor = [System.Drawing.Color]::DarkOrange }
})

$lblUnmatched = New-Object System.Windows.Forms.Label
$lblUnmatched.Text = "Unmatched items:"
$lblUnmatched.Location = New-Object System.Drawing.Point(140, 205)
$lblUnmatched.Size = New-Object System.Drawing.Size(120, 24)
$cmbUnmatched = New-Object System.Windows.Forms.ComboBox
$cmbUnmatched.Location = New-Object System.Drawing.Point(265, 205)
$cmbUnmatched.Size = New-Object System.Drawing.Size(215, 24)
$cmbUnmatched.DropDownStyle = 'DropDownList'
[void]$cmbUnmatched.Items.Add("Set to zero")
[void]$cmbUnmatched.Items.Add("Input a default value")
[void]$cmbUnmatched.Items.Add("Leave untouched")
$cmbUnmatched.SelectedIndex = 0

$lblDefault = New-Object System.Windows.Forms.Label
$lblDefault.Text = "Default value:"
$lblDefault.Location = New-Object System.Drawing.Point(490, 205)
$lblDefault.Size = New-Object System.Drawing.Size(95, 24)
$txtDefault = New-Object System.Windows.Forms.TextBox
$txtDefault.Location = New-Object System.Drawing.Point(585, 205)
$txtDefault.Size = New-Object System.Drawing.Size(80, 24)
$txtDefault.Text = "1"
$lblDefault.Enabled = $false; $txtDefault.Enabled = $false
$cmbUnmatched.Add_SelectedIndexChanged({
    $useDefault = ($cmbUnmatched.SelectedItem -eq "Input a default value")
    $lblDefault.Enabled = $useDefault; $txtDefault.Enabled = $useDefault
})

# Loyalty Level dropdown (applies same level to all items)
$lblLL = New-Object System.Windows.Forms.Label
$lblLL.Text = "Loyalty Level for items"
$lblLL.Location = New-Object System.Drawing.Point(140, 270)
$lblLL.Size = New-Object System.Drawing.Size(220, 24)

$cmbLL = New-Object System.Windows.Forms.ComboBox
$cmbLL.Location = New-Object System.Drawing.Point(365, 270)
$cmbLL.Size = New-Object System.Drawing.Size(115, 24)
$cmbLL.DropDownStyle = 'DropDownList'
[void]$cmbLL.Items.AddRange(@('1','2','3','4'))
$cmbLL.SelectedIndex = 0  # default to "1"

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Build Assort"
$btnRun.Location = New-Object System.Drawing.Point(340, 305)
$btnRun.Size = New-Object System.Drawing.Size(130, 30)
$btnRun.Enabled = $false

$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Log:"
$lblLog.Location = New-Object System.Drawing.Point(15, 345)
$lblLog.Size = New-Object System.Drawing.Size(120, 20)
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(15, 370)
$txtLog.Size = New-Object System.Drawing.Size(735, 280)
$txtLog.Multiline = $true; $txtLog.ScrollBars = 'Vertical'; $txtLog.ReadOnly = $true

$ofd = New-Object System.Windows.Forms.OpenFileDialog
$ofd.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
$sfd = New-Object System.Windows.Forms.SaveFileDialog
$sfd.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"

$btnBrowseIn.Add_Click({
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtIn.Text  = $ofd.FileName
        $txtOut.Text = Get-DefaultOutPath $ofd.FileName
        $btnRun.Enabled = ($txtIn.Text -ne "") -and ($txtOut.Text -ne "")
    }
})
$btnBrowseOut.Add_Click({
    if ($txtIn.Text -and (Test-Path $txtIn.Text)) {
        $sfd.InitialDirectory = [System.IO.Path]::GetDirectoryName($txtIn.Text)
        $sfd.FileName = [System.IO.Path]::GetFileName($txtOut.Text)
    }
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOut.Text = $sfd.FileName
        $btnRun.Enabled = ($txtIn.Text -ne "") -and ($txtOut.Text -ne "")
    }
})
$txtIn.Add_TextChanged({  $btnRun.Enabled = ($txtIn.Text -ne "") -and ($txtOut.Text -ne "") })
$txtOut.Add_TextChanged({ $btnRun.Enabled = ($txtIn.Text -ne "") -and ($txtOut.Text -ne "") })

# --- Parser for items helper ---
function Parse-ItemsHelper {
    param([string]$Path, [System.Windows.Forms.TextBox]$Log)
    Update-Log $Log "Reading items helper: $Path"
    if (-not (Test-Path $Path)) { throw "Could not find: $Path" }
    $raw = Get-Content $Path -Raw

    # Try strict JSON first
    try {
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        $map = @{}

        # Case 1: dictionary of id->level OR items->dictionary
        if ($obj -is [System.Collections.IDictionary]) {
            if ($obj.items -and ($obj.items -is [System.Collections.IDictionary])) {
                $obj = $obj.items
            }
            if ($obj -is [System.Collections.IDictionary]) {
                foreach ($k in $obj.Keys) {
                    if ($k) { $map[[string]$k] = 1 }  # level ignored; GUI overrides
                }
            }
            elseif ($obj.items -and ($obj.items -is [System.Collections.IEnumerable])) {
                foreach ($id in $obj.items) {
                    if ($id) { $map[[string]$id] = 1 }
                }
            }
        }
        # Case 2: array of IDs or objects
        elseif ($obj -is [System.Collections.IEnumerable]) {
            foreach ($e in $obj) {
                if ($e -is [string]) {
                    if ($e) { $map[[string]$e] = 1 }
                } elseif ($e.id) {
                    $map[[string]$e.id] = 1
                } elseif ($e.itemId) {
                    $map[[string]$e.itemId] = 1
                }
            }
        }

        if ($map.Count -gt 0) { return $map }
    } catch { }

    # Fallback: plain text; support one ID per line and messy quoted forms
    $dict = @{}
    $lineIds = 0

    # First, try one ID per line (24 hex)
    $lines = $raw -split "`r?`n"
    foreach ($ln in $lines) {
        $m = [regex]::Match($ln, '(?i)\b([0-9a-f]{24})\b')
        if ($m.Success) {
            $dict[$m.Groups[1].Value] = 1
            $lineIds++
        }
    }
    if ($lineIds -gt 0) { return $dict }

    # Next, try paired forms: "ID": 1 OR ID: 1 OR "ID" 1 (levels ignored)
    $pairRegex = [regex]'(?i)"?([0-9a-f]{24})"?\s*:?\s*(\d+)'
    $matches = $pairRegex.Matches($raw)
    foreach ($m in $matches) {
        $id = $m.Groups[1].Value
        $dict[$id] = 1
    }

    if ($dict.Count -eq 0) { throw "Items helper malformed. Provide a JSON array of IDs or one ID per line." }
    return $dict
}

# --- Core: Build assort ---
function Build-Assort {
    param(
        [string]$InPath,
        [string]$OutPath,
        [ValidateSet('RUB','USD','EUR')][string]$OutCurrency = 'RUB',
        [double]$DiscountPercent = 0.0, # -100..500
        [ValidateSet('Zero','Default','Leave')][string]$UnmatchedMode = 'Zero',
        [double]$DefaultValue = 1.0,
        [object]$CompactJson = $false,
        [ValidateRange(1,4)][int]$LoyaltyLevel = 1,   # selected loyalty level
        [System.Windows.Forms.TextBox]$LogTextBox = $null
    )

    # Backup existing output
    if (Test-Path $OutPath) {
        $bak = "$OutPath.bak"
        Copy-Item -LiteralPath $OutPath -Destination $bak -Force
        Update-Log $LogTextBox "Backed up existing file to: $bak"
    }

    # Parse helper
    $idToLevel = Parse-ItemsHelper -Path $InPath -Log $LogTextBox
    $ids = @($idToLevel.Keys)
    if ($ids.Count -eq 0) { throw "No item IDs parsed from items helper." }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # Tarkov.dev GraphQL (best effort)
    Update-Log $LogTextBox "Querying item data from Tarkov.dev..."
    $query = @'
query($ids:[ID!]) {
  items(ids:$ids) {
    id
    name
    shortName
    buyFor { price currency source }
    sellFor { price currency source }
    avg24hPrice
    basePrice
  }
}
'@
    $payload = @{ query = $query; variables = @{ ids = $ids } } | ConvertTo-Json -Depth 6

    $items = @()
    try {
        $resp = Invoke-RestMethod -Uri "https://api.tarkov.dev/graphql" -Method Post -ContentType "application/json" -Body $payload
        if (-not $resp.data -or -not $resp.data.items) {
            Update-Log $LogTextBox "WARNING: Tarkov.dev returned no items; continuing."
        } else {
            $items = $resp.data.items
            if ($items.Count -eq 0) { Update-Log $LogTextBox "WARNING: Tarkov.dev returned zero items; continuing." }
        }
    } catch {
        Update-Log $LogTextBox ("WARNING: Tarkov.dev request failed: " + $_.Exception.Message + " - continuing.")
    }

    # FX (best effort)
    Update-Log $LogTextBox "Fetching FX (USD base) from ExchangeRate-API..."
    $rates = $null
    try {
        $fx = Invoke-RestMethod -Uri "https://open.er-api.com/v6/latest/USD" -Method Get
        if ($fx.result -eq "success") { $rates = $fx.rates }
        else { Update-Log $LogTextBox ("WARNING: FX API failed: " + ($fx | ConvertTo-Json -Depth 3)) }
    } catch {
        Update-Log $LogTextBox ("WARNING: FX API exception: " + $_.Exception.Message)
    }

    # Currency conversion to OutCurrency
    switch ($OutCurrency) {
        'USD' {
            $usdPerUSD = 1.0
            $usdPerRUB = if ($rates) { 1.0 / [double]$rates.RUB } else { 1.0 }
            $usdPerEUR = if ($rates) { 1.0 / [double]$rates.EUR } else { 1.0 }
            $toOut = { param($rawCur)
                switch (Normalize-Currency $rawCur) {
                    'USD' { $usdPerUSD }
                    'RUB' { $usdPerRUB }
                    'EUR' { $usdPerEUR }
                    default { $usdPerUSD }
                }
            }
        }
        'RUB' {
            $rubPerUSD = if ($rates) { [double]$rates.RUB } else { 1.0 }
            $rubPerEUR = if ($rates) { [double]$rates.RUB / [double]$rates.EUR } else { 1.0 }
            $rubPerRUB = 1.0
            $toOut = { param($rawCur)
                switch (Normalize-Currency $rawCur) {
                    'USD' { $rubPerUSD }
                    'EUR' { $rubPerEUR }
                    'RUB' { $rubPerRUB }
                    default { $rubPerUSD }
                }
            }
        }
        'EUR' {
            $eurPerUSD = if ($rates) { [double]$rates.EUR } else { 1.0 }
            $eurPerRUB = if ($rates) { [double]$rates.EUR / [double]$rates.RUB } else { 1.0 }
            $eurPerEUR = 1.0
            $toOut = { param($rawCur)
                switch (Normalize-Currency $rawCur) {
                    'USD' { $eurPerUSD }
                    'RUB' { $eurPerRUB }
                    'EUR' { $eurPerEUR }
                    default { $eurPerUSD }
                }
            }
        }
    }

    # Price and item name maps
    Update-Log $LogTextBox "Computing trader prices..."
    $priceMap = @{}
    $descMap  = @{}
    foreach ($hid in $ids) { $descMap[$hid] = "UPDATE ITEM TEXT HERE" }

    foreach ($it in $items) {
        $desc = $it.name; if (-not $desc) { $desc = "UPDATE ITEM TEXT HERE" }
        $descMap[$it.id] = $desc

        $offers = @()
        if ($it.buyFor) { $offers += $it.buyFor }
        if (-not $offers -or $offers.Count -eq 0) {
            if ($it.basePrice -and $it.basePrice -gt 0)         { $offers += @{ price = [double]$it.basePrice; currency = 'RUB'; source='basePrice' } }
            elseif ($it.avg24hPrice -and $it.avg24hPrice -gt 0) { $offers += @{ price = [double]$it.avg24hPrice; currency='RUB'; source='avg24h' } }
            elseif ($it.sellFor)                                { $offers += $it.sellFor }
        }

        $bestOut = $null
        foreach ($of in $offers) {
            $cur = [string]$of.currency; if (-not $cur) { $cur = 'USD' }
            $converted = [double]$of.price * (& $toOut $cur)
            if ($bestOut -eq $null -or $converted -lt $bestOut) { $bestOut = $converted }
        }
        if ($bestOut -eq $null) { $bestOut = 0 }

        $d = [double]$DiscountPercent
        $rate = 1.0
        if     ($d -lt 0) { $rate = (100.0 - [math]::Min(100.0, [math]::Abs($d))) / 100.0 }
        elseif ($d -gt 0) { $rate = (100.0 + [math]::Min(500.0, $d)) / 100.0 }

        $priceMap[$it.id] = [int][math]::Round($bestOut * $rate, 0)
    }

    # Ordered top-level object: items -> barter_scheme -> loyal_level_items
    $assort = [ordered]@{
        items             = @()
        barter_scheme     = [ordered]@{}
        loyal_level_items = [ordered]@{}
    }

    # items: in desired field order; StackObjectsCount inside upd
    foreach ($id in $ids) {
        $assort.items += [ordered]@{
            _comment = $descMap[$id]
            _id      = $id
            _tpl     = $id
            slotId   = "hideout"
            parentId = "hideout"
            upd      = @{
                UnlimitedCount    = $true
                StackObjectsCount = 999999999
            }
        }
    }

# barter_scheme: _comment first, then _tpl, then count
$moneyTpl = $CurrencyTplMap[$OutCurrency]
foreach ($id in $ids) {
    $val = $priceMap[$id]
    $hasPrice = ($val -is [int]) -and ($val -gt 0)

    if ($hasPrice) {
        # FORCE nested array: [[{...}]]
        $assort.barter_scheme[$id] = ,( ,(
            [ordered]@{
                _comment = $descMap[$id]
                _tpl     = $moneyTpl
                count    = $val
            }
        ) )
    }
    else {
        switch ($UnmatchedMode) {
            'Zero' {
                # [[{...}]] with count = 0
                $assort.barter_scheme[$id] = ,( ,(
                    [ordered]@{
                        _comment = $descMap[$id]
                        _tpl     = $moneyTpl
                        count    = 0
                    }
                ) )
            }
            'Default' {
                # [[{...}]] with count = DefaultValue
                $assort.barter_scheme[$id] = ,( ,(
                    [ordered]@{
                        _comment = $descMap[$id]
                        _tpl     = $moneyTpl
                        count    = [int][math]::Round($DefaultValue, 0)
                    }
                ) )
            }
            'Leave' { }
        }
    }
}

    # loyal_level_items: apply selected loyalty level to all items
    foreach ($id in $ids) {
        $assort.loyal_level_items[$id] = $LoyaltyLevel
    }

    # Write output
    Update-Log $LogTextBox "Writing: $OutPath"
    $jsonOut = $assort | ConvertTo-Json -Depth 12
    $compactEnabled = $false
    if ($CompactJson -is [bool]) {
        $compactEnabled = $CompactJson
    } elseif ($CompactJson -ne $null -and [string]$CompactJson -ne '') {
        $compactEnabled = [System.Convert]::ToBoolean($CompactJson)
    }
    if ($compactEnabled) { $jsonOut = Minify-Json $jsonOut }
    else { $jsonOut = Format-JsonIndent $jsonOut }
    $jsonOut | Set-Content -Path $OutPath -Encoding utf8
    Update-Log $LogTextBox "Done."
}

# --- Wire Run button ---
$btnRun.Add_Click({
    try {
        $form.UseWaitCursor = $true
        $btnRun.Enabled = $false; $btnBrowseIn.Enabled = $false; $btnBrowseOut.Enabled = $false
        Update-Log $txtLog "Starting..."
        $unmatchedMode = switch ($cmbUnmatched.SelectedItem) {
            'Set to zero'           { 'Zero' }
            'Input a default value' { 'Default' }
            'Leave untouched'       { 'Leave' }
        }
        $defVal = 1.0
        if ($unmatchedMode -eq 'Default') {
            if (-not [double]::TryParse($txtDefault.Text, [ref]$defVal)) { throw "Default value must be numeric." }
            if ($defVal -lt 0) { throw "Default value cannot be negative." }
        }

        # Read selected loyalty level (1..4)
        $selectedLL = [int]$cmbLL.SelectedItem

        Build-Assort -InPath $txtIn.Text `
                     -OutPath $txtOut.Text `
                     -OutCurrency $cmbCur.SelectedItem `
                     -DiscountPercent ([double]$trkDisc.Value) `
                     -UnmatchedMode $unmatchedMode `
                     -DefaultValue $defVal `
                     -CompactJson $false `
                     -LoyaltyLevel $selectedLL `
                     -LogTextBox $txtLog
        [System.Windows.Forms.MessageBox]::Show("Done. Created: $($txtOut.Text)")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error: " + $_.Exception.Message, "Assort Builder")
        Update-Log $txtLog ("ERROR: " + $_.Exception.Message)
    } finally {
        $form.UseWaitCursor = $false
        $btnRun.Enabled = $true; $btnBrowseIn.Enabled = $true; $btnBrowseOut.Enabled = $true
    }
})

# --- Add controls & show ---
$form.Controls.AddRange(@(
    $lblIn, $txtIn, $btnBrowseIn,
    $lblOut, $txtOut, $btnBrowseOut,
    $lblCur, $cmbCur,
    $lblLeft, $lblRight, $lblDisc, $trkDisc,
    $lblUnmatched, $cmbUnmatched, $lblDefault, $txtDefault,
    $lblLL, $cmbLL,
    $btnRun,
    $lblLog, $txtLog
))
[System.Windows.Forms.Application]::EnableVisualStyles()
$form.TopMost = $true
$form.ShowDialog() | Out-Null
