
# --- EXE-safe bootstrap ---
# Only attempt relaunch/STA when running as a .ps1 (PSCommandPath present).
$IsScriptHost = [bool]$PSCommandPath

if ($IsScriptHost -and $PSVersionTable.PSVersion.Major -lt 7) {
  $pwsh = $null
  $cmd = Get-Command pwsh -ErrorAction SilentlyContinue
  if ($cmd) { $pwsh = $cmd.Source }
  if (-not $pwsh) {
    $candidates = @(
      "$env:ProgramFiles\PowerShell\7\pwsh.exe",
      "$env:ProgramFiles\PowerShell\7-preview\pwsh.exe",
      "$env:LOCALAPPDATA\Microsoft\WindowsApps\pwsh.exe"
    )
    foreach ($c in $candidates) { if (Test-Path -LiteralPath $c) { $pwsh = $c; break } }
  }
  if ($pwsh) {
    $argList = @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-File',"`"$PSCommandPath`"")
    if ($args -and $args.Count -gt 0) { $argList += $args }
    Start-Process -FilePath $pwsh -ArgumentList $argList -WorkingDirectory (Split-Path -LiteralPath $PSCommandPath)
    exit
  } else {
    Write-Warning "PowerShell 7 (pwsh.exe) not found; continuing in Windows PowerShell $($PSVersionTable.PSVersion)"
  }
}

# In script host, ensure STA. In EXE builds, ps2exe -STA already sets it.
if ($IsScriptHost -and ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA')) {
  $exe = "$env:Windir\System32\WindowsPowerShell\v1.0\powershell.exe"
  $args = "-NoProfile -ExecutionPolicy Bypass -STA -File `"$PSCommandPath`""
  Start-Process -FilePath $exe -ArgumentList $args -WorkingDirectory (Split-Path $PSCommandPath)
  exit
}

# WinForms assemblies: prefer AssemblyName; avoid LiteralPath fallbacks in EXE.
try {
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
} catch {
  throw "Failed to load WinForms assemblies: $($_.Exception.Message)"
}

# --- Helper functions ---
function Get-DefaultOutPath {
  param([string]$InPath)
  $dir = [System.IO.Path]::GetDirectoryName($InPath)
  $base = [System.IO.Path]::GetFileNameWithoutExtension($InPath)
  [System.IO.Path]::Combine($dir, "${base}_updated.json")
}
function Update-Log {
  param([System.Windows.Forms.TextBox]$tb, [string]$line)
  if ($tb) {
    $tb.AppendText($line + [Environment]::NewLine)
    $tb.SelectionStart = $tb.Text.Length
    $tb.ScrollToCaret()
  }
}

# A small JSON minifier (no whitespace/newlines outside strings).
# Works in Windows PowerShell 5.1 and PowerShell 7+. Keeps all characters inside strings intact.
function Minify-Json {
  param([Parameter(Mandatory=$true)][string]$JsonText)
  $sb = New-Object System.Text.StringBuilder
  $inString = $false
  $escape = $false
  foreach ($ch in $JsonText.ToCharArray()) {
    if ($inString) {
      $null = $sb.Append($ch)
      if ($escape) {
        $escape = $false
      } elseif ($ch -eq '\') {
        $escape = $true
      } elseif ($ch -eq '"') {
        $inString = $false
      }
    } else {
      switch ($ch) {
        '"'  { $inString = $true; $null = $sb.Append($ch) }
        ' '  { } # skip whitespace outside strings
        "`t" { } # skip tabs outside strings
        "`r" { } # skip CR
        "`n" { } # skip LF
        default { $null = $sb.Append($ch) }
      }
    }
  }
  $sb.ToString()
}

# --- Build GUI ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Tarkov Trader Price Updater"
$form.Size = New-Object System.Drawing.Size(700, 540)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Input file
$lblIn = New-Object System.Windows.Forms.Label
$lblIn.Text = "Input JSON file:"
$lblIn.Location = New-Object System.Drawing.Point(15, 20)
$lblIn.Size = New-Object System.Drawing.Size(120, 20)

$txtIn = New-Object System.Windows.Forms.TextBox
$txtIn.Location = New-Object System.Drawing.Point(140, 18)
$txtIn.Size = New-Object System.Drawing.Size(430, 22)

$btnBrowseIn = New-Object System.Windows.Forms.Button
$btnBrowseIn.Text = "Browse..."
$btnBrowseIn.Location = New-Object System.Drawing.Point(580, 16)
$btnBrowseIn.Size = New-Object System.Drawing.Size(90, 26)

# Output file
$lblOut = New-Object System.Windows.Forms.Label
$lblOut.Text = "Output file:"
$lblOut.Location = New-Object System.Drawing.Point(15, 60)
$lblOut.Size = New-Object System.Drawing.Size(120, 20)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = New-Object System.Drawing.Point(140, 58)
$txtOut.Size = New-Object System.Drawing.Size(430, 22)

$btnBrowseOut = New-Object System.Windows.Forms.Button
$btnBrowseOut.Text = "Save As..."
$btnBrowseOut.Location = New-Object System.Drawing.Point(580, 56)
$btnBrowseOut.Size = New-Object System.Drawing.Size(90, 26)

# Currency
$lblCur = New-Object System.Windows.Forms.Label
$lblCur.Text = "Output currency:"
$lblCur.Location = New-Object System.Drawing.Point(15, 100)
$lblCur.Size = New-Object System.Drawing.Size(120, 20)

$cmbCur = New-Object System.Windows.Forms.ComboBox
$cmbCur.Location = New-Object System.Drawing.Point(140, 98)
$cmbCur.Size = New-Object System.Drawing.Size(140, 22)
$cmbCur.DropDownStyle = 'DropDownList'
[void]$cmbCur.Items.AddRange(@('RUB','USD','EUR'))
$cmbCur.SelectedItem = 'RUB' # default

# --- Discount/Markup slider (updated UI with spacing) ---
# Left/Right captions
$lblLeft = New-Object System.Windows.Forms.Label
$lblLeft.Text = "Discount"
$lblLeft.Location = New-Object System.Drawing.Point(140, 135) # spaced below currency
$lblLeft.Size = New-Object System.Drawing.Size(100, 16)

$lblRight = New-Object System.Windows.Forms.Label
$lblRight.Text = "Markup"
$lblRight.Location = New-Object System.Drawing.Point(490, 135) # aligned right of the track
$lblRight.Size = New-Object System.Drawing.Size(100, 16)
$lblRight.TextAlign = 'MiddleRight'

# Center readout
$lblDisc = New-Object System.Windows.Forms.Label
$lblDisc.Text = "0%" # start at 0 (center)
$lblDisc.Location = New-Object System.Drawing.Point(15, 170) # spaced below captions
$lblDisc.Size = New-Object System.Drawing.Size(120, 20)

$trkDisc = New-Object System.Windows.Forms.TrackBar
$trkDisc.Location = New-Object System.Drawing.Point(140, 155) # between captions and readout
$trkDisc.Size = New-Object System.Drawing.Size(350, 45)
$trkDisc.Minimum = -100 # left: Discount down to 100
$trkDisc.Maximum = 500  # right: Markup up to 500
$trkDisc.TickFrequency = 10
$trkDisc.Value = 0 # start center
$trkDisc.Add_Scroll({
  $v = [int]$trkDisc.Value
  if ($v -lt 0) {
    $lblDisc.Text = "Discount: {0}%" -f ([math]::Abs($v))
    $lblDisc.ForeColor = [System.Drawing.Color]::SteelBlue
  } elseif ($v -eq 0) {
    $lblDisc.Text = "0%"
    $lblDisc.ForeColor = [System.Drawing.Color]::Black
  } else {
    $lblDisc.Text = "Markup: {0}%" -f $v
    $lblDisc.ForeColor = [System.Drawing.Color]::DarkOrange
  }
})

# Output formatting options
$chkCompact = New-Object System.Windows.Forms.CheckBox
$chkCompact.Text = "Compact JSON (minify)"
$chkCompact.Location = New-Object System.Drawing.Point(300, 205)
$chkCompact.Size = New-Object System.Drawing.Size(180, 24)
$chkCompact.Checked = $false

# --- Unmatched handling: dropdown with 3 options ---
$lblUnmatched = New-Object System.Windows.Forms.Label
$lblUnmatched.Text = "Unmatched items:"
$lblUnmatched.Location = New-Object System.Drawing.Point(140, 205)
$lblUnmatched.Size = New-Object System.Drawing.Size(120, 24)

$cmbUnmatched = New-Object System.Windows.Forms.ComboBox
$cmbUnmatched.Location = New-Object System.Drawing.Point(265, 205)
$cmbUnmatched.Size = New-Object System.Drawing.Size(215, 24)
$cmbUnmatched.DropDownStyle = 'DropDownList'
[void]$cmbUnmatched.Items.Add("Leave untouched")
[void]$cmbUnmatched.Items.Add("Set to zero")
[void]$cmbUnmatched.Items.Add("Apply slider discount/markup")
$cmbUnmatched.SelectedIndex = 1  # default: "Set to zero" (matches previous behavior)

# Run button
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Run"
$btnRun.Location = New-Object System.Drawing.Point(140, 245)
$btnRun.Size = New-Object System.Drawing.Size(100, 30)
$btnRun.Enabled = $false

# Log box
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Log:"
$lblLog.Location = New-Object System.Drawing.Point(15, 285)
$lblLog.Size = New-Object System.Drawing.Size(120, 20)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(15, 310)
$txtLog.Size = New-Object System.Drawing.Size(655, 180)
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true

# File dialogs
$ofd = New-Object System.Windows.Forms.OpenFileDialog
$ofd.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"

$sfd = New-Object System.Windows.Forms.SaveFileDialog
$sfd.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"

# UI events
$btnBrowseIn.Add_Click({
  if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $txtIn.Text = $ofd.FileName
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
$txtIn.Add_TextChanged({ $btnRun.Enabled = ($txtIn.Text -ne "") -and ($txtOut.Text -ne "") })
$txtOut.Add_TextChanged({ $btnRun.Enabled = ($txtIn.Text -ne "") -and ($txtOut.Text -ne "") })

# --- Conversion function (runs only on button click) ---
function Convert-Items {
  param(
    [string]$InPath,
    [string]$OutPath,
    [ValidateSet('RUB','USD','EUR')][string]$OutCurrency = 'RUB',
    [double]$DiscountPercent = 0.0,  # trackbar sends -100..500
    [ValidateSet('Leave','Zero','ApplyRate')][string]$UnmatchedMode = 'Zero',
    [System.Windows.Forms.TextBox]$LogTextBox = $null,
    [bool]$CompactJson = $false
  )
  Update-Log $LogTextBox "Reading: $InPath"
  if (-not (Test-Path $InPath)) { throw "Could not find: $InPath" }
  if (Test-Path $OutPath) { throw "Refusing to overwrite: $OutPath" }

  $raw = Get-Content $InPath -Raw
  try {
    $fullData = $raw | ConvertFrom-Json -ErrorAction Stop
  } catch {
    throw "JSON parse failed: $($_.Exception.Message)"
  }
  # Extract only the barter_scheme section for processing
  $data = $fullData.barter_scheme
  if (-not $data) { throw "No 'barter_scheme' section found." }
  $ids = $data.PSObject.Properties.Name
  if (-not $ids -or $ids.Count -eq 0) { throw "No item IDs in 'barter_scheme'." }

  # Ensure TLS 1.2
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

  # Query Tarkov.dev (GraphQL)
  Update-Log $LogTextBox "Querying Tarkov.dev..."
  $query = @'
query($ids:[ID!]) {
  items(ids:$ids) {
    id
    name
    buyFor { price currency source }
  }
}
'@
  $payload = @{ query = $query; variables = @{ ids = $ids } } | ConvertTo-Json -Depth 6
  $resp = Invoke-RestMethod -Uri "https://api.tarkov.dev/graphql" -Method Post -ContentType "application/json" -Body $payload
  if (-not $resp.data -or -not $resp.data.items -or $resp.data.items.Count -eq 0) { throw "Tarkov.dev returned no items." }

  # FX (USD base) -> derive per selected OutCurrency
  Update-Log $LogTextBox "Fetching FX (USD base)..."
  $fx = Invoke-RestMethod -Uri "https://open.er-api.com/v6/latest/USD" -Method Get
  if ($fx.result -ne "success") { throw "FX API failed: $($fx | ConvertTo-Json -Depth 3)" }
  $rates = $fx.rates

  # Build conversion map function based on OutCurrency
  switch ($OutCurrency) {
    'USD' {
      $usdPerUSD = 1.0
      $usdPerRUB = 1.0 / [double]$rates.RUB
      $usdPerEUR = 1.0 / [double]$rates.EUR
      $perOut = {
        param($cur)
        switch ($cur.ToUpper()) {
          'USD' { $usdPerUSD }
          '$'   { $usdPerUSD }
          'RUB' { $usdPerRUB }
          '₽'   { $usdPerRUB }
          'EUR' { $usdPerEUR }
          '€'   { $usdPerEUR }
          default { $usdPerUSD }
        }
      }
      Update-Log $LogTextBox ("FX: 1 RUB = {0:N6} USD; 1 EUR = {1:N6} USD" -f $usdPerRUB, $usdPerEUR)
    }
    'RUB' {
      $rubPerUSD = [double]$rates.RUB
      $rubPerEUR = [double]$rates.RUB / [double]$rates.EUR
      $rubPerRUB = 1.0
      $perOut = {
        param($cur)
        switch ($cur.ToUpper()) {
          'USD' { $rubPerUSD }
          '$'   { $rubPerUSD }
          'EUR' { $rubPerEUR }
          '€'   { $rubPerEUR }
          'RUB' { $rubPerRUB }
          '₽'   { $rubPerRUB }
          default { $rubPerUSD }
        }
      }
      Update-Log $LogTextBox ("FX: 1 USD = {0:N2} RUB; 1 EUR = {1:N2} RUB" -f $rubPerUSD, $rubPerEUR)
    }
    'EUR' {
      $eurPerUSD = [double]$rates.EUR
      $eurPerRUB = [double]$rates.EUR / [double]$rates.RUB
      $eurPerEUR = 1.0
      $perOut = {
        param($cur)
        switch ($cur.ToUpper()) {
          'USD' { $eurPerUSD }
          '$'   { $eurPerUSD }
          'RUB' { $eurPerRUB }
          '₽'   { $eurPerRUB }
          'EUR' { $eurPerEUR }
          '€'   { $eurPerEUR }
          default { $eurPerUSD }
        }
      }
      Update-Log $LogTextBox ("FX: 1 USD = {0:N6} EUR; 1 RUB = {1:N6} EUR" -f $eurPerUSD, $eurPerRUB)
    }
  }

  # Build id -> lowest price in selected currency
  Update-Log $LogTextBox "Computing lowest trader prices..."
  $priceMap = @{}
  foreach ($item in $resp.data.items) {
    $best = $null
    foreach ($offer in ($item.buyFor | Where-Object { $_ })) {
      $cur = [string]$offer.currency
      if (-not $cur) { $cur = 'USD' } else { $cur = $cur.ToUpper() }
      $converted = [double]$offer.price * (& $perOut $cur)
      if ($best -eq $null -or $converted -lt $best) { $best = $converted }
    }
    if ($best -eq $null) { $best = 0 } # barter-only/unavailable -> 0
    $priceMap[$item.id] = [math]::Round($best, 2)
  }

  # --- Apply discount/markup and unmatched rule (direction-correct) ---
  $d = [double]$DiscountPercent # -100..500 (negative = discount, positive = markup)
  if ($d -lt 0) {
    $discount = [math]::Min(100.0, [math]::Abs($d)) # cap discount at 100%
    $rate = (100.0 - $discount) / 100.0 # e.g., -30 => 0.70
    Update-Log $LogTextBox ("Applying Discount: {0}% (multiplier {1})" -f $discount, $rate)
  } elseif ($d -gt 0) {
    $markup = [math]::Min(500.0, $d) # cap markup at 500%
    $rate = (100.0 + $markup) / 100.0 # e.g., +50 => 1.50
    Update-Log $LogTextBox ("Applying Markup: {0}% (multiplier {1})" -f $markup, $rate)
  } else {
    $rate = 1.0
    Update-Log $LogTextBox "Applying: 0% (multiplier 1.00)"
  }

  foreach ($prop in $data.PSObject.Properties) {
    $id = $prop.Name
    $blocks = $prop.Value
    $val = $priceMap[$id] # may be $null if ID not returned
    $isMatched = ($val -is [double]) -and ($val -gt 0)
    foreach ($outer in $blocks) {
      foreach ($obj in $outer) {
        if ($obj.PSObject.Properties.Name -contains 'count') {
          if ($isMatched) {
            # Matched item: always apply the rate to the fetched price
            $obj.count = [math]::Round($val * $rate, 0)
          } else {
            switch ($UnmatchedMode) {
              'Leave' {
                # Leave untouched: keep whatever is in the JSON
              }
              'Zero' {
                $obj.count = 0
              }
              'ApplyRate' {
                # Apply discount/markup to the existing count value (if present and numeric),
                # otherwise leave untouched.
                if ($obj.PSObject.Properties.Name -contains 'count' -and
                    ($obj.count -as [double]) -ne $null)
                {
                  $obj.count = [math]::Round(([double]$obj.count) * $rate, 0)
                }
              }
            }
          }
        }
      }
    }
  }

  Update-Log $LogTextBox "Writing output: $OutPath"
  # Put the updated barter_scheme back into the full JSON
  $fullData.barter_scheme = $data

  # Write JSON (pretty or compact)
  $jsonOut = $fullData | ConvertTo-Json -Depth 12
  # If you're on PowerShell 7+, you can replace the next two lines with:
  # $jsonOut = $fullData | ConvertTo-Json -Depth 12 -Compress
  if ($CompactJson) {
    $jsonOut = Minify-Json $jsonOut
  }
  $jsonOut | Set-Content -Path $OutPath -Encoding utf8
	$mode = if ($CompactJson) { "(compact)" } else { "(pretty)" }
	Update-Log $LogTextBox ("✅ Done. $mode")

}

# --- Wire Run button ---
$btnRun.Add_Click({
  try {
    $form.UseWaitCursor = $true
    $btnRun.Enabled = $false
    $btnBrowseIn.Enabled = $false
    $btnBrowseOut.Enabled = $false
    Update-Log $txtLog "Starting..."

    # Map dropdown text to parameter value for unmatched handling
    $unmatchedMode =
      switch ($cmbUnmatched.SelectedItem) {
        'Leave untouched'              { 'Leave' }
        'Set to zero'                  { 'Zero' }
        'Apply slider discount/markup' { 'ApplyRate' }
        default                        { 'Zero' }
      }

    Convert-Items -InPath $txtIn.Text `
      -OutPath $txtOut.Text `
      -OutCurrency $cmbCur.SelectedItem `
      -DiscountPercent ([double]$trkDisc.Value) `
      -UnmatchedMode $unmatchedMode `
      -CompactJson ($chkCompact.Checked) `
      -LogTextBox $txtLog

    [System.Windows.Forms.MessageBox]::Show("Done. Created: $($txtOut.Text)")
  } catch {
    [System.Windows.Forms.MessageBox]::Show("Error: " + $_.Exception.Message, "Tarkov Updater")
    Update-Log $txtLog ("ERROR: " + $_.Exception.Message)
  } finally {
    $form.UseWaitCursor = $false
    $btnRun.Enabled = $true
    $btnBrowseIn.Enabled = $true
    $btnBrowseOut.Enabled = $true
  }
})

# --- Add controls & show the form ---
$form.Controls.AddRange(@(
  $lblIn, $txtIn, $btnBrowseIn,
  $lblOut, $txtOut, $btnBrowseOut,
  $lblCur, $cmbCur,
  $lblLeft, $lblRight,
  $lblDisc, $trkDisc,
  $lblUnmatched, $cmbUnmatched, $chkCompact,
  $btnRun,
  $lblLog, $txtLog
))
[System.Windows.Forms.Application]::EnableVisualStyles()
$form.TopMost = $true
$form.ShowDialog() | Out-Null
