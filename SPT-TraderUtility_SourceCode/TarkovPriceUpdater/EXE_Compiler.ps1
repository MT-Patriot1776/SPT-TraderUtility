
$cs  = 'C:\Users\sadams\Downloads\TarkovUpdaterLauncher.cs'
$ps1 = 'C:\Users\sadams\Downloads\TarkovPrices\TarkovPriceUpdater.ps1'  # your actual script
$ico = 'C:\Users\sadams\Downloads\TarkovPriceUpdater.ico'
$out = 'C:\Users\sadams\Downloads\TarkovPriceUpdater.exe'

# Prefer x64 compiler if available; fall back to x86
$csc = if (Test-Path "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe") {
  "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
} else {
  "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe"
}

# WinForms/Drawing references from matching framework folder
if ($csc -like "*Framework64*") {
  $forms   = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\System.Windows.Forms.dll"
  $drawing = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\System.Drawing.dll"
} else {
  $forms   = "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\System.Windows.Forms.dll"
  $drawing = "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\System.Drawing.dll"
}

# Compile as GUI app, set icon, and EMBED the PS1 with logical name EXACTLY "TarkovPriceUpdater.ps1"
& $csc /nologo /target:winexe /win32icon:"$ico" `
       /out:"$out" `
       /resource:"$ps1",TarkovPriceUpdater.ps1 `
       /reference:"$forms" `
       /reference:"$drawing" `
       "$cs"
