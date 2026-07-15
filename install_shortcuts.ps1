$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Launcher = Join-Path $Root "Run Darwill AI Prospector.bat"
$Icon = Join-Path $Root "darwill_ai_prospector.ico"
$Shell = New-Object -ComObject WScript.Shell

$Desktop = [Environment]::GetFolderPath("Desktop")
$StartMenu = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"

foreach ($Folder in @($Desktop, $StartMenu)) {
    $ShortcutPath = Join-Path $Folder "Darwill AI Prospectorlnk"
    $Shortcut = $Shell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $Launcher
    $Shortcut.WorkingDirectory = $Root
    if (Test-Path $Icon) { $Shortcut.IconLocation = $Icon }
    $Shortcut.Description = "Darwill AI Prospector"
    $Shortcut.Save()
}
