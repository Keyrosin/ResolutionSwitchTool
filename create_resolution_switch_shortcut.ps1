[CmdletBinding()]
param(
    [Parameter()]
    [string]$ScriptPath = $(Join-Path $PSScriptRoot 'resolution_switch.py'),

    [Parameter()]
    [string]$ShortcutPath = $(Join-Path ([Environment]::GetFolderPath('Desktop')) 'Resolution Switch.lnk'),

    [Parameter()]
    [string]$PythonExecutable
)

function Resolve-FullPath {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    return (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
}

try {
    $scriptFullPath = Resolve-FullPath -Path $ScriptPath
}
catch {
    throw "Unable to find the resolution switch script at '$ScriptPath'."
}

if (-not $PythonExecutable) {
    $candidates = @('pythonw.exe', 'pyw.exe', 'python.exe', 'py.exe')
    foreach ($candidate in $candidates) {
        try {
            $command = Get-Command $candidate -ErrorAction Stop
            if ($command) {
                $PythonExecutable = $command.Source
                break
            }
        }
        catch {
            continue
        }
    }
}

if (-not $PythonExecutable) {
    throw 'Python executable not found on PATH. Please specify -PythonExecutable.'
}

try {
    $pythonFullPath = Resolve-FullPath -Path $PythonExecutable
}
catch {
    throw "Unable to resolve Python executable at '$PythonExecutable'."
}

$shortcutDirectory = Split-Path -Parent $ShortcutPath
if (-not $shortcutDirectory) {
    $shortcutDirectory = (Get-Location).Path
}
if (-not (Test-Path -LiteralPath $shortcutDirectory)) {
    New-Item -ItemType Directory -Path $shortcutDirectory -Force | Out-Null
}
$shortcutFullPath = [System.IO.Path]::GetFullPath((Join-Path $shortcutDirectory (Split-Path -Leaf $ShortcutPath)))

$wshShell = New-Object -ComObject WScript.Shell
$shortcut = $wshShell.CreateShortcut($shortcutFullPath)
$shortcut.TargetPath = $pythonFullPath
$shortcut.Arguments = '"' + $scriptFullPath + '"'
$shortcut.WorkingDirectory = Split-Path $scriptFullPath -Parent
$shortcut.WindowStyle = 7
$shortcut.IconLocation = $pythonFullPath + ',0'
$shortcut.Save()

Write-Host "Shortcut created at $shortcutFullPath"
