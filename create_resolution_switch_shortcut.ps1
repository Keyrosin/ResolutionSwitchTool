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

function Get-InterpreterFromPyLauncher {
    $pyLauncher = Get-Command 'py.exe' -ErrorAction SilentlyContinue
    if (-not $pyLauncher) {
        return $null
    }

    try {
        $rawPath = & $pyLauncher.Source -3 -c "import sys; print(sys.executable)" 2>$null
    }
    catch {
        return $null
    }

    if (-not $rawPath) {
        return $null
    }

    $exePath = $rawPath.Trim()
    if (-not $exePath) {
        return $null
    }

    try {
        $resolved = Resolve-FullPath -Path $exePath
    }
    catch {
        return $null
    }

    $windowedSibling = Join-Path (Split-Path $resolved -Parent) 'pythonw.exe'
    if (Test-Path -LiteralPath $windowedSibling) {
        return Resolve-FullPath -Path $windowedSibling
    }

    return $resolved
}

function Get-PythonExecutable {
    param(
        [Parameter()]
        [string]$Preferred
    )

    if ($Preferred) {
        try {
            return Resolve-FullPath -Path $Preferred
        }
        catch {
            throw "Unable to resolve Python executable at '$Preferred'."
        }
    }

    $fromLauncher = Get-InterpreterFromPyLauncher
    if ($fromLauncher) {
        return $fromLauncher
    }

    $pathCandidates = @('pythonw.exe', 'python.exe', 'pyw.exe', 'py.exe')
    foreach ($candidate in $pathCandidates) {
        $command = Get-Command $candidate -ErrorAction SilentlyContinue
        if (-not $command) {
            continue
        }

        $candidatePath = $command.Source

        if ($candidate -eq 'python.exe') {
            $windowedSibling = Join-Path (Split-Path $candidatePath -Parent) 'pythonw.exe'
            if (Test-Path -LiteralPath $windowedSibling) {
                try {
                    return Resolve-FullPath -Path $windowedSibling
                }
                catch {
                    # fall back to python.exe if pythonw.exe path resolution fails
                }
            }
        }

        try {
            return Resolve-FullPath -Path $candidatePath
        }
        catch {
            continue
        }
    }

    throw 'Python executable not found. Please install Python 3 or pass -PythonExecutable with a full path.'
}

try {
    $scriptFullPath = Resolve-FullPath -Path $ScriptPath
}
catch {
    throw "Unable to find the resolution switch script at '$ScriptPath'."
}

$pythonFullPath = Get-PythonExecutable -Preferred $PythonExecutable

$shortcutDirectory = Split-Path -Parent $ShortcutPath
if (-not $shortcutDirectory) {
    $shortcutDirectory = (Get-Location).Path
}
if (-not (Test-Path -LiteralPath $shortcutDirectory)) {
    New-Item -ItemType Directory -Path $shortcutDirectory -Force | Out-Null
}
$shortcutFullPath = [System.IO.Path]::GetFullPath((Join-Path $shortcutDirectory (Split-Path -Leaf $ShortcutPath)))

try {
    $wshShell = New-Object -ComObject WScript.Shell
}
catch {
    throw 'Unable to access the Windows Script Host. Run this script in Windows PowerShell or PowerShell 7 on Windows.'
}

$shortcut = $wshShell.CreateShortcut($shortcutFullPath)
$shortcut.TargetPath = $pythonFullPath
$shortcut.Arguments = '"' + $scriptFullPath + '"'
$shortcut.WorkingDirectory = Split-Path $scriptFullPath -Parent
$shortcut.WindowStyle = 0
$shortcut.IconLocation = $pythonFullPath + ',0'
$shortcut.Save()

Write-Host "Shortcut created at $shortcutFullPath"
