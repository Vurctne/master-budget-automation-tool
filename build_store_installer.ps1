param(
    [switch]$SkipDependencyInstall,
    [string]$InnoSetupCompiler = ''
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$releaseWorkPath = Join-Path $root 'build_release'

function Get-PythonCommand {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        return 'py -3'
    }
    if (Get-Command python -ErrorAction SilentlyContinue) {
        return 'python'
    }
    throw 'Python 3 was not found on this machine.'
}

function Invoke-Python {
    param([string]$Arguments)

    $command = "{0} {1}" -f $script:PythonCommand, $Arguments
    cmd /c $command
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed: $command"
    }
}

function Resolve-InnoSetupCompiler {
    param([string]$ExplicitPath)

    if ($ExplicitPath) {
        if (Test-Path $ExplicitPath) {
            return (Resolve-Path $ExplicitPath).Path
        }
        throw "Inno Setup compiler was not found at: $ExplicitPath"
    }

    $candidates = @(
        "$env:ProgramFiles(x86)\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
    )

    try {
        $registryLocations = @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
        )

        foreach ($registryPath in $registryLocations) {
            Get-ChildItem $registryPath -ErrorAction SilentlyContinue |
                Get-ItemProperty |
                Where-Object { $_.DisplayName -like 'Inno Setup*' -and $_.InstallLocation } |
                ForEach-Object {
                    $candidates += (Join-Path $_.InstallLocation 'ISCC.exe')
                }
        }
    }
    catch {
        # Fall back to the default paths above if registry lookup is unavailable.
    }

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) {
            return (Resolve-Path $candidate).Path
        }
    }

    throw 'Inno Setup 6 was not found. Install it first or pass -InnoSetupCompiler with the ISCC.exe path.'
}

Push-Location $root
try {
    $script:PythonCommand = Get-PythonCommand

    if (-not $SkipDependencyInstall) {
        Write-Host 'Installing Python build dependencies...'
        Invoke-Python '-m pip install -r requirements.txt pyinstaller'
    }

    Write-Host 'Building signed-ready EXE...'
    Invoke-Python ('-m PyInstaller --noconfirm --clean --workpath "{0}" "Master Budget Automation Tool v1.0.2.spec"' -f $releaseWorkPath)

    $compilerPath = Resolve-InnoSetupCompiler -ExplicitPath $InnoSetupCompiler
    Write-Host "Compiling installer with: $compilerPath"
    & $compilerPath 'installer\master_budget_tool.iss'
    if ($LASTEXITCODE -ne 0) {
        throw 'Inno Setup compilation failed.'
    }

    Write-Host ''
    Write-Host 'Store installer build complete.'
    Write-Host 'Installer output folder: dist'
    Write-Host 'Before Store submission, code-sign the installer EXE and upload it to a versioned HTTPS URL.'
}
finally {
    Pop-Location
}
