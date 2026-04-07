param(
    [switch]$SkipDependencyInstall,
    [string]$InnoSetupCompiler = '',
    [string]$SignToolPath = '',
    [string]$CertificateThumbprint = '',
    [string]$PfxPath = '',
    [string]$PfxPassword = '',
    [string]$TimestampUrl = 'http://timestamp.digicert.com'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$releaseWorkPath = Join-Path $root 'build_release'
$appExePath = Join-Path $root 'dist\Master Budget Automation Tool v1.0.2.exe'
$installerExePath = Join-Path $root 'dist\MasterBudgetAutomationTool_Setup_1.0.2.exe'

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

function Resolve-SignToolPath {
    param([string]$ExplicitPath)

    if ($ExplicitPath) {
        if (Test-Path $ExplicitPath) {
            return (Resolve-Path $ExplicitPath).Path
        }
        throw "signtool.exe was not found at: $ExplicitPath"
    }

    $command = Get-Command signtool.exe -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $candidates = @()
    $sdkRoots = @(
        "$env:ProgramFiles(x86)\Windows Kits\10\bin",
        "$env:ProgramFiles\Windows Kits\10\bin"
    )

    foreach ($sdkRoot in $sdkRoots) {
        if (Test-Path $sdkRoot) {
            $candidates += Get-ChildItem $sdkRoot -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue |
                Sort-Object FullName -Descending |
                Select-Object -ExpandProperty FullName
        }
    }

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) {
            return $candidate
        }
    }

    throw 'signtool.exe was not found. Install the Windows SDK or pass -SignToolPath with the full path.'
}

function Test-SigningRequested {
    return -not [string]::IsNullOrWhiteSpace($CertificateThumbprint) -or -not [string]::IsNullOrWhiteSpace($PfxPath)
}

function Invoke-SignFile {
    param(
        [string]$ResolvedSignToolPath,
        [string]$FilePath
    )

    if (-not (Test-Path $FilePath)) {
        throw "File to sign was not found: $FilePath"
    }

    $arguments = @(
        'sign',
        '/fd', 'sha256',
        '/td', 'sha256',
        '/tr', $TimestampUrl
    )

    if ($CertificateThumbprint) {
        $arguments += @('/sha1', $CertificateThumbprint)
    }
    elseif ($PfxPath) {
        $arguments += @('/f', $PfxPath)
        if ($PfxPassword) {
            $arguments += @('/p', $PfxPassword)
        }
    }
    else {
        throw 'Signing was requested, but neither -CertificateThumbprint nor -PfxPath was provided.'
    }

    $arguments += $FilePath

    Write-Host "Signing: $FilePath"
    & $ResolvedSignToolPath @arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Code signing failed for: $FilePath"
    }
}

Push-Location $root
try {
    $script:PythonCommand = Get-PythonCommand
    $signingRequested = Test-SigningRequested
    $resolvedSignToolPath = $null

    if ($signingRequested) {
        $resolvedSignToolPath = Resolve-SignToolPath -ExplicitPath $SignToolPath
    }

    if (-not $SkipDependencyInstall) {
        Write-Host 'Installing Python build dependencies...'
        Invoke-Python '-m pip install -r requirements.txt pyinstaller'
    }

    Write-Host 'Building signed-ready EXE...'
    Invoke-Python ('-m PyInstaller --noconfirm --clean --workpath "{0}" "Master Budget Automation Tool v1.0.2.spec"' -f $releaseWorkPath)
    if ($signingRequested) {
        Invoke-SignFile -ResolvedSignToolPath $resolvedSignToolPath -FilePath $appExePath
    }

    $compilerPath = Resolve-InnoSetupCompiler -ExplicitPath $InnoSetupCompiler
    Write-Host "Compiling installer with: $compilerPath"
    & $compilerPath 'installer\master_budget_tool.iss'
    if ($LASTEXITCODE -ne 0) {
        throw 'Inno Setup compilation failed.'
    }
    if ($signingRequested) {
        Invoke-SignFile -ResolvedSignToolPath $resolvedSignToolPath -FilePath $installerExePath
    }

    Write-Host ''
    Write-Host 'Store installer build complete.'
    Write-Host 'Installer output folder: dist'
    if (-not $signingRequested) {
        Write-Host 'Before Store submission, code-sign the installer EXE and upload it to a versioned HTTPS URL.'
    }
}
finally {
    Pop-Location
}
