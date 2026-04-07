param(
    [switch]$SkipDependencyInstall,
    [string]$MsixConfigPath = 'msix\msix_config.json',
    [string]$MakeAppxPath = '',
    [string]$SignToolPath = '',
    [switch]$CreateTestCertificate,
    [string]$TestCertificatePassword = '',
    [string]$TimestampUrl = 'http://timestamp.digicert.com'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$buildRoot = Join-Path $root 'build_msix'
$pyInstallerWorkPath = Join-Path $buildRoot 'pyinstaller'
$pyInstallerDistPath = Join-Path $buildRoot 'pyinstaller_dist'
$stagingRoot = Join-Path $buildRoot 'staging'
$assetsRoot = Join-Path $stagingRoot 'Assets'
$distRoot = Join-Path $root 'dist_msix'

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

function Resolve-MakeAppxPath {
    param([string]$ExplicitPath)

    if ($ExplicitPath) {
        if (Test-Path $ExplicitPath) {
            return (Resolve-Path $ExplicitPath).Path
        }
        throw "makeappx.exe was not found at: $ExplicitPath"
    }

    $localTool = Join-Path $root '.tools\msixsdk\makeappx.exe'
    if (Test-Path $localTool) {
        return (Resolve-Path $localTool).Path
    }

    $command = Get-Command makeappx.exe -ErrorAction SilentlyContinue
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
            $candidates += Get-ChildItem $sdkRoot -Recurse -Filter makeappx.exe -ErrorAction SilentlyContinue |
                Sort-Object FullName -Descending |
                Select-Object -ExpandProperty FullName
        }
    }

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) {
            return $candidate
        }
    }

    throw 'makeappx.exe was not found. Install the Windows SDK or pass -MakeAppxPath with the full path.'
}

function Resolve-SignToolPath {
    param([string]$ExplicitPath)

    if ($ExplicitPath) {
        if (Test-Path $ExplicitPath) {
            return (Resolve-Path $ExplicitPath).Path
        }
        throw "signtool.exe was not found at: $ExplicitPath"
    }

    $localTool = Join-Path $root '.tools\msixsdk\signtool.exe'
    if (Test-Path $localTool) {
        return (Resolve-Path $localTool).Path
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

function Get-MsixConfig {
    param([string]$ConfigPath)

    $resolvedPath = Join-Path $root $ConfigPath
    if (-not (Test-Path $resolvedPath)) {
        throw "MSIX config file was not found: $resolvedPath"
    }

    return Get-Content $resolvedPath -Raw | ConvertFrom-Json
}

function New-StagingDirectories {
    param([pscustomobject]$Config)

    if (Test-Path $buildRoot) {
        Remove-Item $buildRoot -Recurse -Force
    }

    New-Item -ItemType Directory -Path $buildRoot | Out-Null
    New-Item -ItemType Directory -Path $pyInstallerWorkPath | Out-Null
    New-Item -ItemType Directory -Path $pyInstallerDistPath | Out-Null
    New-Item -ItemType Directory -Path $stagingRoot | Out-Null
    New-Item -ItemType Directory -Path $assetsRoot | Out-Null

    $packageAppRoot = Join-Path $stagingRoot ("VFS\ProgramFilesX64\{0}" -f $Config.installDirectoryName)
    New-Item -ItemType Directory -Path $packageAppRoot -Force | Out-Null
    return $packageAppRoot
}

function Build-AppExecutable {
    param([pscustomobject]$Config)

    Write-Host 'Building application EXE for MSIX packaging...'
    Invoke-Python (
        '-m PyInstaller --noconfirm --clean --workpath "{0}" --distpath "{1}" "{2}"' -f
        $pyInstallerWorkPath,
        $pyInstallerDistPath,
        $Config.pyInstallerSpec
    )

    $builtExe = Join-Path $pyInstallerDistPath $Config.executableName
    if (-not (Test-Path $builtExe)) {
        throw "Expected PyInstaller output was not found: $builtExe"
    }

    return $builtExe
}

function Write-AppxManifest {
    param(
        [pscustomobject]$Config,
        [string]$ManifestPath
    )

    $templatePath = Join-Path $root 'msix\AppxManifest.template.xml'
    $template = Get-Content $templatePath -Raw

    $replacements = @{
        '__IDENTITY_NAME__' = $Config.identityName
        '__PUBLISHER__' = $Config.publisher
        '__VERSION__' = $Config.packageVersion
        '__DISPLAY_NAME__' = $Config.displayName
        '__PUBLISHER_DISPLAY_NAME__' = $Config.publisherDisplayName
        '__DESCRIPTION__' = $Config.description
        '__LANGUAGE__' = $Config.language
        '__PROCESSOR_ARCHITECTURE__' = $Config.processorArchitecture
        '__MIN_VERSION__' = $Config.minVersion
        '__MAX_VERSION_TESTED__' = $Config.maxVersionTested
        '__APPLICATION_ID__' = $Config.applicationId
        '__EXECUTABLE_PATH__' = ('VFS\ProgramFilesX64\{0}\{1}' -f $Config.installDirectoryName, $Config.executableName)
        '__BACKGROUND_COLOR__' = $Config.backgroundColor
    }

    foreach ($key in $replacements.Keys) {
        $template = $template.Replace($key, [string]$replacements[$key])
    }

    Set-Content -Path $ManifestPath -Value $template -Encoding utf8
}

function New-TestCertificate {
    param(
        [pscustomobject]$Config,
        [string]$SignTool,
        [string]$MsixPath
    )

    if (-not $CreateTestCertificate) {
        return
    }

    $subject = $Config.publisher
    $certName = "MSIX Test Certificate - $($Config.displayName)"
    $securePassword = if ($TestCertificatePassword) {
        ConvertTo-SecureString -String $TestCertificatePassword -AsPlainText -Force
    } else {
        ConvertTo-SecureString -String 'Password123!' -AsPlainText -Force
    }

    $cert = New-SelfSignedCertificate `
        -Type Custom `
        -Subject $subject `
        -FriendlyName $certName `
        -KeyUsage DigitalSignature `
        -CertStoreLocation 'Cert:\CurrentUser\My' `
        -TextExtension @('2.5.29.37={text}1.3.6.1.5.5.7.3.3')

    $pfxPath = Join-Path $distRoot "$($Config.identityName)-test-signing.pfx"
    $cerPath = Join-Path $distRoot "$($Config.identityName)-test-signing.cer"
    Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $securePassword | Out-Null
    [System.IO.File]::WriteAllBytes(
        $cerPath,
        $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
    )
    $plainPassword = if ($TestCertificatePassword) { $TestCertificatePassword } else { 'Password123!' }

    Write-Host "Signing MSIX with test certificate: $pfxPath"
    & $SignTool sign /fd sha256 /td sha256 /tr $TimestampUrl /f $pfxPath /p $plainPassword $MsixPath
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to sign MSIX package: $MsixPath"
    }

    Write-Host "Matching CER exported: $cerPath"
}

Push-Location $root
try {
    $script:PythonCommand = Get-PythonCommand
    $config = Get-MsixConfig -ConfigPath $MsixConfigPath

    if (-not $SkipDependencyInstall) {
        Write-Host 'Installing Python build dependencies...'
        Invoke-Python '-m pip install -r requirements.txt pyinstaller'
    }

    if (-not (Test-Path $distRoot)) {
        New-Item -ItemType Directory -Path $distRoot | Out-Null
    }

    $makeAppx = Resolve-MakeAppxPath -ExplicitPath $MakeAppxPath
    $signTool = if ($CreateTestCertificate) { Resolve-SignToolPath -ExplicitPath $SignToolPath } else { $null }
    $packageAppRoot = New-StagingDirectories -Config $config
    $builtExe = Build-AppExecutable -Config $config

    Copy-Item $builtExe -Destination (Join-Path $packageAppRoot $config.executableName) -Force

    & powershell -ExecutionPolicy Bypass -File (Join-Path $root 'msix\generate_msix_assets.ps1') `
        -OutputDirectory $assetsRoot `
        -AccentColor $config.backgroundColor `
        -ShortLabel $config.shortLabel `
        -AppName $config.displayName
    if ($LASTEXITCODE -ne 0) {
        throw 'MSIX visual asset generation failed.'
    }

    Write-AppxManifest -Config $config -ManifestPath (Join-Path $stagingRoot 'AppxManifest.xml')

    $msixFileName = '{0}_{1}_{2}.msix' -f $config.identityName, $config.packageVersion, $config.processorArchitecture
    $msixPath = Join-Path $distRoot $msixFileName
    if (Test-Path $msixPath) {
        Remove-Item $msixPath -Force
    }

    Write-Host "Packing MSIX with: $makeAppx"
    & $makeAppx pack /d $stagingRoot /p $msixPath /o
    if ($LASTEXITCODE -ne 0) {
        throw 'makeappx.exe failed to create the MSIX package.'
    }

    if ($CreateTestCertificate) {
        New-TestCertificate -Config $config -SignTool $signTool -MsixPath $msixPath
    }

    Write-Host ''
    Write-Host "MSIX package created: $msixPath"
    Write-Host 'Before submitting to Partner Center, replace the identity name and publisher in msix_config.json with the values from your reserved Store identity.'
    if (-not $CreateTestCertificate) {
        Write-Host 'If you want to sideload-install the package locally, rerun this script with -CreateTestCertificate.'
    }
}
finally {
    Pop-Location
}
