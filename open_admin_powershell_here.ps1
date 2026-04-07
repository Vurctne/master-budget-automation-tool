param(
    [string]$WorkingDirectory = ''
)

$targetDirectory = if ($WorkingDirectory) {
    $WorkingDirectory
} else {
    Split-Path -Parent $MyInvocation.MyCommand.Path
}

$resolvedDirectory = (Resolve-Path $targetDirectory).Path
$quotedDirectory = $resolvedDirectory.Replace("'", "''")
$command = "Set-Location -LiteralPath '$quotedDirectory'"

Start-Process powershell.exe -Verb RunAs -ArgumentList @(
    '-NoExit',
    '-Command',
    $command
)
