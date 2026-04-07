param(
    [Parameter(Mandatory = $true)]
    [string]$OutputDirectory,
    [string]$AccentColor = '#0F3C5A',
    [string]$ForegroundColor = '#FFFFFF',
    [string]$ShortLabel = 'MB',
    [string]$AppName = 'Master Budget Automation Tool'
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

function New-Color {
    param([string]$HtmlColor)

    return [System.Drawing.ColorTranslator]::FromHtml($HtmlColor)
}

function New-LogoAsset {
    param(
        [string]$Path,
        [int]$Width,
        [int]$Height,
        [float]$FontScale = 0.42
    )

    $bitmap = New-Object System.Drawing.Bitmap $Width, $Height
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $graphics.Clear((New-Color '#FFFFFF'))

    $accentBrush = New-Object System.Drawing.SolidBrush (New-Color $AccentColor)
    $foregroundBrush = New-Object System.Drawing.SolidBrush (New-Color $ForegroundColor)

    $graphics.FillRectangle($accentBrush, 0, 0, $Width, $Height)
    $graphics.FillEllipse(
        (New-Object System.Drawing.SolidBrush (New-Color '#1D648F')),
        [int]($Width * 0.12),
        [int]($Height * 0.12),
        [int]($Width * 0.76),
        [int]($Height * 0.76)
    )

    $fontSize = [Math]::Max(14, [Math]::Round([Math]::Min($Width, $Height) * $FontScale))
    $font = New-Object System.Drawing.Font('Segoe UI', $fontSize, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $textSize = $graphics.MeasureString($ShortLabel, $font)
    $textX = ($Width - $textSize.Width) / 2
    $textY = ($Height - $textSize.Height) / 2
    $graphics.DrawString($ShortLabel, $font, $foregroundBrush, $textX, $textY)

    $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)

    $font.Dispose()
    $foregroundBrush.Dispose()
    $accentBrush.Dispose()
    $graphics.Dispose()
    $bitmap.Dispose()
}

function New-SplashAsset {
    param(
        [string]$Path,
        [int]$Width,
        [int]$Height
    )

    $bitmap = New-Object System.Drawing.Bitmap $Width, $Height
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $graphics.Clear((New-Color $AccentColor))

    $titleBrush = New-Object System.Drawing.SolidBrush (New-Color $ForegroundColor)
    $subtitleBrush = New-Object System.Drawing.SolidBrush (New-Color '#D7EAF6')
    $fontTitle = New-Object System.Drawing.Font('Segoe UI', 34, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $fontSubtitle = New-Object System.Drawing.Font('Segoe UI', 16, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)

    $graphics.FillEllipse(
        (New-Object System.Drawing.SolidBrush (New-Color '#1D648F')),
        32,
        32,
        120,
        120
    )
    $graphics.DrawString($ShortLabel, (New-Object System.Drawing.Font('Segoe UI', 48, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)), $titleBrush, 48, 42)
    $graphics.DrawString($AppName, $fontTitle, $titleBrush, 180, 88)
    $graphics.DrawString('MSIX package preview', $fontSubtitle, $subtitleBrush, 182, 132)

    $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)

    $fontTitle.Dispose()
    $fontSubtitle.Dispose()
    $subtitleBrush.Dispose()
    $titleBrush.Dispose()
    $graphics.Dispose()
    $bitmap.Dispose()
}

New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null

New-LogoAsset -Path (Join-Path $OutputDirectory 'StoreLogo.png') -Width 50 -Height 50 -FontScale 0.34
New-LogoAsset -Path (Join-Path $OutputDirectory 'Square44x44Logo.png') -Width 44 -Height 44 -FontScale 0.34
New-LogoAsset -Path (Join-Path $OutputDirectory 'Square71x71Logo.png') -Width 71 -Height 71 -FontScale 0.36
New-LogoAsset -Path (Join-Path $OutputDirectory 'Square150x150Logo.png') -Width 150 -Height 150 -FontScale 0.40
New-LogoAsset -Path (Join-Path $OutputDirectory 'Wide310x150Logo.png') -Width 310 -Height 150 -FontScale 0.28
New-SplashAsset -Path (Join-Path $OutputDirectory 'SplashScreen.png') -Width 620 -Height 300
