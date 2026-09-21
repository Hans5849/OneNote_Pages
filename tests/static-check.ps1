#requires -Version 5.1
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'OneNote-PageGuides.ps1'
$tokens = $null
$errors = $null
[void][Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$errors)
if (@($errors).Count -gt 0) {
    $errors | ForEach-Object { Write-Error $_.Message }
    exit 1
}

$source = Get-Content -LiteralPath $scriptPath -Raw
$required = @(
    "[string]`$Action = 'Status'",
    "'Configure'",
    '[CmdletBinding(SupportsShouldProcess = $true',
    'function Import-UserSettings',
    'function Save-UserSettings',
    "`$script:SettingsPath",
    'schemaVersion = 3',
    "[ValidateSet('OneNotePdfLetter','Letter','Custom')]",
    "[string]`$PageSizeProfile = 'OneNotePdfLetter'",
    '[double]$PageWidthPoints = 611.40',
    '[double]$PageHeightPoints = 792.84',
    '[double]$GuideWidthPoints = 0.0',
    '[double]$GuideHeightPoints = 0.0',
    '[double]$PageAdvancePoints = 0.0',
    '[double]$OriginXPoints = 0.0',
    '[double]$OriginYPoints = 0.0',
    '# Migration alias only. It supplies PageAdvancePoints, never guide height.',
    '$PageAdvancePoints = $PagePitchPoints',
    'PhysicalSheetWidth=$physicalW; PhysicalSheetHeight=$physicalH',
    'PrintableWidth=$printableW; PrintableHeight=$printableH',
    'GuideWidth=$GuideWidth; GuideHeight=$GuideHeight; PageAdvance=$PageAdvance',
    '$overlap=$GuideHeight-$PageAdvance',
    'Get-Geometry -Profile Letter -Mode PrintArea -GuideHeight 725.72 -PageAdvance 689.33',
    'Legacy PagePitchPoints migrates only to page advance',
    'Overlap has the requested 36.40 display rounding',
    'OutputMarginLeft=$outputLeft; OutputMarginRight=$outputRight',
    'GuideInsetLeft=$outputLeft*$scaleX; GuideInsetRight=$outputRight*$scaleX',
    'profile={5};physicalWidth={6};physicalHeight={7};guideWidth={8};guideHeight={9};overlap={10};originX={11};originY={12};',
    'function Configure-Settings',
    "Write-Host 'Margin examples:'",
    "Write-Host '  Default: left/right 1 inch; top/bottom 0.5 inch.'",
    "Write-Host '  Narrow: 0.5 inch on every side.'",
    "`$OneNote.UpdatePageContent(`$Payload.OuterXml,`$stamp,2,`$false)"
)
foreach ($text in $required) {
    if (-not $source.Contains($text)) { throw "Required source assertion missing: $text" }
}

# Guard the new contract rather than requiring the removed overloaded model.
$removed = @(
    '$Geometry.FrameWidth',
    '$Geometry.FrameHeight',
    '$Geometry.Pitch',
    'if ($Mode -eq ''PrintArea'') { $frameH = $Pitch }'
)
foreach ($text in $removed) {
    if ($source.Contains($text)) { throw "Removed geometry contract unexpectedly restored: $text" }
}

Write-Host 'PASS: PowerShell parser and static safety/configuration assertions.'
