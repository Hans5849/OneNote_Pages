#requires -Version 5.1
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$oldMode = $env:ONENOTE_PAGE_GUIDES_LIBRARY_MODE
$env:ONENOTE_PAGE_GUIDES_LIBRARY_MODE = '1'
try { . (Join-Path $root 'OneNote-PageGuides.ps1') }
finally { $env:ONENOTE_PAGE_GUIDES_LIBRARY_MODE = $oldMode }

$count = 0
function Assert-Equal($Actual,$Expected,[string]$Name,[double]$Tolerance = 0) {
    if ($Actual -is [double] -or $Expected -is [double]) {
        if ([Math]::Abs([double]$Actual-[double]$Expected) -gt $Tolerance) { throw "$Name`: expected $Expected, got $Actual" }
    } elseif ($Actual -cne $Expected) { throw "$Name`: expected '$Expected', got '$Actual'" }
    $script:count++
}
function Assert-True([bool]$Value,[string]$Name) { if (-not $Value) { throw $Name }; $script:count++ }

# Independent guide height/page advance, overlap, and first/later origins.
$g = Get-Geometry -Profile Letter -Mode PrintArea -GuideWidth 500 -GuideHeight 725.72 -PageAdvance 689.33 -OriginX 11 -OriginY 23
Assert-Equal $g.GuideHeight 725.72 'guide height remains independent'
Assert-Equal $g.PageAdvance 689.33 'page advance remains independent'
Assert-Equal $g.Overlap 36.39 'overlap is height minus advance' .000001
$payload = New-GuidePayload -Id 'geometry-test' -Geometry $g -Png ([Convert]::ToBase64String([byte[]](1..40))) -Batch 'batch' -Count 3
$ns = New-NamespaceManager $payload
$positions = @($payload.SelectNodes('/one:Page/one:Image/one:Position',$ns) | ForEach-Object { Read-Point $_.GetAttribute('y') })
Assert-Equal $positions[0] 23 'first sheet uses calibrated origin'
Assert-Equal $positions[1] 712.33 'second sheet advances from first origin' .000001
Assert-Equal $positions[2] 1401.66 'later origins do not add guide height' .000001

# Calibration and conversion of output-sheet margins into guide coordinates.
$g = Get-Geometry -Profile Letter -Mode Paper -GuideWidth 306 -GuideHeight 396 -OriginX 10 -OriginY 20 -TranslateX 3 -TranslateY 4 -Left 1 -Right 1 -Top .5 -Bottom .5
Assert-Equal $g.OriginX 13 'X translation'
Assert-Equal $g.OriginY 24 'Y translation'
Assert-Equal $g.EffectiveScaleX .5 'effective X scale'
Assert-Equal $g.EffectiveScaleY .5 'effective Y scale'
Assert-Equal $g.OutputMarginLeft 72 'horizontal output margin'
Assert-Equal $g.OutputMarginTop 36 'vertical output margin'
Assert-Equal $g.GuideInsetLeft 36 'horizontal converted guide inset'
Assert-Equal $g.GuideInsetTop 18 'vertical converted guide inset'
$guideMargins = Get-Geometry -Profile Letter -GuideWidth 306 -GuideHeight 396 -MarginMode Guide -Left 1 -Top .5
Assert-Equal $guideMargins.GuideInsetLeft 72 'Guide semantics do not scale horizontal margin'
Assert-Equal $guideMargins.GuideInsetTop 36 'Guide semantics do not scale vertical margin'

# Schema migration and per-property command-line precedence.
$saved = @{}
foreach ($name in @('SettingsPath','InvocationParameters','PageAdvancePoints','GuideHeightPoints','MarginLeft','ValueSources','SettingsSchema')) {
    $saved[$name] = Get-Variable -Name $name -Scope Script -ValueOnly
}
$temp = Join-Path ([IO.Path]::GetTempPath()) ('OneNoteGuideTests-' + [guid]::NewGuid().ToString('N'))
try {
    [void][IO.Directory]::CreateDirectory($temp); $script:SettingsPath = Join-Path $temp 'settings.json'
    @{schemaVersion=2;PagePitchPoints=688;GuideHeightPoints=730;MarginLeft=.75} | ConvertTo-Json | Set-Content $script:SettingsPath -Encoding UTF8
    $script:InvocationParameters=@(); $script:PageAdvancePoints=0; $script:GuideHeightPoints=700; $script:MarginLeft=1; $script:ValueSources=@{}
    Import-UserSettings
    Assert-Equal $PageAdvancePoints 688 'legacy pitch migrates to advance'
    Assert-Equal $GuideHeightPoints 730 'migration preserves independent height'
    $script:InvocationParameters=@('MarginLeft','PageAdvancePoints'); $script:MarginLeft=.25; $script:PageAdvancePoints=640
    Import-UserSettings
    Assert-Equal $MarginLeft .25 'explicit margin wins over saved margin'
    Assert-Equal $PageAdvancePoints 640 'explicit advance wins over legacy pitch'
} finally {
    if (Test-Path $temp) { Remove-Item $temp -Recurse -Force }
    foreach ($name in $saved.Keys) { Set-Variable -Name $name -Scope Script -Value $saved[$name] }
}

# Ordinary shortcuts contain only operational arguments; explicit overrides are
# constructed elsewhere and are deliberately frozen.
$shortcut = Get-RefreshShortcutArguments '-NoProfile -File "utility.ps1"' 7 -UseConsolePicker
Assert-Equal $shortcut '-NoProfile -File "utility.ps1" -Action Refresh -Pages 7 -ConsolePicker' 'ordinary shortcut construction'
Assert-True ($shortcut -notmatch 'Guide|Origin|Margin|Calibration') 'ordinary shortcut does not freeze geometry'

Write-Host "PASS: $count behavior-oriented geometry/helper assertions."
