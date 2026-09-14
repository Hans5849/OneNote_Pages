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
    "`$OneNote.UpdatePageContent(`$Payload.OuterXml,`$stamp,2,`$false)"
)
foreach ($text in $required) {
    if (-not $source.Contains($text)) { throw "Required source assertion missing: $text" }
}

Write-Host 'PASS: PowerShell parser and static safety/configuration assertions.'
