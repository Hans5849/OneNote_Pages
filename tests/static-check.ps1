#requires -Version 5.1
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'OneNote-PageGuides.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($scriptPath,[ref]$tokens,[ref]$errors)
if (@($errors).Count) { $errors | ForEach-Object { Write-Error $_.Message }; exit 1 }

# Keep only structural/safety checks here. Geometry, migration, precedence and
# shortcut contracts are exercised through the real helpers below.
$functions = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] },$true).Name)
foreach ($name in @('Get-Geometry','Import-UserSettings','Get-RefreshShortcutArguments','Assert-GuidePayload')) {
    if ($functions -notcontains $name) { throw "Required function missing: $name" }
}
$updateCalls = @($ast.FindAll({ param($node)
    $node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and $node.Member.Value -eq 'UpdatePageContent'
},$true))
if ($updateCalls.Count -eq 0) { throw 'No OneNote UpdatePageContent call found.' }
& (Join-Path $PSScriptRoot 'geometry-helper-tests.ps1')
Write-Host 'PASS: parser, structural safety, and behavior-oriented helper tests.'
