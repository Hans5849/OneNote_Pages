#requires -Version 5.1
<#
.SYNOPSIS
    OneNote Page Guides 2.6.0 - removable visual guides, not native print boundaries.
.DESCRIPTION
    Windows desktop OneNote only. No add-in, subscription, or administrator rights.
    Status is the read-only default. A page is selected ONCE via the notebook
    hierarchy (or -PageId); the Windows COM collection is deliberately not used.

    Add/Refresh inserts background PNGs containing left, right, top and bottom
    boundaries. Refresh stages and verifies new guides BEFORE deleting old ones.
    Only top-level images with this utility's V2 metadata may be deleted.
    A local recovery journal embeds the OLD guide PNG bytes before any write.
    Use Recover after an interrupted/failed change; it does not restore notes.

    Paper mode shows full calibrated sheets plus optional margin rectangles.
    PrintArea mode shows content-sized frames, not physical paper/margins.
    PagePitchPoints is a migration alias for PageAdvancePoints. Guide height and
    page advance are independent. NO interval is guaranteed to match OneNote
    printing. Calibrate against a disposable-page export.
    Remove guides before print/export. Do not edit the page while a write runs.

    Designed for Windows PowerShell 5.1. SelfTest needs Windows/System.Drawing,
    but does not connect to OneNote, install anything, or modify notebooks.
.EXAMPLE
    .\OneNote-PageGuides.ps1 -Action SelfTest
.EXAMPLE
    .\OneNote-PageGuides.ps1 -Action Install -CreateDesktopShortcuts
.EXAMPLE
    .\OneNote-PageGuides.ps1 -Action Configure
.EXAMPLE
    .\OneNote-PageGuides.ps1 -Action Refresh -Pages 2 -WhatIf
.EXAMPLE
    .\OneNote-PageGuides.ps1 -Action Add -Pages 2
.EXAMPLE
    .\OneNote-PageGuides.ps1 -Action Refresh -GuideMode PrintArea -PagePitchPoints 720
.EXAMPLE
    .\OneNote-PageGuides.ps1 -Action Recover -RecoveryFile 'C:\path\transaction.json'
.NOTES
    Version: 2.6.0
    Adds persistent page-size profiles, custom geometry, pitch, and margins.
    Explicit parameters always override saved values for the current invocation.
    Sources: Microsoft OneNote desktop Application interface / Enumerations;
    OneNoteApplication_2013.xsd. No web requests are made by this script.
    Recovery is guide-only and best-effort, not a notebook backup or transaction.
    Full Windows OneNote integration and print calibration require local testing.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [ValidateSet('Status','Add','Refresh','Remove','Recover','Configure','Install','Uninstall','SelfTest')]
    [string]$Action = 'Status',
    [ValidateRange(1,100)][int]$Pages = 10,
    [ValidateSet('Portrait','Landscape')][string]$Orientation = 'Portrait',
    [ValidateSet('Paper','PrintArea')][string]$GuideMode = 'Paper',
    [ValidateSet('OneNotePdfLetter','Letter','Custom')][string]$PageSizeProfile = 'OneNotePdfLetter',
    [ValidateRange(10.0,10000.0)][double]$PageWidthPoints = 611.40,
    [ValidateRange(10.0,10000.0)][double]$PageHeightPoints = 792.84,
    [ValidateRange(0.0,4.0)][double]$MarginLeft = 1.0,
    [ValidateRange(0.0,4.0)][double]$MarginRight = 1.0,
    [ValidateRange(0.0,4.0)][double]$MarginTop = 0.5,
    [ValidateRange(0.0,4.0)][double]$MarginBottom = 0.5,
    [ValidateRange(0.0,900000.0)][double]$OriginXPoints = 0.0,
    [ValidateRange(0.0,900000.0)][double]$OriginYPoints = 0.0,
    [ValidateRange(0.0,10000.0)][double]$GuideWidthPoints = 0.0,
    [ValidateRange(0.0,10000.0)][double]$GuideHeightPoints = 0.0,
    [ValidateRange(0.0,10000.0)][double]$PageAdvancePoints = 0.0,
    # Migration alias only. It supplies PageAdvancePoints, never guide height.
    [ValidateRange(0.0,10000.0)][double]$PagePitchPoints = 0.0,
    [Alias('StartX')][ValidateRange(0.0,900000.0)][double]$LegacyOriginXPoints = 0.0,
    [Alias('StartY')][ValidateRange(0.0,900000.0)][double]$LegacyOriginYPoints = 0.0,
    [switch]$HideMargins,
    [string]$PageId = '',
    [string]$PageFilter = '',
    [switch]$ConsolePicker,
    [string]$RecoveryFile = '',
    [ValidateRange(1,100)][int]$ShortcutPages = 10,
    [switch]$CreateDesktopShortcuts,
    [switch]$ResetSettings
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:Version = '2.6.0'
$script:OneNamespace = 'http://schemas.microsoft.com/office/onenote/2013/onenote'
$script:MetaName = 'OneNotePageGuidesV2'
$script:RecoveryMetaName = 'OneNotePageGuidesRecovery'
$script:InstallDirectory = ''
$script:InstalledScript = ''
$script:BackupDirectory = ''
$script:StartMenuDirectory = ''
$script:SettingsPath = ''
$script:RunningFile = $PSCommandPath
$script:JournalPath = ''
$script:Journal = $null
$script:InvocationParameters = @($PSBoundParameters.Keys)
$script:PageWidthAvailable = $script:InvocationParameters -contains 'PageWidthPoints'
$script:PageHeightAvailable = $script:InvocationParameters -contains 'PageHeightPoints'

if ($script:InvocationParameters -contains 'PagePitchPoints') {
    if ($script:InvocationParameters -contains 'PageAdvancePoints') { throw 'Specify PageAdvancePoints or the legacy PagePitchPoints alias, not both.' }
    $PageAdvancePoints = $PagePitchPoints
}
if ($script:InvocationParameters -contains 'LegacyOriginXPoints') { $OriginXPoints = $LegacyOriginXPoints }
if ($script:InvocationParameters -contains 'LegacyOriginYPoints') { $OriginYPoints = $LegacyOriginYPoints }

function Assert-Windows {
    if ($env:OS -ne 'Windows_NT') {
        throw 'Use Windows PowerShell 5.1 on Windows with desktop OneNote. No changes made.'
    }
}

function Initialize-Paths {
    foreach ($name in @('LOCALAPPDATA','APPDATA')) {
        if ([string]::IsNullOrWhiteSpace([Environment]::GetEnvironmentVariable($name))) {
            throw "Environment variable $name is not available. Run as your normal Windows user."
        }
    }
    $script:InstallDirectory = Join-Path $env:LOCALAPPDATA 'OneNotePageGuides'
    # Keep the old installed filename so existing shortcuts continue to resolve.
    $script:InstalledScript = Join-Path $script:InstallDirectory 'OneNote-PageGuides-V2.ps1'
    $script:BackupDirectory = Join-Path $script:InstallDirectory 'Backups'
    $script:SettingsPath = Join-Path $script:InstallDirectory 'settings.json'
    $script:StartMenuDirectory = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\OneNote Page Guides'
}

function Import-UserSettings {
    if (-not (Test-Path -LiteralPath $script:SettingsPath -PathType Leaf)) { return }
    try {
        $settings = Get-Content -LiteralPath $script:SettingsPath -Raw | ConvertFrom-Json
        $schema = 1
        if ($null -ne $settings.PSObject.Properties['schemaVersion']) { $schema = [int]$settings.schemaVersion }
        if ($schema -lt 1 -or $schema -gt 3) { throw "Unsupported settings schema version $schema." }
        foreach ($name in @('MarginLeft','MarginRight','MarginTop','MarginBottom')) {
            if ($script:InvocationParameters -contains $name) { continue }
            $property = $settings.PSObject.Properties[$name]
            if ($null -eq $property) { continue }
            $value = [double]$property.Value
            if ([double]::IsNaN($value) -or [double]::IsInfinity($value) -or $value -lt 0 -or $value -gt 4) {
                throw "Saved $name must be between 0 and 4 inches."
            }
            Set-Variable -Name $name -Value $value -Scope Script
        }
        if ($script:InvocationParameters -notcontains 'PageSizeProfile' -and
            $null -ne $settings.PSObject.Properties['PageSizeProfile']) {
            $savedProfile = [string]$settings.PageSizeProfile
            if ($savedProfile -notin @('OneNotePdfLetter','Letter','Custom')) { throw 'Saved PageSizeProfile is invalid.' }
            $script:PageSizeProfile = $savedProfile
        }
        foreach ($name in @('PageWidthPoints','PageHeightPoints')) {
            $property = $settings.PSObject.Properties[$name]
            if ($script:InvocationParameters -notcontains $name -and $null -ne $property) {
                $value = [double]$property.Value
                if ([double]::IsNaN($value) -or [double]::IsInfinity($value) -or $value -lt 10 -or $value -gt 10000) {
                    throw "Saved $name must be between 10 and 10000 points."
                }
                Set-Variable -Name $name -Value $value -Scope Script
                Set-Variable -Name ($name.Replace('Points','Available')) -Value $true -Scope Script
            }
        }
        foreach ($name in @('GuideWidthPoints','GuideHeightPoints','PageAdvancePoints','OriginXPoints','OriginYPoints')) {
            if ($script:InvocationParameters -contains $name -or
                ($name -eq 'PageAdvancePoints' -and $script:InvocationParameters -contains 'PagePitchPoints')) { continue }
            $property=$settings.PSObject.Properties[$name]
            if ($null -eq $property) { continue }
            $value=[double]$property.Value
            $maximum=if ($name -in @('OriginXPoints','OriginYPoints')) { 900000 } else { 10000 }
            if ([double]::IsNaN($value) -or [double]::IsInfinity($value) -or $value -lt 0 -or $value -gt $maximum) {
                throw "Saved $name must be 0 through $maximum points."
            }
            Set-Variable -Name $name -Value $value -Scope Script
        }
        # Schema 1/2 migration: pitch maps only to advance. It must never resize a guide.
        if ($script:InvocationParameters -notcontains 'PageAdvancePoints' -and
            $null -eq $settings.PSObject.Properties['PageAdvancePoints'] -and
            $null -ne $settings.PSObject.Properties['PagePitchPoints']) {
            $value=[double]$settings.PagePitchPoints
            if ([double]::IsNaN($value) -or [double]::IsInfinity($value) -or $value -lt 0 -or $value -gt 10000) {
                throw 'Saved PagePitchPoints must be 0 through 10000 points.'
            }
            $script:PageAdvancePoints=$value
        }
    }
    catch {
        throw "Settings file '$($script:SettingsPath)' is invalid. Use -Action Configure -ResetSettings to remove it. $($_.Exception.Message)"
    }
}

function Save-UserSettings {
    [void][IO.Directory]::CreateDirectory($script:InstallDirectory)
    $settings = [ordered]@{
        schemaVersion = 3
        PageSizeProfile = $PageSizeProfile
        PageWidthPoints = $PageWidthPoints
        PageHeightPoints = $PageHeightPoints
        GuideWidthPoints = $GuideWidthPoints
        GuideHeightPoints = $GuideHeightPoints
        PageAdvancePoints = $PageAdvancePoints
        OriginXPoints = $OriginXPoints
        OriginYPoints = $OriginYPoints
        MarginLeft = $MarginLeft
        MarginRight = $MarginRight
        MarginTop = $MarginTop
        MarginBottom = $MarginBottom
    }
    $temporary = $script:SettingsPath + '.tmp'
    $settings | ConvertTo-Json | Set-Content -LiteralPath $temporary -Encoding UTF8
    Move-Item -LiteralPath $temporary -Destination $script:SettingsPath -Force
}

function Read-MarginSetting {
    param([Parameter(Mandatory)][string]$Name, [double]$Current)
    $answer = Read-Host ("{0} in inches (0-4) [{1}]" -f $Name,(Format-Point $Current))
    if ([string]::IsNullOrWhiteSpace($answer)) { return $Current }
    $value = 0.0
    $ok = [double]::TryParse($answer, [Globalization.NumberStyles]::Float,
        [Globalization.CultureInfo]::InvariantCulture, [ref]$value)
    if (-not $ok -or [double]::IsNaN($value) -or [double]::IsInfinity($value) -or $value -lt 0 -or $value -gt 4) {
        throw "Invalid $Name '$answer'. Enter a number from 0 through 4 using a period as the decimal separator."
    }
    return $value
}

function Read-GeometrySetting {
    param([Parameter(Mandatory)][string]$Name, [double]$Current, [switch]$AllowAutomatic)
    $range = if ($AllowAutomatic) { '0 (automatic), or 10-10000' } else { '10-10000' }
    $answer = Read-Host ("{0} in points ({1}) [{2}]" -f $Name,$range,(Format-Point $Current))
    if ([string]::IsNullOrWhiteSpace($answer)) { return $Current }
    $value = 0.0
    $ok = [double]::TryParse($answer,[Globalization.NumberStyles]::Float,
        [Globalization.CultureInfo]::InvariantCulture,[ref]$value)
    if (-not $ok -or [double]::IsNaN($value) -or [double]::IsInfinity($value) -or
        (($value -lt 10 -or $value -gt 10000) -and -not ($AllowAutomatic -and $value -eq 0))) {
        throw "Invalid $Name '$answer'. Enter $range using a period as the decimal separator."
    }
    return $value
}

function Configure-Settings {
    if ($ResetSettings) {
        if (Test-Path -LiteralPath $script:SettingsPath -PathType Leaf) {
            Remove-Item -LiteralPath $script:SettingsPath -Force -Confirm:$false
        }
        Write-Host 'Saved settings removed. Defaults restored: OneNotePdfLetter, automatic pitch, and 1/1/0.5/0.5 inch margins.'
        return
    }
    $explicit = @('MarginLeft','MarginRight','MarginTop','MarginBottom','PageSizeProfile',
        'PageWidthPoints','PageHeightPoints','GuideWidthPoints','GuideHeightPoints','PageAdvancePoints','OriginXPoints','OriginYPoints') |
        Where-Object { $script:InvocationParameters -contains $_ }
    if (@($explicit).Count -eq 0) {
        Write-Host "Current page-size profile: $PageSizeProfile"
        $answer = Read-Host 'Page profile: OneNotePdfLetter, Letter, or Custom (Enter keeps current)'
        if (-not [string]::IsNullOrWhiteSpace($answer)) {
            $match = @('OneNotePdfLetter','Letter','Custom') | Where-Object { $_ -ieq $answer }
            if (@($match).Count -ne 1) { throw "Invalid page profile '$answer'." }
            $script:PageSizeProfile = $match[0]
        }
        if ($PageSizeProfile -eq 'Custom') {
            $script:PageWidthPoints = Read-GeometrySetting 'Portrait page width' $PageWidthPoints
            $script:PageHeightPoints = Read-GeometrySetting 'Portrait page height' $PageHeightPoints
            $script:PageWidthAvailable = $true; $script:PageHeightAvailable = $true
        }
        $script:GuideWidthPoints = Read-GeometrySetting 'Guide width' $GuideWidthPoints -AllowAutomatic
        $script:GuideHeightPoints = Read-GeometrySetting 'Guide height' $GuideHeightPoints -AllowAutomatic
        $script:PageAdvancePoints = Read-GeometrySetting 'Page advance' $PageAdvancePoints -AllowAutomatic
        Write-Host 'Margin examples:'
        Write-Host '  Default: left/right 1 inch; top/bottom 0.5 inch.'
        Write-Host '  Narrow: 0.5 inch on every side.'
        Write-Host 'Press Enter to keep the value shown in brackets.'
        $script:MarginLeft = Read-MarginSetting 'Left margin' $MarginLeft
        $script:MarginRight = Read-MarginSetting 'Right margin' $MarginRight
        $script:MarginTop = Read-MarginSetting 'Top margin' $MarginTop
        $script:MarginBottom = Read-MarginSetting 'Bottom margin' $MarginBottom
    }
    $configuredGeometry = Get-Geometry
    Save-UserSettings
    Write-Host ("Saved geometry: profile {0}; physical {1} x {2} pt; page advance {3}." -f `
        $PageSizeProfile,(Format-Point $configuredGeometry.PhysicalSheetWidth),(Format-Point $configuredGeometry.PhysicalSheetHeight),
        $(if ($PageAdvancePoints -eq 0) { 'automatic' } else { (Format-Point $PageAdvancePoints) + ' pt' }))
    Write-Host ("Saved margins (inches): left {0}, right {1}, top {2}, bottom {3}." -f `
        (Format-Point $MarginLeft),(Format-Point $MarginRight),(Format-Point $MarginTop),(Format-Point $MarginBottom))
    Write-Host 'Refresh and Add now use these values unless parameters are supplied explicitly.'
}

function Format-Point {
    param([double]$Value)
    return $Value.ToString('0.###', [Globalization.CultureInfo]::InvariantCulture)
}

function Read-Point {
    param([string]$Text)
    $value = 0.0
    $ok = [double]::TryParse($Text, [Globalization.NumberStyles]::Float,
        [Globalization.CultureInfo]::InvariantCulture, [ref]$value)
    if (-not $ok -or [double]::IsNaN($value) -or [double]::IsInfinity($value) -or
        $value -lt 0 -or $value -gt 1000000) {
        throw "Invalid OneNote coordinate '$Text'."
    }
    return $value
}

function Read-SafeXml {
    param([Parameter(Mandatory)][string]$Text)
    $settings = [Xml.XmlReaderSettings]::new()
    $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver = $null
    $inputText = [IO.StringReader]::new($Text)
    $reader = $null
    try {
        $reader = [Xml.XmlReader]::Create($inputText, $settings)
        $doc = [Xml.XmlDocument]::new()
        $doc.XmlResolver = $null
        $doc.Load($reader)
        return ,$doc
    }
    finally {
        if ($null -ne $reader) { $reader.Dispose() }
        $inputText.Dispose()
    }
}

function New-NamespaceManager {
    param([Parameter(Mandatory)][xml]$Document)
    $manager = [Xml.XmlNamespaceManager]::new($Document.NameTable)
    $manager.AddNamespace('one', $script:OneNamespace)
    return ,$manager
}

function New-PageDocument {
    param([Parameter(Mandatory)][string]$Id)
    $doc = [Xml.XmlDocument]::new()
    $root = $doc.CreateElement('one', 'Page', $script:OneNamespace)
    $root.SetAttribute('ID', $Id)
    [void]$doc.AppendChild($root)
    return ,$doc
}

function Add-XmlChild {
    param([Parameter(Mandatory)][Xml.XmlElement]$Parent,
          [Parameter(Mandatory)][string]$Name,
          [hashtable]$Attributes = @{}, [string]$Text = '')
    $child = $Parent.OwnerDocument.CreateElement('one', $Name, $script:OneNamespace)
    foreach ($key in $Attributes.Keys) { $child.SetAttribute([string]$key, [string]$Attributes[$key]) }
    if ($Text.Length -gt 0) { $child.InnerText = $Text }
    [void]$Parent.AppendChild($child)
    return ,$child
}

function Release-ComObjectSafe {
    param($Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value) } catch { }
    }
}

function Get-OneNoteApplication {
    try {
        $app = New-Object -ComObject OneNote.Application
        return ,$app
    }
    catch {
        throw "Desktop OneNote COM could not be opened. Run OneNote and PowerShell as the same normal user. $($_.Exception.Message)"
    }
}

function Test-ExcludedHierarchyNode {
    param([Parameter(Mandatory)][Xml.XmlElement]$Node)
    $cursor = $Node
    while ($null -ne $cursor -and $cursor -is [Xml.XmlElement]) {
        foreach ($attribute in @('isRecycleBin','isInRecycleBin','isDeletedPages','isDeleted','isLocked','isReadOnly')) {
            if ($cursor.GetAttribute($attribute) -in @('true','1')) { return $true }
        }
        $cursor = $cursor.ParentNode
    }
    return $false
}

function Convert-HierarchyToPages {
    param([Parameter(Mandatory)][xml]$Document)
    $mgr = New-NamespaceManager $Document
    # Emit records to the pipeline, avoiding @($genericList) binding problems.
    foreach ($node in $Document.SelectNodes('//one:Page', $mgr)) {
        if (Test-ExcludedHierarchyNode $node) { continue }
        $id = $node.GetAttribute('ID')
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        $section = ''
        $notebook = ''
        $viewed = $false
        $groups = @()
        $cursor = $node.ParentNode
        while ($null -ne $cursor -and $cursor -is [Xml.XmlElement]) {
            switch ($cursor.LocalName) {
                'Section' { $section = $cursor.GetAttribute('name') }
                'SectionGroup' { $groups = @($cursor.GetAttribute('name')) + $groups }
                'Notebook' {
                    $notebook = $cursor.GetAttribute('name')
                    $viewed = $cursor.GetAttribute('isCurrentlyViewed') -in @('true','1')
                }
            }
            $cursor = $cursor.ParentNode
        }
        $modified = [datetime]::MinValue
        $raw = $node.GetAttribute('lastModifiedTime')
        if (-not [string]::IsNullOrWhiteSpace($raw)) {
            [void][datetime]::TryParse($raw, [Globalization.CultureInfo]::InvariantCulture,
                [Globalization.DateTimeStyles]::RoundtripKind, [ref]$modified)
        }
        [pscustomobject]@{
            PageName = $node.GetAttribute('name')
            Notebook = $notebook
            Section = (@($groups) + @($section)) -join ' / '
            CurrentlyViewed = $viewed
            LastModified = $modified
            PageId = $id
        }
    }
}

function Get-HierarchyPages {
    param($OneNote)
    [string]$text = ''
    Write-Verbose 'Reading notebook hierarchy (hsPages=4, xs2013=2); not using the Windows collection.'
    $OneNote.GetHierarchy('', 4, [ref]$text, 2)
    if ([string]::IsNullOrWhiteSpace($text)) { throw 'OneNote returned an empty hierarchy.' }
    $doc = Read-SafeXml $text
    if ($doc.DocumentElement.NamespaceURI -ne $script:OneNamespace) {
        throw 'OneNote returned an unexpected XML schema.'
    }
    Convert-HierarchyToPages $doc
}

function Select-TargetPage {
    param([Parameter(Mandatory)][object[]]$AvailablePages,
          [string]$ExplicitId = '', [string]$Filter = '', [switch]$UseConsole)
    if ($ExplicitId.Length -gt 0) {
        $matched = @($AvailablePages | Where-Object { $_.PageId -ceq $ExplicitId })
        if ($matched.Count -ne 1) {
            throw 'The supplied page ID is not a unique live, unlocked page in the open notebooks.'
        }
        return $matched[0]
    }
    $ordered = @($AvailablePages | Sort-Object -Property `
        @{Expression='CurrentlyViewed';Descending=$true}, `
        @{Expression='LastModified';Descending=$true}, Notebook, Section, PageName)
    if ($Filter.Length -gt 0) {
        $ordered = @($ordered | Where-Object {
            ('{0} {1} {2}' -f $_.Notebook,$_.Section,$_.PageName).IndexOf(
                $Filter, [StringComparison]::OrdinalIgnoreCase) -ge 0
        })
    }
    if ($ordered.Count -eq 0) { throw 'No live pages match. Open a notebook or change -PageFilter.' }
    if (-not $UseConsole -and $null -ne (Get-Command Out-GridView -ErrorAction SilentlyContinue)) {
        $gridFailed = $false
        $picked = @()
        try {
            $picked = @($ordered | Out-GridView -Title 'OneNote Page Guides 2.6.0 - choose ONE target page' -OutputMode Single)
        }
        catch {
            $gridFailed = $true
            Write-Verbose "Graphical picker unavailable; using console: $($_.Exception.Message)"
        }
        if (-not $gridFailed) {
            if ($picked.Count -eq 0) { return $null } # Cancel really cancels.
            if ($picked.Count -ne 1) { throw 'Select exactly one page.' }
            return $picked[0]
        }
    }
    $offset = 0
    $search = ''
    :PickerLoop while ($true) {
        $shown = @($ordered | Where-Object {
            ('{0} {1} {2}' -f $_.Notebook,$_.Section,$_.PageName).IndexOf(
                $search, [StringComparison]::OrdinalIgnoreCase) -ge 0
        })
        if ($offset -ge $shown.Count) { $offset = 0 }
        Write-Host ''
        Write-Host ("Matching pages: {0}. N=next, P=previous, F=filter, Q=cancel." -f $shown.Count)
        $end = [Math]::Min($shown.Count, $offset + 20)
        for ($i = $offset; $i -lt $end; $i++) {
            $row = $shown[$i]
            Write-Host ('{0,4}. {1}  [{2} / {3}]' -f ($i+1),$row.PageName,$row.Notebook,$row.Section)
        }
        $answer = (Read-Host 'Select a number or command').Trim()
        switch ($answer.ToUpperInvariant()) {
            'Q' { return $null }
            'N' { if ($end -lt $shown.Count) { $offset += 20 }; continue PickerLoop }
            'P' { $offset = [Math]::Max(0, $offset-20); continue PickerLoop }
            'F' { $search = Read-Host 'Filter text (blank = all)'; $offset = 0; continue PickerLoop }
        }
        $number = 0
        if ([int]::TryParse($answer, [ref]$number) -and $number -ge 1 -and $number -le $shown.Count) {
            return $shown[$number-1]
        }
        Write-Warning 'Enter a listed number, N, P, F, or Q.'
    }
}

function Read-OneNotePage {
    param($OneNote, [Parameter(Mandatory)][string]$Id)
    [string]$text = ''
    $OneNote.GetPageContent($Id, [ref]$text, 0, 2)
    if ([string]::IsNullOrWhiteSpace($text)) { throw 'OneNote returned empty page content.' }
    $doc = Read-SafeXml $text
    if ($doc.DocumentElement.LocalName -ne 'Page' -or
        $doc.DocumentElement.NamespaceURI -ne $script:OneNamespace -or
        $doc.DocumentElement.GetAttribute('ID') -cne $Id) {
        throw 'OneNote returned a different page or an unsupported XML schema. Stopped.'
    }
    [pscustomobject]@{ Id=$Id; Xml=$doc; Name=$doc.DocumentElement.GetAttribute('name') }
}

function Get-PageTimestamp {
    param([Parameter(Mandatory)]$Page)
    $raw = $Page.Xml.DocumentElement.GetAttribute('lastModifiedTime')
    if ([string]::IsNullOrWhiteSpace($raw)) {
        throw 'No page modification timestamp was returned; refusing an unguarded write.'
    }
    $stamp = [Xml.XmlConvert]::ToDateTime($raw, [Xml.XmlDateTimeSerializationMode]::Utc)
    if ($stamp -le [datetime]'1900-01-01') { throw 'Invalid page modification timestamp.' }
    return $stamp
}

function Get-GuideImages {
    param([Parameter(Mandatory)][xml]$Document, [string]$Batch = '')
    $mgr = New-NamespaceManager $Document
    $path = "/one:Page/one:Image[one:Meta[@name='$($script:MetaName)' and starts-with(@content,'v2;')]]"
    foreach ($node in $Document.SelectNodes($path, $mgr)) {
        if ($Batch.Length -gt 0) {
            $meta = $node.SelectSingleNode("one:Meta[@name='$($script:MetaName)']", $mgr)
            if (-not $meta.GetAttribute('content').Contains(';batch=' + $Batch + ';')) { continue }
        }
        Write-Output -NoEnumerate $node
    }
}

function Get-ImageDescriptor {
    param([Parameter(Mandatory)][Xml.XmlElement]$Image)
    $mgr = New-NamespaceManager $Image.OwnerDocument
    $markerNodes = @($Image.SelectNodes("one:Meta[@name='$($script:MetaName)']", $mgr))
    if ($markerNodes.Count -ne 1 -or -not $markerNodes[0].GetAttribute('content').StartsWith('v2;')) {
        throw 'Missing or ambiguous utility metadata; this image will not be modified.'
    }
    $position = $Image.SelectSingleNode('one:Position', $mgr)
    $size = $Image.SelectSingleNode('one:Size', $mgr)
    if ($null -eq $position -or $null -eq $size) { throw 'A guide has no position or size.' }
    [pscustomobject]@{
        Id = $Image.GetAttribute('objectID')
        Marker = $markerNodes[0].GetAttribute('content')
        X = Read-Point $position.GetAttribute('x')
        Y = Read-Point $position.GetAttribute('y')
        Width = Read-Point $size.GetAttribute('width')
        Height = Read-Point $size.GetAttribute('height')
        Background = $Image.GetAttribute('backgroundImage') -in @('true','1')
        Modified = $Image.GetAttribute('lastModifiedTime')
    }
}

function Assert-ImageMatches {
    param([Parameter(Mandatory)][Xml.XmlElement]$Expected,
          [Parameter(Mandatory)][Xml.XmlElement]$Actual,
          [switch]$CheckModified)
    $left = Get-ImageDescriptor $Expected
    $right = Get-ImageDescriptor $Actual
    if ($left.Marker -cne $right.Marker -or $left.Background -ne $right.Background) {
        throw 'Guide metadata or background status changed; operation stopped.'
    }
    foreach ($property in @('X','Y','Width','Height')) {
        if ([Math]::Abs($left.$property - $right.$property) -gt 0.25) {
            throw "Guide $property does not match the expected geometry. No further deletions will be made."
        }
    }
    if ($CheckModified -and $left.Modified.Length -gt 0 -and $left.Modified -cne $right.Modified) {
        throw 'A guide was edited after preflight. Stopped instead of deleting a changed object.'
    }
    if ($right.Width -le 0 -or $right.Height -le 0) { throw 'OneNote returned an empty guide size.' }
}

function Get-Geometry {
    param([string]$PaperOrientation = $Orientation, [string]$Mode = $GuideMode,
          [double]$Left = $MarginLeft, [double]$Right = $MarginRight,
          [double]$Top = $MarginTop, [double]$Bottom = $MarginBottom,
          [double]$GuideWidth = $GuideWidthPoints,
          [double]$GuideHeight = $GuideHeightPoints,
          [double]$PageAdvance = $PageAdvancePoints,
          [double]$OriginX = $OriginXPoints, [double]$OriginY = $OriginYPoints,
          [string]$Profile = $PageSizeProfile,
          [double]$CustomWidth = $PageWidthPoints,
          [double]$CustomHeight = $PageHeightPoints,
          [switch]$CustomDimensionsAvailable)
    foreach ($value in @($Left,$Right,$Top,$Bottom,$GuideWidth,$GuideHeight,$PageAdvance,$OriginX,$OriginY,$CustomWidth,$CustomHeight)) {
        if ([double]::IsNaN($value) -or [double]::IsInfinity($value) -or $value -lt 0) {
            throw 'Geometry must use finite, nonnegative numbers.'
        }
    }
    switch ($Profile) {
        'OneNotePdfLetter' { $portraitW = 611.40; $portraitH = 792.84 }
        'Letter' { $portraitW = 612.0; $portraitH = 792.0 }
        'Custom' {
            $available = $CustomDimensionsAvailable -or ($script:PageWidthAvailable -and $script:PageHeightAvailable)
            if (-not $available) { throw 'Custom profile requires PageWidthPoints and PageHeightPoints from saved settings or command-line parameters.' }
            if ($CustomWidth -lt 10 -or $CustomWidth -gt 10000 -or $CustomHeight -lt 10 -or $CustomHeight -gt 10000) {
                throw 'Custom page width and height must each be 10 to 10000 finite points.'
            }
            $portraitW = $CustomWidth; $portraitH = $CustomHeight
        }
        default { throw "Unknown page-size profile '$Profile'." }
    }
    $physicalW = $portraitW; $physicalH = $portraitH
    if ($PaperOrientation -eq 'Landscape') { $physicalW = $portraitH; $physicalH = $portraitW }
    # These remain physical output-sheet coordinates.
    $outputLeft=72.0*$Left; $outputRight=72.0*$Right
    $outputTop=72.0*$Top; $outputBottom=72.0*$Bottom
    $printableW=$physicalW-$outputLeft-$outputRight
    $printableH=$physicalH-$outputTop-$outputBottom
    if ($printableW -le 0 -or $printableH -le 0) { throw 'Margins leave no usable printable area.' }
    if ($GuideWidth -eq 0) { $GuideWidth = if ($Mode -eq 'Paper') { $physicalW } else { $printableW } }
    if ($GuideHeight -eq 0) { $GuideHeight = if ($Mode -eq 'Paper') { $physicalH } else { $printableH } }
    if ($PageAdvance -eq 0) { $PageAdvance = $GuideHeight }
    foreach ($item in @(@('Guide width',$GuideWidth),@('Guide height',$GuideHeight),@('Page advance',$PageAdvance))) {
        if ($item[1] -lt 10 -or $item[1] -gt 10000) { throw "$($item[0]) must be 10 to 10000 points, or 0 for automatic." }
    }
    $overlap=$GuideHeight-$PageAdvance
    if ($overlap -lt 0) { throw 'Page advance cannot exceed guide height; negative overlap would leave an unrepresented gap.' }
    # Calibrate output-sheet coordinates into the independently sized OneNote guide.
    $scaleX=$GuideWidth/$physicalW; $scaleY=$GuideHeight/$physicalH
    [pscustomobject]@{
        Profile=$Profile; Orientation=$PaperOrientation; Mode=$Mode
        PhysicalSheetWidth=$physicalW; PhysicalSheetHeight=$physicalH
        PrintableWidth=$printableW; PrintableHeight=$printableH
        GuideWidth=$GuideWidth; GuideHeight=$GuideHeight; PageAdvance=$PageAdvance
        OriginX=$OriginX; OriginY=$OriginY; Overlap=$overlap; OverlapRounded=[Math]::Round($overlap,1)
        CalibrationScaleX=$scaleX; CalibrationScaleY=$scaleY
        CalibrationTranslateX=$OriginX; CalibrationTranslateY=$OriginY
        OutputMarginLeft=$outputLeft; OutputMarginRight=$outputRight
        OutputMarginTop=$outputTop; OutputMarginBottom=$outputBottom
        GuideInsetLeft=$outputLeft*$scaleX; GuideInsetRight=$outputRight*$scaleX
        GuideInsetTop=$outputTop*$scaleY; GuideInsetBottom=$outputBottom*$scaleY
    }
}

function New-GuidePng {
    param([Parameter(Mandatory)]$Geometry, [switch]$WithoutMargins)
    Add-Type -AssemblyName System.Drawing
    # Cap long custom frames to avoid allocating huge bitmaps.
    $scale = [Math]::Min(2.0, 4096.0/[Math]::Max($Geometry.GuideWidth,$Geometry.GuideHeight))
    $w = [int][Math]::Ceiling($Geometry.GuideWidth*$scale)
    $h = [int][Math]::Ceiling($Geometry.GuideHeight*$scale)
    $bitmap=$null; $graphics=$null; $brush=$null; $pen=$null; $stream=$null
    try {
        $bitmap = [Drawing.Bitmap]::new($w,$h,[Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $bitmap.SetResolution([single](72.0*$scale),[single](72.0*$scale))
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        $graphics.Clear([Drawing.Color]::Transparent)
        $brush = [Drawing.SolidBrush]::new([Drawing.Color]::FromArgb(190,110,110,110))
        $edge = [single][Math]::Max(1.0,0.75*$scale)
        # Explicit single arguments select one GDI+ overload. All FOUR bounds are drawn.
        $graphics.FillRectangle($brush,[single]0,[single]0,[single]$w,$edge)
        $graphics.FillRectangle($brush,[single]0,[single]($h-$edge),[single]$w,$edge)
        $graphics.FillRectangle($brush,[single]0,[single]0,$edge,[single]$h)
        $graphics.FillRectangle($brush,[single]($w-$edge),[single]0,$edge,[single]$h)
        if ($Geometry.Mode -eq 'Paper' -and -not $WithoutMargins) {
            $pen = [Drawing.Pen]::new([Drawing.Color]::FromArgb(150,140,140,140),[single][Math]::Max(1,0.75*$scale))
            $pen.DashStyle = [Drawing.Drawing2D.DashStyle]::Dash
            $insetWidth=[single](($Geometry.GuideWidth-$Geometry.GuideInsetLeft-$Geometry.GuideInsetRight)*$scale)
            $insetHeight=[single](($Geometry.GuideHeight-$Geometry.GuideInsetTop-$Geometry.GuideInsetBottom)*$scale)
            $graphics.DrawRectangle($pen,[single]($Geometry.GuideInsetLeft*$scale),[single]($Geometry.GuideInsetTop*$scale),
                $insetWidth,$insetHeight)
        }
        $stream = [IO.MemoryStream]::new()
        $bitmap.Save($stream,[Drawing.Imaging.ImageFormat]::Png)
        return [Convert]::ToBase64String($stream.ToArray())
    }
    finally {
        foreach ($resource in @($stream,$pen,$brush,$graphics,$bitmap)) {
            if ($null -ne $resource) { $resource.Dispose() }
        }
    }
}

function Assert-PngData {
    param([Parameter(Mandatory)][string]$Base64)
    $bytes = [Convert]::FromBase64String($Base64)
    $signature = @(137,80,78,71,13,10,26,10)
    if ($bytes.Length -lt 33) { throw 'Guide PNG data is empty or truncated.' }
    for ($i=0; $i -lt 8; $i++) {
        if ($bytes[$i] -ne $signature[$i]) { throw 'Guide data is not a PNG file.' }
    }
}

function New-GuidePayload {
    param([Parameter(Mandatory)][string]$Id, [Parameter(Mandatory)]$Geometry,
          [Parameter(Mandatory)][string]$Png, [Parameter(Mandatory)][string]$Batch,
          [int]$Count = $Pages, [double]$X = $Geometry.OriginX, [double]$Y = $Geometry.OriginY)
    Assert-PngData $Png
    $doc = New-PageDocument $Id
    for ($p=1; $p -le $Count; $p++) {
        $yp = $Y + ($p-1)*$Geometry.PageAdvance
        [void](Read-Point (Format-Point $X))
        [void](Read-Point (Format-Point $yp))
        if ($X+$Geometry.GuideWidth -gt 1000000 -or $yp+$Geometry.GuideHeight -gt 1000000) {
            throw 'The requested guide layout exceeds OneNote coordinate limits.'
        }
        $image = Add-XmlChild $doc.DocumentElement 'Image' @{
            format='png'; backgroundImage='true'; alt="OneNote Page Guides V2 - Letter page $p"
        }
        [void](Add-XmlChild $image 'Position' @{x=(Format-Point $X);y=(Format-Point $yp)})
        [void](Add-XmlChild $image 'Size' @{
            width=(Format-Point $Geometry.GuideWidth);height=(Format-Point $Geometry.GuideHeight);isSetByUser='true'
        })
        $marker = 'v2;version={0};batch={1};page={2};mode={3};advance={4};profile={5};physicalWidth={6};physicalHeight={7};guideWidth={8};guideHeight={9};overlap={10};originX={11};originY={12};' -f `
            $script:Version,$Batch,$p,$Geometry.Mode,(Format-Point $Geometry.PageAdvance),$Geometry.Profile,
            (Format-Point $Geometry.PhysicalSheetWidth),(Format-Point $Geometry.PhysicalSheetHeight),`
            (Format-Point $Geometry.GuideWidth),(Format-Point $Geometry.GuideHeight),(Format-Point $Geometry.Overlap),`
            (Format-Point $Geometry.OriginX),(Format-Point $Geometry.OriginY)
        [void](Add-XmlChild $image 'Meta' @{name=$script:MetaName;content=$marker})
        [void](Add-XmlChild $image 'Data' @{} $Png)
    }
    return ,$doc
}

function Assert-GuidePayload {
    param([Parameter(Mandatory)][xml]$Document, [Parameter(Mandatory)][string]$Id)
    $root = $Document.DocumentElement
    if ($root.LocalName -ne 'Page' -or $root.NamespaceURI -ne $script:OneNamespace -or $root.GetAttribute('ID') -cne $Id) {
        throw 'The payload does not target the pinned OneNote page.'
    }
    foreach ($attribute in $root.Attributes) {
        if ($attribute.Name -notin @('ID','xmlns:one','xmlns')) { throw 'Unexpected page attribute in guide payload.' }
    }
    $images = @(Get-GuideImages $Document)
    $children = @($root.ChildNodes | Where-Object { $_ -is [Xml.XmlElement] })
    if ($images.Count -ne $children.Count) { throw 'Only utility-tagged, top-level guide images are permitted in a payload.' }
    $mgr = New-NamespaceManager $Document
    foreach ($image in $images) {
        foreach ($attribute in $image.Attributes) {
            # OuterXml makes a standalone Image self-contained by declaring its
            # XML namespace. Read-SafeXml + ImportNode preserves that declaration
            # when the recovery image is put back inside a Page document.
            # It is NOT an object ID and must not trigger the overwrite safeguard.
            if ($attribute.NamespaceURI -ceq 'http://www.w3.org/2000/xmlns/') {
                if ($attribute.Name -cnotin @('xmlns:one','xmlns') -or
                    $attribute.Value -cne $script:OneNamespace) {
                    throw ("Unexpected XML namespace declaration '{0}' on a guide image." -f $attribute.Name)
                }
                continue
            }
            if ($attribute.NamespaceURI.Length -ne 0 -or
                $attribute.Name -cnotin @('format','backgroundImage','alt')) {
                throw ("Disallowed image attribute '{0}' in a new/restored guide. Object IDs and other image attributes are not permitted." -f $attribute.Name)
            }
        }
        $names = @($image.ChildNodes | Where-Object { $_ -is [Xml.XmlElement] } | ForEach-Object { $_.LocalName })
        if (($names -join ',') -notmatch '^Position,Size,Meta(,Meta)*,Data$') {
            throw 'Invalid guide child order or unexpected XML content.'
        }
        $descriptor = Get-ImageDescriptor $image
        if ($descriptor.Marker -match ';advance=([^;]+);.*;guideWidth=([^;]+);guideHeight=([^;]+);overlap=([^;]+);') {
            $advance=Read-Point $Matches[1]; $guideWidth=Read-Point $Matches[2]
            $guideHeight=Read-Point $Matches[3]; $overlap=Read-Point $Matches[4]
            if ([Math]::Abs($descriptor.Width-$guideWidth) -gt .01 -or [Math]::Abs($descriptor.Height-$guideHeight) -gt .01 -or
                [Math]::Abs(($guideHeight-$advance)-$overlap) -gt .011 -or $overlap -lt 0) {
                throw 'Guide metadata does not match its independently validated geometry.'
            }
        }
        if (-not $descriptor.Background -or $descriptor.Width -le 0 -or $descriptor.Height -le 0) {
            throw 'Only positive-size background guide images may be inserted.'
        }
        $data = $image.SelectSingleNode('one:Data',$mgr)
        Assert-PngData $data.InnerText
    }
}

function Get-EmbeddedGuidePng {
    param($OneNote, [Parameter(Mandatory)][string]$Id, [Parameter(Mandatory)][Xml.XmlElement]$Image)
    $mgr = New-NamespaceManager $Image.OwnerDocument
    $data = $Image.SelectSingleNode('one:Data',$mgr)
    [string]$png = ''
    if ($null -ne $data -and -not [string]::IsNullOrWhiteSpace($data.InnerText)) {
        $png = $data.InnerText
    }
    else {
        $callback = $Image.SelectSingleNode('one:CallbackID',$mgr)
        if ($null -eq $callback -or [string]::IsNullOrWhiteSpace($callback.GetAttribute('callbackID'))) {
            throw 'A guide has neither PNG bytes nor a callback ID. No destructive changes will be made.'
        }
        $OneNote.GetBinaryPageContent($Id,$callback.GetAttribute('callbackID'),[ref]$png)
    }
    Assert-PngData $png
    return $png
}

function New-ImageRestoreXml {
    param([Parameter(Mandatory)][Xml.XmlElement]$Source,
          [Parameter(Mandatory)][string]$Png, [Parameter(Mandatory)][string]$Id)
    $mgr = New-NamespaceManager $Source.OwnerDocument
    $d = Get-ImageDescriptor $Source
    if (-not $d.Background) { throw 'A marked guide is no longer a background image. Restore it to background first or remove it manually.' }
    $doc = New-PageDocument $Id
    $image = Add-XmlChild $doc.DocumentElement 'Image' @{
        format='png';backgroundImage='true';alt=$Source.GetAttribute('alt')
    }
    [void](Add-XmlChild $image 'Position' @{x=(Format-Point $d.X);y=(Format-Point $d.Y)})
    [void](Add-XmlChild $image 'Size' @{width=(Format-Point $d.Width);height=(Format-Point $d.Height);isSetByUser='true'})
    foreach ($meta in $Source.SelectNodes('one:Meta',$mgr)) {
        [void](Add-XmlChild $image 'Meta' @{name=$meta.GetAttribute('name');content=$meta.GetAttribute('content')})
    }
    [void](Add-XmlChild $image 'Data' @{} $Png)
    Assert-GuidePayload $doc $Id
    return $image.OuterXml
}

function Get-GuideSnapshots {
    param($OneNote, [Parameter(Mandatory)]$Page)
    foreach ($image in @(Get-GuideImages $Page.Xml)) {
        $d = Get-ImageDescriptor $image
        if ([string]::IsNullOrWhiteSpace($d.Id)) { throw 'A guide has no object ID. Cannot safely back it up or delete it.' }
        $png = Get-EmbeddedGuidePng $OneNote $Page.Id $image
        [pscustomobject]@{
            ObjectId=$d.Id
            ObservedXml=$image.OuterXml
            RestoreXml=(New-ImageRestoreXml $image $png $Page.Id)
            RecoveryToken=[guid]::NewGuid().ToString('N')
        }
    }
}

function Get-DataHash {
    param([Parameter(Mandatory)][string]$Base64)
    $hash = [Security.Cryptography.SHA256]::Create()
    try { return [Convert]::ToBase64String($hash.ComputeHash([Convert]::FromBase64String($Base64))) }
    finally { $hash.Dispose() }
}

function Lock-TargetPage {
    param([Parameter(Mandatory)][string]$Id)
    $hash = [Security.Cryptography.SHA256]::Create()
    try { $key = [BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($Id))).Replace('-','') }
    finally { $hash.Dispose() }
    $mutex = [Threading.Mutex]::new($false,('Local\OneNotePageGuides-' + $key))
    try {
        try { $locked = $mutex.WaitOne(0) }
        catch [Threading.AbandonedMutexException] {
            # WaitOne transfers ownership in this case; release rather than guessing
            # whether a crashed writer left a complete transaction.
            $mutex.ReleaseMutex()
            throw 'A prior utility operation ended unexpectedly. Inspect its recovery journal before continuing.'
        }
        if (-not $locked) { throw 'Another Page Guides operation is already using this page.' }
        return ,$mutex
    }
    catch { $mutex.Dispose(); throw }
}

function Assert-UnchangedGuideSet {
    param($OneNote,[string]$Id,[AllowEmptyCollection()][object[]]$Snapshots)
    $page = Read-OneNotePage $OneNote $Id
    $actual = @(Get-GuideImages $page.Xml)
    if ($actual.Count -ne $Snapshots.Count) { throw 'Guide count changed after preflight. No new write was started.' }
    foreach ($snapshot in $Snapshots) {
        $node = Get-ImageById $page.Xml $snapshot.ObjectId
        if ($null -eq $node) { throw 'The guide set changed after preflight. No new write was started.' }
        $original = Read-SafeXml $snapshot.ObservedXml
        Assert-ImageMatches $original.DocumentElement $node -CheckModified
    }
}

function Save-Journal {
    if ($null -eq $script:Journal -or [string]::IsNullOrWhiteSpace($script:JournalPath)) { throw 'Recovery journal not initialized.' }
    $script:Journal.UpdatedUtc = [datetime]::UtcNow.ToString('o')
    $json = $script:Journal | ConvertTo-Json -Depth 12
    $temp = $script:JournalPath + '.tmp'
    [IO.File]::WriteAllText($temp,$json,[Text.UTF8Encoding]::new($true))
    Move-Item -LiteralPath $temp -Destination $script:JournalPath -Force -Confirm:$false
}

function Start-Journal {
    param([string]$Operation, [Parameter(Mandatory)]$Target,
          [AllowEmptyCollection()][object[]]$Before, [string]$Batch = '', [string]$NewPayloadXml = '', $Geometry = $null)
    [void][IO.Directory]::CreateDirectory($script:BackupDirectory)
    $transaction = [guid]::NewGuid().ToString('N')
    $file = '{0}-{1}-{2}.json' -f [datetime]::Now.ToString('yyyyMMdd-HHmmss'),$Operation,$transaction
    $script:JournalPath = Join-Path $script:BackupDirectory $file
    $script:Journal = [pscustomobject]@{
        Format='OneNotePageGuides-Recovery-2.4'; TransactionId=$transaction
        PageId=$Target.PageId; PageName=$Target.PageName; Notebook=$Target.Notebook; Section=$Target.Section
        Action=$Operation; CreatedUtc=[datetime]::UtcNow.ToString('o'); UpdatedUtc=''
        State='Prepared'; NewBatch=$Batch; NewPayloadXml=$NewPayloadXml; Geometry=$Geometry; Before=@($Before); DeletedIds=@(); Error=''
    }
    Save-Journal
    Write-Host "Guide-only recovery journal: $($script:JournalPath)"
}

function Get-ImageById {
    param([xml]$Document,[string]$Id)
    foreach ($node in @(Get-GuideImages $Document)) {
        if ($node.GetAttribute('objectID') -ceq $Id) { return ,$node }
    }
    return $null
}

function Invoke-PageUpdate {
    param($OneNote,[string]$Id,[xml]$Payload)
    Assert-GuidePayload $Payload $Id
    $fresh = Read-OneNotePage $OneNote $Id
    $stamp = Get-PageTimestamp $fresh
    $OneNote.UpdatePageContent($Payload.OuterXml,$stamp,2,$false)
}

function Assert-BatchPresent {
    param($OneNote,[string]$Id,[string]$Batch,[xml]$Expected)
    $wanted = @(Get-GuideImages $Expected)
    $lastError = ''
    for ($attempt=0; $attempt -lt 4; $attempt++) {
        $page = Read-OneNotePage $OneNote $Id
        $actual = @(Get-GuideImages $page.Xml $Batch)
        try {
            if ($actual.Count -ne $wanted.Count) { throw "Expected $($wanted.Count) tagged guides; read back $($actual.Count)." }
            foreach ($image in $wanted) {
                $desc = Get-ImageDescriptor $image
                $match = @($actual | Where-Object { (Get-ImageDescriptor $_).Marker -ceq $desc.Marker })
                if ($match.Count -ne 1) { throw 'The inserted guide markers are missing or not unique.' }
                Assert-ImageMatches $image $match[0]
                if ([string]::IsNullOrWhiteSpace($match[0].GetAttribute('objectID'))) { throw 'An inserted guide has no object ID.' }
            }
            return
        }
        catch { $lastError = $_.Exception.Message }
        if ($attempt -lt 3) { Start-Sleep -Milliseconds 250 }
    }
    throw "Guide round-trip verification failed: $lastError Stopped; use the recovery journal if changes were already made."
}

function Remove-SnapshotObjects {
    param($OneNote,[string]$Id,[AllowEmptyCollection()][object[]]$Snapshots)
    foreach ($snapshot in $Snapshots) {
        # Re-read the PINNED page before EACH destructive call; protect against edits.
        $page = Read-OneNotePage $OneNote $Id
        $actual = Get-ImageById $page.Xml $snapshot.ObjectId
        if ($null -eq $actual) { throw "Expected guide $($snapshot.ObjectId) is no longer present. Stopped." }
        $observed = Read-SafeXml $snapshot.ObservedXml
        Assert-ImageMatches $observed.DocumentElement $actual -CheckModified
        if ($null -ne $snapshot.PSObject.Properties['RestoreXml']) {
            $saved = Read-SafeXml $snapshot.RestoreXml
            $savedMgr = New-NamespaceManager $saved
            $savedPng = $saved.DocumentElement.SelectSingleNode('one:Data',$savedMgr).InnerText
            $livePng = Get-EmbeddedGuidePng $OneNote $Id $actual
            if ((Get-DataHash $savedPng) -cne (Get-DataHash $livePng)) {
                throw 'Guide image bytes changed since backup. Stopped instead of deleting an edited image.'
            }
        }
        $stamp = Get-PageTimestamp $page
        $OneNote.DeletePageContent($Id,$snapshot.ObjectId,$stamp,$false)
        $script:Journal.DeletedIds = @($script:Journal.DeletedIds) + @($snapshot.ObjectId)
        Save-Journal
    }
    $after = Read-OneNotePage $OneNote $Id
    foreach ($snapshot in $Snapshots) {
        if ($null -ne (Get-ImageById $after.Xml $snapshot.ObjectId)) { throw 'A deleted guide is still present; removal did not verify.' }
    }
}

function Write-TargetSummary {
    param($Target)
    Write-Host ("Target: {0} / {1} / {2}" -f $Target.Notebook,$Target.Section,$Target.PageName)
    Write-Host "Page ID: $($Target.PageId)"
}

function Show-GuideStatus {
    param($Page,$Target)
    Write-TargetSummary $Target
    $geometry = Get-Geometry
    $advanceKind = if ($PageAdvancePoints -eq 0) { 'automatic' } else { 'explicit' }
    Write-Host ("Geometry: profile {0}; {1}; physical {2} x {3} pt; printable {4} x {5} pt." -f `
        $geometry.Profile,$geometry.Orientation,(Format-Point $geometry.PhysicalSheetWidth),
        (Format-Point $geometry.PhysicalSheetHeight),(Format-Point $geometry.PrintableWidth),(Format-Point $geometry.PrintableHeight))
    Write-Host ("Guide: {0} x {1} pt; advance {2} pt ({3}); overlap {4} pt (rounded {5:N2}); origin {6}, {7} pt." -f `
        (Format-Point $geometry.GuideWidth),(Format-Point $geometry.GuideHeight),(Format-Point $geometry.PageAdvance),
        $advanceKind,(Format-Point $geometry.Overlap),$geometry.OverlapRounded,(Format-Point $geometry.OriginX),(Format-Point $geometry.OriginY))
    Write-Host ("Calibration: scale {0}, {1}; translation {2}, {3} pt; guide insets L/R/T/B {4}/{5}/{6}/{7} pt." -f `
        (Format-Point $geometry.CalibrationScaleX),(Format-Point $geometry.CalibrationScaleY),
        (Format-Point $geometry.CalibrationTranslateX),(Format-Point $geometry.CalibrationTranslateY),
        (Format-Point $geometry.GuideInsetLeft),(Format-Point $geometry.GuideInsetRight),
        (Format-Point $geometry.GuideInsetTop),(Format-Point $geometry.GuideInsetBottom))
    Write-Host ("Margins (inches): left {0}, right {1}, top {2}, bottom {3}." -f `
        (Format-Point $MarginLeft),(Format-Point $MarginRight),(Format-Point $MarginTop),(Format-Point $MarginBottom))
    $images = @(Get-GuideImages $Page.Xml)
    Write-Host "V2 guide images: $($images.Count)"
    $items = foreach ($image in $images) {
        $d = Get-ImageDescriptor $image
        [pscustomobject]@{ X=$d.X; Y=$d.Y; Width=$d.Width; Height=$d.Height; Background=$d.Background }
    }
    if ($images.Count -gt 0) { $items | Format-Table -AutoSize | Out-Host }
    $mgr = New-NamespaceManager $Page.Xml
    if ($Page.Xml.SelectNodes('/one:Page/one:Outline',$mgr).Count -gt 0 -and $Page.Xml.OuterXml.Contains('ONPG:')) {
        Write-Warning 'Possible V1 text guides detected. V2.4 does not automatically delete outlines.'
    }
    Write-Host 'Status only: no notebook content was changed.'
}

function Read-RecoveryJournal {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw 'Recover requires -RecoveryFile pointing to a V2.4 recovery JSON file.'
    }
    $raw = [IO.File]::ReadAllText((Resolve-Path -LiteralPath $Path).Path)
    $journal = $raw | ConvertFrom-Json
    foreach ($property in @('Format','PageId','Before','TransactionId','NewBatch','NewPayloadXml','State','DeletedIds','Error','UpdatedUtc')) {
        if ($null -eq $journal.PSObject.Properties[$property]) { throw "Invalid recovery journal: missing $property." }
    }
    if ($journal.Format -cne 'OneNotePageGuides-Recovery-2.4') { throw 'Not a V2.4 recovery journal.' }
    if ([string]::IsNullOrWhiteSpace([string]$journal.PageId)) { throw 'Recovery journal contains no page ID.' }
    if ([string]$journal.TransactionId -notmatch '^[a-fA-F0-9]{32}$') { throw 'Invalid recovery transaction ID.' }
    if ([string]$journal.NewBatch -notmatch '^([a-fA-F0-9]{32})?$') { throw 'Invalid recovery batch ID.' }
    if (-not [string]::IsNullOrWhiteSpace([string]$journal.NewPayloadXml)) {
        $newDoc = Read-SafeXml $journal.NewPayloadXml
        Assert-GuidePayload $newDoc $journal.PageId
        $newImages = @(Get-GuideImages $newDoc $journal.NewBatch)
        if ($newImages.Count -ne @(Get-GuideImages $newDoc).Count) { throw 'Recovery batch does not match the saved payload.' }
    }
    elseif (-not [string]::IsNullOrWhiteSpace([string]$journal.NewBatch)) {
        throw 'A staged-batch journal must include its original guide payload.'
    }
    $seen = @{}
    foreach ($snapshot in @($journal.Before)) {
        foreach ($property in @('ObjectId','ObservedXml','RestoreXml','RecoveryToken')) {
            if ($null -eq $snapshot.PSObject.Properties[$property]) { throw "Invalid guide backup: missing $property." }
        }
        if ([string]$snapshot.RecoveryToken -notmatch '^[a-fA-F0-9]{32}$') { throw 'Invalid recovery token.' }
        if ([string]::IsNullOrWhiteSpace([string]$snapshot.ObjectId) -or $seen.ContainsKey([string]$snapshot.ObjectId)) {
            throw 'Recovery backup contains empty or duplicate object IDs.'
        }
        $seen[[string]$snapshot.ObjectId] = $true
        $doc = New-PageDocument $journal.PageId
        $imageDoc = Read-SafeXml $snapshot.RestoreXml
        [void]$doc.DocumentElement.AppendChild($doc.ImportNode($imageDoc.DocumentElement,$true))
        Assert-GuidePayload $doc $journal.PageId
        $observed = Read-SafeXml $snapshot.ObservedXml
        if ($observed.DocumentElement.GetAttribute('objectID') -cne $snapshot.ObjectId) { throw 'Recovery object IDs disagree.' }
        Assert-ImageMatches $imageDoc.DocumentElement $observed.DocumentElement
    }
    return $journal
}

function Find-RestoredImage {
    param([xml]$Document,[string]$Token)
    $mgr = New-NamespaceManager $Document
    foreach ($image in @(Get-GuideImages $Document)) {
        foreach ($meta in $image.SelectNodes("one:Meta[@name='$($script:RecoveryMetaName)']",$mgr)) {
            if ($meta.GetAttribute('content') -ceq $Token) { Write-Output -NoEnumerate $image }
        }
    }
}

function Invoke-GuideRecovery {
    param($OneNote,$Target)
    $id = [string]$script:Journal.PageId
    if ($id -cne $Target.PageId) { throw 'Recovery must use its original page ID.' }
    $script:Journal.State = 'Recovering'
    Save-Journal
    # Restore missing OLD guides first. Never overwrite existing notes or images.
    foreach ($snapshot in @($script:Journal.Before)) {
        $page = Read-OneNotePage $OneNote $id
        $expected = Read-SafeXml $snapshot.RestoreXml
        $original = Get-ImageById $page.Xml $snapshot.ObjectId
        if ($null -ne $original) {
            Assert-ImageMatches $expected.DocumentElement $original
            continue
        }
        $restored = @(Find-RestoredImage $page.Xml $snapshot.RecoveryToken)
        if ($restored.Count -gt 1) { throw 'Duplicate recovered guides found. Stopped for manual inspection.' }
        if ($restored.Count -eq 1) { Assert-ImageMatches $expected.DocumentElement $restored[0]; continue }
        $doc = New-PageDocument $id
        $image = $doc.ImportNode($expected.DocumentElement,$true)
        [void]$doc.DocumentElement.AppendChild($image)
        $mgr = New-NamespaceManager $doc
        # A fresh restore token replaces any token from an earlier recovery.
        foreach ($oldMarker in @($image.SelectNodes("one:Meta[@name='$($script:RecoveryMetaName)']",$mgr))) {
            [void]$image.RemoveChild($oldMarker)
        }
        $marker = $doc.CreateElement('one','Meta',$script:OneNamespace)
        $marker.SetAttribute('name',$script:RecoveryMetaName)
        $marker.SetAttribute('content',[string]$snapshot.RecoveryToken)
        [void]$image.InsertBefore($marker,$image.SelectSingleNode('one:Data',$mgr))
        Invoke-PageUpdate $OneNote $id $doc
        $page = Read-OneNotePage $OneNote $id
        $restored = @(Find-RestoredImage $page.Xml $snapshot.RecoveryToken)
        if ($restored.Count -ne 1) { throw 'Restored guide did not verify. New transaction guides have been left in place.' }
        Assert-ImageMatches $expected.DocumentElement $restored[0]
    }
    # Only after all old guides have verified, remove this transaction's staged batch.
    if (-not [string]::IsNullOrWhiteSpace([string]$script:Journal.NewBatch)) {
        $page = Read-OneNotePage $OneNote $id
        $staged = @()
        $newDoc = Read-SafeXml $script:Journal.NewPayloadXml
        $expectedNew = @(Get-GuideImages $newDoc $script:Journal.NewBatch)
        foreach ($image in @(Get-GuideImages $page.Xml $script:Journal.NewBatch)) {
            $descriptor = Get-ImageDescriptor $image
            $expectedMatch = @($expectedNew | Where-Object { (Get-ImageDescriptor $_).Marker -ceq $descriptor.Marker })
            if ($expectedMatch.Count -ne 1) { throw 'A staged image no longer matches the saved transaction.' }
            Assert-ImageMatches $expectedMatch[0] $image
            $staged += [pscustomobject]@{ObjectId=$image.GetAttribute('objectID');ObservedXml=$image.OuterXml}
        }
        Remove-SnapshotObjects $OneNote $id $staged
    }
    $script:Journal.State = 'Recovered'
    $script:Journal.Error = ''
    Save-Journal
    Write-Host 'Missing old guides were restored and identifiable transaction guides were removed. Ordinary notes were not included.'
    Write-Warning 'Inspect the page visually: recovery cannot identify new images if OneNote discarded their utility metadata.'
}

function New-UtilityShortcut {
    param([string]$Path,[string]$Arguments,[string]$Description)
    $shell=$null; $link=$null
    try {
        $shell = New-Object -ComObject WScript.Shell
        $link = $shell.CreateShortcut($Path)
        $link.TargetPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $link.Arguments = $Arguments
        $link.WorkingDirectory = $script:InstallDirectory
        $link.Description = $Description
        $link.Save()
    }
    finally { Release-ComObjectSafe $link; Release-ComObjectSafe $shell }
}

function Install-Utility {
    if ([string]::IsNullOrWhiteSpace($script:RunningFile) -or -not (Test-Path -LiteralPath $script:RunningFile -PathType Leaf)) {
        throw 'Run this from the saved OneNote-PageGuides.ps1 file.'
    }
    $tokens=$null; $errors=$null
    [void][Management.Automation.Language.Parser]::ParseFile($script:RunningFile,[ref]$tokens,[ref]$errors)
    if (@($errors).Count -gt 0) { throw 'The source script failed PowerShell parser validation.' }
    [void][IO.Directory]::CreateDirectory($script:InstallDirectory)
    $source = [IO.Path]::GetFullPath($script:RunningFile)
    $destination = [IO.Path]::GetFullPath($script:InstalledScript)
    if (-not [string]::Equals($source,$destination,[StringComparison]::OrdinalIgnoreCase)) {
        if (Test-Path -LiteralPath $destination -PathType Leaf) {
            Copy-Item -LiteralPath $destination -Destination ($destination + '.previous') -Force -Confirm:$false
        }
        Copy-Item -LiteralPath $source -Destination $destination -Force -Confirm:$false
    }
    try { Unblock-File -LiteralPath $destination -Confirm:$false -ErrorAction Stop } catch { Write-Verbose $_.Exception.Message }
    [void][IO.Directory]::CreateDirectory($script:StartMenuDirectory)
    $base = '-NoProfile -STA -NoExit -ExecutionPolicy Bypass -File "' + $destination + '"'
    $refresh = $base + (' -Action Refresh -Pages {0} -Orientation {1} -GuideMode {2}' -f $ShortcutPages,$Orientation,$GuideMode)
    foreach ($pair in @(@('StartX',$StartX),@('StartY',$StartY))) {
        $refresh += ' -' + $pair[0] + ' ' + (Format-Point ([double]$pair[1]))
    }
    if ($HideMargins) { $refresh += ' -HideMargins' }
    if ($ConsolePicker) { $refresh += ' -ConsolePicker' }
    $shortcuts = @(
        @('Refresh OneNote Page Guides.lnk',$refresh,'Choose a page and refresh visual guides using saved geometry'),
        @('Remove OneNote Page Guides.lnk',($base+' -Action Remove'),'Choose a page and remove guide images before printing'),
        @('Configure OneNote Page Guide Margins.lnk',($base+' -Action Configure'),'Adjust the geometry and margins used by Add and Refresh'),
        @('OneNote Page Guides Status.lnk',($base+' -Action Status'),'Read-only guide status'),
        @('OneNote Page Guides Self-Test.lnk',($base+' -Action SelfTest'),'Local self-tests without connecting to OneNote')
    )
    foreach ($entry in $shortcuts) {
        New-UtilityShortcut (Join-Path $script:StartMenuDirectory $entry[0]) $entry[1] $entry[2]
    }
    if ($CreateDesktopShortcuts) {
        $desktop = [Environment]::GetFolderPath('Desktop')
        if (-not (Test-Path -LiteralPath $desktop -PathType Container)) { throw 'Windows did not return a usable Desktop folder.' }
        foreach ($entry in $shortcuts[0..1]) {
            New-UtilityShortcut (Join-Path $desktop $entry[0]) $entry[1] $entry[2]
        }
    }
    Write-Host 'OneNote Page Guides 2.6.0 installed. No OneNote content was changed.'
    Write-Host "Installed file: $destination"
    Write-Host 'Shortcuts keep the console open so errors are visible. Close it after use.'
}

function Uninstall-Utility {
    $names = @('Refresh OneNote Page Guides.lnk','Remove OneNote Page Guides.lnk',
        'Configure OneNote Page Guide Margins.lnk','OneNote Page Guides Status.lnk','OneNote Page Guides Self-Test.lnk')
    $locations = @($script:StartMenuDirectory,[Environment]::GetFolderPath('Desktop'))
    foreach ($folder in $locations) {
        foreach ($name in $names) {
            $path = Join-Path $folder $name
            if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }
            $shell=$null; $link=$null
            try {
                $shell = New-Object -ComObject WScript.Shell
                $link = $shell.CreateShortcut($path)
                if ($link.Arguments.IndexOf($script:InstalledScript,[StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    Remove-Item -LiteralPath $path -Force -Confirm:$false
                }
            }
            finally { Release-ComObjectSafe $link; Release-ComObjectSafe $shell }
        }
    }
    if (Test-Path -LiteralPath $script:InstalledScript -PathType Leaf) { Remove-Item -LiteralPath $script:InstalledScript -Force -Confirm:$false }
    if (Test-Path -LiteralPath $script:StartMenuDirectory -PathType Container) {
        if (@(Get-ChildItem -LiteralPath $script:StartMenuDirectory -Force).Count -eq 0) {
            Remove-Item -LiteralPath $script:StartMenuDirectory -Force -Confirm:$false
        }
    }
    Write-Host 'Uninstalled the current script and its shortcuts. Recovery journals and previous copies were preserved.'
    Write-Host 'Guides already in notebooks were NOT removed. Remove those before uninstalling.'
}

function Invoke-SelfTest {
    # No COM or notebook reads. Tests execute the actual helpers; settings tests
    # use an isolated temporary directory and remove it in a finally block.
    $script:TestCount = 0
    function Assert-Test {
        param([bool]$Condition,[string]$Name)
        if (-not $Condition) { throw "SELF-TEST FAILED: $Name" }
        $script:TestCount++
        Write-Host "PASS: $Name"
    }
    $tokens=$null; $errors=$null
    [void][Management.Automation.Language.Parser]::ParseFile($script:RunningFile,[ref]$tokens,[ref]$errors)
    Assert-Test (@($errors).Count -eq 0) 'PowerShell parser (this installed runtime)'
    $savedState = @{}
    foreach ($name in @('InstallDirectory','SettingsPath','InvocationParameters','PageSizeProfile',
        'PageWidthPoints','PageHeightPoints','PagePitchPoints','GuideWidthPoints','GuideHeightPoints','PageAdvancePoints','OriginXPoints','OriginYPoints','PageWidthAvailable','PageHeightAvailable',
        'MarginLeft','MarginRight','MarginTop','MarginBottom','ResetSettings')) {
        $savedState[$name] = Get-Variable -Name $name -Scope Script -ValueOnly
    }
    $testDirectory = Join-Path ([IO.Path]::GetTempPath()) ('OneNotePageGuides-SelfTest-' + [guid]::NewGuid().ToString('N'))
    try {
        [void][IO.Directory]::CreateDirectory($testDirectory)
        $script:InstallDirectory = $testDirectory
        $script:SettingsPath = Join-Path $testDirectory 'settings.json'
        @{schemaVersion=1;MarginLeft=0.75;MarginRight=0.8;MarginTop=0.4;MarginBottom=0.45} |
            ConvertTo-Json | Set-Content -LiteralPath $script:SettingsPath -Encoding UTF8
        $script:InvocationParameters = @()
        $script:PageSizeProfile = 'OneNotePdfLetter'; $script:MarginLeft = 1
        Import-UserSettings
        Assert-Test ($PageSizeProfile -ceq 'OneNotePdfLetter' -and $MarginLeft -eq 0.75) 'Version-1 margin settings use the new calibrated default'
        $before = [IO.File]::ReadAllText($script:SettingsPath)
        $script:InvocationParameters = @('MarginLeft'); $script:MarginLeft = 0.25
        Import-UserSettings
        Assert-Test ($MarginLeft -eq 0.25 -and [IO.File]::ReadAllText($script:SettingsPath) -ceq $before) 'CLI setting overrides saved value without rewriting settings'
        @{schemaVersion=2;PagePitchPoints=689.33} | ConvertTo-Json |
            Set-Content -LiteralPath $script:SettingsPath -Encoding UTF8
        $script:InvocationParameters=@(); $script:PageAdvancePoints=0; $script:GuideHeightPoints=725.72
        Import-UserSettings
        Assert-Test ($PageAdvancePoints -eq 689.33 -and $GuideHeightPoints -eq 725.72) 'Legacy PagePitchPoints migrates only to page advance'
        $script:ResetSettings = $true
        Configure-Settings
        Assert-Test (-not (Test-Path -LiteralPath $script:SettingsPath)) 'ResetSettings removes all saved configuration'
    }
    finally {
        if (Test-Path -LiteralPath $testDirectory) { Remove-Item -LiteralPath $testDirectory -Recurse -Force -Confirm:$false }
        foreach ($name in $savedState.Keys) { Set-Variable -Name $name -Value $savedState[$name] -Scope Script }
    }
    $portrait = Get-Geometry -Profile Letter -PaperOrientation Portrait -Mode Paper -Left 1 -Right 1 -Top .5 -Bottom .5
    Assert-Test ($portrait.PhysicalSheetWidth -eq 612 -and $portrait.PhysicalSheetHeight -eq 792) 'Physical Letter sheet dimensions'
    Assert-Test ($portrait.GuideWidth -eq 612 -and $portrait.GuideHeight -eq 792 -and $portrait.PageAdvance -eq 792) 'Automatic paper guide dimensions and advance'
    Assert-Test ($portrait.OutputMarginLeft -eq 72 -and $portrait.OutputMarginTop -eq 36) 'Margins remain in output-sheet coordinates'
    Assert-Test ($portrait.GuideInsetLeft -eq 72 -and $portrait.GuideInsetTop -eq 36) 'Unit calibration converts margins to guide insets'
    $landscape = Get-Geometry -Profile Letter -PaperOrientation Landscape -Mode Paper -Left 1 -Right 1 -Top .5 -Bottom .5
    Assert-Test ($landscape.PhysicalSheetWidth -eq 792 -and $landscape.PhysicalSheetHeight -eq 612) 'Standard Letter landscape physical sheet'
    $calibrated = Get-Geometry -Profile OneNotePdfLetter -PaperOrientation Portrait -Mode Paper -Left 1 -Right 1 -Top .5 -Bottom .5
    Assert-Test ($calibrated.PrintableWidth -eq 467.4 -and $calibrated.PrintableHeight -eq 720.84) 'Calibrated default-margin printable area'
    $calibratedLandscape = Get-Geometry -Profile OneNotePdfLetter -PaperOrientation Landscape -Mode Paper -Left 1 -Right 1 -Top .5 -Bottom .5
    Assert-Test ($calibratedLandscape.GuideWidth -eq 792.84 -and $calibratedLandscape.GuideHeight -eq 611.4) 'OneNote PDF Letter landscape guide'
    $custom = Get-Geometry -Profile Custom -CustomWidth 500 -CustomHeight 700 -CustomDimensionsAvailable -GuideWidth 480 -GuideHeight 680 -PageAdvance 650 -OriginX 12 -OriginY 34
    Assert-Test ($custom.PhysicalSheetWidth -eq 500 -and $custom.GuideHeight -eq 680 -and $custom.PageAdvance -eq 650 -and $custom.OriginY -eq 34) 'Independent custom geometry properties'
    $failed=$false
    try { [void](Get-Geometry -Profile Custom -CustomWidth 0 -CustomHeight 700 -CustomDimensionsAvailable) } catch { $failed=$true }
    Assert-Test $failed 'Invalid custom dimensions rejected'
    $content = Get-Geometry -Profile Letter -Mode PrintArea -GuideHeight 725.72 -PageAdvance 689.33
    Assert-Test ($content.GuideHeight -eq 725.72 -and $content.PageAdvance -eq 689.33) 'Guide height and page advance coexist independently'
    Assert-Test ([Math]::Abs($content.Overlap-36.39) -lt .0001 -and (Format-Point $content.Overlap) -ceq '36.39') 'Overlap computes as 36.39 points'
    Assert-Test ($content.OverlapRounded -eq 36.4) 'Overlap has the requested 36.40 display rounding'
    $scaled = Get-Geometry -Profile Letter -Mode Paper -GuideWidth 306 -GuideHeight 396 -OriginX 10 -OriginY 20
    Assert-Test ($scaled.CalibrationScaleX -eq .5 -and $scaled.GuideInsetLeft -eq 36 -and $scaled.GuideInsetTop -eq 18 -and $scaled.CalibrationTranslateX -eq 10) 'Margins use explicit calibration scale and translation'
    $failed=$false
    try { [void](Get-Geometry -GuideHeight 700 -PageAdvance 701) } catch { $failed=$true }
    Assert-Test $failed 'Negative overlap rejected'
    $failed=$false
    try { [void](Get-Geometry -PaperOrientation Landscape -Top 4.5 -Bottom 4.5) } catch { $failed=$true }
    Assert-Test $failed 'Reject unusable margins'
    $failed=$false
    try { [void](Read-SafeXml '<!DOCTYPE root [<!ENTITY x SYSTEM "file:///does-not-exist">]><root>&x;</root>') } catch { $failed=$true }
    Assert-Test $failed 'DTD / external entity rejection'
    $hierarchy = Read-SafeXml @'
<one:Notebooks xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote"><one:Notebook name="Test" ID="n" isCurrentlyViewed="true"><one:Section name="Live" ID="s"><one:Page ID="p1" name="Keep" /></one:Section><one:SectionGroup isRecycleBin="true"><one:Section><one:Page ID="p2" name="Deleted" /></one:Section></one:SectionGroup><one:Section isLocked="true"><one:Page ID="p3" name="Locked" /></one:Section></one:Notebook></one:Notebooks>
'@
    $records = @(Convert-HierarchyToPages $hierarchy)
    Assert-Test ($records.Count -eq 1 -and $records[0].PageId -eq 'p1') 'Recycle-bin and locked-page exclusion'
    Assert-Test $records[0].CurrentlyViewed 'Currently viewed notebook is a sort hint only'
    $selected = Select-TargetPage -AvailablePages $records -ExplicitId 'p1'
    Assert-Test ($selected.PageId -eq 'p1') 'Explicit page selection pins the correct ID'
    $failed=$false
    try { [void](Select-TargetPage -AvailablePages $records -ExplicitId 'missing') } catch { $failed=$true }
    Assert-Test $failed 'Unknown/deleted page ID rejected'
    $oldCulture = [Threading.Thread]::CurrentThread.CurrentCulture
    try {
        [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('de-DE')
        Assert-Test ((Format-Point 12.5) -ceq '12.5') 'Invariant decimal formatting'
    }
    finally { [Threading.Thread]::CurrentThread.CurrentCulture = $oldCulture }
    $png = New-GuidePng $portrait
    Assert-PngData $png
    Assert-Test ($png.Length -gt 100) 'PNG generated and signature verified'
    $bytes = [Convert]::FromBase64String($png)
    $stream = [IO.MemoryStream]::new($bytes,$false)
    $bitmap = $null
    try {
        $bitmap = [Drawing.Bitmap]::new($stream)
        $cx = [int]($bitmap.Width/2); $cy=[int]($bitmap.Height/2)
        Assert-Test ($bitmap.GetPixel(0,$cy).A -gt 0 -and $bitmap.GetPixel($bitmap.Width-1,$cy).A -gt 0) 'Visible left/right PNG edges'
        Assert-Test ($bitmap.GetPixel($cx,0).A -gt 0 -and $bitmap.GetPixel($cx,$bitmap.Height-1).A -gt 0) 'Visible top/bottom PNG edges'
        Assert-Test ($bitmap.GetPixel($cx,$cy).A -eq 0) 'Transparent guide interior'
    }
    finally { if ($null -ne $bitmap) { $bitmap.Dispose() }; $stream.Dispose() }
    $batch = [guid]::NewGuid().ToString('N')
    $calibratedPayload = New-GuidePayload -Id 'CALIBRATED' -Geometry $calibrated -Png $png -Batch $batch -Count 6 -X 0 -Y 0
    $calibratedImages = @(Get-GuideImages $calibratedPayload $batch)
    $expectedY = @(0,792.84,1585.68,2378.52,3171.36,3964.20)
    $positionsMatch = $calibratedImages.Count -eq $expectedY.Count
    for ($i=0; $positionsMatch -and $i -lt $expectedY.Count; $i++) {
        $positionsMatch = [Math]::Abs((Get-ImageDescriptor $calibratedImages[$i]).Y - $expectedY[$i]) -lt 0.001
    }
    Assert-Test $positionsMatch 'Six calibrated pages use 792.84-point increments'
    $payload = New-GuidePayload -Id 'TEST&ID' -Geometry $portrait -Png $png -Batch $batch -Count 2 -X 5 -Y 20
    Assert-GuidePayload $payload 'TEST&ID'
    $roundtrip = Read-SafeXml $payload.OuterXml
    $images = @(Get-GuideImages $roundtrip $batch)
    Assert-Test ($images.Count -eq 2) 'Payload XML round-trip and batch metadata'
    Assert-Test ((Get-ImageDescriptor $images[1]).Y -eq 812) 'Second standard-Letter guide uses origin plus page advance'
    $marker = (Get-ImageDescriptor $images[0]).Marker
    Assert-Test ($marker.Contains(';profile=Letter;') -and $marker.Contains(';guideWidth=612;') -and
        $marker.Contains(';guideHeight=792;') -and $marker.Contains(';overlap=0;')) 'New metadata includes profile and effective dimensions'
    $ordinary = Add-XmlChild $roundtrip.DocumentElement 'Image' @{alt='OneNotePageGuidesV2 v2;not-a-guide'}
    Assert-Test (@(Get-GuideImages $roundtrip).Count -eq 2) 'Ordinary images ignored by removal selector'
    $nested = Add-XmlChild $roundtrip.DocumentElement 'Outline'
    [void]$nested.AppendChild($roundtrip.ImportNode($images[0],$true))
    Assert-Test (@(Get-GuideImages $roundtrip).Count -eq 2) 'Nested images excluded from deletion'
    $failed=$false
    try { Assert-GuidePayload $roundtrip 'TEST&ID' } catch { $failed=$true }
    Assert-Test $failed 'Payload refuses non-guide page objects'
    # Use a source that actually carries existing-object attributes; otherwise
    # the no-objectID test would pass without demonstrating that stripping works.
    $restoreSource = Read-SafeXml $images[0].OuterXml
    $restoreSource.DocumentElement.SetAttribute('objectID','EXISTING-GUIDE-DO-NOT-REUSE')
    $restoreSource.DocumentElement.SetAttribute('lastModifiedTime','2026-01-01T00:00:00.000Z')
    $restoreText = New-ImageRestoreXml $restoreSource.DocumentElement $png 'TEST&ID'
    $restoreImage = Read-SafeXml $restoreText
    Assert-Test ($restoreImage.DocumentElement.NamespaceURI -ceq $script:OneNamespace) 'Standalone recovery XML namespace'
    Assert-Test (-not $restoreImage.DocumentElement.HasAttribute('objectID') -and
        -not $restoreImage.DocumentElement.HasAttribute('lastModifiedTime')) 'Restore does not overwrite existing object IDs'
    $restoredPage = New-PageDocument 'TEST&ID'
    [void]$restoredPage.DocumentElement.AppendChild($restoredPage.ImportNode($restoreImage.DocumentElement,$true))
    Assert-GuidePayload $restoredPage 'TEST&ID'
    Assert-Test (@(Get-GuideImages $restoredPage).Count -eq 1) 'Recovery image payload passes structural checks'

    # V2.4.1 regressions: allow ONLY the known namespace declarations and keep
    # blocking existing object IDs, timestamps and arbitrary image attributes.
    Assert-Test ($restoreSource.DocumentElement.GetAttribute('objectID') -ceq 'EXISTING-GUIDE-DO-NOT-REUSE') 'Restore builder leaves the source object ID unchanged'
    $recoveryRoot = $restoredPage.DocumentElement.FirstChild
    Assert-Test ($recoveryRoot.HasAttribute('xmlns:one') -and
        $recoveryRoot.GetAttribute('xmlns:one') -ceq $script:OneNamespace) 'Imported recovery image retains its valid namespace declaration'
    $recoveryRoundtrip = Read-SafeXml $restoredPage.OuterXml
    Assert-GuidePayload $recoveryRoundtrip 'TEST&ID'
    Assert-Test (@(Get-GuideImages $recoveryRoundtrip).Count -eq 1) 'Recovery full-page XML round-trip passes validation'

    foreach ($attributeName in @('objectID','lastModifiedTime','unexpected')) {
        $invalid = Read-SafeXml $restoredPage.OuterXml
        $invalid.DocumentElement.FirstChild.SetAttribute($attributeName,'MUST-BE-REJECTED')
        $failureMessage = ''
        try { Assert-GuidePayload $invalid 'TEST&ID' }
        catch { $failureMessage = $_.Exception.Message }
        Assert-Test ($failureMessage.Contains("Disallowed image attribute '$attributeName'")) ("Recovery payload still rejects '$attributeName'")
    }

    $invalidNamespace = Read-SafeXml $restoredPage.OuterXml
    $declaration = $invalidNamespace.CreateAttribute('xmlns','unexpected','http://www.w3.org/2000/xmlns/')
    $declaration.Value = $script:OneNamespace
    [void]$invalidNamespace.DocumentElement.FirstChild.Attributes.Append($declaration)
    $failureMessage = ''
    try { Assert-GuidePayload $invalidNamespace 'TEST&ID' }
    catch { $failureMessage = $_.Exception.Message }
    Assert-Test ($failureMessage.Contains("Unexpected XML namespace declaration 'xmlns:unexpected'")) 'Recovery payload rejects an unapproved namespace prefix'

    $invalidNamespace = Read-SafeXml $restoredPage.OuterXml
    $declaration = $invalidNamespace.CreateAttribute('','xmlns','http://www.w3.org/2000/xmlns/')
    $declaration.Value = 'urn:not-the-onenote-schema'
    [void]$invalidNamespace.DocumentElement.FirstChild.Attributes.Append($declaration)
    $failureMessage = ''
    try { Assert-GuidePayload $invalidNamespace 'TEST&ID' }
    catch { $failureMessage = $_.Exception.Message }
    Assert-Test ($failureMessage.Contains("Unexpected XML namespace declaration 'xmlns'")) 'Recovery payload rejects a wrong namespace URI'

    $defaultNamespace = Read-SafeXml $restoredPage.OuterXml
    $declaration = $defaultNamespace.CreateAttribute('','xmlns','http://www.w3.org/2000/xmlns/')
    $declaration.Value = $script:OneNamespace
    [void]$defaultNamespace.DocumentElement.FirstChild.Attributes.Append($declaration)
    Assert-GuidePayload $defaultNamespace 'TEST&ID'
    Assert-Test ($defaultNamespace.DocumentElement.FirstChild.GetAttribute('xmlns') -ceq $script:OneNamespace) 'Recovery payload accepts the exact OneNote default namespace'
    $legacy = $restoredPage.DocumentElement.FirstChild
    $legacyMgr = New-NamespaceManager $restoredPage
    $legacy.SelectSingleNode('one:Meta',$legacyMgr).SetAttribute('content','v2;page=1;orientation=Portrait;')
    Assert-Test (@(Get-GuideImages $restoredPage).Count -eq 1) 'V2.0 through V2.3 guide metadata remains supported'
    Assert-Test ((Get-DataHash $png) -ceq (Get-DataHash ($png + "`r`n"))) 'PNG integrity hash ignores base64 whitespace'

    # A local fake implements only the two read/update methods below. No OneNote
    # connection is created. This exercises the real update and verify helpers.
    $empty = New-PageDocument 'TEST&ID'
    $empty.DocumentElement.SetAttribute('lastModifiedTime','2026-01-01T00:00:00.000Z')
    $fake = [pscustomobject]@{PageText=$empty.OuterXml;RequestedIds=@();UpdatedIds=@()}
    $fake | Add-Member -MemberType ScriptMethod -Name GetPageContent -Value {
        param($id,$outputText,$pageInfo,$schema)
        $this.RequestedIds = @($this.RequestedIds) + @([string]$id)
        $outputText.Value = [string]$this.PageText
    }
    $fake | Add-Member -MemberType ScriptMethod -Name UpdatePageContent -Value {
        param($xmlText,$stamp,$schema,$force)
        if ($stamp -le [datetime]'1900-01-01') { throw 'Unguarded timestamp in mock update.' }
        if ($schema -ne 2 -or $force) { throw 'Wrong schema or forced update in mock.' }
        $doc = Read-SafeXml $xmlText
        $this.UpdatedIds = @($this.UpdatedIds) + @($doc.DocumentElement.GetAttribute('ID'))
        $doc.DocumentElement.SetAttribute('lastModifiedTime','2026-01-01T00:00:01.000Z')
        $index = 0
        foreach ($node in @(Get-GuideImages $doc)) {
            $index++
            $node.SetAttribute('objectID',('FAKE-GUIDE-' + $index))
        }
        $this.PageText = $doc.OuterXml
    }
    Invoke-PageUpdate $fake 'TEST&ID' $payload
    Assert-BatchPresent $fake 'TEST&ID' $batch $payload
    Assert-Test ($fake.UpdatedIds.Count -eq 1 -and $fake.UpdatedIds[0] -ceq 'TEST&ID') 'Mock update stays on pinned page ID'
    Assert-Test (@($fake.RequestedIds | Where-Object { $_ -cne 'TEST&ID' }).Count -eq 0) 'All verification reads stay on pinned page'
    $stripped = Read-SafeXml $fake.PageText
    $strippedMgr = New-NamespaceManager $stripped
    foreach ($node in @($stripped.SelectNodes('/one:Page/one:Image/one:Meta',$strippedMgr))) {
        [void]$node.ParentNode.RemoveChild($node)
    }
    $fake.PageText = $stripped.OuterXml
    $failed=$false
    try { Assert-BatchPresent $fake 'TEST&ID' $batch $payload } catch { $failed=$true }
    Assert-Test $failed 'Round-trip verification rejects missing metadata'

    # Exercise the same write helper with a restored, re-imported Image. The
    # fake still performs no COM calls; production validation remains in place.
    $restoredForWrite = Read-SafeXml $recoveryRoundtrip.OuterXml
    Invoke-PageUpdate $fake 'TEST&ID' $restoredForWrite
    Assert-BatchPresent $fake 'TEST&ID' $batch $restoredForWrite
    Assert-Test ($fake.UpdatedIds.Count -eq 2 -and $fake.UpdatedIds[1] -ceq 'TEST&ID') 'Mock recovery update accepts namespaces and stays on the pinned page'
    Write-Host ''
    Write-Host "$($script:TestCount) local self-tests passed. No OneNote content was read or changed."
    Write-Host 'This does NOT test OneNote COM writes, metadata retention, or print/export alignment.'
}

# Entry point. All notebook mutations are behind one high-impact ShouldProcess gate.
$app = $null
$pageMutex = $null
try {
    Assert-Windows
    Initialize-Paths
    if ($Action -in @('Status','Add','Refresh','Install','Configure') -and -not ($Action -eq 'Configure' -and $ResetSettings)) {
        Import-UserSettings
    }
    Write-Host "OneNote Page Guides $($script:Version) - $Action"
    if ($Action -eq 'SelfTest') { Invoke-SelfTest; return }
    if ($Action -eq 'Configure') {
        $operation = if ($ResetSettings) { 'Reset saved geometry and margins' } else { 'Save page geometry and margins' }
        if ($PSCmdlet.ShouldProcess($script:SettingsPath,$operation)) { Configure-Settings }
        return
    }
    if ($Action -eq 'Install') {
        [void](Get-Geometry) # validate shortcut geometry before changing installed files
        if ($PSCmdlet.ShouldProcess($script:InstallDirectory,'Install V2.6.0 and replace utility shortcuts')) {
            Save-UserSettings
            Install-Utility
        }
        return
    }
    if ($Action -eq 'Uninstall') {
        if ($PSCmdlet.ShouldProcess($script:InstallDirectory,'Remove utility and shortcuts; preserve recovery data')) { Uninstall-Utility }
        return
    }
    $app = Get-OneNoteApplication
    $available = @(Get-HierarchyPages $app)
    if ($available.Count -eq 0) { throw 'No live, unlocked pages were returned by desktop OneNote.' }
    if ($Action -eq 'Recover') {
        $journal = Read-RecoveryJournal $RecoveryFile
        if ($PageId.Length -gt 0 -and $PageId -cne $journal.PageId) { throw '-PageId conflicts with the recovery journal.' }
        $target = Select-TargetPage -AvailablePages $available -ExplicitId $journal.PageId
        Write-TargetSummary $target
        if (-not $PSCmdlet.ShouldProcess($target.PageName,'Recover this journal: restore missing old guides, then remove its new batch')) { return }
        $pageMutex = Lock-TargetPage $target.PageId
        $script:Journal = $journal
        $script:JournalPath = (Resolve-Path -LiteralPath $RecoveryFile).Path
        Invoke-GuideRecovery $app $target
        return
    }
    # Exactly ONE selection. Never call the picker again during this operation.
    $target = Select-TargetPage -AvailablePages $available -ExplicitId $PageId -Filter $PageFilter -UseConsole:$ConsolePicker
    if ($null -eq $target) { Write-Host 'Cancelled. No changes made.'; return }
    $pinnedId = [string]$target.PageId
    $page = Read-OneNotePage $app $pinnedId
    if ($Action -eq 'Status') { Show-GuideStatus $page $target; return }
    Write-TargetSummary $target
    $existing = @(Get-GuideImages $page.Xml)
    if ($Action -eq 'Add' -and $existing.Count -gt 0) { throw 'This page already has V2 guides. Use Refresh instead of Add.' }
    if ($Action -eq 'Remove' -and $existing.Count -eq 0) { Write-Host 'No V2 guide images found. Nothing changed.'; return }
    $batch = ''
    $payload = $null
    $geometry = $null
    if ($Action -in @('Add','Refresh')) {
        $geometry = Get-Geometry
        $png = New-GuidePng $geometry -WithoutMargins:$HideMargins
        $batch = [guid]::NewGuid().ToString('N')
        $payload = New-GuidePayload -Id $pinnedId -Geometry $geometry -Png $png -Batch $batch
        Assert-GuidePayload $payload $pinnedId
        Write-Host ('Plan: {0} {1} guides; profile {2}; size {3} x {4} pt; advance {5} pt; overlap {6} pt; origin {7}, {8} pt.' -f `
            $Pages,$GuideMode,$geometry.Profile,$geometry.GuideWidth,$geometry.GuideHeight,$geometry.PageAdvance,$geometry.Overlap,$geometry.OriginX,$geometry.OriginY)
        Write-Warning 'These are uncalibrated visual guides, not native print boundaries. Remove before export/printing.'
    }
    [void](Get-PageTimestamp $page)
    # Preflight includes binary recovery data. WhatIf can read, but never saves or writes.
    $snapshots = @(Get-GuideSnapshots $app $page)
    $description = '{0}: {1} existing guide(s); {2} requested frame(s)' -f $Action,$snapshots.Count,$Pages
    if (-not $PSCmdlet.ShouldProcess(($target.Notebook+' / '+$target.Section+' / '+$target.PageName),$description)) { return }
    $pageMutex = Lock-TargetPage $pinnedId
    Assert-UnchangedGuideSet $app $pinnedId $snapshots
    $newXml = ''
    if ($null -ne $payload) { $newXml = $payload.OuterXml }
    Start-Journal -Operation $Action -Target $target -Before $snapshots -Batch $batch -NewPayloadXml $newXml -Geometry $geometry
    if ($Action -in @('Add','Refresh')) {
        $script:Journal.State = 'StagingNewGuides'; Save-Journal
        Invoke-PageUpdate $app $pinnedId $payload
        Assert-BatchPresent $app $pinnedId $batch $payload
        $script:Journal.State = 'NewGuidesVerified'; Save-Journal
    }
    if ($Action -in @('Refresh','Remove')) {
        $script:Journal.State = 'RemovingOldGuides'; Save-Journal
        Remove-SnapshotObjects $app $pinnedId $snapshots
    }
    if ($Action -in @('Add','Refresh')) { Assert-BatchPresent $app $pinnedId $batch $payload }
    $script:Journal.State = 'Completed'; Save-Journal
    Write-Host "$Action completed and guide objects verified on the pinned page."
    Write-Host "Recovery file retained: $($script:JournalPath)"
}
catch {
    $failure = $_
    if ($null -ne $script:Journal -and $script:JournalPath.Length -gt 0) {
        try { $script:Journal.State='Failed'; $script:Journal.Error=$failure.Exception.Message; Save-Journal } catch { }
        Write-Warning 'The operation was not confirmed complete. Do not repeatedly Refresh a partially changed page.'
        Write-Warning ('To attempt guide-only recovery: & "{0}" -Action Recover -RecoveryFile "{1}"' -f $script:InstalledScript,$script:JournalPath)
        Write-Warning 'Recovery only handles correctly tagged utility images; inspect the page if OneNote dropped metadata.'
    }
    if (-not [string]::IsNullOrWhiteSpace($failure.ScriptStackTrace)) { Write-Verbose $failure.ScriptStackTrace }
    throw $failure
}
finally {
    if ($null -ne $pageMutex) {
        try { $pageMutex.ReleaseMutex() } catch { }
        $pageMutex.Dispose()
    }
    Release-ComObjectSafe $app
}
