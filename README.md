# OneNote Page Guides

A standalone Windows PowerShell utility that adds removable page and margin
guides to desktop OneNote. The guides are transparent PNG background images;
they are **visual aids**, not native print boundaries.

> [!CAUTION]
> Remove guides before printing or exporting. OneNote can print the images and
> include them when determining page extent. Start with a disposable page and
> keep normal OneNote backups.

## Features

- Adjustable left, right, top, and bottom margins (0–4 inches each)
- Persistent page-size profile, custom dimensions, page pitch, and margins
- Calibrated OneNote PDF Letter, exact standard Letter, and custom layouts
- Physical-paper and calibrated print-area modes
- Add, refresh, inspect, and remove operations
- Page picker that pins one explicit page for the whole operation
- Transaction journals and guide-only recovery
- Built-in local tests that do not connect to OneNote
- No add-in, service, administrator install, or network access

## Requirements

- Windows PowerShell 5.1 (not PowerShell 7)
- Desktop OneNote exposing the `OneNote.Application` COM API
- `System.Drawing`, included with supported Windows PowerShell installations

> [!NOTE]
> Desktop OneNote normally stores the user's custom page-templates notebook at
> `%APPDATA%\Microsoft\Templates\My Templates.one` (typically
> `C:\Users\<username>\AppData\Roaming\Microsoft\Templates\My Templates.one`).
> This is separate from OneNote Page Guides' `settings.json`; the utility does
> not modify the templates file.

## Download and verify

Download [`OneNote-PageGuides.ps1`](OneNote-PageGuides.ps1) from the latest
release. Open a normal, non-administrator Windows PowerShell 5.1 window in the
download directory and run:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
Unblock-File .\OneNote-PageGuides.ps1
.\OneNote-PageGuides.ps1 -Action SelfTest
```

Do not continue if the self-test reports a failure.

## Install

```powershell
.\OneNote-PageGuides.ps1 -Action Install -ShortcutPages 10 -CreateDesktopShortcuts
```

The ordinary **Refresh** shortcut stores only operational options (`-Action
Refresh -Pages ...`) and reads geometry from `settings.json` each time it runs.
Use `-CreateOverrideShortcut` during installation only when you intentionally
want an additional, clearly named shortcut with frozen calibrated geometry.

Settings schema 4 stores guide mode, first-page origin, guide size and advance,
calibration scale/translation, calibration verification status, and whether
margins are interpreted in output-sheet or guide coordinates. Schema 1 and 2
files retain their existing page sizes and margins during migration and are
marked `Unverified` until their origin/calibration has been measured.

Installation copies the utility to
`%LOCALAPPDATA%\OneNotePageGuides\OneNote-PageGuides-V2.ps1`. It creates Start
Menu shortcuts for Refresh, Remove, Configure, Status, and Self-Test;
the optional desktop shortcuts are Refresh and Remove. Existing installations
at that compatible path are backed up with a `.previous` suffix.

## Configure geometry and margins

### Interactive, persistent settings

Run the **Configure OneNote Page Guide Margins** Start Menu shortcut (the
compatible shortcut name is retained), or:

```powershell
$guides = "$env:LOCALAPPDATA\OneNotePageGuides\OneNote-PageGuides-V2.ps1"
& $guides -Action Configure
```

The interactive prompt shows the current page profile and permits
`OneNotePdfLetter`, `Letter`, or `Custom`. Custom geometry is entered in points.
It also prompts for page pitch; `0` means automatic. Finally, enter margins in
inches, or press Enter to retain each current value. The settings are saved in
`%LOCALAPPDATA%\OneNotePageGuides\settings.json`; subsequent Add and Refresh
operations—including the Refresh shortcut—use them automatically.

The settings schema is version 2. Existing version-1 files containing only
margins are accepted without deletion: their margins are retained and missing
geometry values use the new `OneNotePdfLetter` and automatic-pitch defaults.

For scripting, supplying any configurable argument skips interactive prompts.
Omitted values retain their saved value (or built-in default):

```powershell
& $guides -Action Configure `
    -MarginLeft 0.75 -MarginRight 0.75 `
    -MarginTop 0.5 -MarginBottom 0.5
```

```powershell
# Use calibrated OneNote PDF geometry permanently
& $guides -Action Configure -PageSizeProfile OneNotePdfLetter

# Use exact physical Letter geometry permanently
& $guides -Action Configure -PageSizeProfile Letter

# Example custom calibration
& $guides -Action Configure `
    -PageSizeProfile Custom `
    -PageWidthPoints 611.4 `
    -PageHeightPoints 792.84 `
    -PagePitchPoints 792.84
```

Reset all saved geometry and margins to `OneNotePdfLetter`, automatic pitch,
and left/right 1 inch plus top/bottom 0.5 inch:

```powershell
& $guides -Action Configure -ResetSettings
```

### One-time overrides

Explicit values override saved settings only for that command and do not alter
`settings.json` unless the action is Configure:

```powershell
& $guides -Action Refresh -Pages 5 `
    -MarginLeft 0.5 -MarginRight 0.5 `
    -MarginTop 0.75 -MarginBottom 0.75
```

Margins must leave a positive content area. Values are validated before any
OneNote write. `-HideMargins` hides the dashed inner margin rectangle in Paper
mode without changing the underlying geometry.

## Use

Running without `-Action` performs the read-only Status action.

```powershell
# Read-only preflight; builds and validates guides without changing the page
& $guides -Action Refresh -Pages 2 -WhatIf

# Add guides only when no recognized V2 guides exist
& $guides -Action Add -Pages 2

# Safely stage and verify replacements, then remove the old set
& $guides -Action Refresh -Pages 10

# Inspect or remove recognized guide images
& $guides -Action Status
& $guides -Action Remove
```

Each invocation asks you to select a page and, for modifications, confirms the
full target. Use `-PageFilter`, `-ConsolePicker`, or an exact `-PageId` when
needed.

## Geometry and calibration

The default `OneNotePdfLetter` profile is **611.40 × 792.84 points** in portrait
(792.84 × 611.40 landscape). These calibrated values come from measured desktop
OneNote PDF exports. They do not redefine physical US Letter paper. The
`Letter` profile remains exactly **612 × 792 points** in portrait (792 × 612
landscape). Paper mode draws the effective sheet with a solid outer frame and
dashed margin rectangle. One inch is 72 points.

The geometry parameters are `PageSizeProfile`, `PageWidthPoints`,
`PageHeightPoints`, and `PagePitchPoints`. Width and height are required for a
`Custom` profile and describe portrait geometry; landscape swaps them after
profile resolution. Profile dimensions normally supply width and height for the
two built-in profiles.

With fresh-install defaults, portrait frames begin at Y positions 0.00,
792.84, 1585.68, 2378.52, 3171.36, and 3964.20 points for a six-page guide set
starting at zero.

```powershell
& $guides -Action Refresh -Orientation Landscape -Pages 10
& $guides -Action Refresh -StartX 36 -StartY 72 -Pages 10
```

PrintArea mode draws only the calculated content frame. `PagePitchPoints`
allows independent vertical calibration. A value of `0` automatically uses the
effective frame height, so an explicit `792.84` is normally unnecessary with
the portrait `OneNotePdfLetter` profile:

```powershell
& $guides -Action Refresh -GuideMode PrintArea -PagePitchPoints 720 -Pages 10
```

Neither mode reads printer settings or guarantees OneNote pagination. Calibrate
with the same export settings and disposable content. A constant offset calls
for `StartY`; accumulating drift calls for a pitch adjustment.

Settings use independent precedence: an explicit command-line value overrides
the corresponding saved value, which overrides the built-in default. Status
reports the effective profile, orientation, frame dimensions, automatic or
explicit pitch, and margins for troubleshooting.

## Recovery

Before a real write, the utility creates a transaction journal under
`%LOCALAPPDATA%\OneNotePageGuides\Backups`. If an operation fails, stop and use
the exact command and journal path printed by the script:

```powershell
& $guides -Action Recover -RecoveryFile 'C:\path\to\transaction.json'
```

Recovery restores missing old guide images and removes identifiable images from
the failed new batch. It is not a notebook backup and cannot restore ordinary
notes or identify an image whose metadata OneNote discarded.

## Uninstall

Remove guides from relevant pages first, then run:

```powershell
& $guides -Action Uninstall
```

The installed script and known shortcuts are removed. Recovery journals,
settings, previous script copies, and guides already present in notebooks are
preserved.

## Project files

- `OneNote-PageGuides.ps1` — supported script
- `tests/static-check.ps1` — parser and safety assertions
- `OneNote-PageGuides-V2.4.1.ps1` and companion files — historical reviewed
  release artifacts retained for traceability
- [`CONTRIBUTING.md`](CONTRIBUTING.md) — development and test guidance
- [`SECURITY.md`](SECURITY.md) — private reporting and local-data notes

## License

[MIT](LICENSE)
