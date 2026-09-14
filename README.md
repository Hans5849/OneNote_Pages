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
- Persistent margin settings with an interactive configuration shortcut
- Letter portrait and landscape layouts
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

Installation copies the utility to
`%LOCALAPPDATA%\OneNotePageGuides\OneNote-PageGuides-V2.ps1`. It creates Start
Menu shortcuts for Refresh, Remove, Configure Margins, Status, and Self-Test;
the optional desktop shortcuts are Refresh and Remove. Existing installations
at that compatible path are backed up with a `.previous` suffix.

## Adjust margins

### Interactive, persistent settings

Run the **Configure OneNote Page Guide Margins** Start Menu shortcut, or:

```powershell
$guides = "$env:LOCALAPPDATA\OneNotePageGuides\OneNote-PageGuides-V2.ps1"
& $guides -Action Configure
```

Enter margins in inches. Press Enter at a prompt to retain its current value.
The settings are saved locally in
`%LOCALAPPDATA%\OneNotePageGuides\settings.json`; subsequent Add and Refresh
operations—including the Refresh shortcut—use them automatically.

For scripting, save all four values without interactive prompts by supplying at
least one margin argument. Omitted values retain their saved value (or the
built-in default):

```powershell
& $guides -Action Configure `
    -MarginLeft 0.75 -MarginRight 0.75 `
    -MarginTop 0.5 -MarginBottom 0.5
```

Reset to the built-in defaults (left/right 1 inch, top/bottom 0.5 inch):

```powershell
& $guides -Action Configure -ResetSettings
```

### One-time overrides

Explicit values override saved margins only for that command:

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

Letter Paper mode draws a 612 × 792 point portrait sheet (792 × 612 landscape),
with a solid outer frame and dashed margin rectangle. One inch is 72 points.

```powershell
& $guides -Action Refresh -Orientation Landscape -Pages 10
& $guides -Action Refresh -StartX 36 -StartY 72 -Pages 10
```

PrintArea mode draws only the calculated content frame. `PagePitchPoints`
allows independent vertical calibration:

```powershell
& $guides -Action Refresh -GuideMode PrintArea -PagePitchPoints 720 -Pages 10
```

Neither mode reads printer settings or guarantees OneNote pagination. Calibrate
with the same export settings and disposable content. A constant offset calls
for `StartY`; accumulating drift calls for a pitch adjustment.

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
