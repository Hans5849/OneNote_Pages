# Contributing

Thanks for helping improve OneNote Page Guides.

## Development requirements

- Windows PowerShell 5.1 on Windows
- Desktop OneNote with the COM API available (integration testing only)
- A disposable notebook page for any live write test

## Before opening a pull request

1. Run `./OneNote-PageGuides.ps1 -Action SelfTest` in Windows PowerShell 5.1.
2. Run `./tests/static-check.ps1`.
3. If OneNote behavior changed, test Add, Refresh, Status, Remove, and recovery
   on a disposable page. Never use important notebook content for a first test.
4. Document user-visible switches and safety implications in `README.md`.

Do not commit recovery journals: they can contain embedded page-guide image data
and local notebook metadata.
