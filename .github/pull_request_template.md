## Summary

- Describe the change.

## Testing

- [ ] `./tests/static-check.ps1`
- [ ] `./OneNote-PageGuides.ps1 -Action SelfTest`
- [ ] Live OneNote testing on a disposable page (when integration changed)

## Safety review

- [ ] Notebook writes remain behind `ShouldProcess`.
- [ ] Changes do not broaden guide deletion selectors.
- [ ] User-facing behavior is documented.
