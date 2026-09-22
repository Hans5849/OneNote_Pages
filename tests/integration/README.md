# Disposable desktop OneNote export fixture

This directory contains a **procedure and non-sensitive manifest**, not a copy
of a user's notebook. Desktop OneNote integration cannot run in CI or be
inferred from unit-test success.

## Create the fixture

1. In desktop OneNote, create a local notebook named exactly **OneNote Page
   Guides Integration Tests**. Create `Fixture Source`, then immediately use
   **Move or Copy > Copy** to make `Fixture Run <timestamp>`. Never run the
   utility against `Fixture Source` or another notebook.
2. On the copied page, create each object in `fixture-manifest.json` at its
   stated point bounds. Use real pen ink for `ink`, a newly taken disposable
   photograph (for example, blank paper with the fixture ID), typed OneNote
   text for `text`, and native shapes/arrows for `diagram`. Place the object's
   exact ID beside/grouped with it. Add `SHEET-ANCHOR` at the same position
   within each of the first, middle, and final sheets. The manifest includes
   wholly enclosed objects and crossings of left, right, top, bottom, and an
   overlapping page-advance boundary.
   Also add each `marginMarks` label with its text bounding box's upper-left at
   the listed point; these registration labels make 72/36-point margins
   measurable in the PDF rather than merely trusting configured values.
3. Add/refresh guides on **only the copy**, then remove them before exporting.
   Export into a new directory below `$env:TEMP`; do not export into this repo.
   Produce `baseline.pdf`, change OneNote's content width, and export
   `changed-width.pdf`. Inspect the first, middle, and final sheets.
4. Run the verifier (PyMuPDF is the only Python dependency):

   ```powershell
   $out = Join-Path $env:TEMP ("OneNoteGuideExport-" + [guid]::NewGuid())
   New-Item $out -ItemType Directory | Out-Null
   python .\tests\integration\verify_exports.py `
     --notebook-name "OneNote Page Guides Integration Tests" `
     --manifest .\tests\integration\fixture-manifest.json `
     --baseline "$out\baseline.pdf" --changed-width "$out\changed-width.pdf" `
     --report "$out\verification.json"
   ```

Only after this command passes, copy `baseline.pdf` to
`$out\representative-corrected.pdf`. Keep PDFs, reports, page IDs, notebook
paths, and photographs out of version control. Record the measured page boxes,
drift, label-width scale change, and limitations from `verification.json` in
the pull-request body.

The verifier checks every produced PDF for page count, equal MediaBox/CropBox,
the declared 72-point horizontal and 36-point vertical margins, cumulative
anchor drift, missing/duplicate labeled objects, guide-cyan pixels, and every
crossing decision. It then requires the changed-content-width export to show a
measurable scale change. PDF rendering cannot prove OneNote object identity;
visual review of ink and photographs on the first, middle, and last sheets is
therefore still mandatory.
