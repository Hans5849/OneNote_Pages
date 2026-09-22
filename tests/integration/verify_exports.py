#!/usr/bin/env python3
"""Verify disposable OneNote PDF exports. Requires PyMuPDF (`pip install pymupdf`)."""
import argparse, json, pathlib, sys, tempfile
import pymupdf as fitz

TOL = 0.75

def close(a, b, tolerance=TOL): return abs(a - b) <= tolerance
def expected_sheets(bounds, guide, count):
    top, bottom = bounds[1], bounds[3]
    return [n + 1 for n in range(count)
            if bottom > guide["originY"] + n * guide["advance"]
            and top < guide["originY"] + n * guide["advance"] + guide["height"]]

def verify(path, manifest):
    doc = fitz.open(path); errors = []
    if doc.page_count != manifest["pageCount"]: errors.append(f"page count {doc.page_count}")
    expected_box = manifest["mediaBox"]
    all_text = []
    for number, page in enumerate(doc):
        media, crop = page.mediabox, page.cropbox
        for label, box in (("MediaBox", media), ("CropBox", crop)):
            got = [box.x0, box.y0, box.x1, box.y1]
            if any(not close(x, y) for x, y in zip(got, expected_box)):
                errors.append(f"page {number+1} {label} {got}")
        all_text.append(page.get_text())
        for label, point in manifest["marginMarks"].items():
            hits=page.search_for(label)
            if len(hits) != 1:
                errors.append(f"page {number+1} needs one {label} margin mark")
                continue
            # Labels are placed with their upper-left corner at the stated
            # registration point; their bounds make margin conversion observable.
            if not close(hits[0].x0, point[0]) or not close(hits[0].y0, point[1]):
                errors.append(f"page {number+1} {label} at {hits[0].x0:.2f},{hits[0].y0:.2f}; expected {point}")
        pix = page.get_pixmap(matrix=fitz.Matrix(1, 1), alpha=False)
        # Guide PNG uses cyan edges. It must have been removed before export.
        samples = pix.samples
        if any(samples[i] < 40 and samples[i+1] > 180 and samples[i+2] > 180
               for i in range(0, len(samples)-2, pix.n)):
            errors.append(f"page {number+1} contains guide-cyan pixels")
    margins = manifest["margins"]
    if (margins != {"left":72,"right":72,"top":36,"bottom":36}): errors.append("manifest margins are not 72/72/36/36 points")
    for obj in manifest["objects"]:
        calculated = expected_sheets(obj["bounds"], manifest["guide"], manifest["pageCount"])
        if calculated != obj["expectedSheets"]:
            errors.append(f"manifest decision {obj['id']}: expected {obj['expectedSheets']}, geometry says {calculated}")
        occurrences = [i+1 for i, text in enumerate(all_text) if obj["id"] in text]
        # Every fixture object has its ID as a nearby text label, including ink/photos.
        if occurrences != obj["expectedSheets"]:
            errors.append(f"{obj['id']} rendered on {occurrences}; expected {obj['expectedSheets']} (missing or duplicate)")
    # Boundary labels provide one measured Y sample per sheet. Compare deltas to
    # page advance so constant offsets are not confused with cumulative drift.
    anchors=[]
    for page in doc:
        hits=page.search_for("SHEET-ANCHOR")
        anchors.append(hits[0].y0 if len(hits)==1 else None)
    if any(v is None for v in anchors): errors.append("each sheet must contain exactly one SHEET-ANCHOR")
    elif max(anchors)-min(anchors) > TOL: errors.append(f"cumulative boundary drift {anchors}")
    return doc, errors

def main():
    p=argparse.ArgumentParser()
    p.add_argument("--notebook-name", required=True, choices=["OneNote Page Guides Integration Tests"])
    p.add_argument("--manifest", required=True); p.add_argument("--baseline", required=True)
    p.add_argument("--changed-width", required=True); p.add_argument("--report", required=True)
    a=p.parse_args(); manifest=json.loads(pathlib.Path(a.manifest).read_text(encoding="utf-8"))
    temp_root=pathlib.Path(tempfile.gettempdir()).resolve()
    artifacts=[pathlib.Path(a.baseline).resolve(),pathlib.Path(a.changed_width).resolve(),pathlib.Path(a.report).resolve()]
    if any(temp_root not in item.parents for item in artifacts):
        p.error("PDFs and report must be below the operating-system temporary directory")
    base, errors=verify(a.baseline,manifest); changed, more=verify(a.changed_width,manifest); errors += more
    ids=[o["id"] for o in manifest["objects"]]
    def union_width(doc):
        boxes=[r for page in doc for ident in ids for r in page.search_for(ident)]
        return sum(r.width for r in boxes)/len(boxes) if boxes else 0
    baseline_width, changed_width=union_width(base),union_width(changed)
    if close(baseline_width,changed_width,.25): errors.append("content-width rerun did not produce a measurable scale change")
    result={"baseline":str(pathlib.Path(a.baseline).resolve()),"changedWidth":str(pathlib.Path(a.changed_width).resolve()),
            "pageCount":base.page_count,"baselineLabelMeanWidth":baseline_width,"changedLabelMeanWidth":changed_width,
            "checks":"MediaBox/CropBox, margins, drift, missing/duplicates, guide absence, crossing decisions, scale change",
            "passed":not errors,"errors":errors,"limitation":"Raster/color and text-label checks complement, but do not prove, OneNote's internal object identity."}
    pathlib.Path(a.report).write_text(json.dumps(result,indent=2)+"\n",encoding="utf-8")
    if errors: print("\n".join("FAIL: "+e for e in errors),file=sys.stderr); return 1
    print("PASS: disposable OneNote exports verified; report: "+a.report); return 0
if __name__ == "__main__": raise SystemExit(main())
