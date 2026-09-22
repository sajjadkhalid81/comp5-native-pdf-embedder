"""
crs_core.py - COMP5 CRS Fill & Zip.

For documents where the discipline engineer has not provided an Excel CRS
(Comment Response Sheet), this fills the standard CTR CRS - Aconex template
from each PDF's own cover page + a tracker Excel (e.g. the "Vendor DC
Workflow" sheet), and packages it with the (untouched) PDF into a
correctly-named ZIP.

Per document:
  {DocNo}_{Rev}_CTR CRS.zip
    {DocNo}_{Rev}.pdf                          <- original PDF bytes, byte-for-byte
    CTR CRS - Aconex_{DocNo}_{Rev}.xlsx         <- filled from the template

Document No / Document Class / Document Title are auto-extracted from each
PDF's own cover page - the PDF is never opened with pikepdf and never
re-saved, so any natives already embedded in it are left completely intact.
Revision comes from the filename (…_{REV}.pdf).

The Return Code is looked up from an uploaded tracker Excel. Two formats
are recognised, auto-detected by header row (searched within the first 20
rows of each sheet, so export-metadata rows above the real header are
skipped):
  - Aconex "Workflow Search Export To Excel" (the trained format): columns
    Document No., Document Revision, Document Title, Step Outcome.
  - Any other tracker with a Document Number/No. column and a
    "...Review Outcome" column (e.g. the older "Saipem Review Outcome").
Rows are matched to the PDF by (Document No., Revision) when a revision
column exists; if there's no exact-revision row, it falls back to any row
for that Document No. and flags the revision mismatch. Multiple rows for
the same key with different outcomes are treated as ambiguous and skipped
with an error rather than guessed.

Reviewer Group (CONTRACTOR REVIEW CODE / who reviewed it) isn't in the
tracker and can't be reliably inferred from the PDF's Discipline field
alone (e.g. Mechanical can be Static or Rotating) - it is left as the
template's own value and always flagged for manual check before the CRS
is sent out.
"""
import io
import os
import re
import zipfile

import openpyxl
from pypdf import PdfReader

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "assets", "CTR_CRS_Template.xlsx")
TEMPLATE_SHEET = "CTR CRS"

RETURN_CODES = {
    "A": "Approved/Reviewed with no comments",
    "B": "Approved/Reviewed with minor comments",
    "C": "Returned with comments",
    "D": "Rejected",
    "I": "Information only / No Review",
}
RETURN_CODE_ROW = {"A": 11, "B": 12, "C": 13, "D": 14, "I": 15}


class CrsError(Exception):
    """User-facing error with a clear message."""
    pass


# ------------------------------------------------------------- extraction

def _clean(text):
    return re.sub(r"\s+", " ", text or "").strip()


def extract_cover_page(pdf_bytes: bytes, filename: str):
    """Pull Document No / Class / Title from the PDF's own cover page,
    and Revision from the filename (…_{REV}.pdf). Read-only - the PDF
    bytes themselves are never touched."""
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        text = reader.pages[0].extract_text() or ""
    except Exception as e:
        raise CrsError(f"{filename}: could not read PDF cover page ({e}).")

    doc_no = ""
    m = re.search(r"COMPANY Document No\.?\s*:?\s*([^\n]+)", text)
    if m:
        # pypdf sometimes inserts stray spaces inside the number itself
        # (e.g. "6951 _24-1553653 -00038") from PDF kerning gaps - strip all
        # internal whitespace rather than stopping at the first one.
        doc_no = re.sub(r"\s+", "", m.group(1))

    doc_class = ""
    m = re.search(r"Document Class\s*:\s*([A-Za-z0-9]+)", text)
    if m:
        doc_class = m.group(1).strip()

    title = ""
    m = re.search(r"Document Title\s*:\s*(.+?)COMPANY Document No", text, re.S)
    if m:
        title = _clean(m.group(1))

    stem = os.path.splitext(os.path.basename(filename))[0]
    rev = ""
    doc_no_from_name = ""
    parts = stem.rsplit("_", 1)
    if len(parts) == 2 and 0 < len(parts[1]) <= 3 and parts[1].isalnum():
        doc_no_from_name, rev = parts[0], parts[1].upper()

    if not doc_no:
        doc_no = doc_no_from_name or stem

    warnings = []
    if not doc_no:
        warnings.append("Document No. not found on cover page or in filename")
    elif doc_no_from_name and doc_no.upper() != doc_no_from_name.upper():
        warnings.append(
            f"Document No. on cover page ('{doc_no}') doesn't match the filename ('{doc_no_from_name}') - using cover page value, please double-check"
        )
    if not rev:
        warnings.append("Revision not found in filename (expected …_<REV>.pdf)")
    if not doc_class:
        warnings.append("Document Class not found on cover page")
    if not title:
        warnings.append("Document Title not found on cover page")

    return {
        "doc_no": doc_no,
        "rev": rev,
        "doc_class": doc_class,
        "title": title,
        "warnings": warnings,
    }


# ----------------------------------------------------------- tracker read

_DOC_NO_HEADERS = {"document number", "document no", "document no."}
_OUTCOME_HEADER_HINTS = ("step outcome", "review outcome")
_REVISION_HEADERS = {"document revision", "revision"}
_TITLE_HEADERS = {"document title", "title"}
_EXTRA_HEADER_HINTS = {"revision status", "remark", "remarks", "step status", "workflow status"}

_HEADER_SCAN_ROWS = 20  # Aconex "Workflow Search Export" has ~9 metadata rows before the real header


def _norm_header(v):
    return _clean(str(v)).lower().rstrip(".") if v is not None else ""


def parse_tracker(tracker_bytes: bytes, tracker_filename: str):
    """Scan every sheet for a header row with a Document No./Number column
    and a Step/Review Outcome column (Document Revision / Document Title /
    extra columns are picked up if present). Returns
    {doc_no: [{"revision": str, "outcome": str, "extra": [str]}, ...]}."""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(tracker_bytes), data_only=True)
    except Exception as e:
        raise CrsError(f"{tracker_filename}: could not open tracker Excel ({e}).")

    lookup = {}
    matched_any_sheet = False

    for ws in wb.worksheets:
        header_row = None
        for row in ws.iter_rows(min_row=1, max_row=min(_HEADER_SCAN_ROWS, ws.max_row)):
            headers = [_norm_header(c.value) for c in row]
            has_doc_no = any(h.rstrip(".") in _DOC_NO_HEADERS for h in headers)
            has_outcome = any(any(hint in h for hint in _OUTCOME_HEADER_HINTS) for h in headers)
            if has_doc_no and has_outcome:
                header_row = row[0].row
                break
        if header_row is None:
            continue

        matched_any_sheet = True
        header_cells = next(ws.iter_rows(min_row=header_row, max_row=header_row))
        headers = [_norm_header(c.value) for c in header_cells]
        doc_no_col = next(i for i, h in enumerate(headers) if h.rstrip(".") in _DOC_NO_HEADERS)
        outcome_col = next(i for i, h in enumerate(headers) if any(hint in h for hint in _OUTCOME_HEADER_HINTS))
        rev_col = next((i for i, h in enumerate(headers) if h in _REVISION_HEADERS), None)
        title_col = next((i for i, h in enumerate(headers) if h in _TITLE_HEADERS), None)
        extra_cols = [i for i, h in enumerate(headers) if h in _EXTRA_HEADER_HINTS]

        for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row):
            if doc_no_col >= len(row):
                continue
            doc_no_val = row[doc_no_col].value
            if not doc_no_val:
                continue
            doc_no = _clean(str(doc_no_val))
            outcome_val = row[outcome_col].value if outcome_col < len(row) else None
            outcome = _clean(str(outcome_val)) if outcome_val else ""
            if not outcome:
                continue
            revision = ""
            if rev_col is not None and rev_col < len(row) and row[rev_col].value:
                revision = _clean(str(row[rev_col].value)).upper()
            title = ""
            if title_col is not None and title_col < len(row) and row[title_col].value:
                title = _clean(str(row[title_col].value))
            extra = []
            for c in extra_cols:
                if c < len(row) and row[c].value:
                    extra.append(_clean(str(row[c].value)))
            lookup.setdefault(doc_no, []).append(
                {"revision": revision, "outcome": outcome, "title": title, "extra": extra}
            )

    if not matched_any_sheet:
        raise CrsError(
            f"{tracker_filename}: no sheet found with a 'Document No.' column and a "
            f"'Step Outcome' / '...Review Outcome' column."
        )

    return lookup


def match_tracker(tracker_lookup: dict, doc_no: str, rev: str):
    """Resolve one document's tracker row(s) to a single outcome entry.
    Returns (entry_dict, note) where note flags a revision fallback, or
    raises CrsError if not found / ambiguous."""
    candidates = tracker_lookup.get(doc_no)
    if not candidates:
        raise CrsError(f"Document No. {doc_no} not found in the tracker")

    have_revisions = any(c["revision"] for c in candidates)
    exact = [c for c in candidates if not have_revisions or c["revision"] == rev.upper()]
    note = None

    if not exact:
        # no row for this exact revision - fall back to whatever revision(s) exist
        outcomes = {c["outcome"] for c in candidates}
        if len(outcomes) > 1:
            revs = ", ".join(sorted({c["revision"] or "?" for c in candidates}))
            raise CrsError(
                f"Document No. {doc_no}: no tracker row for Rev {rev}, and the other revisions "
                f"in the tracker ({revs}) have different outcomes - can't pick one automatically"
            )
        exact = candidates
        revs = ", ".join(sorted({c["revision"] or "?" for c in candidates}))
        note = f"no tracker row for Rev {rev} - using outcome from tracker Rev {revs}"
    else:
        outcomes = {c["outcome"] for c in exact}
        if len(outcomes) > 1:
            raise CrsError(
                f"Document No. {doc_no} Rev {rev}: multiple tracker rows with different outcomes "
                f"({'; '.join(sorted(outcomes))}) - can't pick one automatically"
            )

    chosen = exact[0]
    extra = []
    for c in exact:
        for e in c["extra"]:
            if e not in extra:
                extra.append(e)
    return {"outcome": chosen["outcome"], "title": chosen["title"], "extra": extra}, note


def resolve_return_code(outcome_text: str):
    """'Code D - Rejected' / 'Code D' / 'D - Rejected' -> 'D'. Raises CrsError if unclear."""
    m = re.search(r"\bCode\s*[:\-]?\s*([A-Za-z])\b", outcome_text, re.I)
    if not m:
        m = re.match(r"\s*([A-Za-z])\s*[-:]", outcome_text)
    if not m:
        raise CrsError(f"could not parse a return code letter from '{outcome_text}'")
    code = m.group(1).upper()
    if code not in RETURN_CODES:
        raise CrsError(f"'{code}' parsed from '{outcome_text}' is not a valid return code (A/B/C/D/I)")
    return code


# ------------------------------------------------------------------ fill

def fill_crs(doc_no, rev, doc_class, title, return_code, reviewer_group=""):
    """Fill the CTR CRS template. Returns the .xlsx bytes."""
    if return_code not in RETURN_CODES:
        raise CrsError(f"Invalid return code '{return_code}'.")

    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb[TEMPLATE_SHEET]

    ws["D5"] = doc_no
    ws["H5"] = rev
    ws["K5"] = doc_class
    ws["D6"] = title
    ws["D20"] = (
        "Document sent for review to QELNG in ACONEX and QELNG has provided a "
        f"return code:\n{return_code}-{RETURN_CODES[return_code]}"
    )
    ws[f"G{RETURN_CODE_ROW[return_code]}"] = "X"
    if reviewer_group:
        ws["G21"] = reviewer_group.strip() + "\n"

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# --------------------------------------------------------------- per-doc

def build_one(pdf_bytes: bytes, filename: str, tracker_lookup: dict):
    """One PDF -> (zip_name, zip_bytes, detail_string). Raises CrsError."""
    info = extract_cover_page(pdf_bytes, filename)
    doc_no, rev = info["doc_no"], info["rev"]

    if not doc_no:
        raise CrsError(f"{filename}: could not determine Document No. - skipped.")
    if not rev:
        raise CrsError(
            f"{filename}: could not determine Revision from filename "
            f"(expected …_<REV>.pdf, e.g. …_B.pdf) - skipped."
        )

    try:
        entry, rev_note = match_tracker(tracker_lookup, doc_no, rev)
        return_code = resolve_return_code(entry["outcome"])
    except CrsError as e:
        raise CrsError(f"{filename}: {e} - skipped.")

    title = entry["title"] or info["title"]  # tracker's own title is more reliable when present
    xlsx_bytes = fill_crs(doc_no, rev, info["doc_class"], title, return_code)

    pdf_out_name = f"{doc_no}_{rev}.pdf"
    xlsx_out_name = f"CTR CRS - Aconex_{doc_no}_{rev}.xlsx"
    zip_name = f"{doc_no}_{rev}_CTR CRS.zip"

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(pdf_out_name, pdf_bytes)   # original bytes, untouched
        z.writestr(xlsx_out_name, xlsx_bytes)
    buf.seek(0)

    detail = (f"Tracker outcome '{entry['outcome']}' -> return code {return_code} "
              f"(Doc Class {info['doc_class'] or '?'}), zipped as {zip_name}")
    if rev_note:
        detail += " | ATTENTION: " + rev_note
    if entry["extra"]:
        detail += " | tracker notes: " + "; ".join(entry["extra"])
    detail += " | ATTENTION: reviewer group (CONTRACTOR REVIEW CODE) left as template default - verify/set manually before sending"
    # doc_no/rev "not found" cases already raised above and never reach here;
    # anything left (class/title not found, cover-page/filename mismatch) is new info - show it.
    if info["warnings"]:
        detail += " | " + "; ".join(info["warnings"])

    return zip_name, buf.getvalue(), detail


# ------------------------------------------------------------------ batch

def process_batch(files, tracker_bytes, tracker_filename):
    """files: list of (filename, bytes) - PDFs only.
    Returns (out_filename, out_bytes, results)."""
    try:
        tracker_lookup = parse_tracker(tracker_bytes, tracker_filename)
    except CrsError as e:
        return None, None, [{"file": tracker_filename, "status": "error", "detail": str(e)}]

    outputs, results = [], []
    for fname, fbytes in files:
        try:
            zip_name, zip_bytes, detail = build_one(fbytes, fname, tracker_lookup)
            results.append({"file": fname, "status": "warn", "detail": detail})  # always warn: reviewer group needs a manual check
            outputs.append((zip_name, zip_bytes))
        except CrsError as e:
            results.append({"file": fname, "status": "error", "detail": str(e)})

    if not outputs:
        return None, None, results
    if len(outputs) == 1:
        return outputs[0][0], outputs[0][1], results

    bundle = io.BytesIO()
    used = {}
    with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in outputs:
            if name in used:
                used[name] += 1
                stem, ext = os.path.splitext(name)
                name = f"{stem} ({used[name]}){ext}"
            else:
                used[name] = 1
            z.writestr(name, data)
    bundle.seek(0)
    return "CRS_ZIPS.zip", bundle.read(), results
