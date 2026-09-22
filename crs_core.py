"""
crs_core.py - COMP5 CRS Fill & Zip.

For documents where the discipline engineer has not provided an Excel CRS
(Comment Response Sheet), this fills the standard CTR CRS - Aconex template
from each PDF's own cover page and packages it with the (untouched) PDF
into a correctly-named ZIP.

Per document:
  {DocNo}_{Rev}_CTR CRS.zip
    {DocNo}_{Rev}.pdf                          <- original PDF bytes, byte-for-byte
    CTR CRS - Aconex_{DocNo}_{Rev}.xlsx         <- filled from the template

Return code and reviewer group are supplied once for the whole batch (a
batch handed over by one discipline engineer is normally one review cycle
with one outcome and one reviewer group), and applied to every document.
Document No / Document Class / Document Title are auto-extracted from each
PDF's own cover page - the PDF is never opened with pikepdf and never
re-saved, so any natives already embedded in it are left completely intact.
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
    m = re.search(r"COMPANY Document No\.?\s*:?\s*([A-Za-z0-9_\-]+)", text)
    if m:
        doc_no = m.group(1).strip()

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
    if not rev:
        rev = ""  # left blank - flagged to the user, never guessed

    warnings = []
    if not doc_no:
        warnings.append("Document No. not found on cover page or in filename")
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


# ------------------------------------------------------------------ fill

def fill_crs(doc_no, rev, doc_class, title, return_code, reviewer_group):
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

def build_one(pdf_bytes: bytes, filename: str, return_code: str, reviewer_group: str):
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

    xlsx_bytes = fill_crs(doc_no, rev, info["doc_class"], info["title"],
                           return_code, reviewer_group)

    pdf_out_name = f"{doc_no}_{rev}.pdf"
    xlsx_out_name = f"CTR CRS - Aconex_{doc_no}_{rev}.xlsx"
    zip_name = f"{doc_no}_{rev}_CTR CRS.zip"

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(pdf_out_name, pdf_bytes)   # original bytes, untouched
        z.writestr(xlsx_out_name, xlsx_bytes)
    buf.seek(0)

    detail = f"CRS filled (Doc Class {info['doc_class'] or '?'}, return code {return_code}), zipped as {zip_name}"
    if info["warnings"]:
        detail += " | ATTENTION: " + "; ".join(w for w in info["warnings"] if "Document No" not in w and "Revision" not in w)

    return zip_name, buf.getvalue(), detail


# ------------------------------------------------------------------ batch

def process_batch(files, return_code, reviewer_group):
    """files: list of (filename, bytes). Returns (out_filename, out_bytes, results)."""
    outputs, results = [], []
    for fname, fbytes in files:
        try:
            zip_name, zip_bytes, detail = build_one(fbytes, fname, return_code, reviewer_group)
            st = "warn" if "ATTENTION" in detail else "ok"
            results.append({"file": fname, "status": st, "detail": detail})
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
