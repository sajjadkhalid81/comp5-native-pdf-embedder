# COMP5 Native Embedder - Complete Guide

**Prepared by: Khalid Sajjad | For Internal Use Only**
**COMP5 EPC of NFPS Offshore Compression Complexes (Project No. 033784)**

---

## 1. What This Tool Does

The new document system requires all native files to be embedded inside the
PDF for upload to Aconex, and after review the commented PDF must be
converted back into a ZIP for return to the vendor and COMPANY.

**EMBED (ZIP -> PDF)** - Step 1, vendor submission:
- Input : ZIP containing one main PDF + native files (DWG, XLSX, DOCX, CRS PDF, etc.)
- Output: the main PDF with all natives embedded as attachments
  (visible in Adobe Acrobat's paperclip / Attachments panel)
- A checksum manifest (_MANIFEST.txt) is embedded alongside the natives

**EXTRACT (PDF -> ZIP)** - Step 2, after review:
- Input : the commented/approved PDF downloaded from Aconex
- Output: ZIP named after the document number, containing:
  - the commented PDF with ALL attachments stripped (clean, smaller)
  - every native file, byte-identical to the vendor's originals
  - _VERIFICATION_LOG.txt proving what was extracted and whether it changed

---

## 2. Project Files

| File | Purpose |
|---|---|
| `app.py` | Flask routes: `/` (UI) and `/api/process` (streaming, no in-memory store) |
| `native_core.py` | All embed/extract logic. Standalone module. |
| `templates/index.html` | Web UI: mode selector, drop zone, results panel |
| `requirements.txt` | Pinned, tested versions (pikepdf 10.5.1) |
| `Procfile` | `gunicorn app:app --workers 2 --timeout 120 --bind 0.0.0.0:$PORT` |

**Rules (same as comp5-reports project):**
- The uploaded files here are the authoritative master version.
- Never modify working logic without explicit instruction.
- File input element stays OUTSIDE the drop zone div (Windows fix).
- Document-level dragover/drop preventDefault at top of script.

---

## 3. Built-in Automatic Checks

1. **Checksum manifest** - EMBED records filename, size, MD5 of every native
   inside the PDF. EXTRACT compares each extracted file against it:
   UNCHANGED / CHANGED / NEW / MISSING. Result goes into
   `_VERIFICATION_LOG.txt` inside the output ZIP.
2. **Bookmark & TOC integrity** - Before and after every conversion the tool
   counts all bookmarks (incl. nested) and internal TOC links and verifies
   each destination still points to a valid page. If the source has none,
   it reports "no bookmarks/TOC links in source" (nothing to keep = pass).
   Any loss = amber ATTENTION warning.
3. **Smart host-PDF selection** - If a ZIP contains several PDFs (e.g. main
   document + CRS), the main one is identified by matching the document
   number in the ZIP filename; fallback: the PDF not prefixed "CRS".
   Everything else, including extra PDFs, gets embedded.
4. **Master-ZIP auto-detection** - If an uploaded ZIP contains only .zip
   files, each inner ZIP is treated as one document package automatically.
5. **Batch safety** - A bad file in a batch is reported but never stops the
   good ones. Duplicate output names get " (2)", " (3)" suffixes.

---

## 4. Deploy on Render (same workflow as comp5-reports-web)

### 4.1 GitHub
1. Create a new **private** repository, e.g. `comp5-native-embedder`.
2. Upload these files keeping the structure:
   ```
   app.py
   native_core.py
   requirements.txt
   Procfile
   templates/index.html
   ```
3. Recommended: create a `dev` branch for future changes, merge to `main`
   via PR (same main/dev workflow as the reports app).

### 4.2 Render
1. Render Dashboard -> **New** -> **Web Service** -> connect the repo.
2. Settings:
   - Branch: `main`
   - Build command: `pip install -r requirements.txt`
   - Start command: leave empty (taken from Procfile automatically)
   - Instance type: Free tier is fine to start
3. Deploy. First build takes 2-3 minutes.

### 4.3 If Render caches the wrong pikepdf version
Same trick as the openpyxl issue on the reports app: delete
`requirements.txt` from the repo, commit, then re-add it and commit again.
This forces a clean rebuild.

---

## 5. Run Locally (for testing)

```
pip install -r requirements.txt
python app.py
```
Open http://127.0.0.1:5000

---

## 6. Daily Usage

### Step 1 - Vendor submission arrives
1. Open the app -> **EMBED** mode.
2. Drop the vendor ZIP(s). Works with:
   - single document ZIPs (one per document), many at once
   - a master ZIP containing many document ZIPs (auto-detected)
3. Click **Process Files**.
4. One document -> the embedded PDF downloads directly.
   Many documents -> one `EMBEDDED_PDFS.zip` downloads with all PDFs inside.
5. Check the results panel (green = OK). Upload PDFs to Aconex.

### Step 2 - Review completed
1. Download the commented/approved PDF(s) from Aconex.
2. Open the app -> **EXTRACT** mode.
3. Drop the PDF(s) -> **Process Files**.
4. Each document returns as `<doc-number>.zip` containing the clean
   commented PDF + natives + verification log.
5. **Read the results panel**: green "verified UNCHANGED" = safe to
   transmit. Amber ATTENTION = something changed/missing - investigate
   before sending (the log inside the ZIP shows exactly which file).
6. Transmit ZIPs to vendor and COMPANY.

### Limits
- 300 MB per upload batch (about 100 typical documents; split if larger).
- Server timeout 120 s per request (a 20-document batch takes ~5 s).

---

## 7. Known Considerations (verify once before full rollout)

1. **Aconex round-trip**: confirm Aconex preserves embedded attachments
   when reviewers comment in its own viewer. Test: upload an embedded PDF,
   add a dummy comment, download, run EXTRACT. If natives come back
   "verified UNCHANGED", the workflow is safe. (Commenting in
   Foxit/Adobe is already proven safe - tested with a real document.)
2. **Digital signatures**: stripping attachments rewrites the PDF, which
   breaks digital-signature *validation* (visible stamps remain, but
   Acrobat will report the document as modified after signing). If the
   returned PDF must carry valid verifiable signatures, request the
   no-strip variant (one-line change).

---

## 8. Troubleshooting

| Symptom | Cause / Fix |
|---|---|
| "contains N PDFs and the main document could not be identified" | ZIP has multiple PDFs and none matches the ZIP name / CRS rule. Rename the ZIP to include the main document number, or split. |
| "no embedded files found in this PDF" | The PDF has no attachments - it was never embedded, or Aconex stripped them (see 7.1). |
| Amber "CHANGED" on extract | The native inside the PDF differs from the vendor's original. Check the log, query whether the change was intentional. |
| Amber "bookmark/TOC issue" | Navigation was lost in conversion. Do not upload; report the document. |
| Render build uses old pikepdf | Force rebuild: see 4.3. |
| Drop zone opens the file in the browser | Should not happen (global preventDefault). If editing index.html, keep the two document-level listeners at the top of the script. |
| Worker timeout on huge batches | Split the batch. Timeout is 120 s in the Procfile. |

---

## 9. Version Notes

- Python libs: flask 3.0.3, gunicorn 22.0.0, pikepdf 10.5.1, werkzeug 3.0.3.
- pikepdf 10.5.1 is the exact version all tests were run with - keep it
  pinned unless retesting.
- Test coverage at handover: drawing package (DWG natives, bookmarks),
  SIEMENS doc+CRS packages, master ZIP, DTS package (docx+xlsx natives),
  real commented PDF extract, tamper detection, bookmark-loss detection,
  20-document batch load test, error cases (bad ZIP, wrong type, mixed batch).
