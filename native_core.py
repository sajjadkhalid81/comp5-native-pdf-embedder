"""
native_core.py - COMP5 Native Embedder core logic.

EMBED  : ZIP (one host PDF + native files) -> PDF with natives embedded
EXTRACT: PDF with embedded natives -> ZIP (clean PDF + natives + log)

Features:
- Smart host-PDF selection (doc number match in ZIP name; CRS fallback)
- Master-ZIP auto-expansion (ZIP of document ZIPs)
- MD5 checksum manifest embedded on EMBED, verified on EXTRACT
- Attachments stripped from the PDF placed in the output ZIP
- Bookmark / TOC link integrity check on every conversion
- Duplicate output name handling in batch bundles
"""
import io
import os
import hashlib
import zipfile
from datetime import datetime, timezone
import pikepdf
from pikepdf import AttachedFileSpec

MANIFEST_NAME = "_MANIFEST.txt"
LOG_NAME = "_VERIFICATION_LOG.txt"


class NativeToolError(Exception):
    """User-facing error with a clear message."""
    pass


def _md5(data: bytes) -> str:
    return hashlib.md5(data).hexdigest()


# ------------------------------------------------- BOOKMARK / TOC CHECK

def _nav_check(pdf):
    """Count bookmarks and internal TOC/link annotations and how many
    resolve to a valid page. Returns dict with counts."""
    page_ids = set()
    for p in pdf.pages:
        try:
            page_ids.add(p.obj.objgen)
        except Exception:
            pass

    def lookup_named(name):
        try:
            from pikepdf import NameTree
            nt = NameTree(pdf.Root.Names.Dests)
            if name in nt:
                return nt[name]
        except Exception:
            pass
        try:
            return pdf.Root.Dests[name]
        except Exception:
            return None

    def dest_valid(dest):
        try:
            if dest is None:
                return False
            if isinstance(dest, (pikepdf.Name, pikepdf.String)) or isinstance(dest, str):
                dest = lookup_named(dest)
                if dest is None:
                    return False
            if isinstance(dest, pikepdf.Dictionary) and "/D" in dest:
                dest = dest.D
            if isinstance(dest, pikepdf.Array) and len(dest) > 0:
                return dest[0].objgen in page_ids
            return False
        except Exception:
            return False

    bm_total = bm_ok = 0

    def walk(items):
        nonlocal bm_total, bm_ok
        for it in items:
            bm_total += 1
            dest = it.destination
            if dest is None and it.action is not None:
                try:
                    if it.action.S == pikepdf.Name("/GoTo"):
                        dest = it.action.D
                except Exception:
                    dest = None
            if dest_valid(dest):
                bm_ok += 1
            walk(it.children)

    try:
        with pdf.open_outline() as outline:
            walk(outline.root)
    except Exception:
        pass

    ln_total = ln_ok = 0
    for p in pdf.pages:
        try:
            annots = p.obj.get("/Annots")
            if annots is None:
                continue
            for a in annots:
                try:
                    if a.get("/Subtype") != pikepdf.Name("/Link"):
                        continue
                    dest = a.get("/Dest")
                    if dest is None:
                        act = a.get("/A")
                        if act is not None and act.get("/S") == pikepdf.Name("/GoTo"):
                            dest = act.get("/D")
                        else:
                            continue  # external URI links - not internal nav
                    ln_total += 1
                    if dest_valid(dest):
                        ln_ok += 1
                except Exception:
                    ln_total += 1
        except Exception:
            pass

    return {"bm": bm_total, "bm_ok": bm_ok, "ln": ln_total, "ln_ok": ln_ok}


def _nav_summary(before, after):
    """Compare nav integrity before/after. Returns (ok, text)."""
    ok = (after["bm"] == before["bm"] and after["ln"] == before["ln"]
          and after["bm_ok"] == after["bm"] and after["ln_ok"] == after["ln"])
    if before["bm"] == 0 and before["ln"] == 0:
        return True, "no bookmarks/TOC links in source"
    text = (f"bookmarks {after['bm_ok']}/{before['bm']} intact, "
            f"TOC links {after['ln_ok']}/{before['ln']} working")
    return ok, text


# ---------------------------------------------------------------- EMBED

def embed_zip_to_pdf(zip_bytes: bytes, zip_name: str):
    """ZIP (host PDF + natives) -> (pdf_name, pdf_bytes, embedded, nav_ok, nav_text)."""
    try:
        zf = zipfile.ZipFile(io.BytesIO(zip_bytes))
    except zipfile.BadZipFile:
        raise NativeToolError(f"{zip_name}: not a valid ZIP file.")

    entries = [i for i in zf.infolist()
               if not i.is_dir()
               and not os.path.basename(i.filename).startswith(".")
               and "__MACOSX" not in i.filename]

    pdfs = [i for i in entries if i.filename.lower().endswith(".pdf")]

    if len(pdfs) == 0:
        raise NativeToolError(f"{zip_name}: no PDF found inside the ZIP.")

    # Choose host PDF. All other files (incl. extra PDFs like CRS) get embedded.
    if len(pdfs) == 1:
        host = pdfs[0]
    else:
        zip_stem = os.path.splitext(os.path.basename(zip_name))[0]
        # Rule 1: PDF whose name (without .pdf) appears inside the ZIP filename
        matches = [p for p in pdfs
                   if os.path.splitext(os.path.basename(p.filename))[0] in zip_stem]
        if len(matches) == 1:
            host = matches[0]
        else:
            # Rule 2: exactly one PDF not prefixed CRS
            non_crs = [p for p in pdfs
                       if not os.path.basename(p.filename).upper().startswith("CRS")]
            if len(non_crs) == 1:
                host = non_crs[0]
            else:
                names = ", ".join(os.path.basename(p.filename) for p in pdfs)
                raise NativeToolError(
                    f"{zip_name}: contains {len(pdfs)} PDFs ({names}) and the main "
                    f"document could not be identified automatically.")

    natives = [i for i in entries if i is not host]
    if len(natives) == 0:
        raise NativeToolError(f"{zip_name}: only one PDF and no other files - nothing to embed.")

    pdf_name = os.path.basename(host.filename)

    try:
        pdf = pikepdf.open(io.BytesIO(zf.read(host)))
    except Exception as e:
        raise NativeToolError(f"{zip_name}: could not open {pdf_name} ({e}).")

    nav_before = _nav_check(pdf)

    embedded = []
    manifest_lines = [
        "NATIVE FILE MANIFEST",
        f"Source ZIP : {zip_name}",
        f"PDF        : {pdf_name}",
        f"Embedded on: {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}",
        f"Tool       : COMP5 Native Embedder",
        "",
        "FILENAME | SIZE_BYTES | MD5",
    ]
    for n in natives:
        native_name = os.path.basename(n.filename)
        data = zf.read(n)
        spec = AttachedFileSpec(pdf, data, filename=native_name,
                                description=f"Native file: {native_name}")
        pdf.attachments[native_name] = spec
        embedded.append(native_name)
        manifest_lines.append(f"{native_name} | {len(data)} | {_md5(data)}")

    manifest = "\r\n".join(manifest_lines).encode("utf-8")
    mspec = AttachedFileSpec(pdf, manifest, filename=MANIFEST_NAME,
                             description="Checksum manifest of embedded natives")
    pdf.attachments[MANIFEST_NAME] = mspec

    out = io.BytesIO()
    pdf.save(out)
    pdf.close()
    out.seek(0)
    out_bytes = out.read()

    # verify bookmarks/TOC survived the rewrite
    check = pikepdf.open(io.BytesIO(out_bytes))
    nav_after = _nav_check(check)
    check.close()
    nav_ok, nav_text = _nav_summary(nav_before, nav_after)

    return pdf_name, out_bytes, embedded, nav_ok, nav_text


# ---------------------------------------------------------------- EXTRACT

def _parse_manifest(text: str):
    """Return {filename: (size, md5)} from manifest text."""
    ref = {}
    for line in text.splitlines():
        parts = [p.strip() for p in line.split("|")]
        if len(parts) == 3 and parts[1].isdigit():
            ref[parts[0]] = (int(parts[1]), parts[2].lower())
    return ref


def extract_pdf_to_zip(pdf_bytes: bytes, pdf_name: str):
    """PDF with attachments -> (zip_name, zip_bytes, extracted, summary).
    The PDF written into the ZIP has ALL attachments removed."""
    try:
        pdf = pikepdf.open(io.BytesIO(pdf_bytes))
    except Exception as e:
        raise NativeToolError(f"{pdf_name}: could not open PDF ({e}).")

    nav_before = _nav_check(pdf)
    names = [n for n in pdf.attachments.keys()]
    native_names = [n for n in names if os.path.basename(n) != MANIFEST_NAME]
    if not native_names:
        pdf.close()
        raise NativeToolError(f"{pdf_name}: no embedded files found in this PDF.")

    # read manifest if present
    ref = {}
    if any(os.path.basename(n) == MANIFEST_NAME for n in names):
        for n in names:
            if os.path.basename(n) == MANIFEST_NAME:
                try:
                    ref = _parse_manifest(
                        pdf.attachments[n].get_file().read_bytes().decode("utf-8", "replace"))
                except Exception:
                    ref = {}
                break

    base = os.path.splitext(pdf_name)[0]
    out = io.BytesIO()
    extracted, statuses = [], []
    log = [
        "EXTRACTION VERIFICATION LOG",
        f"Source PDF  : {pdf_name}",
        f"Extracted on: {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}",
        f"Tool        : COMP5 Native Embedder",
        f"Manifest    : {'FOUND - files verified against original checksums' if ref else 'NOT FOUND - sizes/checksums recorded, no comparison possible'}",
        "",
        "FILENAME | SIZE_BYTES | MD5 | STATUS",
    ]

    seen = set()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        extracted_data = {}
        for n in native_names:
            fname = os.path.basename(n)
            extracted_data[fname] = pdf.attachments[n].get_file().read_bytes()

        # remove ALL attachments (natives + manifest) so the PDF in the ZIP is clean
        for n in list(pdf.attachments.keys()):
            del pdf.attachments[n]
        clean_buf = io.BytesIO()
        pdf.save(clean_buf)
        clean_buf.seek(0)
        clean_bytes = clean_buf.read()
        check = pikepdf.open(io.BytesIO(clean_bytes))
        nav_after = _nav_check(check)
        check.close()
        nav_ok, nav_text = _nav_summary(nav_before, nav_after)
        log.insert(5, f"Navigation  : {nav_text}" + ("" if nav_ok else "  <-- ATTENTION"))
        z.writestr(pdf_name, clean_bytes)  # commented PDF, attachments removed

        for fname, data in extracted_data.items():
            z.writestr(fname, data)
            extracted.append(fname)
            seen.add(fname)
            md5 = _md5(data)
            if ref:
                if fname in ref:
                    status = "UNCHANGED" if ref[fname] == (len(data), md5) else "CHANGED"
                else:
                    status = "NEW (not in original submission)"
            else:
                status = "RECORDED"
            statuses.append((fname, status))
            log.append(f"{fname} | {len(data)} | {md5} | {status}")

        for fname in ref:
            if fname not in seen:
                statuses.append((fname, "MISSING"))
                log.append(f"{fname} | - | - | MISSING (was in original, not in PDF)")

        z.writestr(LOG_NAME, "\r\n".join(log).encode("utf-8"))

    pdf.close()
    out.seek(0)

    if not nav_ok:
        statuses.append(("bookmarks/TOC", "DAMAGED"))
    if ref:
        n_ok = sum(1 for _, s in statuses if s == "UNCHANGED")
        flagged = [f"{f} [{s}]" for f, s in statuses if s not in ("UNCHANGED", "RECORDED")]
        summary = f"{len(extracted)} native(s) extracted, {n_ok} verified UNCHANGED, {nav_text}"
        if flagged:
            summary += " | ATTENTION: " + ", ".join(flagged)
    else:
        summary = f"{len(extracted)} native(s) extracted (no manifest - not verified), {nav_text}"

    return f"{base}.zip", out.read(), extracted, summary


# ---------------------------------------------------------------- BATCH

def _expand_master_zip(zip_bytes: bytes):
    """If ZIP contains only .zip entries, return [(inner_name, inner_bytes)].
    Otherwise return None (it is a normal document ZIP)."""
    try:
        zf = zipfile.ZipFile(io.BytesIO(zip_bytes))
    except zipfile.BadZipFile:
        return None
    entries = [i for i in zf.infolist()
               if not i.is_dir()
               and not os.path.basename(i.filename).startswith(".")
               and "__MACOSX" not in i.filename]
    if not entries:
        return None
    if all(i.filename.lower().endswith(".zip") for i in entries):
        return [(os.path.basename(i.filename), zf.read(i)) for i in entries]
    return None


def process_batch(files, mode):
    """files: list of (filename, bytes). mode: 'embed' or 'extract'.
    Returns (out_filename, out_bytes, results)."""
    # Auto-expand master ZIPs in embed mode
    if mode == "embed":
        expanded = []
        for fname, fbytes in files:
            inner = _expand_master_zip(fbytes)
            if inner is not None:
                expanded.extend(inner)
            else:
                expanded.append((fname, fbytes))
        files = expanded

    outputs, results = [], []
    for fname, fbytes in files:
        try:
            if mode == "embed":
                out_name, out_bytes, items, nav_ok, nav_text = embed_zip_to_pdf(fbytes, fname)
                st = "ok" if nav_ok else "warn"
                detail = f"Embedded {len(items)} native(s) + manifest ({nav_text}): " + ", ".join(items)
                if not nav_ok:
                    detail = "ATTENTION - bookmark/TOC issue | " + detail
                results.append({"file": fname, "status": st, "detail": detail})
            else:
                out_name, out_bytes, items, summary = extract_pdf_to_zip(fbytes, fname)
                st = "ok" if "ATTENTION" not in summary else "warn"
                results.append({"file": fname, "status": st, "detail": summary})
            outputs.append((out_name, out_bytes))
        except NativeToolError as e:
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
    label = "EMBEDDED_PDFS.zip" if mode == "embed" else "EXTRACTED_ZIPS.zip"
    return label, bundle.read(), results
