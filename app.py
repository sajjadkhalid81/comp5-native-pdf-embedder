"""
app.py - COMP5 Native Embedder web app.
Three operations:
  EMBED  : ZIP (PDF + natives)  ->  PDF with natives embedded
  EXTRACT: PDF (with natives)   ->  ZIP (clean PDF + natives + log)
  CRS    : PDF(s) (no CRS supplied) -> ZIP per doc (untouched PDF + filled CTR CRS)
Streaming responses - no in-memory job store.
"""
import io
import json
import base64
from urllib.parse import quote
from flask import Flask, render_template, request, jsonify, send_file

from native_core import process_batch
from crs_core import process_batch as crs_process_batch, RETURN_CODES

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 300 * 1024 * 1024  # 300 MB


@app.route("/")
def index():
    return render_template("index.html")


def _read_uploads():
    files = []
    for f in request.files.getlist("files"):
        if f and f.filename:
            files.append((f.filename, f.read()))
    return files


@app.route("/api/process", methods=["POST"])
def process():
    mode = request.form.get("mode", "")
    if mode not in ("embed", "extract"):
        return jsonify({"error": "Invalid mode."}), 400

    files = _read_uploads()
    if not files:
        return jsonify({"error": "No files uploaded."}), 400

    for fname, _ in files:
        low = fname.lower()
        if mode == "embed" and not low.endswith(".zip"):
            return jsonify({"error": f"{fname}: EMBED mode expects ZIP files only."}), 400
        if mode == "extract" and not low.endswith(".pdf"):
            return jsonify({"error": f"{fname}: EXTRACT mode expects PDF files only."}), 400

    out_name, out_bytes, results = process_batch(files, mode)

    if out_bytes is None:
        return jsonify({"error": "All files failed.", "results": results}), 422

    resp = send_file(
        io.BytesIO(out_bytes),
        as_attachment=True,
        download_name=out_name,
        mimetype="application/octet-stream",
    )
    # compact results summary in a header (base64 so it is header-safe)
    compact = [{"file": r["file"][:80], "status": r["status"],
                "detail": r["detail"][:150]} for r in results[:50]]
    if len(results) > 50:
        compact.append({"file": "...", "status": "ok",
                        "detail": f"and {len(results) - 50} more file(s)"})
    payload = json.dumps(compact)
    resp.headers["X-Results"] = base64.b64encode(payload.encode("utf-8")).decode("ascii")
    resp.headers["X-Filename"] = quote(out_name)
    resp.headers["Access-Control-Expose-Headers"] = "X-Results, X-Filename"
    return resp


@app.route("/api/crs", methods=["POST"])
def crs():
    return_code = request.form.get("return_code", "").strip().upper()
    reviewer_group = request.form.get("reviewer_group", "").strip()

    if return_code not in RETURN_CODES:
        return jsonify({"error": "Please select a return code."}), 400

    files = _read_uploads()
    if not files:
        return jsonify({"error": "No files uploaded."}), 400

    for fname, _ in files:
        if not fname.lower().endswith(".pdf"):
            return jsonify({"error": f"{fname}: CRS mode expects PDF files only."}), 400

    out_name, out_bytes, results = crs_process_batch(files, return_code, reviewer_group)

    if out_bytes is None:
        return jsonify({"error": "All files failed.", "results": results}), 422

    resp = send_file(
        io.BytesIO(out_bytes),
        as_attachment=True,
        download_name=out_name,
        mimetype="application/octet-stream",
    )
    compact = [{"file": r["file"][:80], "status": r["status"],
                "detail": r["detail"][:200]} for r in results[:50]]
    if len(results) > 50:
        compact.append({"file": "...", "status": "ok",
                        "detail": f"and {len(results) - 50} more file(s)"})
    payload = json.dumps(compact)
    resp.headers["X-Results"] = base64.b64encode(payload.encode("utf-8")).decode("ascii")
    resp.headers["X-Filename"] = quote(out_name)
    resp.headers["Access-Control-Expose-Headers"] = "X-Results, X-Filename"
    return resp


if __name__ == "__main__":
    app.run(debug=True, port=5000)
