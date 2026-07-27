# COMP5 Native Embedder

Flask web tool for the COMP5 native-file workflow.

- **EMBED (ZIP -> PDF):** vendor ZIP (PDF + natives) -> PDF with natives
  embedded as attachments, ready for Aconex upload.
- **EXTRACT (PDF -> ZIP):** commented PDF -> ZIP with clean PDF + natives +
  verification log, ready to return to vendor and COMPANY.

Includes checksum manifest verification, bookmark/TOC integrity checking,
master-ZIP auto-detection, and batch processing.

See **DEPLOYMENT_GUIDE.md** for the complete setup and usage guide.

## Quick start (local)
    pip install -r requirements.txt
    python app.py
    -> http://127.0.0.1:5000

## Deploy
Push to a private GitHub repo -> Render -> New Web Service -> connect repo.
Build: `pip install -r requirements.txt`. Start command comes from Procfile.
