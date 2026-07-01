# Linux Layout Preview Setup Guide

This guide explains how to set up a Linux server so PowerPoint layout thumbnails are rendered with LibreOffice instead of the generic schematic fallback.

The same steps work for Ubuntu/Debian on ARM64 or AMD64. Use the normal OS packages for the server architecture.

## Expected Result

After setup, this API call:

```bash
curl "http://localhost:8000/export/layouts?refresh_cache=true"
```

should return:

```json
{
  "preview_mode": "libreoffice",
  "renderer": "/usr/bin/soffice",
  "renderer_error": null
}
```

Each layout should also contain a PNG data URL:

```json
{
  "preview": "data:image/png;base64,..."
}
```

If `preview_mode` is `schematic`, LibreOffice rendering failed and the app used the safe fallback.

## 1. Update The Server

```bash
sudo apt-get update
sudo apt-get upgrade -y
```

## 2. Install System Packages

Install Python, LibreOffice Impress, font packages, and basic build/runtime tools:

```bash
sudo apt-get install -y \
  python3 \
  python3-venv \
  python3-pip \
  libreoffice \
  libreoffice-impress \
  fonts-dejavu \
  fonts-liberation \
  fonts-crosextra-caladea \
  fonts-crosextra-carlito \
  fontconfig \
  curl
```

Optional, only if the server has unusual headless/display issues:

```bash
sudo apt-get install -y xvfb
```

If the PowerPoint template uses company fonts, install those fonts on the server too. Missing fonts are the most common reason Linux previews look different from Windows PowerPoint previews.

After adding custom fonts, rebuild the font cache:

```bash
sudo fc-cache -f -v
```

## 3. Verify LibreOffice

Check that the LibreOffice command exists:

```bash
which soffice
```

Expected:

```text
/usr/bin/soffice
```

Check headless mode:

```bash
soffice --headless --version
```

Expected:

```text
LibreOffice x.x.x.x
```

If `soffice` is not found, try:

```bash
which libreoffice
libreoffice --headless --version
```

The app will auto-detect `soffice` or `libreoffice`. You can also force the path with:

```bash
export LIBREOFFICE_PATH=/usr/bin/soffice
```

## 4. Get The Project Onto The Server

Clone or copy the project to the Linux server, then enter the project root:

```bash
cd /path/to/ppt_export
```

Confirm these files exist:

```bash
ls -lh app/main.py
ls -lh app/services/layout_service.py
ls -lh "ppt template export/merkle_template.potx"
```

If the template file is missing, the layout preview API cannot render the real template layouts.

## 5. Create A Python Virtual Environment

```bash
python3 -m venv venv
source venv/bin/activate
python -m pip install --upgrade pip
```

## 6. Install Python Dependencies

For this POC's Linux preview path, install the runtime dependencies:

```bash
pip install \
  fastapi==0.115.6 \
  "uvicorn[standard]==0.34.0" \
  python-pptx==1.0.2 \
  Pillow==11.1.0 \
  PyMuPDF==1.25.1 \
  SQLAlchemy==2.0.36 \
  psycopg2-binary==2.9.10 \
  pydantic==2.10.4
```

`PyMuPDF` provides the `fitz` package, which converts LibreOffice-rendered PDF pages into PNG thumbnails.

If you choose to use the full `requirements.txt`, first verify it is plain UTF-8 text. In this POC snapshot it may be UTF-16, which can confuse `pip` on Linux. The safer preview-focused install is the command above.

Verify the important Python imports:

```bash
python -c "import fastapi, pptx, PIL, fitz; print('imports ok')"
```

Expected:

```text
imports ok
```

## 7. Start The API Server

From the project root:

```bash
source venv/bin/activate
export LIBREOFFICE_PATH=/usr/bin/soffice
uvicorn app.main:app --host 0.0.0.0 --port 8000
```

For a different port:

```bash
uvicorn app.main:app --host 0.0.0.0 --port 8080
```

In production, run the command through your normal process manager, such as `systemd`, supervisor, or your deployment platform.

## 8. Check The Health Endpoint

From another terminal on the server:

```bash
curl "http://localhost:8000/health"
```

Expected:

```json
{"message":"PPT Export Service Running"}
```

## 9. Test LibreOffice Layout Rendering

Force a fresh render:

```bash
curl "http://localhost:8000/export/layouts?refresh_cache=true"
```

For a quick readable check:

```bash
curl -s "http://localhost:8000/export/layouts?refresh_cache=true" \
  | python -m json.tool \
  | grep -E '"preview_mode"|"renderer"|"renderer_error"'
```

Expected:

```text
"preview_mode": "libreoffice",
"renderer": "/usr/bin/soffice",
"renderer_error": null,
```

If the server has the smoke test script:

```bash
python3 scripts/smoke_layout_previews.py http://localhost:8000
```

Expected:

```text
preview_mode=libreoffice
renderer=/usr/bin/soffice
layout_count=35
LibreOffice layout preview smoke test passed.
```

The exact `layout_count` can change if the template changes.

## 10. Open The Frontend

From a browser that can reach the server:

```text
http://<server-host>:8000/layout-picker
```

Expected:

- The layout picker page opens.
- The template layout dropdown is visible.
- Layout thumbnails appear in the dropdown.
- The preview status says LibreOffice rendered the previews.

## 11. Verify Preview Cache

The app stores rendered previews as JSON cache files containing PNG data URLs.

Check:

```bash
ls -lh app/static/layout_picker/preview_cache
```

Expected:

```text
<template-cache-key>.json
```

If you need to force regeneration, call:

```bash
curl "http://localhost:8000/export/layouts?refresh_cache=true"
```

## 12. Test Uploaded Templates

From the frontend:

1. Open `/layout-picker`.
2. Upload a `.pptx` or `.potx` template.
3. Wait for layout previews to refresh.
4. Confirm the dropdown shows the uploaded template's layouts.

Direct API check:

```bash
curl "http://localhost:8000/export/layouts?template_id=<uploaded-template-id>&refresh_cache=true"
```

Expected:

- Layouts are returned for the uploaded template.
- `preview_mode` is `libreoffice`.
- `renderer_error` is `null`.

## 13. Common Issues

### LibreOffice command not found

Check:

```bash
which soffice
which libreoffice
```

Install again if missing:

```bash
sudo apt-get update
sudo apt-get install -y libreoffice libreoffice-impress
```

### PyMuPDF is missing

Check:

```bash
python -c "import fitz; print(fitz.__doc__[:20])"
```

Install:

```bash
source venv/bin/activate
pip install PyMuPDF==1.25.1
```

### API returns `preview_mode: schematic`

Likely causes:

- LibreOffice is not installed or not found.
- `PyMuPDF` is not installed.
- The template file is missing or invalid.
- LibreOffice failed to convert the preview deck to PDF.

Check the API response:

```bash
curl -s "http://localhost:8000/export/layouts?refresh_cache=true" \
  | python -m json.tool \
  | grep -E '"preview_mode"|"renderer_error"'
```

Then check the uvicorn server logs for the detailed LibreOffice error.

### Previews work but look different from Windows

Likely causes:

- Missing fonts on Linux.
- LibreOffice rendering differences from Microsoft PowerPoint.
- PowerPoint-only effects or theme behavior.

Fix:

```bash
sudo fc-match Arial
sudo fc-match Calibri
sudo fc-cache -f -v
```

Install the exact fonts used by the template wherever licensing permits.

## 14. Success Checklist

Setup is complete when all of these are true:

- `soffice --headless --version` works.
- `python -c "import fastapi, pptx, PIL, fitz; print('imports ok')"` works.
- `uvicorn app.main:app --host 0.0.0.0 --port 8000` starts.
- `/health` returns `PPT Export Service Running`.
- `/export/layouts?refresh_cache=true` returns `"preview_mode": "libreoffice"`.
- `renderer_error` is `null`.
- Layout previews start with `data:image/png;base64,`.
- `/layout-picker` shows rendered layout thumbnails.
