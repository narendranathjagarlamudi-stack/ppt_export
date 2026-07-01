# Linux Layout Preview Test Guide

This guide is for testing PowerPoint layout thumbnail previews on a Linux server.

## Purpose

The PPT export can generate `.pptx` files using Python libraries, but layout thumbnail previews need a rendering engine.

On Windows, the POC uses Microsoft PowerPoint if available.

On Linux, the preview renderer should use LibreOffice in headless mode.

If LibreOffice is not available or rendering fails, the app falls back to a simple schematic layout preview.

## 1. Install Required Linux Packages

For Ubuntu/Debian:

```bash
sudo apt-get update
sudo apt-get install -y libreoffice libreoffice-impress fonts-dejavu fonts-liberation
```

Optional, only if the server has display/headless issues:

```bash
sudo apt-get install -y xvfb
```

If the PPT template uses custom company fonts, install those fonts on the server as well. Otherwise thumbnails may render with fallback fonts.

## 2. Verify LibreOffice Headless Mode

Run:

```bash
soffice --headless --version
```

or:

```bash
libreoffice --headless --version
```

Expected result:

```text
LibreOffice x.x.x.x
```

If this command fails, the thumbnail renderer will not be able to use LibreOffice.

## 3. Install Python Dependencies

From the project root:

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
pip install -r linux_preview_requirements.txt
```

If the project uses a different dependency file or internal package process, use that instead.

`linux_preview_requirements.txt` installs `PyMuPDF`, which is used to convert LibreOffice-rendered PDF pages into PNG thumbnails.

## 4. Start the API Server

From the project root:

```bash
source venv/bin/activate
uvicorn app.main:app --host 0.0.0.0 --port 8000
```

If another port is required:

```bash
uvicorn app.main:app --host 0.0.0.0 --port 8080
```

## 5. Open the Frontend Test Page

From a browser that can reach the server:

```text
http://<server-host>:8000/layout-picker
```

Expected result:

- Layout picker page opens.
- Template layout dropdown is visible if a template is available.
- Layout previews should appear inside the dropdown.

## 6. Test Layout API Directly

Run from the Linux server:

```bash
curl "http://localhost:8000/export/layouts?refresh_cache=true"
```

Expected result:

The response should contain layout objects with preview URLs, for example:

```json
{
  "layouts": [
    {
      "index": 0,
      "name": "Title Slide",
      "preview_url": "/static/layout_picker/preview_cache/layout_0.png"
    }
  ]
}
```

Also check these response fields:

```json
{
  "preview_mode": "libreoffice",
  "renderer": "/usr/bin/soffice",
  "renderer_error": null
}
```

If `preview_mode` is `schematic`, LibreOffice rendering did not complete and the app used fallback boxes.

## 7. Verify Thumbnail Cache Files

Check whether preview images were created:

```bash
ls -lh app/static/layout_picker/preview_cache
```

Expected result:

```text
layout_0.png
layout_1.png
layout_2.png
...
```

If files are created and the API returns those image URLs, thumbnail generation is working.

## 8. Confirm LibreOffice Is Actually Being Used

When `refresh_cache=true` is used, the backend should regenerate thumbnails.

If LibreOffice works, previews should look close to the real PowerPoint template.

If LibreOffice fails, previews may still appear, but they will look like plain schematic boxes instead of the real template design.

## 9. Test With Uploaded Template

From the frontend:

1. Open `/layout-picker`.
2. Upload a `.pptx` or `.potx` template.
3. Wait for layout previews to refresh.
4. Check that the dropdown shows layouts from the uploaded template.

Direct API check:

```bash
curl "http://localhost:8000/export/layouts?template_id=<uploaded-template-id>&refresh_cache=true"
```

Expected result:

- Layouts are returned for that uploaded template.
- Preview cache is regenerated.

## 10. Common Issues

### LibreOffice command not found

Check:

```bash
which soffice
which libreoffice
```

Install LibreOffice if missing:

```bash
sudo apt-get install -y libreoffice libreoffice-impress
```

### Previews render but look different from Windows

Likely cause:

- Missing fonts on Linux.
- LibreOffice rendering differences compared to Microsoft PowerPoint.

Fix:

- Install the same fonts used in the PPT template.

### API returns layouts but frontend images are broken

Check static file access:

```text
http://<server-host>:8000/static/layout_picker/preview_cache/layout_0.png
```

If the image does not load, check static file mounting in the FastAPI app.

### Only schematic previews are shown

Likely cause:

- LibreOffice failed.
- `PyMuPDF` is not installed.
- Template conversion failed.
- Cache contains fallback images.

Try:

```bash
pip install -r linux_preview_requirements.txt
curl "http://localhost:8000/export/layouts?refresh_cache=true"
```

Then check server logs for LibreOffice conversion errors.

## 11. Success Criteria

Testing is successful when:

- `soffice --headless --version` works.
- `python -c "import fitz; print(fitz.__doc__[:20])"` works.
- `/export/layouts?refresh_cache=true` returns layout preview URLs.
- API response has `"preview_mode": "libreoffice"`.
- PNG files are created under `app/static/layout_picker/preview_cache`.
- `/layout-picker` shows layout thumbnails in the dropdown.
- Uploaded templates also generate layout previews.
