# PPT Export POC

This POC exports analysis responses into PowerPoint using stored message/session history.

## What Is Included

- FastAPI backend endpoints for layout preview, slide preview, and PPT export.
- Integrated browser UI at `/layout-picker`.
- Heuristic bullet extraction instead of a Snowflake Cortex inference call.
- Support for old final-text responses and newer streamed event responses.
- Multiple chart/event slides from one message ID.
- Chart-level selection from the preview list.
- Multi-session export using `export_items`.
- Per-slide template layout selection when a PPT template is available.
- Template upload support for user-specific templates.
- Config JSON upload support for fallback branding.
- Fallback config mode when no template is available.
- Template palette extraction with black/white filtered out.
- Reusable `PPTXGenerator` class for future integrations.

## Main API Endpoints

- `GET /layout-picker` - browser UI.
- `GET /export/layouts` - returns template layouts, thumbnails, and palette colors.
- `POST /export/preview/{dcs_id}` - previews generated slide count from message IDs.
- `POST /export/ppt/{dcs_id}` - exports the final PPTX.
- `POST /export/templates` - uploads a `.pptx` or `.potx` template and returns a `template_id`.

## Multi-Session Payload

```json
{
  "export_items": [
    {
      "dcs_id": 156,
      "dih_ids": [436, 437]
    },
    {
      "dcs_id": 157,
      "dih_ids": [501]
    }
  ],
  "selected_slide_indices": [0, 2],
  "include_charts": true
}
```

## Class Usage

```python
from app.services.ppt_service import PPTXGenerator

generator = PPTXGenerator(
    template_path="ppt template export/merkle_template.potx",
    config_path="app/config/chart_colors.json",
)

ppt_path = generator.generate(
    presentation_json=presentation_json,
    include_charts=True,
    ppt_config=None,
    selected_colors=None,
)
```

The existing `generate_pptx(...)` function is still available as a backward-compatible wrapper.

## Run Locally

```powershell
.\venv\Scripts\python.exe -m uvicorn app.main:app --reload
```

Open:

```text
http://127.0.0.1:8000/layout-picker
```

## Notes

- If `ppt template export/merkle_template.potx` exists, template layouts are used for design.
- If the template is missing, layout selection is skipped and fallback config is used.
- PowerPoint on Windows or LibreOffice on Linux can render layout thumbnails; otherwise schematic previews are used.
