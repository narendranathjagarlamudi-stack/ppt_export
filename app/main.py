from fastapi import (
    FastAPI,
    HTTPException,
    Depends,
)

from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
import base64
import re
import uuid
from datetime import datetime
from pathlib import Path

from sqlalchemy.orm import Session

from pydantic import BaseModel, Field

from typing import Any, Dict, List, Optional

from app.database.postgres_conn import (
    get_db
)

from app.services.history_service import (
    fetch_chat_contexts
)

from app.services.presentation_service import (
    generate_presentation_content
)

from app.services.ppt_service import (
    PPTXGenerator
)

from app.services.layout_service import (
    get_layouts as get_template_layouts
)


LAYOUT_PICKER_STATIC_DIR = (
    Path(__file__).resolve().parents[1]
    / "app"
    / "static"
    / "layout_picker"
)

UPLOADED_TEMPLATE_DIR = (
    Path(__file__).resolve().parents[1]
    / "app"
    / "uploads"
    / "templates"
)

ALLOWED_TEMPLATE_SUFFIXES = {
    ".pptx",
    ".potx",
}


app = FastAPI(
    title="PPT Export Service"
)

app.mount(
    "/layout-static",
    StaticFiles(directory=LAYOUT_PICKER_STATIC_DIR),
    name="layout-static"
)


@app.get("/health")
async def health():

    return {
        "message":
            "PPT Export Service Running"
    }


@app.get("/")
async def root():

    return RedirectResponse(
        url="/layout-picker"
    )


class ExportItem(BaseModel):

    dcs_id: int

    dih_ids: List[int] = Field(default_factory=list)


class PPTExportRequest(BaseModel):

    dih_ids: List[int] = Field(default_factory=list)

    export_items: Optional[List[ExportItem]] = None

    include_charts: bool = True

    ppt_config: Optional[Dict[str, Any]] = None

    chart_colors: Optional[Dict[str, Any]] = None

    slide_layouts: Optional[Dict[str, Any]] = None

    selected_colors: Optional[List[str]] = None

    selected_slide_indices: Optional[List[int]] = None

    template_id: Optional[str] = None

    filename: Optional[str] = None


class PPTPreviewRequest(BaseModel):

    dih_ids: List[int]

    export_items: Optional[List[ExportItem]] = None


class TemplateUploadRequest(BaseModel):

    filename: str

    content_base64: str


def _safe_filename(value, default_name):

    cleaned = re.sub(
        r"[^A-Za-z0-9_.-]+",
        "_",
        str(value or "").strip()
    ).strip("._")

    return cleaned or default_name


def _resolve_template_path(template_id: Optional[str]):

    if not template_id:
        return None

    safe_template_id = _safe_filename(
        template_id,
        ""
    )

    if not safe_template_id:
        raise HTTPException(
            status_code=400,
            detail="Invalid template_id"
        )

    for suffix in ALLOWED_TEMPLATE_SUFFIXES:
        candidate = UPLOADED_TEMPLATE_DIR / f"{safe_template_id}{suffix}"

        if candidate.exists():
            return candidate

    raise HTTPException(
        status_code=404,
        detail=f"Uploaded template not found: {template_id}"
    )


def _request_items(
    route_dcs_id: int,
    dih_ids: List[int],
    export_items: Optional[List[ExportItem]]
):

    if export_items:
        return export_items

    return [
        ExportItem(
            dcs_id=route_dcs_id,
            dih_ids=dih_ids
        )
    ]


@app.get("/layout-picker", response_class=HTMLResponse)
async def layout_picker():

    html = (
        LAYOUT_PICKER_STATIC_DIR
        / "index.html"
    ).read_text(
        encoding="utf-8"
    )

    return html.replace(
        "/static/",
        "/layout-static/"
    )


@app.get("/export/layouts")
async def export_layouts(
    template_id: Optional[str] = None,
    refresh_cache: bool = False
):

    return get_template_layouts(
        template_path=_resolve_template_path(template_id),
        force_refresh=refresh_cache
    )


@app.post("/export/templates")
async def upload_template(req: TemplateUploadRequest):

    suffix = Path(req.filename).suffix.lower()

    if suffix not in ALLOWED_TEMPLATE_SUFFIXES:
        raise HTTPException(
            status_code=400,
            detail="Template must be a .pptx or .potx file"
        )

    try:
        file_bytes = base64.b64decode(
            req.content_base64,
            validate=True
        )

    except Exception:
        raise HTTPException(
            status_code=400,
            detail="Invalid base64 template content"
        )

    if not file_bytes:
        raise HTTPException(
            status_code=400,
            detail="Template file is empty"
        )

    UPLOADED_TEMPLATE_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    template_id = uuid.uuid4().hex
    template_path = UPLOADED_TEMPLATE_DIR / f"{template_id}{suffix}"

    template_path.write_bytes(
        file_bytes
    )

    return {
        "template_id": template_id,
        "filename": req.filename,
        "template_path": str(template_path)
    }


async def _build_slides_from_history(
    db: Session,
    dcs_id: int,
    dih_ids: List[int],
    slide_layouts: Optional[Dict[str, Any]] = None,
    selected_slide_indices: Optional[List[int]] = None,
):

    chat_contexts = fetch_chat_contexts(

        db=db,

        dcs_id=dcs_id,

        dih_ids=dih_ids
    )

    if not chat_contexts:

        raise HTTPException(
            status_code=404,
            detail="No chat history found"
        )

    slides = []
    slide_metadata = []
    skipped_messages = []
    slide_index = 0
    conn = None

    selected_index_set = (
        set(selected_slide_indices)
        if selected_slide_indices is not None
        else None
    )

    for ctx in chat_contexts:

        slide_json = (
            await generate_presentation_content(

                conn=conn,

                original_query=ctx.get(
                    "query",
                    ""
                ),

                original_response=ctx.get(
                    "response",
                    {}
                ),
            )
        )

        generated_slides = (
            slide_json
            if isinstance(slide_json, list)
            else [slide_json]
        )

        chart_slides = [
            generated_slide
            for generated_slide in generated_slides
            if generated_slide.get("chart_spec")
        ]

        if not chart_slides:

            skipped_messages.append({
                "dih_id": ctx.get("dih_id"),
                "message_id": ctx.get("message_id"),
                "query": ctx.get("query", ""),
                "reason": "No chart found for this message.",
            })

            continue

        for generated_slide in chart_slides:

            current_slide_index = slide_index
            slide_index += 1

            if (
                selected_index_set is not None
                and current_slide_index not in selected_index_set
            ):
                continue

            if slide_layouts:

                selected_layout = None

                for layout_key in (
                    str(current_slide_index),
                    current_slide_index,
                    "default"
                ):

                    if layout_key in slide_layouts:
                        selected_layout = slide_layouts[layout_key]
                        break

                if selected_layout is not None:
                    generated_slide["layout"] = selected_layout

            slides.append(generated_slide)
            slide_metadata.append({
                "index": current_slide_index,
                "dcs_id": ctx.get("dcs_id"),
                "dih_id": ctx.get("dih_id"),
                "message_id": ctx.get("message_id"),
                "query": ctx.get("query", ""),
                "title": (
                    generated_slide
                    .get("slide", {})
                    .get("title", "Key Insights")
                ),
                "bullet_count": len(
                    generated_slide
                    .get("slide", {})
                    .get("bullets", [])
                ),
                "has_chart": bool(
                    generated_slide.get("chart_spec")
                ),
                "event_metadata": generated_slide.get(
                    "event_metadata",
                    {}
                ),
            })

    return slides, slide_metadata, skipped_messages


async def _build_slides_from_request_items(
    db: Session,
    request_items: List[ExportItem],
    slide_layouts: Optional[Dict[str, Any]] = None,
    selected_slide_indices: Optional[List[int]] = None,
):

    all_slides = []
    all_metadata = []
    all_skipped = []
    global_slide_index = 0

    for item in request_items:
        slides, metadata, skipped = await _build_slides_from_history(
            db=db,
            dcs_id=item.dcs_id,
            dih_ids=item.dih_ids,
            slide_layouts=None,
            selected_slide_indices=None
        )

        for slide, meta in zip(slides, metadata):
            current_global_index = global_slide_index
            global_slide_index += 1
            source_index = meta.get("index")

            if (
                selected_slide_indices is not None
                and current_global_index not in selected_slide_indices
            ):
                continue

            if slide_layouts:
                selected_layout = None

                for layout_key in (
                    str(current_global_index),
                    current_global_index,
                    "default"
                ):
                    if layout_key in slide_layouts:
                        selected_layout = slide_layouts[layout_key]
                        break

                if selected_layout is not None:
                    slide["layout"] = selected_layout

            meta = dict(meta)
            meta["index"] = current_global_index
            meta["source_index"] = source_index
            all_slides.append(slide)
            all_metadata.append(meta)

        all_skipped.extend(skipped)

    return all_slides, all_metadata, all_skipped


@app.post("/export/preview/{dcs_id}")
async def export_preview(

    dcs_id: int,

    req: PPTPreviewRequest,

    db: Session = Depends(get_db),
):

    slides, slide_metadata, skipped_messages = await _build_slides_from_request_items(
        db=db,
        request_items=_request_items(
            route_dcs_id=dcs_id,
            dih_ids=req.dih_ids,
            export_items=req.export_items
        )
    )

    return {
        "slide_count": len(slides),
        "slides": slide_metadata,
        "skipped_count": len(skipped_messages),
        "skipped_messages": skipped_messages
    }


@app.post("/export/ppt/{dcs_id}")
async def export_ppt(

    dcs_id: int,

    req: PPTExportRequest,

    db: Session = Depends(get_db),
):

    try:
        request_items = _request_items(
            route_dcs_id=dcs_id,
            dih_ids=req.dih_ids,
            export_items=req.export_items
        )

        slides, slide_metadata, skipped_messages = await _build_slides_from_request_items(
            db=db,
            request_items=request_items,
            slide_layouts=req.slide_layouts,
            selected_slide_indices=req.selected_slide_indices
        )

        if not slides:

            raise HTTPException(
                status_code=400,
                detail={
                    "message": "No chart slides found for the selected message IDs.",
                    "skipped_messages": skipped_messages
                }
            )

        # ============================================
        # FINAL JSON
        # ============================================

        presentation_json = {

            "presentation_title":
                "Analysis Summary",

            "slides":
                slides
        }

        # ============================================
        # GENERATE PPT
        # ============================================

        template_path = _resolve_template_path(
            req.template_id
        )

        generator = PPTXGenerator(
            template_path=template_path
            if template_path
            else PPTXGenerator().template_path
        )

        ppt_path = generator.generate(

            presentation_json=presentation_json,

            include_charts=req.include_charts,

            ppt_config=(
                req.ppt_config
                or req.chart_colors
            ),

            selected_colors=req.selected_colors
        )

        dcs_label = (
            str(request_items[0].dcs_id)
            if len(request_items) == 1
            else "multi_session"
        )
        default_filename = (
            f"analysis_export_{dcs_label}_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        )
        filename = _safe_filename(
            req.filename,
            default_filename
        )

        if not filename.lower().endswith(".pptx"):
            filename = f"{filename}.pptx"

        return FileResponse(

            path=ppt_path,

            media_type=(
                "application/vnd.openxmlformats-"
                "officedocument.presentationml."
                "presentation"
            ),

            filename=filename
        )

    except HTTPException:

        raise

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )

