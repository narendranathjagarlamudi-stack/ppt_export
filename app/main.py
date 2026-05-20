from fastapi import (
    FastAPI,
    HTTPException,
    Depends,
)

from fastapi.responses import FileResponse

from sqlalchemy.orm import Session

from pydantic import BaseModel

from typing import List

from app.database.postgres_conn import (
    get_db
)

from app.database.snowflake_conn import (
    get_snowflake_connection
)

from app.services.history_service import (
    fetch_chat_contexts
)

from app.services.presentation_service import (
    generate_presentation_content
)

from app.services.ppt_service import (
    generate_pptx
)


app = FastAPI(
    title="PPT Export Service"
)


@app.get("/")
async def root():

    return {
        "message":
            "PPT Export Service Running"
    }


class PPTExportRequest(BaseModel):

    dih_ids: List[int]

    include_charts: bool = True


@app.post("/export/ppt/{dcs_id}")
async def export_ppt(

    dcs_id: int,

    req: PPTExportRequest,

    db: Session = Depends(get_db),
):

    conn = None

    try:

        # ============================================
        # FETCH CHAT CONTEXTS
        # ============================================

        chat_contexts = fetch_chat_contexts(

            db=db,

            dcs_id=dcs_id,

            dih_ids=req.dih_ids
        )

        if not chat_contexts:

            raise HTTPException(
                status_code=404,
                detail="No chat history found"
            )

        # ============================================
        # SINGLE SNOWFLAKE CONNECTION
        # ============================================

        conn = get_snowflake_connection()

        slides = []

        # ============================================
        # GENERATE SLIDES
        # ============================================

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

            slides.append(slide_json)

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

        ppt_path = generate_pptx(

            presentation_json=presentation_json,

            include_charts=req.include_charts
        )

        return FileResponse(

            path=ppt_path,

            media_type=(
                "application/vnd.openxmlformats-"
                "officedocument.presentationml."
                "presentation"
            ),

            filename="analysis_export.pptx"
        )

    except HTTPException:

        raise

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )

    finally:

        if conn:

            conn.close()