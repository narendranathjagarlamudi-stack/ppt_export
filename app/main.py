# app/main.py

from fastapi import (
    FastAPI,
    HTTPException,
    Depends,
)

from fastapi.responses import FileResponse

from sqlalchemy.orm import Session

from pydantic import BaseModel

from app.database.postgres_conn import get_db

from app.services.history_service import (
    fetch_chat_context
)

from app.services.presentation_service import (
    generate_presentation_content
)

from app.services.ppt_service import (
    generate_pptx
)


# ============================================
# FASTAPI APP
# ============================================

app = FastAPI(
    title="PPT Export Service"
)


# ============================================
# ROOT ENDPOINT
# ============================================

@app.get("/")
async def root():

    return {
        "message": "PPT Export Service Running"
    }


# ============================================
# REQUEST MODEL
# ============================================

class PPTExportRequest(BaseModel):

    include_charts: bool = True


# ============================================
# EXPORT PPT ENDPOINT
# ============================================

@app.post("/export/ppt/{dcs_id}/{dih_id}")
async def export_ppt(
    dcs_id: int,
    dih_id: int,
    req: PPTExportRequest,
    db: Session = Depends(get_db),
):

    try:

        # ============================================
        # STEP 1: FETCH CHAT CONTEXT
        # ============================================

        chat_context = fetch_chat_context(
            db=db,
            dcs_id=dcs_id,
            dih_id=dih_id
        )

        if not chat_context:

            raise HTTPException(
                status_code=404,
                detail="Chat history not found"
            )

        # ============================================
        # STEP 2: GENERATE PPT CONTENT
        # ============================================

        presentation_json = (
            await generate_presentation_content(
                agent_id=chat_context.get(
                    "agent_id"
                ),

                thread_id=chat_context.get(
                    "thread_id"
                ),

                original_query=chat_context.get(
                    "query",
                    ""
                ),

                original_response=chat_context.get(
                    "response",
                    {}
                ),
            )
        )

        if not presentation_json:

            raise HTTPException(
                status_code=500,
                detail=(
                    "Failed to generate "
                    "presentation content"
                )
            )

        # ============================================
        # DEBUG LOGS
        # ============================================

        print("\n========== PRESENTATION JSON ==========")
        print(presentation_json)
        print("=======================================\n")

        # ============================================
        # STEP 3: GENERATE PPTX
        # ============================================

        ppt_path = generate_pptx(
            presentation_json=presentation_json,
            include_charts=req.include_charts
        )

        # ============================================
        # STEP 4: RETURN PPT FILE
        # ============================================

        return FileResponse(
            path=ppt_path,

            media_type=(
                "application/vnd.openxmlformats-"
                "officedocument.presentationml."
                "presentation"
            ),

            filename=(
                f"analysis_{dih_id}.pptx"
            )
        )

    except HTTPException:

        raise

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )