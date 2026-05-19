# services/history_service.py

from sqlalchemy.orm import Session
from sqlalchemy import text
from typing import Optional, Dict, Any


def fetch_chat_context(
    db: Session,
    dih_id: int,
    dcs_id: int,
) -> Optional[Dict[str, Any]]:
    """
    Fetch complete chat context required for PPT export.

    Returns:
    {
        "dih_id": int,
        "dcs_id": int,
        "query": str,
        "response": dict,
        "message_id": int,
        "thread_id": str,
        "agent_id": str
    }
    """

    query = text("""
        SELECT
            h.dih_id,
            h.dcs_id,
            h.query,
            h.response,
            h.message_id,
            s.thread_id,
            s.agent_name AS agent_id
        FROM "V2".di_chat_history h
        JOIN "V2".di_chat_session s
            ON h.dcs_id = s.dcs_id
        WHERE h.dih_id = :dih_id AND h.dcs_id = :dcs_id
        LIMIT 1
    """)

    result = db.execute(
        query,
        {"dih_id": dih_id, "dcs_id": dcs_id}
    ).mappings().first()

    if not result:
        return None

    return {
        "dih_id": result["dih_id"],
        "dcs_id": result["dcs_id"],
        "query": result["query"],
        "response": result["response"],
        "message_id": result["message_id"],
        "thread_id": result["thread_id"],
        "agent_id": result["agent_id"],
    }