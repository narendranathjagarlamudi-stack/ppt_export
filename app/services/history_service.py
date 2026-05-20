# services/history_service.py

from sqlalchemy.orm import Session
from sqlalchemy import text
from typing import List, Dict, Any


def fetch_chat_contexts(
    db: Session,
    dcs_id: int,
    dih_ids: List[int],
) -> List[Dict[str, Any]]:

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
        WHERE h.dcs_id = :dcs_id
        AND h.dih_id = ANY(:dih_ids)
        ORDER BY h.created_on
    """)

    results = db.execute(
        query,
        {
            "dcs_id": dcs_id,
            "dih_ids": dih_ids
        }
    ).mappings().all()

    output = []

    for result in results:

        output.append({
            "dih_id": result["dih_id"],
            "dcs_id": result["dcs_id"],
            "query": result["query"],
            "response": result["response"],
            "message_id": result["message_id"],
            "thread_id": result["thread_id"],
            "agent_id": result["agent_id"],
        })

    return output