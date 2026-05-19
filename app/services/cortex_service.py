# app/services/cortex_service.py

import json
import snowflake.connector

from app.database.snowflake_conn import (
    get_snowflake_connection
)


SNOWFLAKE_WAREHOUSE = "ARP_WH_SIT"


async def ask_cortex(
    message: str,
    model: str = "llama3-70b"
):
    """
    Stateless Cortex COMPLETE call.

    Used ONLY for PPT summarization / slide generation.

    Does NOT use:
    - thread memory
    - agent API
    - parent_message_id
    - Cortex Analyst tools

    Input:
        Existing final response text

    Output:
        PPT-ready concise JSON/text
    """

    conn = get_snowflake_connection()

    cursor = conn.cursor()

    try:

        # ============================================
        # SELECT WAREHOUSE
        # ============================================

        cursor.execute(
            f"USE WAREHOUSE {SNOWFLAKE_WAREHOUSE}"
        )

        # ============================================
        # CORTEX COMPLETE
        # ============================================

        query = """
        SELECT SNOWFLAKE.CORTEX.COMPLETE(
            %s,
            %s
        )
        """

        cursor.execute(
            query,
            (
                model,
                message
            )
        )

        result = cursor.fetchone()

        if not result:
            raise Exception(
                "Empty response from Cortex COMPLETE"
            )

        response = result[0]

        # Some Cortex responses may already
        # be plain strings

        if isinstance(response, str):
            return response

        return json.dumps(response)

    except Exception as e:

        raise Exception(
            f"Cortex COMPLETE failed: {str(e)}"
        )

    finally:

        cursor.close()
        conn.close()