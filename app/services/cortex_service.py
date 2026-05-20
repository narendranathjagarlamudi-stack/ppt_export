import json


async def ask_cortex(
    conn,
    message: str,
    model: str = "llama3-70b"
):

    cursor = conn.cursor()

    try:

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

        # ============================================
        # DEBUG
        # ============================================

        print("\n========== RAW CORTEX ==========")
        print(response)
        print("================================\n")

        # ============================================
        # STRING RESPONSE
        # ============================================

        if isinstance(response, str):

            return response

        # ============================================
        # DICT RESPONSE
        # ============================================

        if isinstance(response, dict):

            # New Cortex format
            if "choices" in response:

                choices = response.get(
                    "choices",
                    []
                )

                if choices:

                    first = choices[0]

                    # Sometimes:
                    # {"messages":"..."}

                    if isinstance(first, dict):

                        if "messages" in first:

                            return first["messages"]

                        if "message" in first:

                            return first["message"]

            return json.dumps(response)

        return str(response)

    except Exception as e:

        raise Exception(
            f"Cortex COMPLETE failed: {str(e)}"
        )

    finally:

        cursor.close()