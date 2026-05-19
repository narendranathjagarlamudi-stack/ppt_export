from database.snowflake_conn import get_snowflake_connection


def test_connection():

    conn = get_snowflake_connection()

    cursor = conn.cursor()

    try:

        cursor.execute("SELECT CURRENT_USER(), CURRENT_ROLE(), CURRENT_VERSION()")

        result = cursor.fetchone()

        print("\n✅ CONNECTED TO SNOWFLAKE\n")

        print(f"USER: {result[0]}")
        print(f"ROLE: {result[1]}")
        print(f"VERSION: {result[2]}")

    finally:
        cursor.close()
        conn.close()


if __name__ == "__main__":
    test_connection()