import snowflake.connector


def get_snowflake_connection():

    conn = snowflake.connector.connect(
        account="MERKLE-ARP",
        user="NJAGARLAMUDI@MERKLEINC.COM",
        authenticator="externalbrowser",
        role="ARP_SNOWFLAKE_AI",
        warehouse="ARP_WH_SIT",
        database="ARP_SURVEY_SIT",
        schema="MR",
        client_session_keep_alive=True
    )

    return conn