import os
import snowflake.connector




def get_snowflake_connection():

    conn = snowflake.connector.connect(
        account = "MERKLE-ARP",
        user = "NJAGARLAMUDI@MERKLEINC.COM",
        authenticator = "externalbrowser",
        role = "ARP_SNOWFLAKE_AI",
        warehouse = "<none selected>",
        database = "ARP_SURVEY_SIT",
        schema = "MR"
    )

    return conn