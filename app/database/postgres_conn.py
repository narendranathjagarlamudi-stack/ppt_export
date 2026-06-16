# database/postgres_conn.py

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, declarative_base
from sqlalchemy import text
import os
from urllib.parse import quote_plus

POSTGRES_USER = os.getenv("POSTGRES_USER", "arp_dev_1")
POSTGRES_PASSWORD = quote_plus(os.getenv("POSTGRES_PASSWORD", "arpD!9@1$4"))
POSTGRES_HOST = os.getenv("POSTGRES_HOST", "10.23.16.44")
POSTGRES_PORT = os.getenv("POSTGRES_PORT", "5432")
POSTGRES_DB = os.getenv("POSTGRES_DB", "arp_d2i")

DATABASE_URL = (
    f"postgresql+psycopg2://"
    f"{POSTGRES_USER}:{POSTGRES_PASSWORD}"
    f"@{POSTGRES_HOST}:{POSTGRES_PORT}/{POSTGRES_DB}"
)

engine = create_engine(
    DATABASE_URL,
    pool_pre_ping=True,
    pool_size=5,
    max_overflow=10,
)

SessionLocal = sessionmaker(
    autocommit=False,
    autoflush=False,
    bind=engine
)

Base = declarative_base()


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


if __name__ == "__main__":

    db = None
    try:
        db = SessionLocal()

        result = db.execute(text("SELECT 1"))

        print("✅ PostgreSQL connection successful")

        for row in result:
            print(row)

    except Exception as e:
        print("❌ Connection failed")
        print(str(e))

    finally:
        if db:
            db.close()