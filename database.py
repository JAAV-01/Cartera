# database.py
import os
from urllib.parse import quote_plus
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, declarative_base
from dotenv import load_dotenv

load_dotenv()

DB_USER = os.getenv("DB_USER", "desarrollojosea")
DB_PASS = os.getenv("DB_PASS", "Djosea01*")
DB_HOST = os.getenv("DB_HOST", "192.168.1.14")
DB_PORT = os.getenv("DB_PORT", "1433")
DB_NAME = os.getenv("DB_NAME", "cartera_db")
DB_ENCRYPT = os.getenv("DB_ENCRYPT", "yes")              # yes / no
DB_TRUST = os.getenv("DB_TRUSTSERVERCERT", "yes")        # yes / no

# Cadena ODBC robusta (soporta caracteres especiales)
odbc_str = (
    "Driver={ODBC Driver 18 for SQL Server};"
    f"Server={DB_HOST},{DB_PORT};"
    f"Database={DB_NAME};"
    f"UID={DB_USER};"
    f"PWD={DB_PASS};"
    f"Encrypt={DB_ENCRYPT};"
    f"TrustServerCertificate={DB_TRUST};"
)

SQLALCHEMY_DATABASE_URL = "mssql+pyodbc:///?odbc_connect=" + quote_plus(odbc_str)

engine = create_engine(
    SQLALCHEMY_DATABASE_URL,
    pool_pre_ping=True,
    pool_recycle=1800,
    future=True,
)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


# # database.py (PostgreSQL, robusto contra encoding)
# import os
# from sqlalchemy import create_engine
# from sqlalchemy.engine import URL
# from sqlalchemy.orm import sessionmaker, declarative_base
# from dotenv import load_dotenv

# load_dotenv()

# DB_USER = os.getenv("DB_USER", "postgres")
# DB_PASS = os.getenv("DB_PASS", "")
# DB_HOST = os.getenv("DB_HOST", "localhost")
# DB_PORT = os.getenv("DB_PORT", "5432")
# DB_NAME = os.getenv("DB_NAME", "cartera_db")

# # Construir URL sin concatenar strings (evita problemas de UTF-8 / quoting)
# SQLALCHEMY_DATABASE_URL = URL.create(
#     drivername="postgresql+psycopg2",
#     username=DB_USER,      # SQLAlchemy se encarga del quoting/encoding
#     password=DB_PASS,
#     host=DB_HOST,
#     port=int(DB_PORT),
#     database=DB_NAME,
# )

# engine = create_engine(
#     SQLALCHEMY_DATABASE_URL,
#     pool_pre_ping=True,
#     pool_recycle=1800,
#     future=True,
# )

# SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
# Base = declarative_base()

# def get_db():
#     db = SessionLocal()
#     try:
#         yield db
#     finally:
#         db.close()
