import os
import psycopg2
from psycopg2.extras import RealDictCursor
from datetime import datetime
from models import MacroCategory

def get_db_connection():
    return psycopg2.connect(
        host=os.getenv("PGHOST"),
        database=os.getenv("PGDATABASE"),
        user=os.getenv("PGUSER"),
        password=os.getenv("PGPASSWORD"),
        port=os.getenv("PGPORT")
    )

def init_db():
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute('''
        CREATE TABLE IF NOT EXISTS macros (
            id SERIAL PRIMARY KEY,
            title VARCHAR(255),
            description TEXT NOT NULL,
            category VARCHAR(50) DEFAULT 'AI_GENERATED',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    conn.commit()
    cur.close()
    conn.close()

def save_macro(title: str, description: str, category: MacroCategory = MacroCategory.AI_GENERATED) -> int:
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute(
        "INSERT INTO macros (title, description, category) VALUES (%s, %s, %s) RETURNING id",
        (title, description, category)
    )

    macro_id = cur.fetchone()[0]
    conn.commit()
    cur.close()
    conn.close()

    return macro_id

def get_macro_by_id(macro_id: int):
    conn = get_db_connection()
    cur = conn.cursor(cursor_factory=RealDictCursor)

    cur.execute("SELECT * FROM macros WHERE id = %s", (macro_id,))
    macro = cur.fetchone()

    cur.close()
    conn.close()

    return macro

def get_all_macros():
    conn = get_db_connection()
    cur = conn.cursor(cursor_factory=RealDictCursor)

    cur.execute("SELECT * FROM macros ORDER BY created_at DESC")
    macros = cur.fetchall()

    cur.close()
    conn.close()

    return macros

# DB初期化
init_db()