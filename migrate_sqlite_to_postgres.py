import os
import sqlite3
import psycopg2

SQLITE_DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "instance", "wtw_academy.db")
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
if not DATABASE_URL: raise SystemExit("DATABASE_URL is not set.")
TABLES=["admins","students","subjects","questions","batches","batch_students","exams","exam_students","exam_batches","exam_attempts","exam_answers","payments"]
def sqlite_columns(conn, table): return [r[1] for r in conn.execute(f'PRAGMA table_info("{table}")').fetchall()]
def pg_table_exists(pg, table):
    with pg.cursor() as cur:
        cur.execute("SELECT EXISTS (SELECT 1 FROM information_schema.tables WHERE table_schema='public' AND table_name=%s)",(table,)); return cur.fetchone()[0]
def pg_type(t):
    t=(t or "TEXT").upper()
    if "INT" in t: return "INTEGER"
    if "REAL" in t or "FLOA" in t: return "DOUBLE PRECISION"
    if "NUM" in t or "DEC" in t: return "NUMERIC"
    return "TEXT"
def create_table(pg, sqlite, table):
    if pg_table_exists(pg,table): return
    info=sqlite.execute(f'PRAGMA table_info("{table}")').fetchall()
    if not info: return
    defs=[f'"{c[1]}" SERIAL PRIMARY KEY' if c[1]=="id" else f'"{c[1]}" {pg_type(c[2])}' for c in info]
    with pg.cursor() as cur: cur.execute(f'CREATE TABLE "{table}" ({", ".join(defs)})')
    pg.commit()
def add_missing_columns(pg, sqlite, table):
    if not pg_table_exists(pg,table): return
    info=sqlite.execute(f'PRAGMA table_info("{table}")').fetchall()
    with pg.cursor() as cur:
        cur.execute("SELECT column_name FROM information_schema.columns WHERE table_schema='public' AND table_name=%s",(table,))
        existing={r[0] for r in cur.fetchall()}
        for c in info:
            if c[1] not in existing: cur.execute(f'ALTER TABLE "{table}" ADD COLUMN "{c[1]}" {pg_type(c[2])}')
    pg.commit()
def migrate_table(pg, sqlite, table):
    cols=sqlite_columns(sqlite,table)
    if not cols: return
    rows=sqlite.execute(f'SELECT * FROM "{table}"').fetchall()
    print(f"{table}: {len(rows)} SQLite rows")
    if not rows: return
    qc=", ".join(f'"{c}"' for c in cols); ph=", ".join(["%s"]*len(cols))
    sql=f'INSERT INTO "{table}" ({qc}) VALUES ({ph}) ON CONFLICT DO NOTHING'
    with pg.cursor() as cur:
        for row in rows: cur.execute(sql,tuple(row))
    pg.commit()
def reset_sequences(pg):
    for table in TABLES:
        if not pg_table_exists(pg,table): continue
        with pg.cursor() as cur:
            cur.execute("SELECT pg_get_serial_sequence(%s,'id')",(f"public.{table}",)); r=cur.fetchone()
            if not r or not r[0]: continue
            cur.execute(f'SELECT MAX(id) FROM "{table}"'); m=cur.fetchone()[0]
            if m is not None: cur.execute("SELECT setval(%s,%s,true)",(r[0],m))
    pg.commit()
def main():
    if not os.path.exists(SQLITE_DB): raise SystemExit(f"SQLite database not found:\n{SQLITE_DB}")
    print("="*55); print("WTW Academy SQLite -> PostgreSQL"); print("="*55)
    sqlite=sqlite3.connect(SQLITE_DB)
    try:
        pg=psycopg2.connect(DATABASE_URL)
        try:
            for t in TABLES: create_table(pg,sqlite,t)
            for t in TABLES: add_missing_columns(pg,sqlite,t)
            for t in TABLES: migrate_table(pg,sqlite,t)
            reset_sequences(pg)
            print("="*55); print("MIGRATION COMPLETED"); print("="*55)
        finally: pg.close()
    finally: sqlite.close()
if __name__=="__main__": main()
