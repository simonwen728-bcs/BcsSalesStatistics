"""Compatibility adapter for the application's fixed SQL statements.
DATABASE_URL is server-only. Connection failure never falls back to local storage.
"""
import os
import re

DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()


def postgres_sql(sql):
    # SQL here is application-owned, never supplied by clients.
    replace = re.match(r"\s*INSERT OR REPLACE INTO\s+(\w+)\s*\(([^)]+)\)", sql, re.I)
    ignore = re.match(r"\s*INSERT OR IGNORE INTO", sql, re.I)
    if replace:
        table = replace.group(1)
        keys = {"sales": ["id"], "batches": ["id"]}[table]
        cols = [c.strip() for c in replace.group(2).split(',')]
        sql = re.sub(r"INSERT OR REPLACE", "INSERT", sql, count=1, flags=re.I)
        sql += " ON CONFLICT (" + ','.join(keys) + ") DO UPDATE SET "
        sql += ','.join(c + '=EXCLUDED.' + c for c in cols if c not in keys)
    elif ignore:
        sql = re.sub(r"INSERT OR IGNORE", "INSERT", sql, count=1, flags=re.I)
        sql += " ON CONFLICT DO NOTHING"
    # PostgreSQL REAL is 32-bit; preserve SQLite's 64-bit numeric precision.
    sql = re.sub(r'\bREAL\b', 'DOUBLE PRECISION', sql)
    sql = re.sub(r'\bLIKE\b', 'ILIKE', sql)
    return sql.replace('%', '%%').replace('?', '%s')


class PostgresConnection:
    def __init__(self):
        import psycopg2
        from psycopg2.extras import DictCursor
        try:
            self.con = psycopg2.connect(DATABASE_URL, connect_timeout=20,
                                         cursor_factory=DictCursor)
        except psycopg2.Error:
            raise RuntimeError("外部数据库连接失败，请检查 Render 的 DATABASE_URL 和数据库状态。") from None

    def execute(self, sql, params=()):
        cur = self.con.cursor()
        cur.execute(postgres_sql(sql), params)
        return cur

    def executemany(self, sql, rows):
        from psycopg2.extras import execute_batch
        cur = self.con.cursor()
        execute_batch(cur, postgres_sql(sql), rows, page_size=1000)
        return cur

    def executescript(self, script):
        for sql in script.split(';'):
            if sql.strip():
                self.execute(sql)

    def commit(self):
        self.con.commit()

    def close(self):
        self.con.close()
