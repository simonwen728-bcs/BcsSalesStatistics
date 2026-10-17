#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BCS 销售统计 - 单机本地版后端
Python 标准库 only (http.server + sqlite3 + json)，无需 pip 安装任何依赖。

运行:  python server.py            (默认端口 18789)
       python server.py --port 8080
数据:  ./data/bcs_data.db          (SQLite，程序文件夹内)
设置:  ./data/prefs.json           (行高列宽等界面设置，自动保存)
"""
import argparse
import gzip
import io
import json
import os
import re
import shutil
import sqlite3
import sys
import threading
import time
import urllib.parse
import urllib.request
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
DB_PATH = os.path.join(DATA_DIR, "bcs_data.db")
PREFS_PATH = os.path.join(DATA_DIR, "prefs.json")
SHEETS_API_URL = os.environ.get(
    "SHEETS_API_URL",
    "https://script.google.com/macros/s/AKfycbwmHqWobogKZ6CEAoyUPzn6sMorU_ZfigS-wRKt-goWuirAggXQb_BUiHB4FjTbknPR/exec?sheet=Pull",
)
# 报价工具共享文档的 Pull 接口（合并接口）：石材库存直接读这里的 Pull，不再用销售统计表里的 StoneStock 页签。
# Pull 11 列：No.=0, Description=1, Item Type=2, Qty.(On Hand)=3, Available for Sales Qty.=4,
# Qty.(Transit)=5, Unit Price=6, Blocked=7, Only PO=8, Color=9, Category=10
QUOTE_PULL_API_URL = os.environ.get(
    "QUOTE_PULL_API_URL",
    "https://script.google.com/macros/s/AKfycbxNRW6qfXFVunloGzMmFbw0WgaHYqs2OHpA40O-fpsm2WKv_1E7s-eXV-0pX-X6Slw8pQ/exec",
)
UI_DIR = os.path.join(BASE_DIR, "ui")

os.makedirs(DATA_DIR, exist_ok=True)

DB_LOCK = threading.RLock()
PREFS_LOCK = threading.RLock()


# ================= 数据库 =================

def connect():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    with DB_LOCK:
        con = connect()
        try:
            con.executescript("""
            CREATE TABLE IF NOT EXISTS sales (
              id TEXT PRIMARY KEY, batch_id TEXT NOT NULL, month TEXT NOT NULL,
              item_no TEXT DEFAULT '', description TEXT DEFAULT '', color TEXT DEFAULT '',
              name TEXT DEFAULT '', type TEXT DEFAULT '', category TEXT DEFAULT '',
              product_type TEXT DEFAULT '', sold_date TEXT DEFAULT '', invoice TEXT DEFAULT '',
              ledger TEXT DEFAULT '', cust_no TEXT DEFAULT '', cust_name TEXT DEFAULT '',
              rep1 TEXT DEFAULT '', rep1_name TEXT DEFAULT '', rep2 TEXT DEFAULT '',
              rep2_name TEXT DEFAULT '', qty REAL DEFAULT 0, amount REAL DEFAULT 0, cost REAL DEFAULT 0
            );
            CREATE TABLE IF NOT EXISTS customers (
              no TEXT PRIMARY KEY, name TEXT DEFAULT '', city TEXT DEFAULT '',
              state TEXT DEFAULT '', phone TEXT DEFAULT '', contact TEXT DEFAULT '',
              type TEXT DEFAULT '', email TEXT DEFAULT '', sales1 TEXT DEFAULT '',
              created_at TEXT DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS stock (
              item_no TEXT PRIMARY KEY, description TEXT DEFAULT '',
              available REAL DEFAULT 0, in_transit REAL DEFAULT 0
            );
            CREATE TABLE IF NOT EXISTS reps (
              kind TEXT NOT NULL, code TEXT NOT NULL, name TEXT DEFAULT '',
              group_ TEXT DEFAULT '', locked INTEGER DEFAULT 0, deleted INTEGER DEFAULT 0,
              PRIMARY KEY (kind, code)
            );
            CREATE TABLE IF NOT EXISTS batches (
              id TEXT PRIMARY KEY, month TEXT NOT NULL, file TEXT DEFAULT '',
              rows INTEGER DEFAULT 0, imported_at TEXT DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS meta (
              key TEXT PRIMARY KEY, value TEXT DEFAULT ''
            );
            CREATE INDEX IF NOT EXISTS idx_sales_month ON sales(month);
            CREATE INDEX IF NOT EXISTS idx_sales_date ON sales(sold_date);
            CREATE INDEX IF NOT EXISTS idx_sales_item ON sales(item_no);
            CREATE INDEX IF NOT EXISTS idx_sales_cust ON sales(cust_no);
            CREATE INDEX IF NOT EXISTS idx_sales_batch ON sales(batch_id);
            """)
            con.commit()
        finally:
            con.close()


# ================= 工具 =================

def row_to_sale(r):
    return {
        "id": r["id"], "month": r["month"] or "", "itemNo": r["item_no"] or "",
        "description": r["description"] or "", "color": r["color"] or "",
        "name": r["name"] or "", "type": r["type"] or "",
        "category": r["category"] or "", "productType": r["product_type"] or "",
        "soldDate": r["sold_date"] or "", "invoice": r["invoice"] or "",
        "ledger": r["ledger"] or "", "custNo": r["cust_no"] or "",
        "custName": r["cust_name"] or "", "rep1": r["rep1"] or "",
        "rep1Name": r["rep1_name"] or "", "rep2": r["rep2"] or "",
        "rep2Name": r["rep2_name"] or "", "qty": r["qty"] or 0,
        "amount": r["amount"] or 0, "cost": r["cost"] or 0,
        "batchId": r["batch_id"] or "",
    }


def row_to_customer(r):
    return {
        "no": r["no"] or "", "name": r["name"] or "", "city": r["city"] or "",
        "state": r["state"] or "", "phone": r["phone"] or "",
        "contact": r["contact"] or "", "type": r["type"] or "",
        "email": r["email"] or "", "sales1": r["sales1"] or "",
        "created": r["created_at"] or "",
    }


def row_to_stock(r):
    return {
        "itemNo": r["item_no"] or "", "description": r["description"] or "",
        "available": r["available"] or 0, "inTransit": r["in_transit"] or 0,
    }


def row_to_rep(r):
    return {
        "code": r["code"] or "", "name": r["name"] or "",
        "group": r["group_"] or "", "locked": bool(r["locked"]),
    }


def get_reps_group_map(con):
    """rep1 code(大写) -> group(大写)，用于销售行的 code/group 富化。"""
    m = {}
    for r in con.execute("SELECT code, group_ FROM reps WHERE kind='s1' AND (deleted IS NULL OR deleted=0)"):
        m[(r["code"] or "").strip().upper()] = (r["group_"] or "").strip().upper()
    return m


def enrich_sale(s, group_map):
    code = (s.get("rep1") or "").strip().upper()
    s["code"] = code
    s["group"] = group_map.get(code, "")
    return s


def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


# ================= prefs (界面设置:行高列宽等) =================

def prefs_load():
    with PREFS_LOCK:
        try:
            with open(PREFS_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}


def prefs_save(d):
    with PREFS_LOCK:
        tmp = PREFS_PATH + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False)
        os.replace(tmp, PREFS_PATH)


# ================= API 实现 =================

class Api:
    # ---- 销售记录 ----
    def query_sales(self, offset=0, limit=200, filters_json=""):
        try:
            f = json.loads(filters_json) if filters_json else {}
        except Exception:
            f = {}
        where, params = [], []
        q = (f.get("q") or "").strip()
        if q:
            like = "%" + q + "%"
            where.append("(item_no LIKE ? OR description LIKE ? OR color LIKE ? OR type LIKE ?"
                         " OR invoice LIKE ? OR cust_no LIKE ? OR cust_name LIKE ?"
                         " OR rep1 LIKE ? OR rep2 LIKE ? OR month LIKE ?)")
            params += [like] * 10
        if f.get("start"):
            where.append("sold_date >= ?"); params.append(f["start"])
        if f.get("end"):
            where.append("sold_date <= ?"); params.append(f["end"])
        if f.get("custNo"):
            where.append("cust_no LIKE ?"); params.append("%" + f["custNo"] + "%")
        if f.get("itemNo"):
            where.append("item_no LIKE ?"); params.append("%" + f["itemNo"] + "%")
        if f.get("invoice"):
            where.append("invoice LIKE ?"); params.append("%" + f["invoice"] + "%")
        if f.get("type"):
            where.append("type = ?"); params.append(f["type"])
        # 前端会传这些条件，之前服务端静默忽略（描述/客户名/颜色/业务员筛选不起作用）
        if f.get("custName"):
            where.append("cust_name LIKE ?"); params.append("%" + f["custName"] + "%")
        if f.get("description"):
            where.append("description LIKE ?"); params.append("%" + f["description"] + "%")
        if f.get("color"):
            where.append("color LIKE ?"); params.append("%" + f["color"] + "%")
        if f.get("rep1"):
            where.append("rep1 LIKE ?"); params.append("%" + f["rep1"] + "%")
        w = ("WHERE " + " AND ".join(where)) if where else ""
        with DB_LOCK:
            con = connect()
            try:
                total = con.execute("SELECT COUNT(*) FROM sales " + w, params).fetchone()[0]
                rows = con.execute(
                    "SELECT * FROM sales " + w + " ORDER BY sold_date DESC, invoice, id LIMIT ? OFFSET ?",
                    params + [limit, offset]).fetchall()
                gm = get_reps_group_map(con)
            finally:
                con.close()
        return {"total": total, "rows": [enrich_sale(row_to_sale(r), gm) for r in rows]}

    def update_sale(self, sale_id, values_json):
        try:
            values = json.loads(values_json) if isinstance(values_json, str) else (values_json or {})
        except Exception:
            return {"ok": False, "error": "bad values"}
        allowed = {"item_no": "itemNo", "description": "description", "color": "color",
                   "type": "type", "sold_date": "soldDate", "invoice": "invoice",
                   "cust_no": "custNo", "cust_name": "custName", "rep1": "rep1",
                   "rep1_name": "rep1Name", "rep2": "rep2", "rep2_name": "rep2Name",
                   "qty": "qty", "amount": "amount", "cost": "cost", "month": "month"}
        sets, params = [], []
        for col, key in allowed.items():
            if key in values:
                sets.append(col + " = ?")
                params.append(values[key])
        if not sets:
            return {"ok": False, "error": "nothing to update"}
        params.append(sale_id)
        with DB_LOCK:
            con = connect()
            try:
                con.execute("UPDATE sales SET " + ", ".join(sets) + " WHERE id=?", params)
                con.commit()
            finally:
                con.close()
        return {"ok": True}

    # ---- 导入 ----
    def apply_sales_import(self, payload_json):
        try:
            p = json.loads(payload_json) if isinstance(payload_json, str) else (payload_json or {})
        except Exception:
            return {"ok": False, "error": "bad payload"}
        month = (p.get("month") or "").strip()
        file = p.get("file") or ""
        records = p.get("records") or []
        if not re.match(r"^\d{4}-\d{2}$", month):
            return {"ok": False, "error": "bad month"}
        batch_id = "b_" + month
        with DB_LOCK:
            con = connect()
            try:
                con.execute("DELETE FROM sales WHERE batch_id=?", (batch_id,))
                n = 0
                for i, r in enumerate(records):
                    rid = "%s_%d" % (batch_id, i)
                    m = (r.get("soldDate") or "")[:7]
                    if not re.match(r"^\d{4}-\d{2}$", m):
                        m = month
                    con.execute(
                        """INSERT OR REPLACE INTO sales
                        (id,batch_id,month,item_no,description,color,name,type,category,product_type,
                         sold_date,invoice,ledger,cust_no,cust_name,rep1,rep1_name,rep2,rep2_name,qty,amount,cost)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (rid, batch_id, m, r.get("itemNo") or "", r.get("description") or "",
                         r.get("color") or "", r.get("name") or "", r.get("type") or "",
                         r.get("category") or "", r.get("productType") or "",
                         r.get("soldDate") or "", r.get("invoice") or "", r.get("ledger") or "",
                         r.get("custNo") or "", r.get("custName") or "",
                         r.get("rep1") or "", r.get("rep1Name") or "",
                         r.get("rep2") or "", r.get("rep2Name") or "",
                         _num(r.get("qty")), _num(r.get("amount")), _num(r.get("cost"))))
                    n += 1
                con.execute(
                    "INSERT OR REPLACE INTO batches (id,month,file,rows,imported_at) VALUES (?,?,?,?,?)",
                    (batch_id, month, file, n, now_str()))
                synced = self._sync_reps_from_records(con)
                con.commit()
            finally:
                con.close()
        return {"ok": True, "rows": n, "repsSynced": synced}

    def _sync_reps_from_records(self, con):
        """从销售记录收集 rep1/rep2，补入 reps 表(不覆盖已有)。返回 {s1, s2} 新增数。"""
        have_s1 = {r[0] for r in con.execute("SELECT code FROM reps WHERE kind='s1'")}
        have_s2 = {r[0] for r in con.execute("SELECT code FROM reps WHERE kind='s2'")}
        add_s1, add_s2 = 0, 0
        for code, name in con.execute("SELECT DISTINCT rep1, rep1_name FROM sales WHERE rep1<>''"):
            c = (code or "").strip().upper()
            if c and c not in have_s1:
                con.execute("INSERT OR IGNORE INTO reps (kind,code,name,group_) VALUES ('s1',?,?,?)",
                            (c, (name or "").strip(), ""))
                have_s1.add(c); add_s1 += 1
        for code, name in con.execute("SELECT DISTINCT rep2, rep2_name FROM sales WHERE rep2<>''"):
            c = (code or "").strip().upper()
            if c and c not in have_s2:
                con.execute("INSERT OR IGNORE INTO reps (kind,code,name,group_) VALUES ('s2',?,?,?)",
                            (c, (name or "").strip(), ""))
                have_s2.add(c); add_s2 += 1
        return {"s1": add_s1, "s2": add_s2}

    def delete_batch(self, batch_id):
        with DB_LOCK:
            con = connect()
            try:
                con.execute("DELETE FROM sales WHERE batch_id=?", (batch_id,))
                con.execute("DELETE FROM batches WHERE id=?", (batch_id,))
                con.commit()
            finally:
                con.close()
        return {"ok": True}

    def list_batches(self):
        with DB_LOCK:
            con = connect()
            try:
                rows = con.execute("SELECT * FROM batches ORDER BY month").fetchall()
            finally:
                con.close()
        return [{"id": r["id"], "month": r["month"], "file": r["file"] or "",
                 "rows": r["rows"] or 0, "importedAt": r["imported_at"] or ""} for r in rows]

    # ---- 统计上下文(报表/客户动态用) ----
    def get_stats_context(self, start="", end=""):
        """返回报表所需的全部数据。走 HTTP+gzip，不再经过 pywebview 桥(原卡死根因)。"""
        where, params = [], []
        if start:
            where.append("sold_date >= ?"); params.append(start)
        if end:
            where.append("sold_date <= ?"); params.append(end)
        w = ("WHERE " + " AND ".join(where)) if where else ""
        with DB_LOCK:
            con = connect()
            try:
                sales = con.execute("SELECT * FROM sales " + w + " ORDER BY sold_date, invoice, id",
                                    params).fetchall()
                gm = get_reps_group_map(con)
                reps_s1 = [row_to_rep(r) for r in
                           con.execute("SELECT * FROM reps WHERE kind='s1' AND (deleted IS NULL OR deleted=0) ORDER BY code")]
                reps_s2 = [row_to_rep(r) for r in
                           con.execute("SELECT * FROM reps WHERE kind='s2' AND (deleted IS NULL OR deleted=0) ORDER BY code")]
                stock = [row_to_stock(r) for r in con.execute("SELECT * FROM stock ORDER BY item_no")]
                customers = [row_to_customer(r) for r in con.execute("SELECT * FROM customers ORDER BY no")]
            finally:
                con.close()
        return {
            "sales": [enrich_sale(row_to_sale(r), gm) for r in sales],
            "reps": {"s1": reps_s1, "s2": reps_s2},
            "stock": stock,
            "customers": customers,
        }

    def get_lookup_data(self):
        with DB_LOCK:
            con = connect()
            try:
                items = [r[0] for r in con.execute(
                    "SELECT DISTINCT item_no FROM sales WHERE item_no<>'' ORDER BY item_no LIMIT 5000")]
                customers = [{"no": r["no"], "name": r["name"]}
                             for r in con.execute("SELECT no,name FROM customers ORDER BY no LIMIT 5000")]
                invoices = [r[0] for r in con.execute(
                    "SELECT DISTINCT invoice FROM sales WHERE invoice<>'' ORDER BY invoice DESC LIMIT 5000")]
                types = [r[0] for r in con.execute(
                    "SELECT DISTINCT type FROM sales WHERE type<>'' ORDER BY type")]
                colors = [r[0] for r in con.execute(
                    "SELECT DISTINCT color FROM sales WHERE color<>'' ORDER BY color")]
                # 城市下拉（前端 fCity 模糊用 lk.cities）：之前漏返回，城市联想恒为空
                cities = [r[0] for r in con.execute(
                    "SELECT DISTINCT city FROM customers WHERE city<>'' ORDER BY city")]
            finally:
                con.close()
        return {"items": items, "customers": customers, "invoices": invoices,
                "types": types, "colors": colors, "cities": cities}

    def search_invoices(self, q=""):
        # 发票号模糊搜索走服务端：get_lookup_data 的发票列表 LIMIT 5000 会截掉老发票
        q = (q or "").strip()
        with DB_LOCK:
            con = connect()
            try:
                like = "%" + q + "%"
                rows = con.execute(
                    "SELECT DISTINCT invoice FROM sales WHERE invoice LIKE ? ORDER BY invoice DESC LIMIT 20",
                    (like,)).fetchall()
            finally:
                con.close()
        return [r[0] for r in rows]

    def search_descs(self, q=""):
        # 描述模糊搜索走服务端：lookup 里没有描述数据
        q = (q or "").strip()
        with DB_LOCK:
            con = connect()
            try:
                like = "%" + q + "%"
                rows = con.execute(
                    "SELECT DISTINCT description FROM sales WHERE description LIKE ? ORDER BY description LIMIT 20",
                    (like,)).fetchall()
            finally:
                con.close()
        return [r[0] for r in rows if r[0]]

    # ---- 客户 ----
    def get_all_customers(self, query=""):
        q = (query or "").strip()
        with DB_LOCK:
            con = connect()
            try:
                if q:
                    like = "%" + q + "%"
                    rows = con.execute(
                        "SELECT * FROM customers WHERE no LIKE ? OR name LIKE ? OR city LIKE ? ORDER BY no",
                        (like, like, like)).fetchall()
                else:
                    rows = con.execute("SELECT * FROM customers ORDER BY no").fetchall()
            finally:
                con.close()
        return [row_to_customer(r) for r in rows]

    def upsert_customer(self, values_json):
        v = json.loads(values_json) if isinstance(values_json, str) else (values_json or {})
        no = (v.get("no") or "").strip()
        if not no:
            return {"ok": False, "error": "no required"}
        with DB_LOCK:
            con = connect()
            try:
                con.execute(
                    """INSERT INTO customers (no,name,city,state,phone,contact,type,email,sales1,created_at)
                    VALUES (?,?,?,?,?,?,?,?,?,?)
                    ON CONFLICT(no) DO UPDATE SET
                    name=excluded.name, city=excluded.city, state=excluded.state,
                    phone=excluded.phone, contact=excluded.contact, type=excluded.type,
                    email=excluded.email, sales1=excluded.sales1, created_at=excluded.created_at""",
                    (no, v.get("name") or "", v.get("city") or "", v.get("state") or "",
                     v.get("phone") or "", v.get("contact") or "", v.get("type") or "",
                     v.get("email") or "", (v.get("sales1") or "").strip().upper(),
                     v.get("created") or ""))
                con.commit()
            finally:
                con.close()
        return {"ok": True}

    def delete_customer(self, no):
        with DB_LOCK:
            con = connect()
            try:
                con.execute("DELETE FROM customers WHERE no=?", (no,))
                con.commit()
            finally:
                con.close()
        return {"ok": True}

    # ---- 库存 ----
    def get_all_stock(self, query=""):
        q = (query or "").strip()
        with DB_LOCK:
            con = connect()
            try:
                if q:
                    like = "%" + q + "%"
                    rows = con.execute(
                        "SELECT * FROM stock WHERE item_no LIKE ? OR description LIKE ? ORDER BY item_no",
                        (like, like)).fetchall()
                else:
                    rows = con.execute("SELECT * FROM stock ORDER BY item_no").fetchall()
            finally:
                con.close()
        return [row_to_stock(r) for r in rows]

    def upsert_stock(self, values_json):
        v = json.loads(values_json) if isinstance(values_json, str) else (values_json or {})
        item_no = (v.get("itemNo") or "").strip()
        if not item_no:
            return {"ok": False, "error": "itemNo required"}
        with DB_LOCK:
            con = connect()
            try:
                con.execute(
                    """INSERT INTO stock (item_no,description,available,in_transit) VALUES (?,?,?,?)
                    ON CONFLICT(item_no) DO UPDATE SET
                    description=excluded.description, available=excluded.available,
                    in_transit=excluded.in_transit""",
                    (item_no, v.get("description") or "",
                     _num(v.get("available")), _num(v.get("inTransit"))))
                con.commit()
            finally:
                con.close()
        return {"ok": True}

    def delete_stock(self, item_no):
        with DB_LOCK:
            con = connect()
            try:
                con.execute("DELETE FROM stock WHERE item_no=?", (item_no,))
                con.commit()
            finally:
                con.close()
        return {"ok": True}

    # ---- 销售员 ----
    def get_all_reps(self):
        with DB_LOCK:
            con = connect()
            try:
                s1 = [row_to_rep(r) for r in con.execute(
                    "SELECT * FROM reps WHERE kind='s1' AND (deleted IS NULL OR deleted=0) ORDER BY code")]
                s2 = [row_to_rep(r) for r in con.execute(
                    "SELECT * FROM reps WHERE kind='s2' AND (deleted IS NULL OR deleted=0) ORDER BY code")]
            finally:
                con.close()
        return {"s1": s1, "s2": s2}

    def upsert_rep(self, kind, values_json):
        v = json.loads(values_json) if isinstance(values_json, str) else (values_json or {})
        kind = "s1" if kind == "s1" else "s2"
        code = (v.get("code") or "").strip().upper()
        if not code:
            return {"ok": False, "error": "code required"}
        with DB_LOCK:
            con = connect()
            try:
                con.execute(
                    """INSERT INTO reps (kind,code,name,group_,deleted) VALUES (?,?,?,?,0)
                    ON CONFLICT(kind,code) DO UPDATE SET
                    name=excluded.name, group_=excluded.group_, deleted=0""",
                    (kind, code, v.get("name") or "", v.get("group") or ""))
                con.commit()
            finally:
                con.close()
        return {"ok": True}

    def delete_rep(self, kind, code):
        with DB_LOCK:
            con = connect()
            try:
                con.execute("UPDATE reps SET deleted=1 WHERE kind=? AND code=?", (kind, code))
                con.commit()
            finally:
                con.close()
        return {"ok": True}

    # ---- 主数据导入 ----
    def apply_master_import(self, payload_json):
        p = json.loads(payload_json) if isinstance(payload_json, str) else (payload_json or {})
        with DB_LOCK:
            con = connect()
            try:
                for c in (p.get("customers") or []):
                    no = (c.get("no") or c.get("custNo") or "").strip()
                    if not no:
                        continue
                    con.execute(
                        """INSERT INTO customers (no,name,city,state,phone,contact,type,email,sales1,created_at)
                        VALUES (?,?,?,?,?,?,?,?,?,?)
                        ON CONFLICT(no) DO UPDATE SET
                        name=excluded.name, city=excluded.city, state=excluded.state,
                        phone=excluded.phone, contact=excluded.contact, type=excluded.type,
                        email=excluded.email, sales1=excluded.sales1, created_at=excluded.created_at""",
                        (no, c.get("name") or c.get("custName") or "", c.get("city") or "",
                         c.get("state") or "", c.get("phone") or "", c.get("contact") or "",
                         c.get("type") or "", c.get("email") or "",
                         (c.get("sales1") or "").strip().upper(), c.get("created") or c.get("created_at") or ""))
                for s in (p.get("stock") or []):
                    item_no = (s.get("itemNo") or s.get("item_no") or "").strip()
                    if not item_no:
                        continue
                    con.execute(
                        """INSERT INTO stock (item_no,description,available,in_transit) VALUES (?,?,?,?)
                        ON CONFLICT(item_no) DO UPDATE SET
                        description=excluded.description, available=excluded.available,
                        in_transit=excluded.in_transit""",
                        (item_no, s.get("description") or "",
                         _num(s.get("available")), _num(s.get("inTransit") or s.get("in_transit"))))
                # reps 两种形状都接受：同步路径传扁平列表 [{kind,code,...}]，
                # 前端主数据导入/演示数据传 {s1:[...], s2:[...]}；字典形状先展平成列表
                reps_in = p.get("reps") or []
                if isinstance(reps_in, dict):
                    flat = []
                    for _kind in ("s1", "s2"):
                        for _r in (reps_in.get(_kind) or []):
                            _rr = dict(_r) if isinstance(_r, dict) else {}
                            _rr.setdefault("kind", _kind)
                            flat.append(_rr)
                    reps_in = flat
                for r in reps_in:
                    kind = "s1" if (r.get("kind") or "s1") == "s1" else "s2"
                    code = (r.get("code") or "").strip().upper()
                    if not code:
                        continue
                    con.execute(
                        """INSERT INTO reps (kind,code,name,group_,deleted) VALUES (?,?,?,?,0)
                        ON CONFLICT(kind,code) DO UPDATE SET
                        name=excluded.name, group_=excluded.group_, deleted=0""",
                        (kind, code, r.get("name") or "", r.get("group") or ""))
                con.commit()
            finally:
                con.close()
        return {"ok": True}

    def clear_all_data(self):
        with DB_LOCK:
            con = connect()
            try:
                for t in ("sales", "customers", "stock", "reps", "batches"):
                    con.execute("DELETE FROM " + t)
                con.commit()
            finally:
                con.close()
        return {"ok": True}

    def initialize_data(self):
        """初始化:补全销售2 / 补齐销售分组 / 刷新查询数据源。
        与前端 initializeFlow 对齐，返回 {sales2Filled, newS1, newS2}（之前只返回 counts，初始化实际没做）。"""
        init_db()
        with DB_LOCK:
            con = connect()
            try:
                counts = {t: con.execute("SELECT COUNT(*) FROM " + t).fetchone()[0]
                          for t in ("sales", "customers", "stock", "reps", "batches")}
                # 补销售2：rep2 为空且 rep1 有值时，rep2=rep1，并补 rep2_name
                cur = con.execute(
                    "UPDATE sales SET rep2=rep1, "
                    "rep2_name=COALESCE(NULLIF(rep2_name,''), NULLIF(rep1_name,''), 'Ex Sales Rep') "
                    "WHERE (rep2 IS NULL OR rep2='') AND rep1<>''")
                filled = cur.rowcount or 0
                # 补 rep2_name：rep2 有值但 rep2_name 为空时取 rep1_name，仍空则 'Ex Sales Rep'
                con.execute(
                    "UPDATE sales SET rep2_name=COALESCE(NULLIF(rep1_name,''), 'Ex Sales Rep') "
                    "WHERE rep2<>'' AND (rep2_name IS NULL OR rep2_name='')")
                # 补齐销售分组：销售记录里出现过的 code，reps 表没有则补入（后出现的同名覆盖先出现的，与 exe 逻辑一致）
                have_s1 = {r[0] for r in con.execute("SELECT UPPER(code) FROM reps WHERE kind='s1'")}
                have_s2 = {r[0] for r in con.execute("SELECT UPPER(code) FROM reps WHERE kind='s2'")}
                uniq = {}
                for code, name in con.execute(
                        "SELECT rep1, rep1_name FROM sales WHERE rep1<>'' ORDER BY id"):
                    uniq[code.strip().upper()] = (name or "").strip() or code.strip().upper()
                for code, name in con.execute(
                        "SELECT rep2, rep2_name FROM sales WHERE rep2<>'' ORDER BY id"):
                    uniq[code.strip().upper()] = (name or "").strip() or code.strip().upper()
                new_s1, new_s2 = [], []
                for code in sorted(uniq):
                    if code not in have_s1:
                        con.execute("INSERT OR IGNORE INTO reps (kind,code,name,group_) VALUES ('s1',?,?,?)",
                                    (code, uniq[code], ""))
                        new_s1.append(code)
                    if code not in have_s2:
                        con.execute("INSERT OR IGNORE INTO reps (kind,code,name,group_) VALUES ('s2',?,?,?)",
                                    (code, uniq[code], ""))
                        new_s2.append(code)
                con.commit()
            finally:
                con.close()
        return {"ok": True, "sales2Filled": filled, "newS1": new_s1, "newS2": new_s2,
                "counts": counts}

    # ---- meta / prefs ----
    def get_meta(self):
        with DB_LOCK:
            con = connect()
            try:
                cols = [d[0] for d in con.execute("SELECT * FROM meta LIMIT 1").description or []]
                if "key" in cols:
                    rows = con.execute("SELECT key, value FROM meta").fetchall()
                    meta = {r[0]: r[1] for r in rows}
                else:
                    rows = con.execute("SELECT k, v FROM meta").fetchall()
                    meta = {r[0]: r[1] for r in rows}
            finally:
                con.close()
        prefs = prefs_load()
        # 前端 set_pref 存到 prefs.json;旧版 meta 表里的设置也合并进来
        # 前端启动时从 meta.prefs 恢复界面设置(行高列宽等)
        for k in ("sheetLayouts", "userStyles", "navCollapsed", "sidebarCollapsed"):
            if k in meta and k not in prefs:
                prefs[k] = meta[k]
        meta["prefs"] = prefs
        # lang 放顶层(前端直接读 meta.lang)
        if "lang" in prefs:
            meta["lang"] = prefs["lang"]
        elif "lang" in meta:
            meta["lang"] = meta["lang"]
        return meta

    def set_pref(self, key, value):
        prefs = prefs_load()
        prefs[key] = value
        prefs_save(prefs)
        return {"ok": True}

    def set_lang(self, lang):
        prefs = prefs_load()
        prefs["lang"] = lang
        prefs_save(prefs)
        return {"ok": True}

    def close_app(self):
        # 网页版由浏览器关闭页面；接口保留兼容（exe 版真正退出），前端退出时调用
        return {"ok": True}

    # ---- 备份 ----
    def backup_inspect(self):
        with DB_LOCK:
            con = connect()
            try:
                counts = {t: con.execute("SELECT COUNT(*) FROM " + t).fetchone()[0]
                          for t in ("sales", "customers", "stock", "reps", "batches")}
            finally:
                con.close()
        size = os.path.getsize(DB_PATH) if os.path.exists(DB_PATH) else 0
        return {"ok": True, "counts": counts, "size": size,
                "mtime": time.strftime("%Y-%m-%d %H:%M", time.localtime(os.path.getmtime(DB_PATH)))
                if os.path.exists(DB_PATH) else ""}

    def backup_restore(self, payload_json):
        # 网页版走 /api/backup_upload 直接传文件;此接口保留兼容
        return {"ok": False, "error": "请使用备份上传功能"}

    def sync_from_sheets(self):
        """从 Google Sheets JSON 接口同步销售数据,按 soldDate 月份分组导入。
        数据源: SHEETS_API_URL 环境变量,默认为 V66 销售统计表的 Pull 工作表。"""
        url = SHEETS_API_URL
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "BCS-Sync/1.0"})
            with urllib.request.urlopen(req, timeout=120) as resp:
                raw = resp.read()
        except Exception as e:
            return {"ok": False, "error": "拉取失败: %s: %s" % (type(e).__name__, e)}
        try:
            data = json.loads(raw.decode("utf-8"))
        except Exception as e:
            return {"ok": False, "error": "JSON 解析失败: %s" % e}
        if not data.get("ok"):
            return {"ok": False, "error": data.get("error") or "接口返回失败"}
        rows = data.get("rows") or []
        # 按 soldDate 月份 (YYYY-MM) 分组
        months = {}
        skipped = 0
        for r in rows:
            try:
                sold = str(r[7] or "").strip()[:10]
            except Exception:
                sold = ""
            m = sold[:7] if re.match(r"^\d{4}-\d{2}$", sold[:7]) else ""
            if not m:
                skipped += 1
                continue
            rec = {
                "itemNo": str(r[0] or ""), "description": str(r[1] or ""),
                "color": str(r[2] or ""), "name": str(r[3] or ""),
                "type": str(r[4] or ""), "category": str(r[5] or ""),
                "productType": str(r[6] or ""), "soldDate": sold,
                "invoice": str(r[8] or ""), "ledger": str(r[9] or ""),
                "custNo": str(r[10] or ""), "custName": str(r[11] or ""),
                "rep1": str(r[12] or ""), "rep1Name": str(r[13] or ""),
                "rep2": str(r[14] or ""), "rep2Name": str(r[15] or ""),
                "qty": _num(r[16]), "amount": _num(r[17]), "cost": _num(r[18]),
            }
            months.setdefault(m, []).append(rec)
        total = 0
        results = {}
        for m in sorted(months.keys()):
            recs = months[m]
            res = self.apply_sales_import({"month": m, "records": recs, "file": "sheets-sync"})
            if not res.get("ok"):
                return {"ok": False, "error": "导入 %s 失败: %s" % (m, res.get("error"))}
            total += res.get("rows", 0)
            results[m] = res.get("rows", 0)
        return {"ok": True, "months": len(months), "rows": total,
                "skipped": skipped, "byMonth": results}

    def _fetch_sheet(self, sheet):
        """拉取指定工作表的 {headers, rows}。成功返回 (data, None),失败返回 (None, error)。"""
        base = SHEETS_API_URL.split("?")[0]
        url = base + "?sheet=" + urllib.parse.quote(sheet)
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "BCS-Sync/1.0"})
            with urllib.request.urlopen(req, timeout=120) as resp:
                raw = resp.read()
        except Exception as e:
            return None, "拉取 %s 失败: %s: %s" % (sheet, type(e).__name__, e)
        try:
            data = json.loads(raw.decode("utf-8"))
        except Exception as e:
            return None, "%s JSON 解析失败: %s" % (sheet, e)
        if not data.get("ok"):
            return None, "%s 接口返回失败: %s" % (sheet, data.get("error") or "")
        return data, None

    def _fetch_quote_sheet(self, sheet):
        """拉取报价工具共享文档指定工作表的 {headers, rows}（合并接口）。成功返回 (data, None)，失败返回 (None, error)。"""
        base = QUOTE_PULL_API_URL.split("?")[0]
        url = base + "?sheet=" + urllib.parse.quote(sheet)
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "BCS-Sync/1.0"})
            with urllib.request.urlopen(req, timeout=120) as resp:
                raw = resp.read()
        except Exception as e:
            return None, "拉取报价 %s 失败: %s: %s" % (sheet, type(e).__name__, e)
        try:
            data = json.loads(raw.decode("utf-8"))
        except Exception as e:
            return None, "报价 %s JSON 解析失败: %s" % (sheet, e)
        if not data.get("ok"):
            return None, "报价 %s 接口返回失败: %s" % (sheet, data.get("error") or "")
        return data, None

    @staticmethod
    def _cell(r, i):
        try:
            v = r[i]
        except Exception:
            return ""
        return "" if v is None else str(v)

    @staticmethod
    def _clean_date(v):
        # 统一成 YYYY-MM-DD：ISO（含时间后缀）取日期部分；美式 M/D/YYYY 转 ISO；无法解析的原样返回
        s = (v or "").strip()
        if not s:
            return s
        for fmt, txt in (("%Y-%m-%d", s[:10]), ("%m/%d/%Y", s)):
            try:
                return datetime.strptime(txt, fmt).strftime("%Y-%m-%d")
            except Exception:
                continue
        return s

    def _restore_reps_from_v66_if_empty(self):
        """业务分组恢复：仅当本地 reps 表没有任何 group_ 非空行时，才从 V66 Group 表恢复。
        本地有 Group 数据时不动（手动改的不会被覆盖）；容器重置后 reps 被销售同步补回但 group_ 全空，
        此时从 V66 恢复 Group。返回 (restored_count, error)。"""
        try:
            con = connect()
            try:
                n = con.execute(
                    "SELECT COUNT(*) FROM reps WHERE (deleted IS NULL OR deleted=0) AND group_ IS NOT NULL AND group_<>''"
                ).fetchone()[0]
            finally:
                con.close()
        except Exception as e:
            return 0, "检查 reps 表失败: %s" % e
        if n > 0:
            return 0, None  # 本地已有 Group，不动
        data, err = self._fetch_sheet("Group")
        if err:
            return 0, err
        reps = []
        for r in (data.get("rows") or []):
            # Sales1：A=Code B=Name C=Group 1；Sales2：F=Code G=Name H=Group 2（两栏并排）
            c1 = self._cell(r, 0).strip()
            if c1 and c1.upper() not in ("SALES1 CODE", "CODE"):
                reps.append({"kind": "s1", "code": c1, "name": self._cell(r, 1), "group": self._cell(r, 2)})
            c2 = self._cell(r, 5).strip() if len(r) > 5 else ""
            if c2 and c2.upper() not in ("SALES2 CODE", "CODE"):
                reps.append({"kind": "s2", "code": c2, "name": self._cell(r, 6) if len(r) > 6 else "",
                             "group": self._cell(r, 7) if len(r) > 7 else ""})
        return reps, None

    def sync_master_from_sheets(self, scope='all'):
        """从 Google Sheets 同步主数据:报价文档 Customer -> customers, 报价 Pull 的 SL* 行 -> stock, V66 Group -> reps（仅本地 Group 为空时恢复）。
        客户资料与石材库存都直接读报价工具共享文档，不再读销售统计表里的 Customer / StoneStock 页签。
        scope: 'all'（默认，客户+库存全同步）、'customers'（只同步客户）、'stock'（只同步库存）。
        独立同步只写对应表（apply_master_import 是 upsert，不动其他表）。"""
        scope = (scope or 'all').strip().lower()
        payload = {"customers": [], "stock": [], "reps": []}
        # --- Customer（报价工具共享文档的 Customer 页签）---
        # 报价 Customer 列：No.=0, Name=1, City=3, State=4, Salesperson Code=7,
        # Phone No.=10, Cell Phone No.=11, Contact=12, Creation Date=13, Customer Type=16, E-Mail=17
        if scope in ('all', 'customers'):
            data, err = self._fetch_quote_sheet("Customer")
            if err:
                return {"ok": False, "error": err}
            for r in (data.get("rows") or []):
                no = self._cell(r, 0).strip()
                if not no:
                    continue
                phone = self._cell(r, 10).strip() or self._cell(r, 11).strip()
                payload["customers"].append({
                    "no": no, "name": self._cell(r, 1),
                    "city": self._cell(r, 3), "state": self._cell(r, 4),
                    "phone": phone, "contact": self._cell(r, 12),
                    "type": self._cell(r, 16), "email": self._cell(r, 17),
                    "sales1": self._cell(r, 7), "created": self._clean_date(self._cell(r, 13)),
                })
        # --- 石材库存：报价工具共享文档的 Pull，SL* 行，只进有库存的 ---
        if scope in ('all', 'stock'):
            data, err = self._fetch_quote_sheet("Pull")
            if err:
                return {"ok": False, "error": err}
            for r in (data.get("rows") or []):
                item_no = self._cell(r, 0).strip()
                if not item_no or not item_no.upper().startswith("SL"):
                    continue
                available = _num(r[4] if len(r) > 4 else "")
                in_transit = _num(r[5] if len(r) > 5 else "")
                if available <= 0 and in_transit <= 0:
                    continue  # 无库存的行不进
                payload["stock"].append({
                    "itemNo": item_no, "description": self._cell(r, 1),
                    "available": available, "inTransit": in_transit,
                })
        # 业务分组恢复：仅本地 Group 为空时从 V66 Group 表恢复；本地有 Group 时不动。
        # 注意：Group 在 V66 里（不是报价文档），走 V66 接口 _fetch_sheet。
        if scope in ('all',):
            reps_restored, reps_err = self._restore_reps_from_v66_if_empty()
            if reps_err:
                return {"ok": False, "error": reps_err}
            if isinstance(reps_restored, list):
                payload["reps"] = reps_restored
        # 同步 = 导入式更新（与原来的导入功能一致，只是来源改成远程报价文档）：
        # 新增的加上，已有的用报价文档数据更新，不在报价文档里的本地行保留（不清空）
        res = self.apply_master_import(payload)
        if not res.get("ok"):
            return {"ok": False, "error": "主数据导入失败: %s" % res.get("error")}
        return {"ok": True, "customers": len(payload["customers"]),
                "stock": len(payload["stock"]), "reps": len(payload["reps"])}


def _num(v):
    try:
        if v is None or v == "":
            return 0
        if isinstance(v, str):
            v = v.replace(",", "").replace("$", "").strip()
            if v.startswith("(") and v.endswith(")"):
                v = "-" + v[1:-1]
        return float(v)
    except Exception:
        return 0


API = Api()
API_METHODS = {name: getattr(API, name) for name in dir(API) if not name.startswith("_")}


# ================= HTTP =================

MIME = {
    ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8", ".json": "application/json; charset=utf-8",
    ".png": "image/png", ".jpg": "image/jpeg", ".svg": "image/svg+xml",
    ".woff": "font/woff", ".woff2": "font/woff2", ".ttf": "font/ttf",
    ".ico": "image/x-icon", ".map": "application/json",
}


class Handler(BaseHTTPRequestHandler):
    server_version = "BCSLocal/1.0"

    def log_message(self, fmt, *args):
        sys.stderr.write("[%s] %s\n" % (time.strftime("%H:%M:%S"), fmt % args))

    def _send(self, body: bytes, ctype="application/json; charset=utf-8", status=200, gzip_ok=True):
        accept = self.headers.get("Accept-Encoding", "")
        if gzip_ok and len(body) > 102400 and "gzip" in accept:
            body = gzip.compress(body, compresslevel=6)
            self.send_response(status)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Encoding", "gzip")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        else:
            self.send_response(status)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

    def _json(self, obj, status=200):
        self._send(json.dumps(obj, ensure_ascii=False).encode("utf-8"), status=status)

    def do_GET(self):
        path = urllib.parse.urlparse(self.path).path
        if path == "/api/backup_download":
            return self._backup_download()
        if path in ("/", "/index.html"):
            return self._static("index.html")
        # 静态文件
        rel = urllib.parse.unquote(path.lstrip("/"))
        safe = os.path.normpath(os.path.join(UI_DIR, rel))
        if not safe.startswith(UI_DIR) or not os.path.isfile(safe):
            self._json({"error": "not found"}, 404)
            return
        self._static_rel(safe)

    def _static(self, name):
        self._static_rel(os.path.join(UI_DIR, name))

    def _static_rel(self, full):
        ext = os.path.splitext(full)[1].lower()
        ctype = MIME.get(ext, "application/octet-stream")
        try:
            with open(full, "rb") as f:
                body = f.read()
        except OSError:
            self._json({"error": "not found"}, 404)
            return
        self._send(body, ctype)

    def _backup_download(self):
        if not os.path.exists(DB_PATH):
            self._json({"ok": False, "error": "no db"}, 404)
            return
        with DB_LOCK:
            tmp = DB_PATH + ".dl"
            shutil.copy2(DB_PATH, tmp)
        try:
            with open(tmp, "rb") as f:
                body = f.read()
        finally:
            try:
                os.remove(tmp)
            except OSError:
                pass
        fname = "bcs_backup_%s.db" % time.strftime("%Y%m%d_%H%M%S")
        self.send_response(200)
        self.send_header("Content-Type", "application/octet-stream")
        self.send_header("Content-Disposition", "attachment; filename=%s" % fname)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        path = urllib.parse.urlparse(self.path).path
        if path == "/api/backup_upload":
            return self._backup_upload()
        if not path.startswith("/api/"):
            self._json({"error": "not found"}, 404)
            return
        method = path[len("/api/"):]
        fn = API_METHODS.get(method)
        if not fn:
            self._json({"error": "unknown method: " + method}, 404)
            return
        try:
            length = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            length = 0
        raw = self.rfile.read(length) if length else b""
        try:
            payload = json.loads(raw.decode("utf-8")) if raw else {}
        except Exception:
            payload = {}
        args = payload.get("args", [])
        try:
            result = fn(*args)
        except TypeError as e:
            self._json({"error": "bad args: %s" % e}, 400)
            return
        except Exception as e:
            self._json({"error": "%s: %s" % (type(e).__name__, e)}, 500)
            return
        # 兼容前端:python 侧返回 JSON 字符串时自动解析 —— 我们直接返回对象即可
        self._json(result if result is not None else {"ok": True})

    def _backup_upload(self):
        length = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(length) if length else b""
        if len(raw) < 100 or not raw.startswith(b"SQLite format 3\x00"):
            self._json({"ok": False, "error": "不是有效的 SQLite 数据库文件"})
            return
        with DB_LOCK:
            bak = DB_PATH + ".bak"
            if os.path.exists(DB_PATH):
                shutil.copy2(DB_PATH, bak)
            with open(DB_PATH, "wb") as f:
                f.write(raw)
            # 简单校验
            try:
                con = sqlite3.connect(DB_PATH)
                con.execute("SELECT count(*) FROM sqlite_master").fetchone()
                con.close()
            except Exception as e:
                if os.path.exists(bak):
                    shutil.copy2(bak, DB_PATH)
                self._json({"ok": False, "error": "数据库校验失败: %s" % e})
                return
        self._json({"ok": True})


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=int(os.environ.get("PORT", 18789)))
    ap.add_argument("--no-browser", action="store_true")
    args = ap.parse_args()

    init_db()
    if os.environ.get("AUTO_SYNC") == "1":
        print("AUTO_SYNC=1,正在从 Google Sheets 同步数据...")
        # 销售从 V66 读（V66 数据是最新的，服务重启后自动恢复销售）
        try:
            res = Api().sync_from_sheets()
            print("销售同步结果: %s" % json.dumps(res, ensure_ascii=False)[:500])
        except Exception as e:
            print("销售同步失败: %s: %s" % (type(e).__name__, e))
        # 客户/库存从报价文档读（不在 V66 里）
        try:
            res = Api().sync_master_from_sheets()
            print("主数据同步结果: %s" % json.dumps(res, ensure_ascii=False)[:500])
        except Exception as e:
            print("主数据同步失败: %s: %s" % (type(e).__name__, e))
    srv = ThreadingHTTPServer(("0.0.0.0", args.port), Handler)
    url = "http://0.0.0.0:%d/" % args.port
    print("=" * 60)
    print("BCS 销售统计单机版已启动")
    print("访问地址: %s" % url)
    print("数据库:   %s" % DB_PATH)
    print("=" * 60)
    if not args.no_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n已退出")


if __name__ == "__main__":
    main()
