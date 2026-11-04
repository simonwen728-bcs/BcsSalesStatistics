"""Read-only product metadata for Door/ASET SKU mix analysis."""
import hashlib
import json
from pathlib import Path
import re
import threading
import time


def _text(value):
    return re.sub(r"\s+", " ", str(value or "").strip()).upper()


def install(api, methods):
    source = Path(__file__).with_name("cabinet_sku_master.json")
    state = {"items": None, "last_fetch": 0, "error": ""}
    lock = threading.RLock()

    def metadata(codes):
        with lock:
            if state["items"] is None:
                with source.open(encoding="utf-8") as f:
                    state["items"] = json.load(f)["items"]
            items = state["items"]
            missing = [code for code in codes if code not in items]
            # The supplied full Pull has Name; the existing 11-column API does
            # not. Known names remain available without a remote network wait.
            if missing and time.time() - state["last_fetch"] > 60:
                state["last_fetch"] = time.time()
                data, error = api._fetch_quote_sheet("Pull")
                state["error"] = error or ""
                if not error:
                    headers = [_text(h) for h in data.get("headers", [])]
                    rows = data.get("rows", [])
                    if not isinstance(rows, list) or not rows or "NO." not in headers:
                        raise ValueError("Pull返回格式异常，请稍后重新生成SKU分析。")
                    indices = {name: headers.index(_text(name)) for name in
                               ["No.", "Name", "Description", "Color", "Category"]
                               if _text(name) in headers}
                    for row in rows:
                        if not isinstance(row, list):
                            continue
                        def cell(name):
                            i = indices.get(name)
                            return _text(row[i]) if i is not None and i < len(row) else ""
                        code = cell("No.")
                        if not code.startswith(("CC", "ASET")):
                            continue
                        prior = items.get(code, {})
                        values = {k: cell(k) or prior.get(k, "") for k in
                                  ["Name", "Description", "Color", "Category"]}
                        if not values["Name"] and values["Color"]:
                            desc = values["Description"]
                            color = values["Color"]
                            pos = desc.find(color)
                            if pos > 0:
                                values["Name"] = desc[:pos].strip()
                        if code not in items or cell("Name"):
                            items[code] = values
            missing = [code for code in codes if code not in items]
            if missing:
                suffix = "；" + state["error"] if state["error"] else ""
                raise ValueError("Pull缺少SKU品名映射：" + ", ".join(missing[:8]) + suffix)
            selected = {code: items[code] for code in codes}
            revision = hashlib.sha256(json.dumps(selected, ensure_ascii=False,
                                                sort_keys=True).encode()).hexdigest()[:24]
            return selected, revision

    def get_cabinet_sku_context(start="", end="", known_revision="", known_master_revision=""):
        original = api.get_product_context(start, end)
        fields = original["cols"]
        code_col = fields.index("itemNo")
        codes = sorted({_text(row[code_col]) for row in original["rows"]
                        if _text(row[code_col]).startswith(("CC", "ASET"))})
        selected, master_revision = metadata(codes)
        if known_revision == original["revision"] and known_master_revision == master_revision:
            return {"unchanged": True, "revision": original["revision"],
                    "cabinetItemRevision": master_revision}
        result = dict(original)
        result.update(cabinetItems=selected, cabinetItemRevision=master_revision)
        return result

    methods["get_cabinet_sku_context"] = get_cabinet_sku_context
