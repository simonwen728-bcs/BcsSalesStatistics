"""Validate the whole incoming month before any DELETE or sales-rep expansion."""
import math
import re
from datetime import date


def valid_rep(code, name=''):
    code = str(code or '').strip()
    if not code:
        return True
    if not re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9_-]{0,63}', code):
        return False
    norm = re.sub(r'[\s._#]', '', code).lower()
    if norm in ('repcode', 'rep1code', 'rep2code', 'rep1', 'rep2', 'sales1', 'sales2'):
        return False
    return not (code.isdigit() and re.fullmatch(r'C\d+', str(name or '').strip(), re.I))


def validate(records, month):
    if not isinstance(records, list) or not records:
        return '没有有效销售明细，原月份未改动。'
    for i, row in enumerate(records, 1):
        if not isinstance(row, dict):
            return f'第{i}条不是销售记录，原月份未改动。'
        sold = row.get('soldDate')
        try:
            if not isinstance(sold, str) or not re.fullmatch(r'\d{4}-\d{2}-\d{2}', sold):
                raise ValueError()
            date.fromisoformat(sold)
        except ValueError:
            return f'第{i}条销售日期无效，可能混入表头或参数，原月份未改动。'
        if sold[:7] != month:
            return f'第{i}条销售日期与导入月份不一致，原月份未改动。'
        if not (str(row.get('itemNo') or '').strip() or str(row.get('invoice') or '').strip()):
            return f'第{i}条缺少商品和发票，原月份未改动。'
        if not valid_rep(row.get('rep1'), row.get('rep1Name')) or not valid_rep(row.get('rep2'), row.get('rep2Name')):
            return f'第{i}条销售代码异常，已阻止错位数据污染分组，原月份未改动。'
        for key in ('qty', 'amount', 'cost'):
            raw = row.get(key)
            if raw in (None, ''):
                continue
            if isinstance(raw, str):
                raw = raw.replace(',', '').replace('$', '').strip()
                if raw.startswith('(') and raw.endswith(')'):
                    raw = '-' + raw[1:-1]
            try:
                if isinstance(raw, bool) or not math.isfinite(float(raw)):
                    raise ValueError()
            except (TypeError, ValueError, OverflowError):
                return f'第{i}条{key}不是有效数字，原月份未改动。'
    return ''
