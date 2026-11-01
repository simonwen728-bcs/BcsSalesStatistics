/* Sales import: detect actual sheet/header; never infer sales columns by position. */
(function (root) {
  'use strict';
  const norm = value => String(value == null ? '' : value).toLowerCase().replace(/[\s._#]/g, '').trim();
  function describe(row, map) {
    const keys = row.map(value => map[norm(value)] || map[String(value || '').trim()] || null);
    const set = new Set(keys.filter(Boolean));
    const valid = set.size >= 8 && ['itemNo', 'soldDate', 'invoice', 'qty', 'amount'].every(key => set.has(key));
    return { keys, valid, score: set.size };
  }
  function dateText(value) {
    if (!(value instanceof Date)) return value == null ? '' : value;
    if (isNaN(value.getTime())) return '';
    return value.getFullYear() + '-' + String(value.getMonth() + 1).padStart(2, '0') + '-' + String(value.getDate()).padStart(2, '0');
  }
  function select(wb, XLSX, map) {
    let candidates = [];
    wb.SheetNames.forEach((name, index) => {
      const ws = wb.Sheets[name];
      if (!ws || !ws['!ref']) return;
      const meta = (wb.Workbook?.Sheets || []).find(s => s.name === name) || (wb.Workbook?.Sheets || [])[index] || {};
      const extent = XLSX.utils.decode_range(ws['!ref']);
      for (let ri = extent.s.r; ri <= Math.min(extent.e.r, extent.s.r + 99); ri++) {
        const row = [];
        for (let ci = 0; ci <= Math.min(extent.e.c, 99); ci++) row.push(ws[XLSX.utils.encode_cell({r: ri, c: ci})]?.v ?? '');
        const desc = describe(row, map);
        if (desc.valid) {
          candidates.push({ name, ws, header: ri, hidden: !!meta.Hidden,
            rank: desc.score * 10 + (/^data$/i.test(name) ? 5 : 0) });
          break;
        }
      }
    });
    if (candidates.some(c => !c.hidden)) candidates = candidates.filter(c => !c.hidden);
    candidates.sort((a, b) => b.rank - a.rank);
    if (!candidates.length) throw new Error('未找到销售明细表头（商品、销售日期、发票、数量、金额），未导入。');
    if (candidates.length > 1 && candidates[0].rank === candidates[1].rank)
      throw new Error('有多个销售明细表，无法确定应导入哪张：' + candidates.filter(c => c.rank === candidates[0].rank).map(c => c.name).join(', '));
    const chosen = candidates[0];
    const aoa = XLSX.utils.sheet_to_json(chosen.ws, {header: 1, defval: '', raw: true, blankrows: false, range: chosen.header});
    const headers = (aoa.shift() || []).map(v => String(v == null ? '' : v).trim());
    const rows = aoa.map(row => headers.map((_, ci) => dateText(row[ci])));
    return {headers, rows, sourceSheet: chosen.name, headerRow: chosen.header + 1};
  }
  function numeric(value, row, key) {
    if (value == null || value === '') return 0;
    let text = typeof value === 'string' ? value.replace(/[$,\s]/g, '') : value;
    if (typeof text === 'string' && /^\(.*\)$/.test(text)) text = '-' + text.slice(1, -1);
    const result = Number(text);
    if (!Number.isFinite(result)) throw new Error('明细第' + row + '行的' + key + '不是有效数字，未导入。');
    return result;
  }
  function validDate(value) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
    const parsed = new Date(value + 'T00:00:00Z');
    return !isNaN(parsed.getTime()) && parsed.toISOString().slice(0, 10) === value;
  }
  function parse(headers, rows, map, helpers) {
    let desc = describe(headers || [], map);
    if (!desc.valid) {
      const all = [headers || [], ...(rows || [])];
      const found = all.slice(0, 100).findIndex(row => describe(row || [], map).valid);
      if (found < 0) throw new Error('销售表头不完整或列无法识别，已阻止按位置猜测导入。');
      headers = all[found]; rows = all.slice(found + 1); desc = describe(headers, map);
    }
    const recs = [];
    (rows || []).forEach((arr, idx) => {
      if (!arr || arr.every(v => v == null || v === '')) return;
      if (describe(arr, map).valid) return; // repeated export header
      const rec = {name: '', category: '', productType: '', ledger: '', rep1Name: '', rep2Name: '', rep1Group: '', rep2Group: '', cost: 0};
      desc.keys.forEach((key, ci) => {
        if (!key) return;
        const value = arr[ci];
        if (['qty', 'amount', 'cost'].includes(key)) rec[key] = numeric(value, idx + 1, key);
        else if (key === 'soldDate') rec[key] = helpers.date(value);
        else rec[key] = String(value == null ? '' : value).trim();
      });
      if (/^(?:total|subtotal|grand total|合计|总计)$/i.test(rec.itemNo || '') && !rec.invoice && !rec.soldDate) return;
      if (!(rec.itemNo || rec.invoice || rec.amount)) return;
      if (!validDate(rec.soldDate || '')) throw new Error('明细第' + (idx + 1) + '行销售日期无效，可能含参数、公式或错位列，未导入。');
      for (const key of ['rep1', 'rep2']) {
        const code = rec[key] || '';
        if (code && (!/^[A-Za-z0-9][A-Za-z0-9_-]{0,63}$/.test(code) || /^(?:repcode|rep1code|rep2code|rep1|rep2|sales1|sales2)$/.test(norm(code))))
          throw new Error('明细第' + (idx + 1) + '行销售代码格式异常，可能把客户名称当成销售代码，未导入。');
      }
      if (/^\d+$/.test(rec.rep1 || '') && /^C\d+$/i.test(rec.rep1Name || ''))
        throw new Error('明细第' + (idx + 1) + '行流水号和客户编号落入销售字段，已阻止错位导入。');
      recs.push(rec);
    });
    recs.unmapped = headers.map((h, ci) => ({h: String(h || '').trim(), ci}))
      .filter(x => x.h && !desc.keys[x.ci] && (rows || []).some(row => row && row[x.ci] != null && row[x.ci] !== '')).map(x => x.h);
    return recs;
  }
  root.SALES_IMPORT_GUARD = {select, parse, describe};
})(typeof window === 'undefined' ? globalThis : window);
