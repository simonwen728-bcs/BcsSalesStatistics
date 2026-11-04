/* Cabinet quantities grouped by Color + resolved Name; Door and ASET count 1:1. */
window.SKU_MIX = (() => {
  'use strict';
  const clean = v => String(v == null ? '' : v).trim().replace(/\s+/g, ' ').toUpperCase();
  const escape = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

  function resolveModel(master, salesDescription) {
    const name = clean(master.Name).replace(/\s+DOOR$/, '').trim();
    if (!name) return '';
    const candidates = name.split(/\/(?=[A-Z]+\d)/);
    if (candidates.length === 1) return name;
    const desc = clean(salesDescription || master.Description);
    const hits = candidates.filter(c => new RegExp('^' + escape(c) + '(?:\\s|$)').test(desc));
    if (hits.length === 1) return hits[0];
    if (name === 'B15/VB15') return /^B15(?:\s|$)/.test(desc) ? 'B15' : 'VB15';
    return '';
  }

  function generate(data, opts, t) {
    const H = window.STATS.helpers;
    const zh = t('skuBestSeller') === 'SKU畅销分析';
    const range = H.range(opts.start, opts.end);
    if (!range.success) return { error: range.message };
    const master = opts.cabinetItems || {};
    const groups = new Map(), colors = new Map(), errors = new Set();
    const groupMap = H.sales1Groups(opts.reps || { s1: [], s2: [] });
    let included = 0;
    for (const r of data) {
      const date = H.date(r.soldDate);
      if (!date || !H.inRange(date, range) || H.isIC(clean(r.rep1), groupMap)) continue;
      const code = clean(r.itemNo);
      if (!/^(?:CC|ASET)/.test(code)) continue;
      const item = master[code];
      if (!item) { errors.add(code); continue; }
      const kind = code.startsWith('ASET') ? 'ASET' : clean(item.Category) === 'DOOR' ? 'DOOR' : '';
      if (!kind) continue;
      const model = resolveModel(item, r.description), color = clean(item.Color);
      if (!model || !color) { errors.add(code); continue; }
      const key = JSON.stringify([color, model]);
      if (!groups.has(key)) groups.set(key, { color, model, door: 0, aset: 0, total: 0 });
      if (!colors.has(color)) colors.set(color, 0);
      const qty = H.amount(r.qty), value = groups.get(key);
      value[kind === 'DOOR' ? 'door' : 'aset'] += qty;
      value.total += qty;
      colors.set(color, colors.get(color) + qty);
      included++;
    }
    if (errors.size) return { error: (zh ? '以下SKU的Name、颜色或共用型号无法确认，请更新Pull品名资料：' : 'Missing or ambiguous SKU name/color mapping: ') + [...errors].slice(0, 8).join(', ') };
    const values = [[...(zh ? ['颜色', 'Item', '门数量', 'ASET数量', '合计数量', '本颜色合计', '本颜色占比'] : ['Color', 'Item', 'Door Qty', 'ASET Qty', 'Total Qty', 'Color Total Qty', 'Share of Color'])]];
    const list = [...groups.values()].sort((a, b) => (colors.get(b.color) - colors.get(a.color)) || a.color.localeCompare(b.color) || b.total - a.total || a.model.localeCompare(b.model, undefined, { numeric: true }));
    let door = 0, aset = 0;
    for (const item of list) {
      const colorTotal = colors.get(item.color);
      values.push([item.color, item.model, item.door, item.aset, item.total, colorTotal, colorTotal === 0 ? '' : item.total / colorTotal]);
      door += item.door; aset += item.aset;
    }
    values.push([zh ? '合计' : 'Total', '', door, aset, door + aset, '', '']);
    const last = values.length;
    return {
      sheet: 'SKU畅销分析', values, merges: [],
      numFmt: [{ range: 'C2:F' + last, fmt: '#,##0;-#,##0;0' }, { range: 'G2:G' + last, fmt: '0.00%;-0.00%;0.00%' }],
      bold: ['A1:G1', 'A' + last + ':G' + last],
      bg: [{ range: 'A1:G1', color: '#d9eaf7' }, { range: 'A' + last + ':G' + last, color: '#d9e6f3' }],
      align: [], widths: [[1, 250], [2, 200], [3, 120], [4, 120], [5, 140], [6, 160], [7, 140]],
      freeze: { rows: 1, cols: 2 }, skuMixAlgorithm: 1,
      info: (opts.start || (zh ? '全部历史' : 'All history')) + (opts.end ? ' - ' + opts.end : '') + (zh ? '；仅DOOR＋ASET，按1:1合并；排除IC；数量保留退货正负号。' : '; Door + ASET, counted 1:1; IC excluded; returns retain negative quantity.'),
      summary: [[zh ? '颜色数' : 'Colors', colors.size], [zh ? 'Item数' : 'Items', list.length], [zh ? '合计数量' : 'Total Qty', door + aset]],
    };
  }
  return { generate, resolveModel };
})();
