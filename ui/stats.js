// Statistics engine: faithful JS port of google-excel.appScript (V59) reports.
// Each generator returns a spec: {sheet, values, merges, numFmt, bold, bg, align, widths, freeze}
window.STATS = (function () {
  'use strict';

  // ---------------- SYS common (port of SYS in appScript) ----------------
  function date(v) {
    if (v === null || v === undefined || v === '') return '';
    if (v instanceof Date && !isNaN(v.getTime())) {
      return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0') + '-' + String(v.getDate()).padStart(2, '0');
    }
    if (typeof v === 'number') {
      if (v >= 20000 && v <= 80000) {           // Excel 日期序列号（1954~2088 年）
        const d = new Date(Math.round((v - 25569) * 86400000));
        return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0');
      }
      const s8 = String(Math.round(v));
      if (/^\d{8}$/.test(s8)) return s8.slice(0, 4) + '-' + s8.slice(4, 6) + '-' + s8.slice(6, 8);
      return '';
    }
    const s = String(v).trim();
    const MON = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
    let m = s.match(/^(\d{4})[-\/.年](\d{1,2})[-\/.月](\d{1,2})日?/);                       // 2024-06-29 / 2024/6/29 / 2024年6月29日
    if (m) return m[1] + '-' + m[2].padStart(2, '0') + '-' + m[3].padStart(2, '0');
    m = s.match(/^(\d{1,2})[-/ ]([A-Za-z]{3,9})\.?[-/, ]+(\d{2,4})$/);                     // 06-Jan-2026
    if (m && MON[m[2].slice(0, 3).toLowerCase()]) {
      const y = Number(m[3]);
      return (y < 100 ? y + 2000 : y) + '-' + String(MON[m[2].slice(0, 3).toLowerCase()]).padStart(2, '0') + '-' + m[1].padStart(2, '0');
    }
    m = s.match(/^([A-Za-z]{3,9})\.?[-/ ]+(\d{1,2}),?\s+(\d{2,4})$/);                      // Jan 7, 2026
    if (m && MON[m[1].slice(0, 3).toLowerCase()]) {
      const y = Number(m[3]);
      return (y < 100 ? y + 2000 : y) + '-' + String(MON[m[1].slice(0, 3).toLowerCase()]).padStart(2, '0') + '-' + String(m[2]).padStart(2, '0');
    }
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);                                      // m/d/yyyy 美式；首段>12 按日/月
    if (m) {
      let y = Number(m[3]); if (y < 100) y += 2000;
      let mo = Number(m[1]), d = Number(m[2]);
      if (mo > 12 && d <= 12) { mo = Number(m[2]); d = Number(m[1]); }
      return y + '-' + String(mo).padStart(2, '0') + '-' + String(d).padStart(2, '0');
    }
    m = s.match(/^(\d{4})(\d{2})(\d{2})$/);                                                // 20240629
    if (m) return m[1] + '-' + m[2] + '-' + m[3];
    return '';
  }

  function amount(v) {
    if (typeof v === 'string') {
      v = v.replace(/\$/g, '').replace(/,/g, '').replace(/\(/g, '-').replace(/\)/g, '').trim();
    }
    return Number(v) || 0;
  }

  function range(startStr, endStr) {
    const rawStart = startStr, rawEnd = endStr;
    const start = rawStart ? date(rawStart) : '';
    const end = rawEnd ? date(rawEnd) : '';
    if (rawStart && !start) return { success: false, message: 'BAD_START' };
    if (rawEnd && !end) return { success: false, message: 'BAD_END' };
    if (start && end && start > end) return { success: false, message: 'BAD_ORDER' };
    return { success: true, start, end };
  }

  function inRange(dt, rg) {
    if (!dt) return false;
    if (rg.start && dt < rg.start) return false;
    if (rg.end && dt > rg.end) return false;
    return true;
  }

  function sales1Groups(reps) {
    const map = {};
    (reps.s1 || []).forEach((r) => {
      const code = String(r.code || '').trim().toUpperCase();
      const group = String(r.group || '').trim().toUpperCase();
      if (code) map[code] = group;
    });
    return map;
  }

  function isIC(code, groupMap) {
    code = String(code || '').trim().toUpperCase();
    return code === 'IC' || groupMap[code] === 'IC';
  }

  function normType(t) {
    return String(t || '').trim().toUpperCase().replace(/\s+/g, '');
  }
  function isCabinetType(t) {
    const x = normType(t);
    return x === 'CABINET' || x === 'CABINETOTHERS';
  }

  // ---------------- A1 helpers ----------------
  function colLetter(n) {
    // 1-based
    let s = '';
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }
  function a1(r, c) { return colLetter(c) + r; }
  function rng(r1, c1, r2, c2) { return a1(r1, c1) + ':' + a1(r2, c2); }

  function fmtMoney(pattern) { return pattern || '#,##0.00'; }

  function baseSpec(sheet) {
    return { sheet, values: [], merges: [], numFmt: [], bold: [], bg: [], align: [], widths: [], freeze: null };
  }

  function monthKeysBetween(startYm, endYm) {
    const keys = [];
    let y = Number(startYm.split('-')[0]), m = Number(startYm.split('-')[1]);
    const ey = Number(endYm.split('-')[0]), em = Number(endYm.split('-')[1]);
    while (y < ey || (y === ey && m <= em)) {
      keys.push(y + '-' + String(m).padStart(2, '0'));
      m++; if (m > 12) { m = 1; y++; }
    }
    return keys;
  }

  function ytdWindow() {
    // TTM 口径与全局一致：上个月末往前 12 个月（数据滞后一个月）
    const now = new Date();
    const pm = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const start = new Date(pm.getFullYear(), pm.getMonth() - 11, 1);
    const endLast = new Date(pm.getFullYear(), pm.getMonth() + 1, 0);
    const pad = (n) => String(n).padStart(2, '0');
    return { start: start.getFullYear() + '-' + pad(start.getMonth() + 1) + '-01',
      end: endLast.getFullYear() + '-' + pad(endLast.getMonth() + 1) + '-01',
      ytdEndInclusive: endLast.getFullYear() + '-' + pad(endLast.getMonth() + 1) + '-' + pad(endLast.getDate()) };
  }

  // ============================================================
  // 1) 年度统计 (annual) — sheet 年度销售总表
  // Fixed: reads ALL history (no date range), groups by year/month,
  // YTD = rolling 12 months, grand total.
  // ============================================================
  function annual(data, opts, t) {
    const spec = baseSpec('年度销售总表');
    // 年度统计固定读全部历史（用户口径：不需要时间段）
    // categories: distinct Type values in data
    const catSet = new Set();
    data.forEach((r) => {
      const c = String(r.type || '').trim().toUpperCase();
      if (c) catSet.add(c);
    });
    const categories = Array.from(catSet).sort();
    const makeStat = () => {
      const cats = {};
      categories.forEach((k) => (cats[k] = 0));
      return { total: 0, ic: 0, net: 0, cats };
    };
    const addRow = (stat, r) => {
      const amt = amount(r.amount);
      const sales1 = String(r.rep1 || '').trim().toUpperCase();
      stat.total += amt;
      if (sales1 === 'IC') stat.ic += amt;
      else {
        stat.net += amt;
        const cat = String(r.type || '').trim().toUpperCase();
        if (cat in stat.cats) stat.cats[cat] += amt;
      }
    };

    const years = {};
    const grand = makeStat();
    data.forEach((r) => {
      const dt = date(r.soldDate);
      if (!dt) return;
      const year = dt.substring(0, 4);
      const month = dt.substring(0, 7);
      if (!years[year]) years[year] = { total: makeStat(), months: {} };
      if (!years[year].months[month]) years[year].months[month] = makeStat();
      addRow(years[year].total, r);
      addRow(years[year].months[month], r);
      addRow(grand, r);
    });

    const yearList = Object.keys(years).sort();
    if (!yearList.length) return { error: t('statsNoData') };

    const yw = ytdWindow();
    const ytd = makeStat();
    data.forEach((r) => {
      const dt = date(r.soldDate);
      if (!dt) return;
      // rolling 12 months: [ytdStart, ytdEnd]
      if (dt < yw.start || dt > yw.ytdEndInclusive) return;
      addRow(ytd, r);
    });

    // header: [期间, 合计, IC, Net, (CAT, CAT %)*] —— 统计表固定中文，不随界面语言
    const header = ['期间', '合计', 'IC', 'Net'];
    categories.forEach((c) => { header.push(c, c + ' %'); });
    spec.values.push(header);

    const boldRows = [];
    const yearGroups = [];   // 年份行(1基)与该年月份行(1基)列表——供前端点击年份行折叠/展开
    yearList.forEach((year) => {
      spec.values.push([year, years[year].total.total, years[year].total.ic, years[year].total.net,
        ...categories.flatMap((c) => [years[year].total.cats[c] || 0, years[year].total.net ? (years[year].total.cats[c] || 0) / years[year].total.net : 0])]);
      boldRows.push(spec.values.length);
      const g = { row: spec.values.length, months: [] };
      Object.keys(years[year].months).sort().forEach((month) => {
        const m = years[year].months[month];
        spec.values.push([month, m.total, m.ic, m.net, ...categories.flatMap((c) => [m.cats[c] || 0, m.net ? (m.cats[c] || 0) / m.net : 0])]);
        g.months.push(spec.values.length);
      });
      yearGroups.push(g);
    });
    spec.yearGroups = yearGroups;   // 注意：不要给 annual 设 spec.info（提示走 toast）——info 会加 2 行偏移使 yearGroups 行号错位
    spec.values.push(['TTM', ytd.total, ytd.ic, ytd.net, ...categories.flatMap((c) => [ytd.cats[c] || 0, ytd.net ? (ytd.cats[c] || 0) / ytd.net : 0])]);
    boldRows.push(spec.values.length);
    spec.values.push(['合计', grand.total, grand.ic, grand.net, ...categories.flatMap((c) => [grand.cats[c] || 0, grand.net ? (grand.cats[c] || 0) / grand.net : 0])]);
    boldRows.push(spec.values.length);

    const lastRow = spec.values.length;
    const lastCol = header.length;
    spec.bold.push('A1:' + colLetter(lastCol) + '1');
    boldRows.forEach((r) => spec.bold.push('A' + r + ':' + colLetter(lastCol) + r));
    const n = lastRow - 1;
    spec.numFmt.push({ range: 'B2:' + colLetter(lastCol) + lastRow, fmt: '#,##0.00' });
    // % columns: every 2nd col starting from 6 (F,H,...)
    for (let c = 6; c <= lastCol; c += 2) {
      spec.numFmt.push({ range: colLetter(c) + '2:' + colLetter(c) + lastRow, fmt: '0.00%' });
      spec.align.push({ range: colLetter(c) + '2:' + colLetter(c) + lastRow, horiz: 'right' });
    }
    spec.widths = [[1, 130], ...Array.from({ length: lastCol - 1 }, (_, i) => [i + 2, 115])];
    spec.freeze = { rows: 1, cols: 0 };
    spec.bg.push({ range: 'A1:' + colLetter(lastCol) + '1', color: '#1f4e78' });
    return spec;
  }

  // ============================================================
  // 2) 销售业绩 (sales performance) — sheet 销售业绩
  // Left: Sales1 grouped by Group (House first) with subtotals.
  // Right: inside Sales2 matrix; House-only inside subtotal; HO 合计.
  // ============================================================
  function salesPerformance(data, opts, t) {
    const rg = range(opts.start, opts.end);
    if (!rg.success) return { error: t('statEmpty').replace('{msg}', rg.message) };
    const spec = baseSpec('销售业绩');

    const reps = opts.reps || { s1: [], s2: [] };
    const sales1Info = {};
    const insideSales2 = new Set();
    const sales2Names = {};
    (reps.s1 || []).forEach((r) => {
      const s1 = String(r.code || '').trim().toUpperCase();
      const name = String(r.name || '').trim();
      const group = String(r.group || '').trim();
      if (s1) sales1Info[s1] = { name: name || s1, group: group || (s1 === 'HO' ? 'House' : 'Other Branch') };
    });
    (reps.s2 || []).forEach((r) => {
      const s2 = String(r.code || '').trim().toUpperCase();
      const name = String(r.name || '').trim();
      const group = String(r.group || '').trim().toLowerCase();
      if (s2) {
        sales2Names[s2] = name || s2;
        if (group === 'inside') insideSales2.add(s2);   // 只认分组表数据，不再写死 LP（用户要求，组别变化自动跟随）
      }
    });

    const sales1Amount = {};
    const matrix = {};
    const sales2Totals = {};
    let grandSales1 = 0;
    let grandInside = 0;
    insideSales2.forEach((s2) => (sales2Totals[s2] = 0));

    data.forEach((r) => {
      const dt = date(r.soldDate);
      if (!inRange(dt, rg)) return;
      const sales1 = String(r.rep1 || '').trim().toUpperCase();
      const sales2 = String(r.rep2 || '').trim().toUpperCase();
      const amt = amount(r.amount);
      if (!sales1 || !amt || sales1 === 'IC') return;
      if (!sales1Info[sales1]) {
        sales1Info[sales1] = { name: String(r.rep1Name || sales1).trim(), group: sales1 === 'HO' ? 'House' : 'Other Branch' };
      }
      const info = sales1Info[sales1];
      if (!info.group || info.group.trim().toUpperCase() === 'IC') return;
      sales1Amount[sales1] = (sales1Amount[sales1] || 0) + amt;
      grandSales1 += amt;
      if (insideSales2.has(sales2)) {
        if (!matrix[sales1]) matrix[sales1] = {};
        matrix[sales1][sales2] = (matrix[sales1][sales2] || 0) + amt;
        sales2Totals[sales2] = (sales2Totals[sales2] || 0) + amt;
        grandInside += amt;
      }
    });

    const sales2List = Array.from(insideSales2).filter((s2) => sales2Totals[s2] !== 0).sort();
    const groupMap = {};
    Object.keys(sales1Amount).forEach((code) => {
      const g = sales1Info[code].group;
      if (!groupMap[g]) groupMap[g] = [];
      groupMap[g].push(code);
    });
    const isHouseGroup = (g) => String(g || '').trim().toLowerCase() === 'house';
    const groups = Object.keys(groupMap).sort((a, b) => {
      const ha = isHouseGroup(a), hb = isHouseGroup(b);
      if (ha && !hb) return -1;
      if (!ha && hb) return 1;
      return a.localeCompare(b);
    });

    const leftRows = [['部门1', '销售1', '金额', '部门占比', '总额占比']];
    const subtotalRowIdx = [];

    groups.forEach((groupName) => {
      const codes = groupMap[groupName];
      codes.sort((a, b) => {
        if (isHouseGroup(groupName)) {
          if (a === 'HO' && b !== 'HO') return -1;
          if (a !== 'HO' && b === 'HO') return 1;
        }
        return sales1Info[a].name.localeCompare(sales1Info[b].name);
      });
      const groupTotal = codes.reduce((sum, c) => sum + sales1Amount[c], 0);
      codes.forEach((code) => {
        const amt = sales1Amount[code];
        leftRows.push([groupName, sales1Info[code].name, amt, groupTotal ? amt / groupTotal : 0, grandSales1 ? amt / grandSales1 : 0]);
      });
      leftRows.push([groupName + ' 小计', '', groupTotal, 1, grandSales1 ? groupTotal / grandSales1 : 0]);
      subtotalRowIdx.push(leftRows.length);
    });

    // right side（组名大小写不敏感：用户库里可能是 HOUSE/House/other branch 等写法，原版只认 'House' 会整块落空）
    const houseKey = Object.keys(groupMap).find((g) => isHouseGroup(g)) || null;
    const houseCodes = houseKey ? groupMap[houseKey] : [];
    const houseInsideTotals = {};
    sales2List.forEach((s2) => {
      houseInsideTotals[s2] = houseCodes.reduce((sum, c) => sum + (matrix[c] ? matrix[c][s2] || 0 : 0), 0);
    });
    const grandHouseInside = sales2List.reduce((sum, s2) => sum + (houseInsideTotals[s2] || 0), 0);

    // 右块（各销售2列+Inside 小计+两个占比行）完全由分组表驱动：没有任何 Inside 销售2 时整块省略，
    // 只保留 Sale 1 总计行（无 HO 合计/占比的概念）
    const rightRows = [];
    if (sales2List.length) {
      rightRows.push(['', ...sales2List.map((s2) => sales2Names[s2] || s2)]);
      for (let r = 1; r < leftRows.length; r++) {
        const deptLabel = String(leftRows[r][0] || '');
        const sales1Name = String(leftRows[r][1] || '');
        if (deptLabel.endsWith(' 小计') || deptLabel.endsWith(' Subtotal')) {
          const gname = deptLabel.replace(/ (小计|Subtotal)\s*$/, '').trim().toLowerCase();
          if (gname === 'house') {
            rightRows.push(['Inside 小计', ...sales2List.map((s2) => houseInsideTotals[s2] || 0)]);
          } else {
            rightRows.push(new Array(1 + sales2List.length).fill(''));
          }
          continue;
        }
        let code = '';
        Object.keys(sales1Info).some((c) => {
          if (sales1Info[c].name === sales1Name) { code = c; return true; }
          return false;
        });
        const rr = [''];
        sales2List.forEach((s2) => rr.push(code && matrix[code] ? matrix[code][s2] || 0 : 0));
        rightRows.push(rr);
      }
    }
    leftRows.push(['Sale 1 总计', '', grandSales1, '', 1]);
    if (sales2List.length) {
      rightRows.push(['部门占比', ...sales2List.map((s2) => (grandHouseInside ? (houseInsideTotals[s2] || 0) / grandHouseInside : 0))]);
      leftRows.push(['HO 合计', '', grandHouseInside, '', grandSales1 ? grandHouseInside / grandSales1 : 0]);
      rightRows.push(['总额占比', ...sales2List.map((s2) => (grandSales1 ? (houseInsideTotals[s2] || 0) / grandSales1 : 0))]);
    }

    const rowCount = Math.max(leftRows.length, rightRows.length);
    while (leftRows.length < rowCount) leftRows.push(['', '', '', '', '']);
    while (rightRows.length < rowCount) rightRows.push(sales2List.length ? new Array(1 + sales2List.length).fill('') : []);
    const output = leftRows.map((row, i) => row.concat(rightRows[i]));
    spec.values = output;

    const lastRow = output.length;
    const lastCol = output[0].length;
    spec.bold.push(rng(1, 1, 1, lastCol));
    spec.bg.push({ range: rng(1, 1, 1, lastCol), color: '#1f4e78' });
    if (lastRow > 1) {
      spec.numFmt.push({ range: 'C2:C' + lastRow, fmt: '#,##0.00' });
      spec.numFmt.push({ range: 'D2:E' + lastRow, fmt: '0.00%' });
      if (sales2List.length) {
        // 右块结构：F=行标签列，G..=各销售2数值列 → 金额/百分比从第 7 列起（原从 F 起差一列，最后一列 0 裸露）
        const c0 = colLetter(7), c1 = colLetter(6 + sales2List.length);
        spec.numFmt.push({ range: c0 + '2:' + c1 + (lastRow - 2), fmt: '#,##0.00' });
        spec.numFmt.push({ range: c0 + (lastRow - 1) + ':' + c1 + lastRow, fmt: '0.00%' });
      }
    }
    subtotalRowIdx.forEach((r) => {
      spec.bold.push(rng(r, 1, r, lastCol));
      spec.bg.push({ range: rng(r, 1, r, lastCol), color: '#d9eaf7' });
    });
    spec.bold.push(rng(lastRow - 1, 1, lastRow, lastCol));
    spec.bg.push({ range: rng(lastRow - 1, 1, lastRow, lastCol), color: '#d9eaf7' });
    spec.freeze = { rows: 1, cols: 2 };
    spec.widths = Array.from({ length: lastCol }, (_, c) => [c + 1, [150, 150, 180, 100, 100, 150][c] || 150]);
    return spec;
  }

  // ============================================================
  // 3) 客户订单统计 (customer orders) — sheet 客户订单统计
  // ============================================================
  function customerOrders(data, opts, t) {
    const rg = range(opts.start, opts.end);
    if (!rg.success) return { error: t('statEmpty').replace('{msg}', rg.message) };
    const spec = baseSpec('客户订单统计');
    const reps = opts.reps || { s1: [], s2: [] };
    const sales1Names = {}, sales1GroupsMap = {}, sales2Names = {};
    const insideSales2 = new Set();
    (reps.s1 || []).forEach((r) => {
      const s1 = String(r.code || '').trim().toUpperCase();
      if (s1) {
        sales1Names[s1] = String(r.name || s1).trim();
        sales1GroupsMap[s1] = String(r.group || '').trim().toUpperCase();
      }
    });
    (reps.s2 || []).forEach((r) => {
      const s2 = String(r.code || '').trim().toUpperCase();
      const name = String(r.name || s2).trim();
      const group = String(r.group || '').trim().toLowerCase();
      if (s2) {
        sales2Names[s2] = name;
        if (group === 'inside') insideSales2.add(s2);   // 只认分组表数据，不再写死 LP
      }
    });

    const newStat = () => ({ invoices: new Set(), customers: new Set(), total: 0, cabinet: 0, stone: 0 });
    const s1Stats = {}, s2Stats = {};
    const s1Total = newStat(), s2Total = newStat();

    data.forEach((r) => {
      const dt = date(r.soldDate);
      if (!inRange(dt, rg)) return;
      const invoice = String(r.invoice || '').trim().toUpperCase();
      const customer = String(r.custNo || '').trim().toUpperCase();
      const sales1 = String(r.rep1 || '').trim().toUpperCase();
      const sales2 = String(r.rep2 || '').trim().toUpperCase();
      const amt = amount(r.amount);
      if (!sales1 || sales1 === 'IC' || sales1GroupsMap[sales1] === 'IC' || !amt) return;
      const isCabinet = isCabinetType(r.type);
      if (!sales1Names[sales1]) sales1Names[sales1] = String(r.rep1Name || sales1).trim();
      if (!s1Stats[sales1]) s1Stats[sales1] = newStat();
      const s1 = s1Stats[sales1];
      if (invoice) { s1.invoices.add(invoice); s1Total.invoices.add(invoice); }
      if (customer) { s1.customers.add(customer); s1Total.customers.add(customer); }
      s1.total += amt; s1Total.total += amt;
      if (isCabinet) { s1.cabinet += amt; s1Total.cabinet += amt; } else { s1.stone += amt; s1Total.stone += amt; }
      if (insideSales2.has(sales2)) {
        if (!s2Stats[sales2]) s2Stats[sales2] = newStat();
        const s2 = s2Stats[sales2];
        if (invoice) { s2.invoices.add(invoice); s2Total.invoices.add(invoice); }
        if (customer) { s2.customers.add(customer); s2Total.customers.add(customer); }
        s2.total += amt; s2Total.total += amt;
        if (isCabinet) { s2.cabinet += amt; s2Total.cabinet += amt; } else { s2.stone += amt; s2Total.stone += amt; }
      }
    });

    const rows = [['Sales 1', 'Sales Order', 'Customers', 'Total Sales', '%', 'Cabinets', '%', 'Stone', '%']];
    const groupPriority = { HOUSE: 1, OUTSIDE: 2, 'OTHER BRANCH': 3 };
    const getGroup = (code) => String(sales1GroupsMap[code] || (code === 'HO' ? 'HOUSE' : 'OTHER BRANCH')).trim().toUpperCase();
    const addStat = (target, source) => {
      source.invoices.forEach((v) => target.invoices.add(v));
      source.customers.forEach((v) => target.customers.add(v));
      target.total += source.total; target.cabinet += source.cabinet; target.stone += source.stone;
    };
    const s1Row = (code, x) => [sales1Names[code] || code, x.invoices.size, x.customers.size, x.total,
      s1Total.total ? x.total / s1Total.total : 0, x.cabinet, s1Total.cabinet ? x.cabinet / s1Total.cabinet : 0, x.stone, s1Total.stone ? x.stone / s1Total.stone : 0];
    const subRow = (x) => ['Sub Total:', x.invoices.size, x.customers.size, x.total, '', x.cabinet, x.total ? x.cabinet / x.total : 0, x.stone, x.total ? x.stone / x.total : 0];

    const s1List = Object.keys(s1Stats).sort((a, b) => {
      const gA = getGroup(a), gB = getGroup(b);
      const pA = groupPriority[gA] || 99, pB = groupPriority[gB] || 99;
      if (pA !== pB) return pA - pB;
      if (pA === 99 && gA !== gB) return gA.localeCompare(gB);
      if (gA === 'HOUSE') {
        if (a === 'HO' && b !== 'HO') return -1;
        if (a !== 'HO' && b === 'HO') return 1;
      }
      return String(sales1Names[a] || a).localeCompare(String(sales1Names[b] || b), undefined, { sensitivity: 'base' });
    });

    const subtotalRows = [];
    let currentGroup = null, currentGroupStat = null;
    s1List.forEach((code) => {
      const groupName = getGroup(code);
      if (currentGroup !== null && groupName !== currentGroup) {
        rows.push(subRow(currentGroupStat));
        subtotalRows.push(rows.length);
        rows.push(new Array(9).fill(''));
        currentGroupStat = newStat();
      }
      if (currentGroup === null) currentGroupStat = newStat();
      currentGroup = groupName;
      const x = s1Stats[code];
      rows.push(s1Row(code, x));
      addStat(currentGroupStat, x);
    });
    if (currentGroupStat) {
      rows.push(subRow(currentGroupStat));
      subtotalRows.push(rows.length);
    }
    rows.push(['合计', s1Total.invoices.size, s1Total.customers.size, s1Total.total, '',
      s1Total.cabinet, s1Total.total ? s1Total.cabinet / s1Total.total : 0, s1Total.stone, s1Total.total ? s1Total.stone / s1Total.total : 0]);
    const s1GrandRow = rows.length;
    rows.push(new Array(9).fill(''));
    const s2HeaderRow = rows.length + 1;
    rows.push(['Sales 2', 'Sales Order', 'Customers', 'Total Sales', '%', 'Cabinets', '%', 'Stone', '%']);
    const s2List = Object.keys(s2Stats).sort((a, b) => (sales2Names[a] || a).localeCompare(sales2Names[b] || b));
    s2List.forEach((code) => {
      const x = s2Stats[code];
      rows.push([sales2Names[code] || code, x.invoices.size, x.customers.size, x.total,
        s2Total.total ? x.total / s2Total.total : 0, x.cabinet, s2Total.cabinet ? x.cabinet / s2Total.cabinet : 0, x.stone, s2Total.stone ? x.stone / s2Total.stone : 0]);
    });
    rows.push(['合计', s2Total.invoices.size, s2Total.customers.size, s2Total.total, '',
      s2Total.cabinet, s2Total.total ? s2Total.cabinet / s2Total.total : 0, s2Total.stone, s2Total.total ? s2Total.stone / s2Total.total : 0]);
    spec.values = rows;

    const lastRow = rows.length;
    spec.bold.push(rng(1, 1, 1, 9));
    spec.bg.push({ range: rng(1, 1, 1, 9), color: '#1f4e78' });
    spec.bold.push(rng(s2HeaderRow, 1, s2HeaderRow, 9));
    spec.bg.push({ range: rng(s2HeaderRow, 1, s2HeaderRow, 9), color: '#1f4e78' });
    subtotalRows.forEach((r) => {
      spec.bold.push(rng(r, 1, r, 9));
      spec.bg.push({ range: rng(r, 1, r, 9), color: '#d9eaf7' });
    });
    spec.bold.push(rng(s1GrandRow, 1, s1GrandRow, 9));
    spec.bg.push({ range: rng(s1GrandRow, 1, s1GrandRow, 9), color: '#d9eaf7' });
    spec.bold.push(rng(lastRow, 1, lastRow, 9));
    spec.bg.push({ range: rng(lastRow, 1, lastRow, 9), color: '#d9eaf7' });
    for (let r = 2; r <= lastRow; r++) {
      if (r === s2HeaderRow || rows[r - 1].every((v) => v === '')) continue;
      spec.numFmt.push({ range: 'B' + r + ':C' + r, fmt: '#,##0' });
      spec.numFmt.push({ range: 'D' + r + ':D' + r, fmt: '#,##0.00' });
      if (rows[r - 1][4] !== '') spec.numFmt.push({ range: 'E' + r + ':E' + r, fmt: '0.00%' });
      spec.numFmt.push({ range: 'F' + r + ':F' + r, fmt: '#,##0.00' });
      spec.numFmt.push({ range: 'G' + r + ':G' + r, fmt: '0.00%' });
      spec.numFmt.push({ range: 'H' + r + ':H' + r, fmt: '#,##0.00' });
      spec.numFmt.push({ range: 'I' + r + ':I' + r, fmt: '0.00%' });
    }
    spec.widths = [150,100,100,180,100,150,100,150,100].map((w,c)=>[c+1,w]);
    spec.freeze = { rows: 1, cols: 0 };
    return spec;
  }

  // ============================================================
  // 4) 销售查询/分类统计 (category query) — sheet 销售查询
  // mode: amount|qty ; category: ALL|<TYPE>
  // ============================================================
  function categoryQuery(data, opts, t) {
    const mode = opts.mode === 'qty' ? 'qty' : 'amount';
    const filter = String(opts.category || 'ALL').trim().toUpperCase();
    const rg = range(opts.start, opts.end);
    if (!rg.success) return { error: t('statEmpty').replace('{msg}', rg.message) };
    const spec = baseSpec('销售查询');

    const allYearsSet = new Set();
    let minDate = '', maxDate = '';
    data.forEach((r) => {
      const dt = date(r.soldDate);
      if (!dt) return;
      allYearsSet.add(dt.substring(0, 4));
      if (!minDate || dt < minDate) minDate = dt;
      if (!maxDate || dt > maxDate) maxDate = dt;
    });
    const allYears = Array.from(allYearsSet).sort((a, b) => Number(a) - Number(b));
    if (!allYears.length) return { error: t('statEmpty').replace('{msg}', 'no dates') };

    const effectiveStart = rg.start || minDate;
    const effectiveEnd = rg.end || maxDate;
    const monthKeys = monthKeysBetween(effectiveStart.substring(0, 7), effectiveEnd.substring(0, 7));

    const yw = ytdWindow();
    const ytdStart = yw.start, ytdEnd = yw.ytdEndInclusive;

    const items = {};
    const newItem = () => ({ category: '', item: '', description: '', total: 0, years: {}, months: {}, ytd: 0, selectedSeen: false, yearSeen: new Set(), monthSeen: new Set(), ytdSeen: false });
    data.forEach((r) => {
      const dt = date(r.soldDate);
      if (!dt) return;
      const itemNo = String(r.itemNo || '').trim();
      if (!itemNo) return;
      const description = String(r.description || '').trim();
      const category = String(r.type || '').trim();
      if (filter !== 'ALL' && category.trim().toUpperCase() !== filter) return;
      const value = mode === 'qty' ? amount(r.qty) : amount(r.amount);
      if (!items[itemNo]) {
        items[itemNo] = newItem();
        items[itemNo].item = itemNo;
        items[itemNo].category = category;
        items[itemNo].description = description;
      }
      const x = items[itemNo];
      if (dt >= ytdStart && dt <= ytdEnd) { x.ytd += value; x.ytdSeen = true; }
      if (!inRange(dt, rg)) return;
      x.selectedSeen = true;
      x.total += value;
      const year = dt.substring(0, 4);
      const monthKey = dt.substring(0, 7);
      x.years[year] = (x.years[year] || 0) + value;
      x.months[monthKey] = (x.months[monthKey] || 0) + value;
      x.yearSeen.add(year);
      x.monthSeen.add(monthKey);
    });

    const itemList = Object.values(items)
      .filter((x) => x.selectedSeen)
      .sort((a, b) => {
        const ca = String(a.category || '').toUpperCase();
        const cb = String(b.category || '').toUpperCase();
        if (ca !== cb) return ca.localeCompare(cb);
        return String(a.item).localeCompare(String(b.item), undefined, { numeric: true, sensitivity: 'base' });
      });
    if (!itemList.length) return { error: t('statEmpty').replace('{msg}', 'category') };

    const yearStartCol = 5;
    const ytdCol = yearStartCol + allYears.length;
    const monthStartCol = ytdCol + 1;
    const totalCols = 4 + allYears.length + 1 + monthKeys.length;

    const header1 = new Array(totalCols).fill('');
    const header2 = new Array(totalCols).fill('');
    header1[0] = t('colType') === '类别' ? '类别' : 'Category';
    header1[1] = t('colItemNo') === '品名' ? '品名' : 'Item';
    header1[2] = t('colDesc') === '描述' ? '描述' : 'Description';
    header1[3] = '合计';
    allYears.forEach((y, i) => (header1[yearStartCol - 1 + i] = y));
    header1[ytdCol - 1] = 'TTM';
    monthKeys.forEach((key, i) => {
      const parts = key.split('-');
      const year = parts[0];
      const month = Number(parts[1]);
      const idx = monthStartCol - 1 + i;
      header2[idx] = month + (t('colType') === '类别' ? '月' : 'M');
      if (i === 0 || monthKeys[i - 1].substring(0, 4) !== year) header1[idx] = year;
    });
    spec.values.push(header1, header2);

    const merges = ['A1:A2', 'B1:B2', 'C1:C2', 'D1:D2'];
    allYears.forEach((y, i) => merges.push(rng(1, yearStartCol + i, 2, yearStartCol + i)));
    merges.push(rng(1, ytdCol, 2, ytdCol));
    let gs = 0;
    while (gs < monthKeys.length) {
      const year = monthKeys[gs].substring(0, 4);
      let ge = gs;
      while (ge + 1 < monthKeys.length && monthKeys[ge + 1].substring(0, 4) === year) ge++;
      if (ge - gs + 1 > 1) merges.push(rng(1, monthStartCol + gs, 1, monthStartCol + ge));
      gs = ge + 1;
    }
    spec.merges = merges;

    const rows = itemList.map((x) => {
      const row = new Array(totalCols).fill('');
      row[0] = x.category;
      row[1] = x.item;
      row[2] = x.description;
      row[3] = x.total;
      allYears.forEach((y, i) => { if (x.yearSeen.has(y)) row[yearStartCol - 1 + i] = x.years[y] || 0; });
      if (x.ytdSeen) row[ytdCol - 1] = x.ytd;
      monthKeys.forEach((mk, i) => { if (x.monthSeen.has(mk)) row[monthStartCol - 1 + i] = x.months[mk] || 0; });
      return row;
    });
    spec.values.push(...rows);

    const totalRow = new Array(totalCols).fill('');
    totalRow[0] = '合计';
    for (let c = 3; c < totalCols; c++) totalRow[c] = rows.reduce((sum, row) => sum + (typeof row[c] === 'number' ? row[c] : 0), 0);
    const totalRowNo = rows.length + 3;
    while (spec.values.length < totalRowNo - 1) spec.values.push(new Array(totalCols).fill(''));
    spec.values.push(totalRow);

    const fmt = mode === 'qty' ? '#,##0' : '#,##0.00';   // 整数列不带小数点；zhFmt 会补隐藏零段
    spec.numFmt.push({ range: 'D3:' + colLetter(totalCols) + totalRowNo, fmt });
    spec.bold.push(rng(totalRowNo, 1, totalRowNo, totalCols));
    spec.bg.push({ range: rng(totalRowNo, 1, totalRowNo, totalCols), color: '#d9e6f3' });
    spec.bold.push(rng(1, 1, 2, totalCols));
    spec.bg.push({ range: rng(1, 1, 2, totalCols), color: '#D9E8F5' });
    spec.align.push({ range: rng(1, 1, 2, totalCols), horiz: 'center' });
    // 数据区对齐：金额表右齐，数量表居中（D列起所有数字列）
    spec.align.push({ range: 'D3:' + colLetter(totalCols) + totalRowNo, horiz: mode === 'qty' ? 'center' : 'right' });
    // 列宽写死（用户要求）：A=100, B=100, C=200, D-H(合计/年/TTM)=120, I起月列=100
    spec.widths = [[1, 100], [2, 100], [3, 200], [4, 120]];
    for (let c = yearStartCol; c <= ytdCol; c++) spec.widths.push([c, 120]);
    for (let c = monthStartCol; c <= totalCols; c++) spec.widths.push([c, 100]);
    spec.freeze = { rows: 2, cols: 3 };
    return spec;
  }

  // ============================================================
  // 5) 客户统计 (customer statistics) — sheet 客户统计
  // scope: all|cabinet|stone
  // ============================================================
  function customerStats(data, opts, t) {
    const scope = ['cabinet', 'stone'].includes(opts.scope) ? opts.scope : 'all';
    const rg = range(opts.start, opts.end);
    if (!rg.success) return { error: t('statEmpty').replace('{msg}', rg.message) };
    const spec = baseSpec('客户统计');
    const groupMap = sales1Groups(opts.reps || { s1: [], s2: [] });

    const newCustomer = () => ({ customerNo: '', customerName: '', sales1Amounts: {}, sales2Amounts: {}, totalAmount: 0, invoices: new Set(), cabinetAmount: 0, cabinetQty: 0, cabinetInvoices: new Set(), stoneAmount: 0, stoneQty: 0, stoneInvoices: new Set() });
    const customers = {};
    const total = { amount: 0, invoices: new Set(), cabinetAmount: 0, cabinetQty: 0, cabinetInvoices: new Set(), stoneAmount: 0, stoneQty: 0, stoneInvoices: new Set() };

    data.forEach((r) => {
      const sales1Code = String(r.rep1 || '').trim().toUpperCase();
      const sales2Code = String(r.rep2 || '').trim().toUpperCase();
      if (isIC(sales1Code, groupMap)) return;
      const dt = date(r.soldDate);
      if (!inRange(dt, rg)) return;
      const customerNo = String(r.custNo || '').trim();
      if (!customerNo) return;
      const customerName = String(r.custName || '').trim();
      const invoice = String(r.invoice || '').trim().toUpperCase();
      const isCabinet = isCabinetType(r.type);
      if (scope === 'cabinet' && !isCabinet) return;
      if (scope === 'stone' && isCabinet) return;
      const qty = amount(r.qty);
      const amt = amount(r.amount);
      if (!customers[customerNo]) {
        customers[customerNo] = newCustomer();
        customers[customerNo].customerNo = customerNo;
        customers[customerNo].customerName = customerName;
      }
      const x = customers[customerNo];
      if (!x.customerName && customerName) x.customerName = customerName;
      if (sales1Code) x.sales1Amounts[sales1Code] = (x.sales1Amounts[sales1Code] || 0) + amt;
      if (sales1Code === 'HO' && sales2Code) x.sales2Amounts[sales2Code] = (x.sales2Amounts[sales2Code] || 0) + amt;
      x.totalAmount += amt;
      total.amount += amt;
      if (invoice) { x.invoices.add(invoice); total.invoices.add(invoice); }
      if (isCabinet) {
        x.cabinetAmount += amt; x.cabinetQty += qty;
        total.cabinetAmount += amt; total.cabinetQty += qty;
        if (invoice) { x.cabinetInvoices.add(invoice); total.cabinetInvoices.add(invoice); }
      } else {
        x.stoneAmount += amt; x.stoneQty += qty;
        total.stoneAmount += amt; total.stoneQty += qty;
        if (invoice) { x.stoneInvoices.add(invoice); total.stoneInvoices.add(invoice); }
      }
    });

    const mainSales1 = (x) => {
      let best = '', bestAmt = -Infinity;
      Object.keys(x.sales1Amounts).forEach((c) => {
        if (x.sales1Amounts[c] > bestAmt) { bestAmt = x.sales1Amounts[c]; best = c; }
      });
      return best;
    };
    const mainSales2 = (x) => {
      if (mainSales1(x) !== 'HO') return '';
      let best = '', bestAmt = -Infinity;
      Object.keys(x.sales2Amounts || {}).forEach((c) => {
        const a = x.sales2Amounts[c] || 0;
        if (a > bestAmt) { bestAmt = a; best = c; }
      });
      return best;
    };

    const list = Object.values(customers)
      .filter((x) => x.invoices.size > 0 || x.totalAmount !== 0)
      .sort((a, b) => b.totalAmount - a.totalAmount);
    if (!list.length) return { error: t('statEmpty').replace('{msg}', 'customers') };

    spec.values.push(
      ['客户编号', '客户名称', '销售1', '销售2',
        '合计', '', '', 'CABINET', '', '', '', '', 'STONE', '', '', '', ''],
      ['', '', '', '', '总金额', '占比', '总单数',
        '金额', '金额占比', '数量', '数量占比', '单数',
        '金额', '金额占比', '数量', '数量占比', '单数']
    );
    spec.merges = ['A1:A2', 'B1:B2', 'C1:C2', 'D1:D2', 'E1:G1', 'H1:L1', 'M1:Q1'];

    spec.values.push(['合计', '', '', '', total.amount, '', total.invoices.size,
      total.cabinetAmount, total.amount ? total.cabinetAmount / total.amount : 0, total.cabinetQty, '', total.cabinetInvoices.size,
      total.stoneAmount, total.amount ? total.stoneAmount / total.amount : 0, total.stoneQty, '', total.stoneInvoices.size]);

    list.forEach((x) => {
      const s1 = mainSales1(x);
      spec.values.push([x.customerNo, x.customerName, s1, s1 === 'HO' ? mainSales2(x) : '',
        x.totalAmount, total.amount ? x.totalAmount / total.amount : 0, x.invoices.size,
        x.cabinetAmount, total.amount ? x.cabinetAmount / total.amount : 0, x.cabinetQty, total.cabinetQty ? x.cabinetQty / total.cabinetQty : 0, x.cabinetInvoices.size,
        x.stoneAmount, total.amount ? x.stoneAmount / total.amount : 0, x.stoneQty, total.stoneQty ? x.stoneQty / total.stoneQty : 0, x.stoneInvoices.size]);
    });

    const lastRow = spec.values.length;
    spec.bold.push(rng(1, 1, 2, 17));
    spec.bg.push({ range: rng(1, 1, 2, 17), color: '#d9e6f3' });
    spec.bold.push(rng(3, 1, 3, 17));
    spec.bg.push({ range: rng(3, 1, 3, 17), color: '#d9bd8b' });
    [5, 8, 13].forEach((c) => spec.numFmt.push({ range: colLetter(c) + '3:' + colLetter(c) + lastRow, fmt: '#,##0.00' }));
    [6, 9, 11, 14, 16].forEach((c) => spec.numFmt.push({ range: colLetter(c) + '3:' + colLetter(c) + lastRow, fmt: '0.00%' }));
    [7, 12, 17].forEach((c) => spec.numFmt.push({ range: colLetter(c) + '3:' + colLetter(c) + lastRow, fmt: '#,##0' }));
    [10, 15].forEach((c) => spec.numFmt.push({ range: colLetter(c) + '3:' + colLetter(c) + lastRow, fmt: '#,##0' }));
    spec.align.push({ range: 'A3:D' + lastRow, horiz: 'left' });
    spec.align.push({ range: 'E3:Q' + lastRow, horiz: 'right' });
    spec.align.push({ range: 'C1:D' + lastRow, horiz: 'center' });
    spec.widths = [[1, 115], [2, 320], [3, 75], [4, 75], [5, 135], [6, 88], [7, 92], [8, 135], [9, 100], [10, 100], [11, 100], [12, 88], [13, 135], [14, 100], [15, 100], [16, 100], [17, 88]];
    spec.freeze = { rows: 3, cols: 2 };
    return spec;
  }

  // ============================================================
  // 6) 橱柜颜色统计 (cabinet colors) — sheet 橱柜颜色统计
  // ============================================================
  // ============================================================
  // 6) 橱柜颜色统计 (cabinet colors) — sheet 橱柜颜色统计
  // V65 新口径（客户 2026-09 确认）：UNIVERSAL 不再按整体比例分摊，
  // 而是「按 Item 拆分」——每条 UNIVERSAL 销售（Item X）按同一 Item 的
  // 非 UNIVERSAL 颜色构成（数量按占比+最大余数取整、金额按金额占比）
  // 拆到具体颜色；该 Item 没有非 UNIVERSAL 记录时退回全局颜色占比。
  // 「包含 UNIVERSAL」时 UNIVERSAL 仍单独成行、不拆分。
  // ============================================================
  function cabinetColors(data, opts, t) {
    const rg = range(opts.start, opts.end);
    if (!rg.success) return { error: t('statEmpty').replace('{msg}', rg.message) };
    const includeUniversal = opts.includeUniversal === true;
    const spec = baseSpec('橱柜颜色统计');
    const groupMap = sales1Groups(opts.reps || { s1: [], s2: [] });

    const colorOf = (raw) => String(raw || '').trim() || '(空白)';
    // 颜色汇总：normalPerItem[item][colorKey] / universalPerItem[item]
    const normalPerItem = {};
    const universalPerItem = {};
    const colorMeta = {};   // colorKey -> 显示名
    let universalQty = 0, universalAmount = 0;
    let normalQty = 0, normalAmount = 0;
    // 全局颜色合计（Item 级拆分缺参照时的回退权重）
    const globalColor = {};

    data.forEach((r) => {
      const sales1Code = String(r.rep1 || '').trim().toUpperCase();
      if (isIC(sales1Code, groupMap)) return;
      const dt = date(r.soldDate);
      if (!inRange(dt, rg)) return;
      if (!isCabinetType(r.type)) return;
      const rawColor = String(r.color || '').trim();
      const normalized = rawColor.toUpperCase();
      const qty = amount(r.qty);
      const amt = amount(r.amount);
      const item = String(r.itemNo || '').trim().toUpperCase() || '(NO ITEM)';
      if (normalized === 'UNIVERSAL') {
        if (!universalPerItem[item]) universalPerItem[item] = { qty: 0, amount: 0 };
        universalPerItem[item].qty += qty;
        universalPerItem[item].amount += amt;
        universalQty += qty; universalAmount += amt;
        return;
      }
      const color = colorOf(rawColor);
      const key = color.toUpperCase();
      colorMeta[key] = color;
      if (!normalPerItem[item]) normalPerItem[item] = {};
      if (!normalPerItem[item][key]) normalPerItem[item][key] = { qty: 0, amount: 0 };
      normalPerItem[item][key].qty += qty;
      normalPerItem[item][key].amount += amt;
      if (!globalColor[key]) globalColor[key] = { qty: 0, amount: 0 };
      globalColor[key].qty += qty;
      globalColor[key].amount += amt;
      normalQty += qty;
      normalAmount += amt;
    });

    // 汇总容器：颜色 → {qty, amount}
    const totals = {};
    const addTo = (key, qty, amt) => {
      if (!totals[key]) totals[key] = { qty: 0, amount: 0 };
      totals[key].qty += qty;
      totals[key].amount += amt;
    };
    Object.keys(globalColor).forEach((key) => addTo(key, globalColor[key].qty, globalColor[key].amount));

    // 数量拆分：最大余数法，保证拆分后整数数量之和严格等于目标
    const splitQty = (weights, target) => {
      // weights: [{key, weight}]（取绝对值加权），target 整数（带符号）
      const sign = target < 0 ? -1 : 1;
      const abs = Math.abs(target);
      const weightTotal = weights.reduce((s, w) => s + Math.abs(w.weight), 0);
      if (!abs || weightTotal <= 0) return null;
      const alloc = weights.map((w, idx) => {
        const exact = (abs * Math.abs(w.weight)) / weightTotal;
        const base = Math.floor(exact);
        return { key: w.key, idx, base, frac: exact - base };
      });
      let used = alloc.reduce((s, a) => s + a.base, 0);
      let remainder = abs - used;
      alloc.sort((a, b) => b.frac - a.frac || a.idx - b.idx);
      for (let i = 0; i < remainder; i++) alloc[i % alloc.length].base += 1;
      alloc.sort((a, b) => a.idx - b.idx);
      const out = {};
      alloc.forEach((a) => (out[a.key] = sign * a.base));
      return out;
    };

    if (!includeUniversal) {
      const hasAnyNormal = Object.keys(globalColor).length > 0;
      if (!hasAnyNormal) {
        return { error: t('statEmpty').replace('{msg}', 'cabinet colors') };
      }
      Object.keys(universalPerItem).forEach((item) => {
        const u = universalPerItem[item];
        const perItem = normalPerItem[item] || {};
        let keys = Object.keys(perItem);
        let weightsQty, weightsAmt;
        if (keys.length) {
          // Item 级拆分：优先用该 Item 自己的颜色构成
          weightsQty = keys.map((k) => ({ key: k, weight: perItem[k].qty }));
          weightsAmt = keys.map((k) => ({ key: k, weight: perItem[k].amount !== 0 ? perItem[k].amount : perItem[k].qty }));
        } else {
          // 该 Item 没有非 UNIVERSAL 记录 → 退回全局颜色占比
          keys = Object.keys(globalColor);
          weightsQty = keys.map((k) => ({ key: k, weight: globalColor[k].qty }));
          weightsAmt = keys.map((k) => ({ key: k, weight: globalColor[k].amount !== 0 ? globalColor[k].amount : globalColor[k].qty }));
        }
        // 金额拆分：按权重占比分配 UNIVERSAL 金额（金额可为小数，不取整）
        const amtWeightTotal = weightsAmt.reduce((s, w) => s + Math.abs(w.weight), 0);
        if (u.amount !== 0 && amtWeightTotal > 0) {
          weightsAmt.forEach((w) => {
            addTo(w.key, 0, u.amount * (Math.abs(w.weight) / amtWeightTotal) * (w.weight < 0 ? -1 : 1));
          });
        }
        // 数量拆分：最大余数取整
        const qtyAlloc = splitQty(weightsQty, Math.round(u.qty));
        if (qtyAlloc) Object.keys(qtyAlloc).forEach((k) => addTo(k, qtyAlloc[k], 0));
      });
    } else if (universalQty !== 0 || universalAmount !== 0) {
      addTo('UNIVERSAL', universalQty, universalAmount);
      colorMeta['UNIVERSAL'] = 'UNIVERSAL';
    }

    let list = Object.keys(totals).map((key) => ({ color: colorMeta[key] || key, qty: totals[key].qty, amount: totals[key].amount }));
    if (!list.length) return { error: t('statEmpty').replace('{msg}', 'cabinet colors') };

    const totalQty = includeUniversal ? normalQty + universalQty : list.reduce((s, x) => s + x.qty, 0);
    const totalAmount = includeUniversal ? normalAmount + universalAmount : list.reduce((s, x) => s + x.amount, 0);

    list.sort((a, b) => {
      const au = String(a.color || '').toUpperCase() === 'UNIVERSAL';
      const bu = String(b.color || '').toUpperCase() === 'UNIVERSAL';
      if (au && !bu) return 1;
      if (!au && bu) return -1;
      if (b.qty !== a.qty) return b.qty - a.qty;
      return String(a.color).localeCompare(String(b.color), undefined, { sensitivity: 'base' });
    });

    spec.values.push(['颜色', '数量', '数量占比', '金额', '金额占比']);
    list.forEach((x) => spec.values.push([x.color, x.qty, totalQty ? x.qty / totalQty : 0, x.amount, totalAmount ? x.amount / totalAmount : 0]));
    spec.values.push(['合计', totalQty, '', totalAmount, '']);
    const lastRow = spec.values.length;
    spec.bold.push(rng(1, 1, 1, 5));
    spec.bg.push({ range: rng(1, 1, 1, 1), color: '#ffe699' });
    spec.bg.push({ range: rng(1, 2, 1, 5), color: '#d9e6f3' });
    spec.numFmt.push({ range: 'B2:B' + lastRow, fmt: '#,##0' });
    spec.numFmt.push({ range: 'C2:C' + (lastRow - 1), fmt: '0.00%' });
    spec.numFmt.push({ range: 'D2:D' + lastRow, fmt: '#,##0.00' });
    spec.numFmt.push({ range: 'E2:E' + (lastRow - 1), fmt: '0.00%' });
    spec.align.push({ range: 'B2:E' + lastRow, horiz: 'right' });
    spec.align.push({ range: 'A2:A' + lastRow, horiz: 'left' });
    spec.bold.push(rng(lastRow, 1, lastRow, 5));
    spec.bg.push({ range: rng(lastRow, 1, lastRow, 5), color: '#d9bd8b' });
    spec.widths = [[1, 180], [2, 105], [3, 125], [4, 140], [5, 125]];
    spec.freeze = { rows: 1, cols: 0 };
    spec.info = includeUniversal ? (LANG_UNIVERSAL_INCLUDE) : LANG_UNIVERSAL_SPLIT;
    return spec;
  }
  const LANG_UNIVERSAL_INCLUDE = '包含 UNIVERSAL：单独成行，不参与拆分';
  const LANG_UNIVERSAL_SPLIT = 'UNIVERSAL 已按 Item 的颜色构成拆分到各颜色（数量按占比取整，金额按金额占比）';

  // ============================================================
  // 7) 石材销量统计 (stone sales) — sheet 石材销量统计
  // type: SLAB (default) or other non-cabinet type
  // ============================================================
  function stoneSales(data, opts, t) {
    const rg = range(opts.start, opts.end);
    if (!rg.success) return { error: t('statEmpty').replace('{msg}', rg.message) };
    let selectedType = String(opts.type || 'SLAB').trim().toUpperCase();
    if (selectedType === 'CABINET' || selectedType === 'CABINETOTHERS') {
      return { error: t('statEmpty').replace('{msg}', 'CABINET not allowed') };
    }
    const spec = baseSpec('石材销量统计');
    const groupMap = sales1Groups(opts.reps || { s1: [], s2: [] });

    // 石材销售统计：仅计可用库存，不含在途数量。
    const stockMap = new Map();
    (opts.stock || []).forEach((s) => {
      const key = String(s.itemNo || '').trim().toUpperCase();
      if (!key) return;
      stockMap.set(key, { itemNo: s.itemNo, description: String(s.description || ''), stock: Number(s.available) || 0 });
    });
    // description map from sales (Search_List equivalent)
    const descMap = new Map();
    data.forEach((r) => {
      const k = String(r.itemNo || '').trim().toUpperCase();
      const d = String(r.description || '').trim();
      if (k && d && !descMap.has(k)) descMap.set(k, d);
    });

    const items = {};
    let grandQty = 0, grandAmount = 0;
    data.forEach((r) => {
      const sales1Code = String(r.rep1 || '').trim().toUpperCase();
      if (isIC(sales1Code, groupMap)) return;
      const dt = date(r.soldDate);
      if (!inRange(dt, rg)) return;
      const type = String(r.type || '').trim().toUpperCase();
      if (type !== selectedType) return;
      const itemNo = String(r.itemNo || '').trim();
      if (!itemNo) return;
      const description = String(r.description || '').trim();
      const qty = amount(r.qty);
      const amt = amount(r.amount);
      const key = itemNo.toUpperCase();
      if (!items[key]) items[key] = { category: selectedType, itemNo, description, qty: 0, amount: 0 };
      const x = items[key];
      if (!x.description && description) x.description = description;
      x.qty += qty;
      x.amount += amt;
      grandQty += qty;
      grandAmount += amt;
    });

    stockMap.forEach((inv, key) => {
      const searchDesc = descMap.get(key) || '';
      if (items[key]) {
        const existing = items[key];
        if (!existing.description) existing.description = searchDesc || inv.description || '';
        existing.stock = inv.stock;
      } else {
        items[key] = { category: 'SLAB', itemNo: inv.itemNo, description: searchDesc || inv.description || '', qty: 0, amount: 0, stock: inv.stock };
      }
    });
    Object.keys(items).forEach((key) => {
      const x = items[key];
      const stockKey = String(x.itemNo || '').trim().toUpperCase();
      const stock = x.stock !== undefined ? Number(x.stock) || 0 : stockMap.has(stockKey) ? Number(stockMap.get(stockKey).stock) || 0 : 0;
      if (stock === 0 && (Number(x.qty) || 0) === 0) delete items[key];
    });

    const list = Object.values(items).sort((a, b) => {
      if (b.qty !== a.qty) return b.qty - a.qty;
      return String(a.itemNo).localeCompare(String(b.itemNo), undefined, { numeric: true, sensitivity: 'base' });
    });
    if (!list.length) return { error: t('statEmpty').replace('{msg}', selectedType) };

    spec.values.push(['类别', '品名', '描述', '库存数量', '销售数量', '金额', '%(Top 30)', '平均单价']);
    list.forEach((x, index) => {
      const ratio = index < 30 && grandQty ? x.qty / grandQty : '';
      const stockKey = String(x.itemNo || '').trim().toUpperCase();
      const stock = x.stock !== undefined ? x.stock : stockMap.has(stockKey) ? stockMap.get(stockKey).stock : '';
      // 平均单价 = 金额 / 销售数量（客户手工列转正）；无数量时留空
      const unitPrice = (Number(x.qty) || 0) > 0 ? (Number(x.amount) || 0) / (Number(x.qty) || 0) : '';
      spec.values.push([x.category, x.itemNo, x.description, stock, x.qty, x.amount, ratio, unitPrice]);
    });
    const grandStock = list.reduce((sum, x) => {
      const stockKey = String(x.itemNo || '').trim().toUpperCase();
      const stock = x.stock !== undefined ? x.stock : stockMap.has(stockKey) ? stockMap.get(stockKey).stock : '';
      return sum + (stock === '' ? 0 : Number(stock) || 0);
    }, 0);
    const grandPrice = grandQty ? grandAmount / grandQty : '';
    spec.values.push(['合计', '', '', grandStock, grandQty, grandAmount, '', grandPrice]);
    const lastRow = spec.values.length;
    spec.bold.push(rng(1, 1, 1, 8));
    spec.bg.push({ range: rng(1, 1, 1, 1), color: '#ffe699' });
    spec.bg.push({ range: rng(1, 2, 1, 8), color: '#d9e6f3' });
    spec.numFmt.push({ range: 'D2:E' + lastRow, fmt: '#,##0' });
    spec.numFmt.push({ range: 'F2:F' + lastRow, fmt: '#,##0.00' });
    const ratioRows = Math.min(30, list.length);
    if (ratioRows > 0) spec.numFmt.push({ range: 'G2:G' + (1 + ratioRows), fmt: '0.00%' });
    spec.numFmt.push({ range: 'H2:H' + lastRow, fmt: '#,##0.00' });
    spec.align.push({ range: 'A2:C' + lastRow, horiz: 'left' });
    spec.align.push({ range: 'D2:E' + lastRow, horiz: 'center' });
    spec.align.push({ range: 'F2:F' + lastRow, horiz: 'right' });
    spec.align.push({ range: 'G2:G' + lastRow, horiz: 'center' });
    spec.align.push({ range: 'H2:H' + lastRow, horiz: 'right' });
    spec.bold.push(rng(lastRow, 1, lastRow, 8));
    spec.bg.push({ range: rng(lastRow, 1, lastRow, 8), color: '#d9bd8b' });
    spec.widths = [[1, 125], [2, 130], [3, 340], [4, 100], [5, 100], [6, 150], [7, 80], [8, 110]];
    spec.freeze = { rows: 1, cols: 0 };
    return spec;
  }

  // ============================================================
  // 8) 客户动态分析 — 纯数据版（供「客户动态」页：左排名 / 中趋势表 / 右图表）
  // opts.period: ALL|LAST_MONTH|THIS_YEAR|YTD
  // ============================================================
  function customerDashboardData(data, opts, reps) {
    const period = ['LAST_MONTH', 'THIS_YEAR', 'YTD'].includes(opts.period) ? opts.period : 'ALL';
    const groupMap = sales1Groups(reps || { s1: [], s2: [] });
    const now = new Date();
    const last = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const monthKeys = [];
    for (let i = 11; i >= 0; i--) {
      const d = new Date(last.getFullYear(), last.getMonth() - i, 1);
      monthKeys.push(d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0'));
    }
    const monthSet = new Set(monthKeys);
    const dtext = (d) => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    const periods = {
      ALL: { start: '', end: '' },
      LAST_MONTH: { start: dtext(new Date(now.getFullYear(), now.getMonth() - 1, 1)), end: dtext(new Date(now.getFullYear(), now.getMonth(), 0)) },
      THIS_YEAR: { start: dtext(new Date(now.getFullYear(), 0, 1)), end: dtext(now) },
      YTD: (() => { const w = ytdWindow(); return { start: w.start, end: w.ytdEndInclusive }; })(),
    };
    const rankStats = { ALL: {}, LAST_MONTH: {}, THIS_YEAR: {}, YTD: {} };
    const customers = {};
    data.forEach((r) => {
      const sales1Code = String(r.rep1 || '').trim().toUpperCase();
      if (isIC(sales1Code, groupMap)) return;
      const dt = date(r.soldDate);
      if (!dt) return;
      const no = String(r.custNo || '').trim();
      if (!no) return;
      const name = String(r.custName || '').trim();
      const sales1 = sales1Code, sales2 = String(r.rep2 || '').trim().toUpperCase();
      const amt = amount(r.amount);
      const isCab = isCabinetType(r.type);
      if (!customers[no]) {
        customers[no] = { no, name, totalAmount: 0, sales1Amounts: {}, sales2Amounts: {}, years: {},
          months: monthKeys.reduce((m, k) => ((m[k] = { total: 0, cabinet: 0, stone: 0 }), m), {}) };
      }
      const x = customers[no];
      if (!x.name && name) x.name = name;
      if (sales1) x.sales1Amounts[sales1] = (x.sales1Amounts[sales1] || 0) + amt;
      if (sales1 === 'HO' && sales2) x.sales2Amounts[sales2] = (x.sales2Amounts[sales2] || 0) + amt;
      x.totalAmount += amt;
      const yr = dt.substring(0, 4);
      if (!x.years[yr]) x.years[yr] = { total: 0, cabinet: 0, stone: 0 };
      x.years[yr].total += amt;
      if (isCab) x.years[yr].cabinet += amt; else x.years[yr].stone += amt;
      const mk = dt.substring(0, 7);
      if (monthSet.has(mk)) {
        x.months[mk].total += amt;
        if (isCab) x.months[mk].cabinet += amt; else x.months[mk].stone += amt;
      }
      Object.keys(periods).forEach((p) => {
        const pg = periods[p];
        if (pg.start && dt < pg.start) return;
        if (pg.end && dt > pg.end) return;
        if (!rankStats[p][no]) rankStats[p][no] = { no, name, totalAmount: 0, sales1Amounts: {}, sales2Amounts: {} };
        const b = rankStats[p][no];
        b.totalAmount += amt;
        if (sales1) b.sales1Amounts[sales1] = (b.sales1Amounts[sales1] || 0) + amt;
        if (sales1 === 'HO' && sales2) b.sales2Amounts[sales2] = (b.sales2Amounts[sales2] || 0) + amt;
      });
    });
    const main1 = (amts) => {
      let best = '', bestAmt = -Infinity;
      Object.keys(amts).forEach((c) => { if (amts[c] > bestAmt) { bestAmt = amts[c]; best = c; } });
      return best;
    };
    const rankings = {};
    Object.keys(rankStats).forEach((p) => {
      rankings[p] = Object.values(rankStats[p])
        .sort((a, b) => (b.totalAmount !== a.totalAmount ? b.totalAmount - a.totalAmount : String(a.name).localeCompare(String(b.name), undefined, { sensitivity: 'base' })))
        .map((x) => {
          const s1 = main1(x.sales1Amounts);
          return [x.no, x.name, x.totalAmount, s1, s1 === 'HO' ? main1(x.sales2Amounts) : ''];
        });
    });
    return { period, monthKeys, customers, rankings, periodLabels: { ALL: '全部', LAST_MONTH: '上月', THIS_YEAR: '今年', YTD: 'TTM' } };
  }

  // ============================================================
  // 9) 销售历史查询 (salesHistory) — 对应原 runSalesHistory（第15段）
  // opts: {start, end, custNo, itemNo, invoice}；查询优先级 Invoice > Item > Customer（精确匹配）
  // 三种视图：Customer(按发票汇总) / Invoice(发票明细) / Item(品名流水)
  // ============================================================
  function salesHistory(data, opts, t) {
    const norm = (v) => String(v || '').trim().toUpperCase();
    const invoiceF = norm(opts.invoice);
    const itemF = norm(opts.itemNo);
    const custF = norm(opts.custNo);
    const sales1F = norm(opts.sales);
    const rg = range(opts.start, opts.end);
    if (!rg.success) return { error: rg.message };
    const inP = (dt) => { if (!rg.start && !rg.end) return true; return inRange(dt, rg); };
    // 原版业务规则：销售历史查询一律排除 IC（只有年度销售总表保留 IC）
    const gmapIC = sales1Groups(opts.reps || { s1: [], s2: [] });
    const src = [];
    data.forEach((r) => {
      const dt = date(r.soldDate);
      if (!dt || !inP(dt)) return;
      // 组合查询（原版口径）：填了哪些条件就必须同时满足哪些（AND）
      if (invoiceF && norm(r.invoice) !== invoiceF) return;
      if (itemF && norm(r.itemNo) !== itemF) return;
      if (custF && norm(r.custNo) !== custF) return;
      const rep1 = String(r.rep1 || '').trim();
      if (sales1F && norm(rep1) !== sales1F) return;
      if (isIC(norm(rep1), gmapIC)) return;
      const rawCat = String(r.type || '').trim();
      const up = rawCat.toUpperCase().replace(/\s+/g, '');
      // V64 成本规则：SLAB 的 S 列是单位成本需 ×数量；其他类别 S 列已是整行成本
      const unitCost = Math.abs(amount(r.cost));
      src.push({
        item: String(r.itemNo || '').trim(), desc: String(r.description || '').trim(),
        dt, invoice: String(r.invoice || '').trim(), cust: String(r.custNo || '').trim(),
        name: String(r.custName || '').trim(), rep: String(r.rep1 || '').trim(), s2: String(r.rep2 || '').trim(),
        color: String(r.color || '').trim(), rawCat,
        cat: (up === 'CABINET' || up === 'CABINETOTHERS') ? 'Cabinet' : 'Stone',
        qty: amount(r.qty), amt: amount(r.amount),
        cost: up === 'SLAB' ? unitCost * amount(r.qty) : unitCost,
      });
    });
    if (!src.length) return { error: 'NO_MATCH' };
    const fmtDate = (s) => { const p = String(s).split('-'); return p.length === 3 ? p[1] + '/' + p[2] + '/' + p[0] : s; };
    const periodText = (rg.start ? fmtDate(rg.start) : 'All History') + ' - ' + (rg.end ? fmtDate(rg.end) : 'Today');
    const spec = baseSpec('销售历史查询');
    spec.historyView = sales1F ? 'sales' : invoiceF ? 'invoice' : itemF ? 'item' : custF ? 'customer' : 'all';
    spec.freeze = {rows:1, cols:0};
    const money = '#,##0.00;-#,##0.00;';
    const qtyFmt = '#,##0;-#,##0;';

    // 合并工具：按发票号码合并（原版口径）——同发票多行合计数量/金额/成本、日期取最近、客户取第一个非空
    const mergeByInvoice = (withSales) => {
      const grouped = new Map();
      src.forEach((x) => {
        const key = norm(x.invoice);
        if (!grouped.has(key)) grouped.set(key, { dt: x.dt, invoice: x.invoice, cats: new Set(), qty: 0, amt: 0, cost: 0, cust: '', name: '', s1: new Set(), s2: new Set() });
        const g = grouped.get(key);
        if (x.dt > g.dt) g.dt = x.dt;
        g.cats.add(x.cat); g.qty += x.qty; g.amt += x.amt; g.cost += x.cost;
        if (!g.cust && x.cust) g.cust = x.cust;
        if (!g.name && x.name) g.name = x.name;
        if (x.rep) g.s1.add(x.rep);
        if (x.s2) g.s2.add(x.s2);
      });
      return Array.from(grouped.values()).sort((a, b) => b.dt.localeCompare(a.dt) || String(a.invoice).localeCompare(String(b.invoice)));
    };
    const catOf = (cats) => {
      if (cats.has('Cabinet') && cats.has('Stone')) return 'Cabinet / Stone';
      return cats.has('Cabinet') ? 'Cabinet' : 'Stone';
    };
    const totalRow = (qtyIdx, amtIdx, costIdx, costCol1) => {
      const row = ['Total', '', '',
        spec.values.slice(1).reduce((s2, r) => s2 + (Number(r[qtyIdx]) || 0), 0),
        spec.values.slice(1).reduce((s2, r) => s2 + (Number(r[amtIdx]) || 0), 0)];
      while (row.length < costCol1 - 1) row.push('');
      row.push(spec.values.slice(1).reduce((s2, r) => s2 + (Number(r[costIdx]) || 0), 0));
      return row;
    };

    if (!custF && !itemF && !invoiceF && !sales1F) {
      // 无附加条件：按发票号码合并（原版口径），最近在前；Sales 1 / Sales 2 去重并列
      const rows = mergeByInvoice(true);
      spec.values.push(['Sold Date', 'Invoice No.', 'Category', 'Qty', 'Amount', 'Customer No', 'Customer Name', 'Sales 1', 'Sales 2', 'Cost']);
      rows.forEach((g) => {
        spec.values.push([g.dt, g.invoice, catOf(g.cats), g.qty, g.amt, g.cust, g.name, Array.from(g.s1).join(' / '), Array.from(g.s2).join(' / '), g.cost]);
      });
      const n = spec.values.length;
      spec.values.push(totalRow(3, 4, 9, 10));
      spec.numFmt.push({ range: 'D2:D' + (n + 1), fmt: qtyFmt },
        { range: 'E2:E' + (n + 1), fmt: money }, { range: 'J2:J' + (n + 1), fmt: money });
      spec.align = [{ range: 'A2:C' + n, horiz: 'left' }, { range: 'F2:I' + n, horiz: 'left' }];
      spec.freeze = { rows: 1, cols: 0 };
      spec.bold.push(rng(1, 1, 1, 10));
      spec.bg.push({ range: rng(1, 1, 1, 10), color: '#d9eaf7' });
      spec.bold.push(rng(n + 1, 1, n + 1, 10));
      spec.bg.push({ range: rng(n + 1, 1, n + 1, 10), color: '#d9e6f3' });
      spec.info = 'All Sales | Period: ' + periodText + ' | ' + rows.length + ' invoices';
      return spec;
    }

    if (sales1F) {
      // SALES 1 查询（原版口径）：按发票号码合并，可与客户/品名/发票条件组合（AND）
      const rows = mergeByInvoice(false);
      spec.values.push(['Sold Date', 'Invoice No.', 'Category', 'Qty', 'Amount', 'Customer No', 'Customer Name', 'Cost']);
      rows.forEach((g) => {
        spec.values.push([g.dt, g.invoice, catOf(g.cats), g.qty, g.amt, g.cust, g.name, g.cost]);
      });
      const n = spec.values.length;
      spec.values.push(totalRow(3, 4, 7, 8));
      spec.numFmt.push({ range: 'D2:D' + (n + 1), fmt: qtyFmt },
        { range: 'E2:E' + (n + 1), fmt: money }, { range: 'H2:H' + (n + 1), fmt: money });
      spec.align = [{ range: 'A2:C' + n, horiz: 'left' }, { range: 'F2:G' + n, horiz: 'left' }];
      spec.freeze = { rows: 1, cols: 0 };
      spec.bold.push(rng(1, 1, 1, 8));
      spec.bg.push({ range: rng(1, 1, 1, 8), color: '#d9eaf7' });
      spec.bold.push(rng(n + 1, 1, n + 1, 8));
      spec.bg.push({ range: rng(n + 1, 1, n + 1, 8), color: '#d9e6f3' });
      const repInfo = ((opts.reps && opts.reps.s1) || []).find((r) => norm(r.code) === sales1F);
      const repName = repInfo ? (repInfo.name || '') : '';
      spec.info = 'Sales 1: ' + sales1F + (repName ? ' | ' + repName : '') + ' | Period: ' + periodText + ' | ' + rows.length + ' invoices';
      return spec;
    }

    if (custF && !itemF && !invoiceF) {
      // CUSTOMER 视图：客户信息入 info，明细按 Customer+Invoice 合并
      const grouped = new Map();
      src.forEach((x) => {
        const key = norm(x.cust) + '||' + norm(x.invoice);
        if (!grouped.has(key)) grouped.set(key, { dt: x.dt, invoice: x.invoice, cats: new Set(), qty: 0, amt: 0, rep: x.rep, cost: 0 });
        const g = grouped.get(key);
        if (x.dt > g.dt) g.dt = x.dt;
        g.cats.add(x.cat); g.qty += x.qty; g.amt += x.amt; g.cost += x.cost;
        if (!g.rep && x.rep) g.rep = x.rep;
      });
      spec.values.push(['Sold Date', 'Invoice No.', 'Category', 'Qty', 'Amount', 'Rep Code', 'Cost']);
      Array.from(grouped.values()).sort((a, b) => b.dt.localeCompare(a.dt)).forEach((g) => {
        let cat = 'Stone';
        if (g.cats.has('Cabinet') && g.cats.has('Stone')) cat = 'Cabinet / Stone';
        else if (g.cats.has('Cabinet')) cat = 'Cabinet';
        spec.values.push([fmtDate(g.dt), g.invoice, cat, g.qty, g.amt, g.rep, g.cost]);
      });
      const n = spec.values.length;
      spec.values.push(['Total', '', '', spec.values.slice(1).reduce((s, r) => s + r[3], 0), spec.values.slice(1).reduce((s, r) => s + r[4], 0), '', spec.values.slice(1).reduce((s, r) => s + r[6], 0)]);
      spec.numFmt.push({ range: 'D2:D' + n, fmt: qtyFmt }, { range: 'E2:E' + (n + 1), fmt: money }, { range: 'G2:G' + (n + 1), fmt: money });
      spec.bold.push(rng(n + 1, 1, n + 1, 7));
      spec.bg.push({ range: rng(n + 1, 1, n + 1, 7), color: '#d9e6f3' });
      spec.info = src[0].cust + ' | ' + src[0].name + ' | Period: ' + periodText;
      return spec;
    }

    if (invoiceF) {
      // INVOICE 视图：同 Invoice 内 Item+Color+Category 相同才合并；Price=Amount/Qty 反算
      const grouped = new Map();
      src.forEach((x) => {
        const key = norm(x.invoice) + '||' + norm(x.item) + '||' + norm(x.color) + '||' + norm(x.rawCat);
        if (!grouped.has(key)) grouped.set(key, { item: x.item, desc: x.desc, color: x.color, cat: x.rawCat, qty: 0, amt: 0, cost: 0 });
        const g = grouped.get(key);
        g.qty += x.qty; g.amt += x.amt; g.cost += x.cost;
        if (!g.desc && x.desc) g.desc = x.desc;
      });
      spec.values.push(['序号', 'Item #', 'Description', 'Color', 'Cat', 'Qty', 'Price', 'Amount', 'Cost']);
      Array.from(grouped.values()).forEach((g, i) => {
        spec.values.push([i + 1, g.item, g.desc, g.color, g.cat, g.qty, g.qty !== 0 ? g.amt / g.qty : '', g.amt, g.cost]);
      });
      const n = spec.values.length;
      const d = spec.values.slice(1);
      spec.values.push(['Total', '', '', '', '', d.reduce((s, r) => s + r[5], 0), '', d.reduce((s, r) => s + r[7], 0), d.reduce((s, r) => s + r[8], 0)]);
      spec.numFmt.push({ range: 'A2:A' + n, fmt: '0' }, { range: 'F2:F' + (n + 1), fmt: qtyFmt },
        { range: 'G2:H' + (n + 1), fmt: money }, { range: 'I2:I' + (n + 1), fmt: money });
      spec.bold.push(rng(n + 1, 1, n + 1, 9));
      spec.bg.push({ range: rng(n + 1, 1, n + 1, 9), color: '#d9e6f3' });
      const f = src[0];
      spec.info = 'Invoice #: ' + f.invoice + '    Date: ' + fmtDate(f.dt) + '\n' +
        'Customer #: ' + f.cust + '    Customer Name: ' + f.name + '    Sales Rep 1: ' + f.rep + '    Period: ' + periodText;
      return spec;
    }

    // ITEM 视图：按 Invoice+Item 合并的流水
    const grouped = new Map();
    src.forEach((x) => {
      const key = norm(x.invoice) + '||' + norm(x.item);
      if (!grouped.has(key)) grouped.set(key, { item: x.item, desc: x.desc, dt: x.dt, invoice: x.invoice, qty: 0, amt: 0, cust: x.cust, name: x.name, rep: x.rep, cost: 0 });
      const g = grouped.get(key);
      g.qty += x.qty; g.amt += x.amt; g.cost += x.cost;
      if (x.dt > g.dt) g.dt = x.dt;
    });
    spec.values.push(['Item No.', 'Description', 'Sold Date', 'Invoice No.', 'Qty', 'Amount', 'Customer No.', 'Customer Name', 'Rep Code', 'Cost']);
    Array.from(grouped.values()).sort((a, b) => b.dt.localeCompare(a.dt)).forEach((g) => {
      spec.values.push([g.item, g.desc, fmtDate(g.dt), g.invoice, g.qty, g.amt, g.cust, g.name, g.rep, g.cost]);
    });
    const n = spec.values.length;
    const d = spec.values.slice(1);
    spec.values.push(['Total', '', '', '', d.reduce((s, r) => s + r[4], 0), d.reduce((s, r) => s + r[5], 0), '', '', '', d.reduce((s, r) => s + r[9], 0)]);
    spec.numFmt.push({ range: 'E2:E' + (n + 1), fmt: qtyFmt }, { range: 'F2:F' + (n + 1), fmt: money }, { range: 'J2:J' + (n + 1), fmt: money });
    spec.bold.push(rng(n + 1, 1, n + 1, 10));
    spec.bg.push({ range: rng(n + 1, 1, n + 1, 10), color: '#d9e6f3' });
    spec.info = 'Item: ' + src[0].item + (src[0].desc ? ' | ' + src[0].desc : '') + ' | Period: ' + periodText;
    return spec;
  }

  // ============================================================
  // 10) 客户类型统计 (customerType) — 对应原 runCustomerTypeAnalysis（第11段）
  // 需要 Start/End；客户类型取自 Customer 主数据（无则 UNASSIGNED）；
  // Salesperson=IC 或类型=INTERCOMPANY 的客户不参与。
  // ============================================================
  function customerType(data, opts, t) {
    const rg = range(opts.start, opts.end);
    if (!rg.success) return { error: rg.message };
    // 不强制要求时间段：留空 = 全部数据（快速按钮可选上月/今年/TTM/重置）
    const custMap = new Map();
    (opts.customers || []).forEach((c) => {
      const no = String(c.no || '').trim().toUpperCase();
      if (!no) return;
      let type = String(c.type || '').trim().toUpperCase();
      if (!type) type = 'UNASSIGNED';
      const sales1 = String(c.sales1 || '').trim().toUpperCase();
      custMap.set(no, { type, isIC: sales1 === 'IC' || type === 'INTERCOMPANY' });
    });
    const stats = new Map();
    let grandSales = 0;
    const grandCustomers = new Set();
    data.forEach((r) => {
      const dt = date(r.soldDate);
      if (!inRange(dt, rg)) return;
      const no = String(r.custNo || '').trim().toUpperCase();
      if (!no) return;
      const c = custMap.get(no);
      if (c && c.isIC) return;
      const type = c ? c.type : 'UNASSIGNED';
      if (!stats.has(type)) stats.set(type, { amt: 0, custs: new Set() });
      const s = stats.get(type);
      s.amt += amount(r.amount);
      s.custs.add(no);
    });
    stats.forEach((s) => { grandSales += s.amt; s.custs.forEach((c) => grandCustomers.add(c)); });
    const preferredOrder = ['BUILDER/DEVELOPER', 'DEALER/SHOWROOM', 'DESIGNER/INVESTOR', 'GENERAL CONTRACTOR',
      'IMPORTER/DISTRIBUTOR', 'INSTALLER/FABRICATOR', 'PROPERTY MANAGEMENT', 'UNASSIGNED'];
    const types = Array.from(stats.keys()).sort((a, b) => {
      const ai = preferredOrder.indexOf(a), bi = preferredOrder.indexOf(b);
      if (ai !== -1 && bi !== -1) return ai - bi;
      if (ai !== -1) return -1;
      if (bi !== -1) return 1;
      return a.localeCompare(b);
    });
    const spec = baseSpec('客户类型统计');
    spec.values.push(['Customer Type', 'Sales Amount', 'Sales %', 'Number Of Customer', 'Customer %']);
    types.forEach((type) => {
      const s = stats.get(type);
      spec.values.push([type, s.amt, grandSales !== 0 ? s.amt / grandSales : 0, s.custs.size, grandCustomers.size !== 0 ? s.custs.size / grandCustomers.size : 0]);
    });
    spec.values.push(['Grand Total', grandSales, grandSales !== 0 ? 1 : 0, grandCustomers.size, grandCustomers.size !== 0 ? 1 : 0]);
    const n = spec.values.length;
    spec.numFmt.push({ range: 'B2:B' + n, fmt: '#,##0.00;-#,##0.00;' }, { range: 'C2:C' + n, fmt: '0.00%' },
      { range: 'D2:D' + n, fmt: '#,##0;-#,##0;' }, { range: 'E2:E' + n, fmt: '0.00%' });
    spec.bold.push(rng(n, 1, n, 5));
    spec.bg.push({ range: rng(n, 1, n, 5), color: '#d9eaf7' });
    spec.align = [{ range: 'A2:A' + n, horiz: 'left' }, { range: 'B2:B' + n, horiz: 'right' }, { range: 'C2:E' + n, horiz: 'center' }];
    spec.info = (opts.start || '') + ' - ' + (opts.end || '');
    return spec;
  }

  // ============================================================
  // 11) 客户开户统计 (salesAccount) — 对应原 runSalesAccountAnalysis（第12段）
  // 不受日期影响：系统截止月 = 最新销售日期所在月；
  // Total/年度列/YTD 按开户日期（Customer.created），Sleeping=近3月无购买，Never=从未购买。
  // ============================================================
  function salesAccount(data, opts, t) {
    // 时间段（快速按钮）：有范围时只统计范围内的销售（Active/Sleeping 窗口随之收窄），重置=全部
    const rgS = range(opts.start, opts.end);
    if (!rgS.success) return { error: rgS.message };
    if (opts.start || opts.end) {
      data = data.filter((r) => { const dt = date(r.soldDate); return dt && inRange(dt, rgS); });
    }
    const customers = opts.customers || [];
    const groupMap = sales1Groups(opts.reps || { s1: [], s2: [] });
    const salesMap = new Map();
    (opts.reps && opts.reps.s1 || []).forEach((r) => {
      const code = String(r.code || '').trim().toUpperCase();
      const g = String(r.group || '').trim().toUpperCase();
      if (!code || code === 'IC' || g === 'IC') return;
      salesMap.set(code, { name: String(r.name || '').trim() || code, group: g });
    });
    if (!salesMap.size) return { error: 'NO_SALES1' };
    if (!(opts.customers || []).length) return { error: 'NO_CUSTOMERS' };
    let latest = '';
    const everPurchased = new Set();
    data.forEach((r) => {
      const dt = date(r.soldDate);
      if (!dt) return;
      if (dt > latest) latest = dt;
      const no = String(r.custNo || '').trim().toUpperCase();
      if (no) everPurchased.add(no);
    });
    if (!latest) return { error: 'NO_DATA' };
    const sysY = Number(latest.slice(0, 4)), sysM = Number(latest.slice(5, 7));
    const monthStart = (y, m, back) => { const d = new Date(y, m - 1 - back, 1); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-01'; };
    const recent3Start = monthStart(sysY, sysM, 2);
    const ytdStart = monthStart(sysY, sysM, 11);
    const recentPurchased = new Set();
    data.forEach((r) => {
      const dt = date(r.soldDate);
      if (dt && dt >= recent3Start) recentPurchased.add(String(r.custNo || '').trim().toUpperCase());
    });
    const yearsSet = new Set();
    const custs = [];
    customers.forEach((c) => {
      const no = String(c.no || '').trim().toUpperCase();
      if (!no) return;
      const salesCode = String(c.sales1 || '').trim().toUpperCase();
      if (!salesMap.has(salesCode)) return;
      const created = date(c.created);
      custs.push({ no, salesCode, created });
      if (created) yearsSet.add(Number(created.slice(0, 4)));
    });
    const years = Array.from(yearsSet).sort((a, b) => a - b);
    const stats = new Map();
    salesMap.forEach((info, code) => {
      const yc = {};
      years.forEach((y) => (yc[y] = 0));
      stats.set(code, { name: info.name, group: info.group, total: 0, yc, ytd: 0, months: [0, 0, 0], sleeping: 0, never: 0 });
    });
    custs.forEach((c) => {
      const s = stats.get(c.salesCode);
      if (!s) return;
      s.total++;
      if (c.created) {
        const y = Number(c.created.slice(0, 4));
        if (Object.prototype.hasOwnProperty.call(s.yc, y)) s.yc[y]++;
        if (c.created >= ytdStart) s.ytd++;
        const ym = c.created.slice(0, 7);
        [2, 1, 0].forEach((back, i) => {
          const st = monthStart(sysY, sysM, back).slice(0, 7);
          if (ym === st) s.months[i]++;
        });
      }
      if (!everPurchased.has(c.no)) s.never++;
      else if (!recentPurchased.has(c.no)) s.sleeping++;
    });
    const groupPriority = { HOUSE: 1, OUTSIDE: 2, 'OUTSIDE SALES': 2, 'OUTSIDE SALE': 2, OTHER: 3, 'OTHER BRANCH': 3 };
    const rows = Array.from(stats.entries()).map(([code, x]) => Object.assign({ code }, x)).sort((a, b) => {
      const pa = groupPriority[String(a.group || '').toUpperCase()] || 99;
      const pb = groupPriority[String(b.group || '').toUpperCase()] || 99;
      if (pa !== pb) return pa - pb;
      return String(a.name).localeCompare(String(b.name), undefined, { sensitivity: 'base' });
    });
    const monthLabels = [2, 1, 0].map((back) => {
      const d = new Date(sysY, sysM - 1 - back, 1);
      return ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][d.getMonth()] + '-' + String(d.getFullYear()).slice(2);
    });
    const headers = ['Sales 1', 'Total'].concat(years.map(String)).concat(['YTD']).concat(monthLabels).concat(['Sleeping', 'Never']);
    const spec = baseSpec('客户开户统计');
    spec.values.push(headers);
    let prevGroup = '';
    rows.forEach((x) => {
      const g = String(x.group || '').trim().toUpperCase();
      if (prevGroup && g !== prevGroup) spec.values.push(headers.map(() => ''));
      spec.values.push([x.name, x.total].concat(years.map((y) => x.yc[y] || 0)).concat([x.ytd]).concat(x.months).concat([x.sleeping, x.never]));
      prevGroup = g;
    });
    const grand = { total: 0, yc: {}, ytd: 0, months: [0, 0, 0], sleeping: 0, never: 0 };
    years.forEach((y) => (grand.yc[y] = 0));
    rows.forEach((x) => {
      grand.total += x.total;
      years.forEach((y) => (grand.yc[y] += x.yc[y] || 0));
      grand.ytd += x.ytd;
      for (let i = 0; i < 3; i++) grand.months[i] += x.months[i];
      grand.sleeping += x.sleeping; grand.never += x.never;
    });
    spec.values.push(['Grand Total', grand.total].concat(years.map((y) => grand.yc[y])).concat([grand.ytd]).concat(grand.months).concat([grand.sleeping, grand.never]));
    const n = spec.values.length;
    const cols = headers.length;
    spec.numFmt.push({ range: 'B2:' + colL(cols) + n, fmt: '#,##0;-#,##0;' });
    spec.bold.push(rng(n, 1, n, cols));
    spec.bg.push({ range: rng(n, 1, n, cols), color: '#d9eaf7' });
    spec.align = [{ range: 'A2:A' + n, horiz: 'left' }, { range: 'B2:' + colL(cols) + n, horiz: 'center' }];
    spec.info = (opts.start || opts.end ? 'Period: ' + (opts.start || '...') + ' - ' + (opts.end || '...') + ' | ' : '')
      + 'System As Of ' + latest;
    return spec;
  }

  // ============================================================
  // 12) 客户活跃统计 (customerActivity) — 对应原 runCustomerActivityAnalysis（第13段）
  // opts.mode: detail|sales1|city|customertype|sleeping|never|new
  // 一个维度一张独立 Excel 表（用户要求完全隔离，便于给客户配不同样式）：
  // 每个模式有自己的工作表名，由 app.js 侧的独立报表项（actDetail/actBySales1/...）分别调用。
  const ACT_SHEET = {
    detail: '客户活跃统计', sales1: '活跃-按销售1', city: '活跃-按城市', customertype: '活跃-按客户类型',
    sleeping: '沉睡客户', never: '从未购买', new: '新开户',
  };
  // opts.mode: detail|sales1|city|customertype|sleeping|never|new
  // Active=Start~End 有销售(默认系统近3月)；New=End 月向前3完整月开户；
  // Sleeping=系统近3月无购买；Never=截至系统截止月从未购买。
  // ============================================================
  function customerActivity(data, opts, t) {
    const mode = ['detail', 'sales1', 'city', 'customertype', 'sleeping', 'never', 'new'].includes(opts.mode) ? opts.mode : 'detail';
    const customers = opts.customers || [];
    // 客户主数据为空：detail 无从构建（只剩标题/口径行），给明确指引而不是空表
    if (!customers.length) return { error: 'NO_CUSTOMERS' };
    // 模糊查询（可选）：客户/销售/城市，命中=子串包含（不区分大小写）；不填=全部客户。
    // 本报表不受起始时间限制：时间窗只用于 Active/Sleeping/New 的状态判定，行集始终是全部客户。
    const fCust = String(opts.fCust || '').trim().toLowerCase();
    const fSales = String(opts.fSales || '').trim().toLowerCase();
    const fCity = String(opts.fCity || '').trim().toLowerCase();
    const hit = (v, q) => !q || String(v || '').toLowerCase().includes(q);
    const groupMap = sales1Groups(opts.reps || { s1: [], s2: [] });
    const salesMap = new Map();
    (opts.reps && opts.reps.s1 || []).forEach((r) => {
      const code = String(r.code || '').trim().toUpperCase();
      const g = String(r.group || '').trim().toUpperCase();
      if (!code || code === 'IC' || g === 'IC') return;
      salesMap.set(code, { name: String(r.name || '').trim() || code, group: g });
    });
    let latest = '', earliest = '';
    data.forEach((r) => {
      const dt = date(r.soldDate);
      if (!dt) return;
      if (dt > latest) latest = dt;
      if (!earliest || dt < earliest) earliest = dt;
    });
    if (!latest) return { error: 'NO_DATA' };
    const sysY = Number(latest.slice(0, 4)), sysM = Number(latest.slice(5, 7));
    const monthStart = (y, m, back) => { const d = new Date(y, m - 1 - back, 1); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-01'; };
    const monthEnd = (y, m) => { const d = new Date(y, m, 0); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); };
    // 近期窗口=数据最后月份往前 6 个月（用户口径，考虑数据滞后：以数据内最后日期为基准，非今天）。
    // 例：数据到 2026-08 → 窗口 2026-03-01 ~ 2026-08-31；Active 默认窗口与 Sleeping 判定都用它。
    const systemRecentStart = monthStart(sysY, sysM, 5);
    const activeStart = opts.start || systemRecentStart;
    const activeEnd = opts.end || monthEnd(sysY, sysM);
    if (activeStart > activeEnd) return { error: 'BAD_ORDER' };
    // New=数据最后月份往前 3 个完整月（同样以数据末月为基准）。例：数据到 2026-08 → 2026-06-01 ~ 2026-08-31
    const newEnd = monthEnd(Number(activeEnd.slice(0, 4)), Number(activeEnd.slice(5, 7)));
    const newStart = monthStart(Number(activeEnd.slice(0, 4)), Number(activeEnd.slice(5, 7)), 2);
    const custs = new Map();
    customers.forEach((c) => {
      const id = String(c.no || '').trim().toUpperCase();
      if (!id) return;
      const salesCode = String(c.sales1 || '').trim().toUpperCase();
      const salesInfo = salesMap.get(salesCode);
      if (salesCode === 'IC' || (salesInfo && String(salesInfo.group || '').toUpperCase() === 'IC')) return;
      const created = date(c.created);
      custs.set(id, {
        id, name: String(c.name || '').trim(), salesCode,
        sales1: salesInfo ? salesInfo.name : (salesCode || 'Unassigned'),
        salesGroup: salesInfo ? String(salesInfo.group || '').trim() : '',
        city: String(c.city || '').trim() || 'Unknown',
        customerType: String(c.type || '').trim() || 'Unknown',
        created, amt: 0, orders: 0, lastSale: '', activeIn: false, recentSale: false,
      });
    });
    data.forEach((r) => {
      const dt = date(r.soldDate);
      if (!dt || dt > latest) return;
      const c = custs.get(String(r.custNo || '').trim().toUpperCase());
      if (!c) return;
      if (!c.name && r.custName) c.name = String(r.custName).trim();
      c.orders += 1;
      c.amt += amount(r.amount);
      if (dt > c.lastSale) c.lastSale = dt;
      if (dt >= activeStart && dt <= activeEnd) c.activeIn = true;
      if (dt >= systemRecentStart) c.recentSale = true;
    });
    const daysBetween = (a, b) => Math.floor((new Date(b) - new Date(a)) / 86400000);
    const detail = [];
    custs.forEach((c) => {
      if (!hit(c.id, fCust) && !hit(c.name, fCust)) return;
      if (!hit(c.salesCode, fSales) && !hit(c.sales1, fSales)) return;
      if (!hit(c.city, fCity)) return;
      const isNever = !c.lastSale;
      const isSleeping = !!(c.lastSale && !c.recentSale);
      const isNew = !!(c.created && c.created >= newStart && c.created <= newEnd);
      let status = '';
      if (isNever) status = 'Never';
      else if (isSleeping) status = 'Sleeping';
      else if (c.activeIn) status = 'Active';
      else status = 'Inactive';
      detail.push({
        id: c.id, name: c.name, sales1: c.sales1, salesGroup: c.salesGroup, city: c.city,
        customerType: c.customerType, created: c.created, amt: c.amt, orders: c.orders,
        lastSale: c.lastSale, days: c.lastSale ? daysBetween(c.lastSale, latest) : '',
        status, isActive: c.activeIn, isNew, isSleeping, isNever,
      });
    });
    const summary = { total: detail.length, active: 0, sleeping: 0, never: 0, newCount: 0 };
    detail.forEach((x) => {
      if (x.isActive) summary.active++;
      if (x.isSleeping) summary.sleeping++;
      if (x.isNever) summary.never++;
      if (x.isNew) summary.newCount++;
    });
    const statusOrder = { Active: 1, Sleeping: 2, Never: 3, New: 4 };
    const detailRow = (x) => [x.id, x.name, x.sales1, x.city, x.created || '', x.amt, x.orders, x.lastSale || '', x.days, x.status];
    const DETAIL_HEAD = ['No.', 'Customer Name', 'Sales Rep', 'City', 'Since', 'Sales Amount', 'Orders', 'LSD', 'DSLS', 'Status'];
    const spec = baseSpec(ACT_SHEET[mode] || '客户活跃统计');
    const money = '#,##0.00;-#,##0.00;';
    const pct = '0.0%';

    if (mode === 'detail') {
      spec.values.push(DETAIL_HEAD);
      detail.slice().sort((a, b) => {
        const s = String(a.sales1).localeCompare(String(b.sales1), undefined, { sensitivity: 'base' });
        if (s !== 0) return s;
        const st = (statusOrder[a.status] || 99) - (statusOrder[b.status] || 99);
        if (st !== 0) return st;
        if (b.amt !== a.amt) return b.amt - a.amt;
        return String(a.name).localeCompare(String(b.name), undefined, { sensitivity: 'base' });
      }).forEach((x) => spec.values.push(detailRow(x)));
    } else if (mode === 'sales1' || mode === 'city' || mode === 'customertype') {
      const keyName = mode === 'sales1' ? 'sales1' : mode === 'city' ? 'city' : 'customerType';
      const stats = new Map();
      detail.forEach((x) => {
        const key = String(x[keyName] || 'Unknown').trim() || 'Unknown';
        if (!stats.has(key)) stats.set(key, { name: key, total: 0, active: 0, sleeping: 0, never: 0, newCount: 0, totalSales: 0, activeSales: 0, sleepingSales: 0 });
        const s = stats.get(key);
        s.total++; s.totalSales += Number(x.amt) || 0;
        if (x.isActive) { s.active++; s.activeSales += Number(x.amt) || 0; }
        if (x.isSleeping) { s.sleeping++; s.sleepingSales += Number(x.amt) || 0; }
        if (x.isNever) s.never++;
        if (x.isNew) s.newCount++;
      });
      const head = [mode === 'sales1' ? 'Sales 1' : mode === 'city' ? 'City' : 'Customer Type',
        'Total', 'Active', 'Sleeping', 'Never', 'New', 'Active %', 'Sleeping %', 'Never %', 'Historical Sales', 'Sleeping Sales'];
      spec.values.push(head);
      let list = Array.from(stats.values());
      if (mode === 'sales1') {
        const byName = new Map();
        salesMap.forEach((info, code) => { if (!byName.has(info.name)) byName.set(info.name, String(info.group || '').toUpperCase()); });
        const rank = { HOUSE: 1, OUTSIDE: 2, 'OUTSIDE SALES': 2, 'OUTSIDE SALE': 2, OTHER: 3, 'OTHER BRANCH': 3 };
        list.sort((a, b) => {
          const pa = rank[byName.get(a.name) || ''] || 99;
          const pb = rank[byName.get(b.name) || ''] || 99;
          if (pa !== pb) return pa - pb;
          return String(a.name).localeCompare(String(b.name), undefined, { sensitivity: 'base' });
        });
        let prevRank = null;
        list.forEach((s) => {
          const r = rank[byName.get(s.name) || ''] || 99;
          if (prevRank !== null && r !== prevRank) spec.values.push(head.map(() => ''));
          spec.values.push([s.name, s.total, s.active, s.sleeping, s.never, s.newCount,
            s.total ? s.active / s.total : 0, s.total ? s.sleeping / s.total : 0, s.total ? s.never / s.total : 0, s.totalSales, s.sleepingSales]);
          prevRank = r;
        });
      } else {
        list.sort((a, b) => b.total - a.total || String(a.name).localeCompare(String(b.name), undefined, { sensitivity: 'base' }));
        list.forEach((s) => spec.values.push([s.name, s.total, s.active, s.sleeping, s.never, s.newCount,
          s.total ? s.active / s.total : 0, s.total ? s.sleeping / s.total : 0, s.total ? s.never / s.total : 0, s.totalSales, s.sleepingSales]));
      }
      // 合计行
      const tot = list.reduce((acc, s) => {
        acc.total += s.total; acc.active += s.active; acc.sleeping += s.sleeping; acc.never += s.never;
        acc.newCount += s.newCount; acc.totalSales += s.totalSales; acc.sleepingSales += s.sleepingSales;
        return acc;
      }, { total: 0, active: 0, sleeping: 0, never: 0, newCount: 0, totalSales: 0, sleepingSales: 0 });
      const totalName = 'Grand Total';
      const totalRow = [totalName, tot.total, tot.active, tot.sleeping, tot.never, tot.newCount,
        tot.total ? tot.active / tot.total : 0, tot.total ? tot.sleeping / tot.total : 0,
        tot.total ? tot.never / tot.total : 0, tot.totalSales, tot.sleepingSales];
      spec.values.push(totalRow);
      const n = spec.values.length;
      // 列序：A分组 B Total C Active D Sleeping E Never F New G Active% H Sleeping% I Never% J Hist.K Sales L Sleeping Sales
      spec.numFmt.push({ range: 'B2:F' + n, fmt: '#,##0;-#,##0;' }, { range: 'G2:I' + n, fmt: pct },
        { range: 'J2:J' + n, fmt: money }, { range: 'K2:K' + n, fmt: money });
      const lastIdx = spec.values.length;
      spec.bold.push(rng(lastIdx, 1, lastIdx, 11));
      spec.bg.push({ range: rng(lastIdx, 1, lastIdx, 11), color: '#d9eaf7' });
      spec.freeze = { rows: 1, cols: 1 };
    } else {
      // sleeping / never / new 名单
      const list = detail.filter((x) => mode === 'new' ? x.isNew : mode === 'sleeping' ? x.isSleeping : x.isNever);
      if (mode === 'sleeping') {
        list.sort((a, b) => b.amt - a.amt || (Number(b.days) || 0) - (Number(a.days) || 0));
      } else {
        list.sort((a, b) => (b.created || '').localeCompare(a.created || '') ||
          String(a.name).localeCompare(String(b.name), undefined, { sensitivity: 'base' }));
      }
      spec.values.push(DETAIL_HEAD);
      list.forEach((x) => spec.values.push(detailRow(x)));
    }
    const n = spec.values.length;
    const isGrouped = mode === 'sales1' || mode === 'city' || mode === 'customertype';
    if (spec.values.length > 1 && !isGrouped) {
      // 明细模式列序：A No. B Name C SalesRep D City E Since F Amount G Orders H LSD I DSLS J Status
      spec.numFmt.push({ range: 'F2:F' + n, fmt: money });
      spec.numFmt.push({ range: 'G2:G' + n, fmt: '#,##0;-#,##0;' });
      spec.numFmt.push({ range: 'I2:I' + n, fmt: '#,##0;-#,##0;' });
      spec.align = [{ range: 'B2:E' + n, horiz: 'left' }, { range: 'G2:J' + n, horiz: 'center' }];
    }
    if (!spec.freeze) spec.freeze = { rows: 1, cols: 0 };   // 首行锁定
    spec.bold.push(rng(1, 1, 1, spec.values[0].length));
    spec.bg.push({ range: rng(1, 1, 1, spec.values[0].length), color: '#d9eaf7' });
    spec.widths = [[1, 120], [2, 240], [3, 140], [4, 120], [5, 110], [6, 130], [7, 80], [8, 110], [9, 80], [10, 100], [11, 130], [12, 130]];
    const windowText = mode === 'sleeping' ? 'Sleeping Window: ' + systemRecentStart + ' - ' + latest
      : mode === 'never' ? 'System History: ' + (earliest || '') + ' - ' + latest
      : mode === 'new' ? 'New Account Window: ' + newStart + ' - ' + newEnd
      : 'Active: ' + activeStart + ' - ' + activeEnd + ' | New: ' + newStart + ' - ' + newEnd + ' | Sleeping Window: ' + systemRecentStart + ' - ' + latest;
    spec.info = windowText;
    spec.summary = [
      ['Total', summary.total], ['Active', summary.active], ['Sleeping', summary.sleeping],
      ['Never', summary.never], ['New', summary.newCount],
      ['Active %', summary.total ? summary.active / summary.total : 0],
      ['Sleeping %', summary.total ? summary.sleeping / summary.total : 0],
      ['Never %', summary.total ? summary.never / summary.total : 0],
      ['New %', summary.total ? summary.newCount / summary.total : 0],
    ];
    return spec;
  }

  // ============================================================
  // 13) 沉睡客户唤醒统计 (sleepingReactivation) — 原版 Sleeping Customer Reactivation
  // 不使用时间段；扫描全部历史。同客户可多次唤醒，每次"进入 Sleeping 后再次购买"记一行。
  // Sleeping Since = 上一次购买日 + 6 个日历月（月末自动钳制：08/31+6月→次年02/28，闰年02/29）
  // Sleeping Days  = Reactivated Date - Sleeping Since
  // 同一天多张发票只算一个购买日；Sales 1/2 Code = 唤醒当天 Pull 的历史 Code（非客户表当前值）
  // Total Orders/Total Sales = 截至数据最新日期的客户历史累计（Orders=唯一发票数+空票行数）
  // IC / INTERCOMPANY 排除。
  // ============================================================
  function sleepingReactivation(data, opts, t) {
    const customers = opts.customers || [];
    if (!customers.length) return { error: 'NO_CUSTOMERS' };
    const normId = (v) => String(v || '').trim().toUpperCase();
    const salesMap = new Map();
    (opts.reps && opts.reps.s1 || []).forEach((r) => {
      const code = String(r.code || '').trim().toUpperCase();
      if (code) salesMap.set(code, { group: String(r.group || '').trim().toUpperCase() });
    });
    let systemEnd = '';
    data.forEach((r) => { const dt = date(r.soldDate); if (dt && dt > systemEnd) systemEnd = dt; });
    if (!systemEnd) return { error: 'NO_DATA' };

    // 有效客户宇宙：沿用客户+分组规则，排除 IC / INTERCOMPANY
    const customerMap = new Map();
    customers.forEach((c) => {
      const id = normId(c.no);
      if (!id) return;
      const currentSales1 = String(c.sales1 || '').trim().toUpperCase();
      const customerType = String(c.type || '').trim().toUpperCase();
      const creationDate = date(c.created);
      if (currentSales1 === 'IC' || customerType === 'INTERCOMPANY') return;
      const salesInfo = salesMap.get(currentSales1);
      if (!salesInfo) return;
      if (salesInfo.group === 'IC') return;
      if (creationDate && creationDate > systemEnd) return;
      customerMap.set(id, { id, name: String(c.name || '').trim(), totalSales: 0, invoiceSet: new Set(), blankInvoiceCount: 0, purchasesByDay: new Map() });
    });

    // 全历史扫描：累计 Total Orders / Total Sales，并按 客户+购买日 聚合（同日多票=一个购买日）
    data.forEach((r) => {
      const soldDate = date(r.soldDate);
      if (!soldDate || soldDate > systemEnd) return;
      const c = customerMap.get(normId(r.custNo));
      if (!c) return;
      if (!c.name) { const pn = String(r.custName || '').trim(); if (pn) c.name = pn; }
      const invoice = normId(r.invoice);
      if (invoice) c.invoiceSet.add(invoice); else c.blankInvoiceCount++;
      c.totalSales += amount(r.amount);
      let p = c.purchasesByDay.get(soldDate);
      if (!p) { p = { date: soldDate, sales1Code: '', sales2Code: '' }; c.purchasesByDay.set(soldDate, p); }
      if (!p.sales1Code) p.sales1Code = String(r.rep1 || '').trim().toUpperCase();
      if (!p.sales2Code) p.sales2Code = String(r.rep2 || '').trim().toUpperCase();
    });

    // 6 个日历月（非 180 天）：日月进位、日钳制到目标月最后一天
    const addCalendarMonths = (ds, months) => {
      const parts = ds.split('-');
      const y = Number(parts[0]), m = Number(parts[1]), d = Number(parts[2]);
      const total = (m - 1) + months;
      const ty = y + Math.floor(total / 12);
      const tm = ((total % 12) + 12) % 12;
      const lastDay = new Date(ty, tm + 1, 0).getDate();
      return ty + '-' + String(tm + 1).padStart(2, '0') + '-' + String(Math.min(d, lastDay)).padStart(2, '0');
    };
    const dayDiff = (a, b) => Math.round((new Date(b) - new Date(a)) / 86400000);

    const events = [];
    customerMap.forEach((c) => {
      const purchases = Array.from(c.purchasesByDay.values()).sort((a, b) => a.date.localeCompare(b.date));
      if (purchases.length < 2) return;
      const totalOrders = c.invoiceSet.size + c.blankInvoiceCount;
      for (let i = 1; i < purchases.length; i++) {
        const prev = purchases[i - 1], cur = purchases[i];
        const sleepingSince = addCalendarMonths(prev.date, 6);
        if (cur.date < sleepingSince) continue;   // 未满 Sleeping 期就回购，不算唤醒
        events.push({
          customerNo: c.id, customerName: c.name,
          sleepingSince, reactivatedDate: cur.date,
          sleepingDays: Math.max(0, dayDiff(sleepingSince, cur.date)),
          sales1Code: cur.sales1Code || '', sales2Code: cur.sales2Code || '',
          totalOrders, totalSales: c.totalSales,
        });
      }
    });
    // 最近唤醒在前；同日按 Total Sales 高→低；再按客户名
    events.sort((a, b) => b.reactivatedDate.localeCompare(a.reactivatedDate)
      || (b.totalSales - a.totalSales)
      || String(a.customerName || '').localeCompare(String(b.customerName || '')));

    const spec = baseSpec('沉睡客户唤醒统计');
    const money = '#,##0.00;-#,##0.00;';
    spec.values.push(['Customer No.', 'Customer Name', 'Sleeping Since', 'Reactivated Date', 'Sleeping Days', 'Sales 1 Code', 'Sales 2 Code', 'Total Orders', 'Total Sales']);
    events.forEach((e) => {
      spec.values.push([e.customerNo, e.customerName, e.sleepingSince, e.reactivatedDate, e.sleepingDays, e.sales1Code, e.sales2Code, e.totalOrders, e.totalSales]);
    });
    const n = spec.values.length;
    if (n > 1) {
      spec.numFmt.push({ range: 'E2:E' + n, fmt: '#,##0;-#,##0;' },
        { range: 'H2:H' + n, fmt: '#,##0;-#,##0;' },
        { range: 'I2:I' + n, fmt: money });
      spec.align = [{ range: 'A2:D' + n, horiz: 'left' }, { range: 'F2:G' + n, horiz: 'left' }];
    }
    spec.freeze = { rows: 1, cols: 0 };
    spec.bold.push(rng(1, 1, 1, 9));
    spec.bg.push({ range: rng(1, 1, 1, 9), color: '#d9eaf7' });
    spec.widths = [[1, 110], [2, 240], [3, 130], [4, 140], [5, 120], [6, 110], [7, 110], [8, 110], [9, 150]];
    spec.info = 'Sleeping Customer Reactivation | Scans full history (not limited by period) | ' + (n - 1) + ' events';
    return spec;
  }

  // 列号 1 基 → 字母（numFmt 范围用）
  function colL(n) {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  // 客户活跃统计拆分：一个维度一个报表（各自独立 Excel 工作表），mode 由报表项固定
  const actMode = (m) => (data, opts, t) => customerActivity(data, Object.assign({}, opts, { mode: m }), t);

  return {
    annual, salesPerformance, customerOrders, categoryQuery,
    customerStats, cabinetColors, stoneSales, customerDashboardData,
    salesHistory, customerType, salesAccount, customerActivity,
    actDetail: actMode('detail'), actBySales1: actMode('sales1'), actByCity: actMode('city'),
    actByType: actMode('customertype'), actSleeping: actMode('sleeping'), actNever: actMode('never'), actNew: actMode('new'),
    sleepingReactivation, actReact: sleepingReactivation,
    // 报表ID → 专属工作表名（与 baseSpec 名一一对应）：未生成时也用这个名字铺空表，保证菜单=表格一一对应
    sheetNames: {
      annual: '年度销售总表', salesPerformance: '销售业绩', customerOrders: '客户订单统计',
      categoryQuery: '销售查询', customerStats: '客户统计', cabinetColors: '橱柜颜色统计',
      stoneSales: '石材销量统计', salesHistory: '销售历史查询', customerType: '客户类型统计',
      salesAccount: '客户开户统计', customerActivity: '客户活跃统计',
      actDetail: ACT_SHEET.detail, actBySales1: ACT_SHEET.sales1, actByCity: ACT_SHEET.city,
      actByType: ACT_SHEET.customertype, actSleeping: ACT_SHEET.sleeping, actNever: ACT_SHEET.never, actNew: ACT_SHEET.new,
      actReact: '沉睡客户唤醒统计',
    },
    helpers: { date, amount, range, inRange, sales1Groups, isIC, normType, isCabinetType },
  };
})();
