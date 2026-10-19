// BCS Sales Statistics — main application (pywebview + SQLite + Univer)
/* global I18N, STATS, UniverCore, UniverCoreFacade, UniverEngineRender, UniverEngineFormula,
   UniverUi, UniverSheets, UniverSheetsUi, UniverDocs, UniverDocsUi, UniverSheetsFormula,
   UniverSheetsFormulaUi, UniverSheetsNumfmt, UniverSheetsNumfmtUi, UniverDesign,
   UniverDesignZhCN, UniverUiZhCN, UniverSheetsZhCN, UniverSheetsUiZhCN, UniverDocsUiZhCN,
   UniverSheetsFormulaUiZhCN, UniverSheetsNumfmtUiZhCN,
   UniverDesignEnUS, UniverUiEnUS, UniverSheetsEnUS, UniverSheetsUiEnUS, UniverDocsUiEnUS,
   UniverSheetsFormulaUiEnUS, UniverSheetsNumfmtUiEnUS */
(function () {
  'use strict';

  // ================= i18n =================
  let LANG = 'en';
  function t(key) {
    return (I18N[LANG] && I18N[LANG][key]) || I18N.en[key] || key;
  }
  function tf(key, vars) {
    let s = t(key);
    Object.keys(vars || {}).forEach((k) => (s = s.replace('{' + k + '}', vars[k])));
    return s;
  }

  // ================= demo data =================
  function makeDemoData() {
    const reps = {
      s1: [
        { code: 'HO', name: 'HOUSE ACCT', group: 'House' },
        { code: 'SW', name: 'SIMON WEN', group: 'House' },
        { code: 'AB', name: 'Aliasgher Bandali', group: 'OutSide' },
        { code: 'AD', name: 'Ana E. Diaz Rojas', group: 'OutSide' },
        { code: 'KC', name: 'Kevin Canning', group: 'IC' },
      ],
      s2: [
        { code: 'LP', name: 'Lawreen Paulemond', group: 'inside' },
        { code: 'JM', name: 'Jane Moore', group: 'inside' },
        { code: 'SW', name: 'SIMON WEN', group: '' },
      ],
    };
    const customers = [
      { no: 'C00011', name: 'BCS STONE & CABINETS', city: 'Dallas', state: 'TX', phone: '', contact: 'DORIS TAN', type: 'INTERCOMPANY', email: '' },
      { no: 'C00101', name: 'ANAHEIM KITCHEN & BATH', city: 'Anaheim', state: 'CA', phone: '714-399-1588', contact: 'JASON LI', type: 'KITCHEN', email: 'jason@akb.com' },
      { no: 'C00132', name: 'GRANITE DEPOT LLC', city: 'Denver', state: 'CO', phone: '303-555-0132', contact: 'MARIA G.', type: 'STONE YARD', email: 'sales@granitedepot.com' },
      { no: 'C00147', name: 'COASTAL CABINETRY', city: 'Tampa', state: 'FL', phone: '813-555-0147', contact: 'BILL M.', type: 'CABINET', email: '' },
    ];
    const stock = [
      { itemNo: 'SL0073', description: '2cm Gran Violet', available: 7, inTransit: 0 },
      { itemNo: 'SL0129', description: '2cm New Blue', available: 12, inTransit: 3 },
      { itemNo: 'SL0151', description: '2cm Golden River', available: 36, inTransit: 0 },
      { itemNo: 'SL1917', description: '3cm QTZ Jumbo Sparkling', available: 0, inTransit: 0 },
    ];
    const mk = (month, i, o) => Object.assign({
      id: 'demo_' + month + '_' + i, batch: 'b_' + month, month,
      itemNo: '', description: '', color: '', name: '', type: 'SLAB', category: '', productType: '',
      soldDate: month + '-15', invoice: 'INV9' + String(1000 + i), ledger: '',
      custNo: 'C00101', custName: 'ANAHEIM KITCHEN & BATH', rep1: 'SW', rep1Name: 'SIMON WEN',
      rep2: 'LP', rep2Name: 'Lawreen Paulemond', qty: 1, amount: 500, cost: 0,
    }, o);
    const sales = [];
    const custs = ['C00101', 'C00132', 'C00147', 'C00011'];
    const reps1 = ['SW', 'AB', 'AD', 'HO'];
    const custNameMap = { C00101: 'ANAHEIM KITCHEN & BATH', C00132: 'GRANITE DEPOT LLC', C00147: 'COASTAL CABINETRY', C00011: 'BCS STONE & CABINETS' };
    let n = 0;
    ['2026-07', '2026-08'].forEach((month, mi) => {
      for (let i = 0; i < 24; i++) {
        const isCab = i % 3 === 0;
        const isIC = i === 23;
        sales.push(mk(month, n++, {
          itemNo: isCab ? 'CA031' + (i % 5) : 'SL007' + (i % 5),
          description: isCab ? 'B36 WHITE SHAKER' : '2cm Gran Violet',
          color: isCab ? 'WHITE SHAKER' : '',
          type: isCab ? 'CABINET' : 'SLAB',
          soldDate: month + '-' + String(3 + (i % 25)).padStart(2, '0'),
          custNo: custs[i % 4],
          custName: custNameMap[custs[i % 4]],
          rep1: isIC ? 'IC' : reps1[i % 4],
          rep1Name: isIC ? 'Intercompany' : '',
          rep2: i % 2 ? 'LP' : 'JM',
          qty: 1 + (i % 6),
          amount: 200 + i * 137 + mi * 500,
        }));
      }
    });
    return {
      version: 1, lang: 'en', seq: 10000,
      batches: [
        { id: 'b_2026-07', month: '2026-07', file: 'Sales Record 202607.xlsx', rows: 24, importedAt: '2026-08-05 09:00' },
        { id: 'b_2026-08', month: '2026-08', file: 'Sales Record 202608.xlsx', rows: 24, importedAt: '2026-09-05 09:00' },
      ],
      sales, customers, stock, reps,
    };
  }

  // ================= bridge =================
  let BRIDGE = null;
  let USING_MOCK = false;

  function makeMockBridge() {
    let D = null;
    try { D = JSON.parse(localStorage.getItem('bcs_mock_data') || 'null'); } catch (e) { /* ignore */ }
    // 默认空库：exe 启动时 pywebview 桥接注入前的 mock 阶段会用这份数据渲染第一屏，
    // 若默认演示数据会在真机上闪出"残留数据"（载入演示仅经菜单显式触发）
    if (!D || D.version !== 1) D = { version: 1, batches: [], sales: [], customers: [], stock: [], reps: { s1: [], s2: [] }, lang: 'en' };
    const persist = () => { try { localStorage.setItem('bcs_mock_data', JSON.stringify(D)); } catch (e) { /* ignore */ } };
    let snap = null; // 备份快照（浏览器模拟用）
    const matches = (r, fields, q) => {
      const s = q.toLowerCase();
      return fields.some((f) => String(r[f] || '').toLowerCase().includes(s));
    };
    const byMonthDesc = (a, b) => (b.month || '').localeCompare(a.month || '') || (a.id < b.id ? 1 : -1);
    let mockPrefs = {};
    try { mockPrefs = JSON.parse(localStorage.getItem('bcs_mock_prefs') || '{}'); } catch (e) { /* ignore */ }
    const persistPrefs = () => { try { localStorage.setItem('bcs_mock_prefs', JSON.stringify(mockPrefs)); } catch (e) { /* ignore */ } };
    return {
      async get_meta() {
        return { lang: D.lang || 'en', prefs: mockPrefs, sales: D.sales.length, batches: D.batches.length, customers: D.customers.length, stock: D.stock.length };
      },
      async set_pref(key, value) { mockPrefs[key] = value; persistPrefs(); return 'ok'; },
      async close_app() { return 'ok'; },
      async set_lang(lang) { D.lang = lang; persist(); return 'ok'; },
      async query_sales(offset, limit, filters) {
        let f = filters || {};
        if (typeof f === 'string') { try { f = JSON.parse(f); } catch (e) { f = {}; } }
        let list = D.sales;
        const q = String(f.q || '').toLowerCase();
        if (q) list = list.filter((r) => matches(r, ['itemNo', 'description', 'color', 'type', 'invoice', 'custNo', 'custName', 'rep1', 'rep2', 'month'], q));
        if (f.start) list = list.filter((r) => r.soldDate >= f.start);
        if (f.end) list = list.filter((r) => r.soldDate <= f.end);
        if (f.custNo) list = list.filter((r) => String(r.custNo || '').toLowerCase().includes(String(f.custNo).toLowerCase()));
        if (f.custName) list = list.filter((r) => String(r.custName || '').toLowerCase().includes(String(f.custName).toLowerCase()));
        if (f.type) list = list.filter((r) => String(r.type || '').toLowerCase().includes(String(f.type).toLowerCase()));
        if (f.color) list = list.filter((r) => String(r.color || '').toLowerCase().includes(String(f.color).toLowerCase()));
        if (f.rep1) list = list.filter((r) => String(r.rep1 || '').toLowerCase().includes(String(f.rep1).toLowerCase()));
        if (f.description) list = list.filter((r) => String(r.description || '').toLowerCase().includes(String(f.description).toLowerCase()));
        if (f.invoice) list = list.filter((r) => String(r.invoice || '').toLowerCase().includes(String(f.invoice).toLowerCase()));
        if (f.itemNo) list = list.filter((r) => String(r.itemNo || '').toLowerCase().includes(String(f.itemNo).toLowerCase()));
        list = list.slice().sort(byMonthDesc);
        const page = list.slice(offset, offset + Math.min(limit, 60000));
        const cols = ['id', 'month', 'itemNo', 'description', 'color', 'type', 'soldDate', 'invoice',
          'custNo', 'custName', 'rep1', 'rep1Name', 'rep2', 'rep2Name', 'qty', 'amount', 'cost'];
        return { total: list.length, cols, rows: page.map((r) => cols.map((k) => r[k] === undefined ? '' : r[k])) };
      },
      async update_sale(id, fieldsJson) {
        const r = D.sales.find((x) => x.id === id);
        if (r) Object.assign(r, JSON.parse(fieldsJson));
        persist();
        return 'ok';
      },
      async apply_sales_import(payloadJson) {
        const p = JSON.parse(payloadJson);
        const batchId = 'b_' + p.month;
        D.batches = D.batches.filter((b) => b.month !== p.month);
        D.sales = D.sales.filter((r) => r.batch !== batchId);
        p.records.forEach((r, i) => { r.id = batchId + '_' + i; r.batch = batchId; r.month = p.month; });
        D.batches.push({ id: batchId, month: p.month, file: p.file, rows: p.records.length, importedAt: new Date().toISOString().slice(0, 16).replace('T', ' ') });
        D.sales.push(...p.records);
        // 部门列 → 分组同步（只增量：新代码补入、库内组别为空才采用；姓名/已有组别保留）
        const canon = { house: 'House', outside: 'OutSide', other: 'Other Branch', 'other branch': 'Other Branch', ic: 'IC' };
        const up = (kind, code, g) => {
          const list = D.reps[kind];
          const hit = list.find((x) => String(x.code).toUpperCase() === code);
          if (hit) { if (!hit.deleted && g && !hit.group) hit.group = g; }
          else list.push({ code, name: '', group: g });
        };
        const m1 = {}, m2 = {};
        p.records.forEach((r) => {
          const c1 = String(r.rep1 || '').trim().toUpperCase();
          const g1 = String(r.rep1Group || '').trim();
          if (c1) m1[c1] = g1 ? (canon[g1.toLowerCase()] || g1) : (c1 === 'IC' ? 'IC' : '');
          const c2 = String(r.rep2 || '').trim().toUpperCase();
          const g2 = String(r.rep2Group || '').trim();
          if (c2) m2[c2] = g2 ? (g2.toLowerCase() === 'inside' ? 'inside' : g2) : '';
        });
        Object.keys(m1).forEach((c) => up('s1', c, m1[c]));
        Object.keys(m2).forEach((c) => up('s2', c, m2[c]));
        // 多出来的 Pull 列尽量用库内数据套上（与 store._backfill_pull_columns 同规则）：
        // Rep.Name←分组表；Rep.2←Rep.1、Rep.2 Name←Rep.1 Name→分组表→Ex Sales Rep；
        // Name/Category/ProductType←同品名库内已知属性；Ledger/Cost 无来源留空
        const attr = {};
        D.sales.forEach((r) => {
          if (r.batch !== batchId && (r.name || r.category || r.productType))
            attr[r.itemNo] = attr[r.itemNo] || { name: r.name || '', category: r.category || '', productType: r.productType || '' };
        });
        D.sales.filter((r) => r.batch === batchId).forEach((r) => {
          const h1 = D.reps.s1.find((x) => x.code === String(r.rep1 || '').toUpperCase());
          if (!r.rep1Name && r.rep1 && h1 && h1.name) r.rep1Name = h1.name;
          if (!r.rep2 && r.rep1) r.rep2 = r.rep1;
          if (!r.rep2Name && r.rep2) {
            if (r.rep1Name) r.rep2Name = r.rep1Name;
            else {
              const h2 = D.reps.s2.find((x) => x.code === String(r.rep2).toUpperCase());
              r.rep2Name = (h2 && h2.name) || 'Ex Sales Rep';
            }
          }
          const a = attr[r.itemNo];
          if (a) { if (!r.name) r.name = a.name; if (!r.category) r.category = a.category; if (!r.productType) r.productType = a.productType; }
        });
        const n1 = Object.keys(m1).length, n2 = Object.keys(m2).length;
        persist();
        return { ok: true, month: p.month, rows: p.records.length, repsSynced: { s1: n1, s2: n2 } };
      },
      async delete_batch(batchId) {
        D.sales = D.sales.filter((r) => r.batch !== batchId);
        D.batches = D.batches.filter((x) => x.id !== batchId);
        persist();
        return { ok: true };
      },
      async list_batches() {
        return D.batches.slice().sort((a, b) => b.month.localeCompare(a.month));
      },
      async get_stats_context(start, end) {
        const sales = D.sales.filter((r) => (!start || r.soldDate >= start) && (!end || r.soldDate <= end));
        const cols = ['month', 'itemNo', 'description', 'color', 'type', 'soldDate', 'invoice',
          'custNo', 'custName', 'rep1', 'rep2', 'qty', 'amount', 'cost'];
        return { cols, rows: sales.map((r) => cols.map((k) => r[k])), reps: D.reps, stock: D.stock,
          customers: D.customers.map((c) => ({ no: c.no, name: c.name, city: c.city, state: c.state,
            phone: c.phone, contact: c.contact, type: c.type, email: c.email, sales1: c.sales1 || '', created: c.created || '' })) };
      },
      async top_customers(n) {
        const map = {};
        D.sales.forEach((r) => {
          if (String(r.rep1 || '').toUpperCase() === 'IC' || !r.custNo) return;
          map[r.custNo] = map[r.custNo] || { no: r.custNo, name: r.custName, total: 0 };
          map[r.custNo].total += Number(r.amount) || 0;
        });
        return Object.values(map).sort((a, b) => b.total - a.total).slice(0, n);
      },
      async get_lookup_data() {
        const uniq = (arr) => Array.from(new Set(arr.filter(Boolean)));
        const reps = [...D.reps.s1, ...D.reps.s2].filter((r) => !r.deleted)
          .map((r) => ({ code: r.code, name: r.name || '' }));
        const cities = uniq(D.customers.map((c) => String(c.city || '').trim())).sort();
        return {
          items: uniq(D.sales.map((r) => String(r.itemNo || '').trim())).sort().slice(0, 20000)
            .map((code) => ({ code, name: (D.sales.find((r) => String(r.itemNo || '').trim() === code) || {}).description || '' })),
          customers: uniq(D.sales.map((r) => String(r.custNo || '').trim())).sort().slice(0, 20000)
            .map((code) => ({ code, name: (D.sales.find((r) => String(r.custNo || '').trim() === code) || {}).custName || '' })),
          invoices: uniq(D.sales.map((r) => String(r.invoice || '').trim())).sort().reverse().slice(0, 20000),
          types: uniq(D.sales.map((r) => String(r.type || '').trim())).sort(),
          colors: uniq(D.sales.map((r) => String(r.color || '').trim())).sort(),
          reps,
          cities,
        };
      },
      async export_sheet_xlsx(payloadJson) {
        let p = {};
        try { p = typeof payloadJson === 'string' ? JSON.parse(payloadJson) : (payloadJson || {}); } catch (e) { /* ignore */ }
        return { ok: true, path: 'browser://downloads/' + (p.name || 'Sheet') + ' + demo.xlsx', rows: (p.values || []).length };
      },
      async get_all_customers(search) {
        let list = D.customers;
        if (search) list = list.filter((r) => matches(r, ['no', 'name', 'city', 'state', 'phone', 'contact', 'type', 'email'], search));
        return list.slice(0, 5000);
      },
      async upsert_customer(json) { const c = JSON.parse(json); const i = D.customers.findIndex((x) => x.no === c.no); if (i >= 0) D.customers[i] = c; else D.customers.push(c); persist(); return 'ok'; },
      async delete_customer(no) { D.customers = D.customers.filter((x) => x.no !== no); persist(); return 'ok'; },
      async get_all_stock(search) {
        let list = D.stock;
        if (search) list = list.filter((r) => matches(r, ['itemNo', 'description'], search));
        return list.slice(0, 5000);
      },
      async upsert_stock(json) { const s = JSON.parse(json); const i = D.stock.findIndex((x) => x.itemNo === s.itemNo); if (i >= 0) D.stock[i] = s; else D.stock.push(s); persist(); return 'ok'; },
      async delete_stock(itemNo) { D.stock = D.stock.filter((x) => x.itemNo !== itemNo); persist(); return 'ok'; },
      async get_all_reps() {
        return { s1: D.reps.s1.filter((r) => !r.deleted), s2: D.reps.s2.filter((r) => !r.deleted) };
      },
      async upsert_rep(kind, json) {
        const r = JSON.parse(json);
        const i = D.reps[kind].findIndex((x) => x.code === r.code);
        if (i >= 0) D.reps[kind][i] = r; else D.reps[kind].push(r);
        persist();
        return 'ok';
      },
      async delete_rep(kind, code) {
        // 墓碑语义与 store 一致：删除后导入/初始化不再加回；手工新增同代码复活
        const hit = D.reps[kind].find((x) => x.code === code);
        if (hit) { hit.deleted = true; hit.locked = true; }
        else D.reps[kind].push({ code, name: '', group: '', locked: true, deleted: true });
        persist();
        return 'ok';
      },
      async apply_master_import(json) {
        const p = JSON.parse(json);
        // 与 store 语义一致：字段显式传入即覆盖（空数组=清空），未传入则不动
        if ('customers' in p) D.customers = p.customers || [];
        if ('stock' in p) D.stock = p.stock || [];
        // 销售分组只增量：新代码补入；已有代码的姓名/组别仅在为空时补，永不清空、永不覆盖
        const inc = (kind, list) => (list || []).forEach((r) => {
          const hit = D.reps[kind].find((x) => x.code === r.code);
          if (hit) { if (!hit.name && r.name) hit.name = r.name; if (!hit.group && r.group) hit.group = r.group; }
          else D.reps[kind].push({ code: r.code, name: r.name || '', group: r.group || '' });
        });
        if (p.reps) { inc('s1', p.reps.s1); inc('s2', p.reps.s2); }
        persist();
        return 'ok';
      },
      async clear_all_data() {
        D.sales = []; D.batches = []; D.customers = []; D.stock = []; D.reps = { s1: [], s2: [] };
        persist();
        return 'ok';
      },
      async import_sales_files() {
        // 测试钩子：浏览器演示模式下可用 setMockSalesFiles 喂真实文件数据（exe 真桥不受影响）
        if (S.mockSalesFiles && S.mockSalesFiles.length) return { ok: true, files: S.mockSalesFiles };
        const mk = (month, day, cust, custName, amt, inv, item) =>
          ['SLAB', item, 'Mock desc ' + inv, '', month + '-' + day, inv, cust, custName, 'AB', '', 'LP', '', '2', String(amt)];
        const files = [];
        // 月度文件（文件名识别月份）
        files.push({
          ok: true, fileName: 'Sales Record 202609.xlsx',
          headers: ['类别', '品名', '描述', '颜色', '销售日期', '发票号码', '客户编号', '客户名称', '销售1', '部门1', '销售2', '部门2', '数量', '金额'],
          rows: [mk('2026-09', '05', 'C00101', 'ANAHEIM KITCHEN & BATH', 411, 'INVM9001', 'SL0071')],
        });
        // 跨月历史大文件（按销售日期拆分）
        const rows2 = [];
        [['2026-05', 'INMH7001', 901], ['2026-05', 'INMH7002', 902], ['2026-06', 'INMH7003', 903]].forEach(([m, inv, amt]) => {
          rows2.push(['TILE', 'TL7701', 'Mock history', '', m + '-10', inv, 'C00132', 'GRANITE DEPOT LLC', 'AB', '', 'JM', '', '7', String(amt)]);
        });
        files.push({
          ok: true, fileName: 'History_2026.xlsx',
          headers: ['类别', '品名', '描述', '颜色', '销售日期', '发票号码', '客户编号', '客户名称', '销售1', '部门1', '销售2', '部门2', '数量', '金额'],
          rows: rows2,
        });
        return { ok: true, files };
      },
      async initialize_data() {
        let filled = 0;
        D.sales.forEach((r) => {
          if (!r.rep2 && r.rep1) { r.rep2 = r.rep1; filled++; }
          if (r.rep2 && !r.rep2Name) r.rep2Name = r.rep1Name || 'Ex Sales Rep';
        });
        const have = (kind) => new Set(D.reps[kind].map((x) => String(x.code).toUpperCase()));   // 含墓碑行
        const h1 = have('s1'), h2 = have('s2');
        const uniq = {};
        D.sales.forEach((r) => {
          if (r.rep1) uniq[r.rep1.toUpperCase()] = r.rep1Name || r.rep1;
          if (r.rep2) uniq[r.rep2.toUpperCase()] = r.rep2Name || r.rep2;
        });
        const newS1 = [], newS2 = [];
        Object.keys(uniq).sort().forEach((code) => {
          if (!h1.has(code)) { D.reps.s1.push({ code, name: uniq[code], group: '' }); newS1.push(code); }
          if (!h2.has(code)) { D.reps.s2.push({ code, name: uniq[code], group: '' }); newS2.push(code); }
        });
        persist();
        return { ok: true, sales2Filled: filled, newS1, newS2 };
      },
      async import_master_file() { return { ok: false, error: 'mock' }; },
      async import_customers_file() {
        return { ok: true, fileName: 'mock_customers.xlsx',
          customers: [{ no: 'C90001', name: 'MOCK CUSTOMER A', city: '', state: '', phone: '', contact: '', type: '', email: '' },
                      { no: 'C90002', name: 'MOCK CUSTOMER B', city: '', state: '', phone: '', contact: '', type: '', email: '' }] };
      },
      async import_stock_file() {
        return { ok: true, fileName: 'mock_stock.xlsx',
          stock: [{ itemNo: 'MOCK001', description: 'Mock stock item', available: 5, inTransit: 1 }] };
      },
      async export_backup(optionsJson) {
        let o = {};
        try { o = typeof optionsJson === 'string' ? JSON.parse(optionsJson) : (optionsJson || {}); } catch (e) { /* ignore */ }
        snap = JSON.parse(JSON.stringify(D));
        snap.__kind = o.mode || 'full';
        return { ok: true, path: 'memory://backup (' + snap.__kind + ')' };
      },
      async backup_inspect() {
        if (!snap) return { ok: false, error: 'browser: export backup first' };
        const months = {};
        snap.sales.forEach((r) => { months[r.month] = (months[r.month] || 0) + 1; });
        return {
          ok: true, kind: snap.__kind || 'full', path: 'memory://backup',
          months: Object.keys(months).sort().map((m) => ({ month: m, rows: months[m] })),
          salesTotal: snap.sales.length,
          customers: snap.customers.length, stock: snap.stock.length,
          reps: (snap.reps.s1 || []).length + (snap.reps.s2 || []).length,
        };
      },
      async backup_restore(payloadJson) {
        const p = typeof payloadJson === 'string' ? JSON.parse(payloadJson) : payloadJson;
        if (!snap) return { ok: false, error: 'no backup' };
        if (p.mode === 'replace') {
          D = JSON.parse(JSON.stringify(snap));
        } else {
          for (const b of (snap.batches || [])) {
            const exists = D.batches.some((x) => x.month === b.month);
            if (exists && p.monthConflict === 'skip') continue;
            D.batches = D.batches.filter((x) => x.month !== b.month);
            D.batches.push(b);
            D.sales = D.sales.filter((r) => r.month !== b.month);
            D.sales.push(...snap.sales.filter((r) => r.month === b.month));
          }
          if (p.masterConflict === 'overwrite') {
            if (snap.customers.length) D.customers = snap.customers.slice();
            if (snap.stock.length) D.stock = snap.stock.slice();
            if ((snap.reps.s1 || []).length || (snap.reps.s2 || []).length) D.reps = JSON.parse(JSON.stringify(snap.reps));
          }
        }
        persist();
        return { ok: true, mode: p.mode };
      },
    };
  }

  if (window.pywebview && window.pywebview.api) {
    BRIDGE = window.pywebview.api;
  } else {
    // 浏览器开发模式先挂 Mock；若实际运行在 pywebview 宿主中，boot 时会重新检测并切换到真实桥接
    BRIDGE = makeMockBridge();
    USING_MOCK = true;
  }

  let costToken='', costDeadline=0;
  function costVisible(){return !!costToken && Date.now()<costDeadline;}
  async function checkCostAccess(){
    if(!costVisible())return false;
    const status=await api('cost_status');
    if(!status.unlocked){costToken='';location.reload();return false;}
    return true;
  }
  async function unlockCost(){
    if(await checkCostAccess()){toast(LANG==='zh'?'管理权限已解锁':'Cost is unlocked');return;}
    const data=await formModal(LANG==='zh'?'解锁管理权限':'Unlock management',[
      {key:'password',label:LANG==='zh'?'管理密码':'Management password',type:'password',value:''}
    ]);
    if(!data)return;
    const result=await api('unlock_cost',data.password);
    if(!result.ok){toast(result.error==='COST_PASSWORD_NOT_CONFIGURED'?
      '请在 Render 设置 COST_VIEW_PASSWORD':result.error==='TOO_MANY_ATTEMPTS'?'尝试过多，请一分钟后再试':'密码不正确');return;}
    costToken=result.token;costDeadline=Date.now()+result.expiresIn*1000;
    // Drop privilege-dependent in-memory results and rebuild the current view.
    S.ctxCache=null;S.dataToken++;S.sheetToken={};
    delete S.statResults.salesHistory;delete S.reportsLoaded.salesHistory;
    S.costRefreshHistory=true;
    await clearAllStatSheets();
    buildMenus();
    navigate(S.nav);
    toast(LANG==='zh'?'管理权限已解锁，当前页面已更新':'Cost unlocked; view refreshed');
  }
  function hideCostColumn(spec){
    if(costVisible() || !spec?.values?.length)return spec;
    const col=spec.values[0].findIndex(h=>/^(cost|成本)$/i.test(String(h).trim()));
    if(col>=0)spec.values=spec.values.map(row=>row.filter((_,i)=>i!==col));
    return spec;
  }
  // A fresh page always starts locked. Tokens exist only in this page's memory.
  setInterval(()=>{if(costToken && !costVisible())location.reload();},1000);
  document.addEventListener('visibilitychange',()=>{if(!document.hidden && costToken)checkCostAccess();});

  // 统一调用：python 侧返回 JSON 字符串时自动解析
  // 远程版：pywebview 不可用时走 HTTP POST /api/<method>
  async function api(method) {
    var args = Array.prototype.slice.call(arguments, 1);
    // 如果是真实 pywebview 桥接且有该方法，走桥接（本地 exe 版）
    if (!USING_MOCK && BRIDGE && BRIDGE[method]) {
      const r = await BRIDGE[method].apply(BRIDGE, args);
      if (typeof r === 'string') {
        try { return JSON.parse(r); } catch (e) { return r; }
      }
      return r;
    }
    // 否则走 HTTP（远程网页版）
    const resp = await fetch('/api/' + method, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'X-Cost-Token': costVisible()?costToken:'' },
      body: JSON.stringify({ args: args }),
    });
    if(costToken && method!=='unlock_cost' && resp.headers.get('X-Cost-Unlocked')==='0'){
      costToken='';location.reload();throw new Error('成本权限已过期，请重新解锁');
    }
    const text = await resp.text();
    try { return JSON.parse(text); } catch (e) { return text; }
  }

  // ================= state =================
  const S = {
    nav: 'sales',
    statDef: null,
    statOpts: {},          // 每个统计报表记住上次的条件
    statResults: {},       // 每个统计报表缓存最近一次结果（列表式）
    statSheetName: {},     // 报表ID → 其专属工作表名（缓存清空后仍能把视图切回自己的表）
    yearToggle: null,      // 年度销售总表折叠状态（点击年份行切换该年月份行显示）
    statTriggered: {},     // 报表ID → 最近触发的功能按钮文字（置灰标记，指示当前表格来源）
    sheetLayouts: {},      // 工作表名 → {w:列宽[], h:行高{行号:高}, fz:冻结}（用户调整的布局自动记忆，重启恢复）
    userStyles: {},        // 工作表名 → {"行:列": 样式对象}（用户手工修饰的单元格格式差异，重启恢复）
    appliedStyles: {},     // 工作表名 → 渲染+回放完成时的样式基线（定时差分用）
    isRendering: false,    // renderSpec 进行中标志（定时快照避让）
    dataToken: 0,          // 数据变更计数（导航时表内容未变则跳过重渲染，毫秒级切换）
    sheetToken: {},        // 工作表名 → 渲染时对应的 (dataToken+查询条件)，用于跳过重渲染
    sheetCtxCache: {},     // 工作表名 → 该页最近一次全量渲染的 sheetCtx；activateIfFresh 短路返回时恢复，避免切回页面后 sheetCtx 为空导致行选不中、Edit/Delete 点了没反应
    lastRenderedSheet: '', // 最近渲染的表（切换/渲染前/关闭时只同步它，避免全表扫描卡顿）
    pendingSaves: [],      // 进行中的异步保存（关闭前 __flushMemory 等它们全部落库）
    appliedLayouts: {},    // 工作表名 → 本会话最后一次应用的列宽（快照基线，仅内存）
    hubTab: {},            // 每个大模块页当前选中的报表选项卡
    sheetCtx: null,        // 当前工作表数据上下文 {kind, rows, startRow, colCount}
    sheetSel: -1,          // 当前选中行（0 基，数据区行号偏移后）
    lastHLRow: null,
    ctxCache: null,
    salesFilters: {},      // 销售记录组合查询条件
    salesPage: 0,
    custQuery: '',
    invQuery: '',
    pageSize: 200,
    histSel: new Set(),
    u: null,               // Univer 工作簿实例（统计报表以 Excel 表格呈现）
    sidebarCollapsed: false,
    navCollapsed: new Set(),
    dash: { period: 'ALL', customerNo: '', cache: null, chart: null, rowMap: {} },
    dashEngines: null,     // 客户动态两块面板各自的 Univer 实例（懒加载）
  };

  function setStatus(msg) { const el = $('#status'); if (el) el.textContent = msg; }
  function now2() {
    const d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0') + ' ' + String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0');
  }
  function saved() { setStatus(t('saved') + ' · ' + now2()); }

  // ================= tiny dom helpers =================
  const $ = (sel) => document.querySelector(sel);
  function esc(v) {
    return String(v === null || v === undefined ? '' : v).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }
  function fmtMoney(v) { return (Number(v) || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  // 展示用：0 / 无数据显示为空（不要单独显示 0、0.00）
  function fmtMoneyZ(v) { return Number(v) === 0 ? '' : fmtMoney(v); }

  // 桥接返回的 gz 压缩包解压（>1MB 的查询结果走 gzip+base64 传输）
  async function gunzipJson(res) {
    if (!res || !res.gz) return res;
    const bin = atob(res.gz);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('gzip'));
    const parsed = JSON.parse(await new Response(stream).text());
    return Object.assign({}, res, { gz: undefined }, parsed);
  }

  // 行数组还原成对象数组（cols/rows 瘦身格式；兼容旧对象格式）
  function rowsToObjects(res, listKey) {
    if (!res || !res.cols || !res.rows) return res;
    const cols = res.cols;
    const list = res.rows.map((arr) => { const o = {}; cols.forEach((k, i) => { o[k] = arr[i]; }); return o; });
    delete res.rows;   // 注意：先 delete 源键再赋值——listKey 为 'rows' 时顺序颠倒会把转换结果删掉
    delete res.cols;
    res[listKey] = list;
    return res;
  }

  async function normalizeCtx(res) {
    return rowsToObjects(await gunzipJson(res), 'sales');
  }

  async function normalizeQuery(res) {
    return rowsToObjects(await gunzipJson(res), 'rows');
  }

  // ================= get_stats_context 缓存（数据变更后由 invalidateCtx 失效） =================
  function getCtx(start, end) {
    const key = (start || '') + '|' + (end || '');
    if (S.ctxCache && S.ctxCache.key === key) return S.ctxCache.data;
    return null;
  }
  function putCtx(start, end, data) {
    S.ctxCache = { key: (start || '') + '|' + (end || ''), data };
  }
  function invalidateCtx() {
    S.ctxCache = null; S.dataToken++;
    Object.values(S.statResults).forEach((r) => { r.stale = true; });
  }
  function reportState(spec) {
    if (!spec) return t('statNotGen');
    return tf('statStateDone', { t: spec.doneAt || '' }) +
      (spec.saveFailed ? (LANG === 'zh' ? ' · 保存失败，请重新生成重试' : ' · Save failed; generate again to retry') :
       spec.stale ? (LANG === 'zh' ? ' · 数据已更新，请重新生成' : ' · Data changed; please regenerate') :
       (LANG === 'zh' ? ' · 已保存' : ' · Saved'));
  }

  // ================= loading 反馈 =================
  function loadBar(on) {
    const b = document.getElementById('loadbar');
    if (!b) return;
    if (on) { b.className = 'loadbar on'; }
    else { b.className = 'loadbar done'; setTimeout(() => { if (b.className === 'loadbar done') b.className = 'loadbar'; }, 350); }
  }
  function statOverlay(on, text) {
    const mount = $('#univerMount');
    if (!mount) return;
    let ov = document.getElementById('stat-loading');
    if (on) {
      if (!ov) {
        ov = document.createElement('div');
        ov.id = 'stat-loading';
        ov.className = 'loading-overlay';
        ov.innerHTML = '<div class="spin"></div><div class="txt"></div>';
        mount.appendChild(ov);
      }
      ov.querySelector('.txt').textContent = text || t('loadingGen');
    } else if (ov) ov.remove();
  }
  // 让浏览器先把 loading 画出来再做重活
  function nextPaint() { return new Promise((r) => setTimeout(r, 40)); }

  // ================= fuzzy query dropdown（对应原脚本 销售历史查询 Sidebar 动态模糊下拉） =================
  // 一次读取下拉主数据，之后所有模糊筛选在页面本地完成；最多返回 20 条；选中后回填。
  let LOOKUP = null;
  let REPS_CACHE = null;
  async function getLookup() {
    if (LOOKUP) return LOOKUP;
    try { LOOKUP = await api('get_lookup_data'); } catch (e) { LOOKUP = {}; }
    LOOKUP.items = LOOKUP.items || [];
    LOOKUP.customers = LOOKUP.customers || [];
    LOOKUP.invoices = LOOKUP.invoices || [];
    LOOKUP.types = LOOKUP.types || [];
    LOOKUP.colors = LOOKUP.colors || [];
    return LOOKUP;
  }
  async function getRepsCache() {
    if (REPS_CACHE) return REPS_CACHE;
    try { REPS_CACHE = await api('get_all_reps'); } catch (e) { REPS_CACHE = { s1: [], s2: [] }; }
    return REPS_CACHE;
  }
  const fuzzyHit = (text, ...fields) => {
    const s = String(text || '').toLowerCase();
    return fields.some((f) => String(f || '').toLowerCase().includes(s));
  };
  const cap20 = (list) => list.slice(0, 20);

  // 销售历史查询等精确条件：回填编码（label 显示名称便于辨认）
  // 后端 customers 字段是 no（不是 code），兼容两种写法
  const custNo = (c) => c.no || c.code || '';
  async function fuzzyCustomersCode(text) {
    const lk = await getLookup();
    return cap20(lk.customers.filter((c) => fuzzyHit(text, custNo(c), c.name))
      .map((c) => ({ v: custNo(c), label: c.name, sub: custNo(c) })));
  }
  // 本地主数据来源
  async function fuzzyCustomers(text) {
    const lk = await getLookup();
    return cap20(lk.customers.filter((c) => fuzzyHit(text, custNo(c), c.name))
      .map((c) => ({ v: c.name, label: c.name, sub: custNo(c), data: { no: custNo(c) } })));
  }
  // 后端 items 是纯字符串数组（不是 {code,name} 对象）
  async function fuzzyItems(text) {
    const lk = await getLookup();
    const norm = (c) => (typeof c === 'string' ? { code: c, name: '' } : c);
    return cap20(lk.items.filter((c) => { const n = norm(c); return fuzzyHit(text, n.code, n.name); })
      .map((c) => { const n = norm(c); return { v: n.code, label: n.code, sub: n.name }; }));
  }
  // 发票号走服务端实时搜索：lookup 只缓存最新 5000 张，老发票会被 LIMIT 截掉
  async function fuzzyInvoices(text) {
    try {
      const list = await api('search_invoices', text);
      return cap20((Array.isArray(list) ? list : []).map((v) => ({ v, label: v })));
    } catch (e) { return []; }
  }
  // 描述模糊搜索走服务端：lookup 里没有描述数据
  async function fuzzyDescs(text) {
    try {
      const list = await api('search_descs', text);
      return cap20((Array.isArray(list) ? list : []).map((v) => ({ v, label: v })));
    } catch (e) { return []; }
  }
  async function fuzzySimple(list, text) {
    return cap20(list.filter((v) => fuzzyHit(text, v)).map((v) => ({ v, label: v })));
  }
  async function fuzzyRep1(text) {
    const reps = await getRepsCache();
    return cap20((reps.s1 || []).filter((r) => fuzzyHit(text, r.code, r.name))
      .map((r) => ({ v: r.code, label: r.code, sub: r.name })));
  }

  // 把输入框包上模糊下拉：输入停止 0.3s 弹出候选；↑↓/Enter/Esc 键盘可选；点选后回填
  function attachFuzzy(input, opts) {
    if (!input || input._fuzzy) return;
    input._fuzzy = true;
    input.setAttribute('autocomplete', 'off');
    const wrap = document.createElement('div');
    wrap.className = 'fuzzy-wrap';
    input.parentNode.insertBefore(wrap, input);
    wrap.appendChild(input);
    const pop = document.createElement('div');
    pop.className = 'fuzzy-pop';
    wrap.appendChild(pop);
    let tm = null, items = [], act = -1, seq = 0;
    const hide = () => { pop.classList.remove('show'); act = -1; };
    const pick = (it) => {
      if (!it) return;
      input.value = it.v !== undefined && it.v !== '' ? it.v : it.label;
      hide();
      if (opts.onPick) opts.onPick(it, input);
    };
    const render = (list, emptyText) => {
      items = list; act = -1;
      if (!list.length) {
        pop.innerHTML = '<div class="fuzzy-empty">' + esc(emptyText || t('noMatches')) + '</div>';
      } else {
        pop.innerHTML = list.map((it, i) =>
          '<div class="fuzzy-item" data-fi="' + i + '">' + esc(it.label) +
          (it.sub ? '<span class="sub">' + esc(it.sub) + '</span>' : '') + '</div>').join('');
        pop.querySelectorAll('[data-fi]').forEach((el) => {
          el.addEventListener('mousedown', (e) => { e.preventDefault(); pick(items[Number(el.dataset.fi)]); });
        });
      }
      pop.classList.add('show');
    };
    const run = async () => {
      const my = ++seq;
      let list = [];
      try { list = (await opts.fetch(input.value.trim())) || []; } catch (e) { list = []; }
      if (my === seq) render(list);
    };
    input.addEventListener('input', () => {
      clearTimeout(tm);
      tm = setTimeout(run, 300);
    });
    input.addEventListener('focus', () => { if (!input.value.trim()) run(); });
    input.addEventListener('blur', () => setTimeout(hide, 150));
    input.addEventListener('keydown', (e) => {
      if (!pop.classList.contains('show')) return;
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault();
        if (!items.length) return;
        act = e.key === 'ArrowDown' ? Math.min(act + 1, items.length - 1) : Math.max(act <= 0 ? 0 : act - 1, 0);
        pop.querySelectorAll('.fuzzy-item').forEach((el, i) => el.classList.toggle('act', i === act));
      } else if (e.key === 'Enter') {
        if (act >= 0 && items[act]) pick(items[act]);  // 不拦截：回填后继续触发原 Enter 查询
        else hide();
      } else if (e.key === 'Escape') {
        hide();
      }
    });
  }

  // ================= toast & modal =================
  function toast(msg) {
    const box = $('#toast');
    box.textContent = msg;
    box.classList.add('show');
    clearTimeout(box._tm);
    box._tm = setTimeout(() => box.classList.remove('show'), 3200);
  }

  function closeModal() {
    document.querySelectorAll('.modal-wrap').forEach((w) => w.remove());
  }
  function showModal(title, bodyEl, buttons) {
    closeModal();
    const wrap = document.createElement('div');
    wrap.className = 'modal-wrap';
    const buttonsHtml = buttons.map((b, i) => '<button class="btn ' + (b.primary ? 'primary' : '') + '" data-mi="' + i + '">' + esc(b.label) + '</button>').join('');
    wrap.innerHTML = '<div class="modal"><div class="modal-title"><span>' + esc(title) + '</span><button class="modal-x" aria-label="close">×</button></div><div class="modal-body"></div><div class="modal-foot">' + buttonsHtml + '</div></div>';
    wrap.querySelector('.modal-body').appendChild(bodyEl);
    const close = () => wrap.remove();
    wrap.querySelector('.modal-x').onclick = close;
    wrap.querySelectorAll('[data-mi]').forEach((btn) => {
      btn.onclick = () => {
        const b = buttons[Number(btn.dataset.mi)];
        if (b.onClick) b.onClick(close);
        else close();
      };
    });
    document.body.appendChild(wrap);
    return close;
  }
  function confirmModal(message, okLabel) {
    return new Promise((resolve) => {
      const body = document.createElement('div');
      body.className = 'confirm-text';
      body.textContent = message;
      showModal(t('confirm'), body, [
        { label: t('cancel') },
        {
          label: okLabel || t('ok'), primary: true, onClick: (close) => {
            close();
            resolve(true);
          },
        },
      ]);
      const obs = new MutationObserver(() => {
        if (!document.body.contains(body)) { obs.disconnect(); resolve(false); }
      });
      obs.observe(document.body, { childList: true, subtree: false });
    });
  }
  function formModal(title, fields, okLabel) {
    return new Promise((resolve) => {
      const body = document.createElement('div');
      body.className = 'form';
      fields.forEach((f) => {
        const row = document.createElement('label');
        row.className = 'form-row';
        if (f.type === 'select') {
          row.innerHTML = '<span>' + esc(f.label) + '</span><select data-k="' + f.key + '">' + f.options.map((o) => '<option value="' + esc(o[0]) + '"' + (String(f.value) === String(o[0]) ? ' selected' : '') + '>' + esc(o[1]) + '</option>').join('') + '</select>';
        } else if (f.type === 'check') {
          row.innerHTML = '<span></span><label class="chk"><input type="checkbox" data-k="' + f.key + '"' + (f.value ? ' checked' : '') + '> ' + esc(f.label) + '</label>';
          row.className = 'form-row full';
        } else {
          row.innerHTML = '<span>' + esc(f.label) + '</span><input data-k="' + f.key + '" type="' + (f.type || 'text') + '" value="' + esc(f.value === undefined ? '' : f.value) + '" placeholder="' + esc(f.placeholder || '') + '">';
        }
        body.appendChild(row);
      });
      showModal(title, body, [
        { label: t('cancel') },
        {
          label: okLabel || t('ok'), primary: true, onClick: (close) => {
            const values = {};
            body.querySelectorAll('[data-k]').forEach((el) => {
              values[el.dataset.k] = el.type === 'checkbox' ? el.checked : el.value;
            });
            close(); resolve(values);
          },
        },
      ]);
      const obs = new MutationObserver(() => {
        if (!document.body.contains(body)) { obs.disconnect(); resolve(null); }
      });
      obs.observe(document.body, { childList: true, subtree: false });
    });
  }

  // ================= statistics =================
  // ================= Univer integration（统计报表以 Excel 工作表呈现） =================
  // Univer 应用工厂：工作台主画布与客户动态两块面板（列表/分析）共用同一套插件注册。
  // 一个画布一次只显示一张工作表，所以左右面板各需一个独立实例。
  function makeUniverApp(containerId) {
    const { LocaleType, merge, Univer, UniverInstanceType } = UniverCore;
    const { FUniver } = UniverCoreFacade;
    const { UniverRenderEnginePlugin } = UniverEngineRender;
    const { UniverFormulaEnginePlugin } = UniverEngineFormula;
    const { UniverUIPlugin } = UniverUi;
    const { UniverSheetsPlugin } = UniverSheets;
    const { UniverSheetsUIPlugin } = UniverSheetsUi;
    const { UniverDocsPlugin } = UniverDocs;
    const { UniverDocsUIPlugin } = UniverDocsUi;
    const { UniverSheetsFormulaPlugin } = UniverSheetsFormula;
    const { UniverSheetsFormulaUIPlugin } = UniverSheetsFormulaUi;
    const { UniverSheetsNumfmtPlugin } = UniverSheetsNumfmt;
    const { UniverSheetsNumfmtUIPlugin } = UniverSheetsNumfmtUi;
    const { defaultTheme: theme } = UniverDesign;

    const univer = new Univer({
      theme,
      locale: LocaleType.ZH_CN,
      locales: {
        [LocaleType.ZH_CN]: merge(
          UniverDesignZhCN, UniverUiZhCN, UniverSheetsZhCN, UniverSheetsUiZhCN,
          UniverDocsUiZhCN, UniverSheetsFormulaUiZhCN, UniverSheetsNumfmtUiZhCN,
        ),
        [LocaleType.EN_US]: merge(
          UniverDesignEnUS, UniverUiEnUS, UniverSheetsEnUS, UniverSheetsUiEnUS,
          UniverDocsUiEnUS, UniverSheetsFormulaUiEnUS, UniverSheetsNumfmtUiEnUS,
        ),
      },
    });
    univer.registerPlugin(UniverRenderEnginePlugin);
    univer.registerPlugin(UniverFormulaEnginePlugin);
    univer.registerPlugin(UniverUIPlugin, { container: containerId, footer: !['dashListMount', 'dashAnaMount'].includes(containerId) });
    univer.registerPlugin(UniverDocsPlugin);
    univer.registerPlugin(UniverDocsUIPlugin);
    univer.registerPlugin(UniverSheetsPlugin);
    univer.registerPlugin(UniverSheetsUIPlugin);
    univer.registerPlugin(UniverSheetsFormulaPlugin);
    univer.registerPlugin(UniverSheetsFormulaUIPlugin);
    univer.registerPlugin(UniverSheetsNumfmtPlugin);
    univer.registerPlugin(UniverSheetsNumfmtUIPlugin);
    univer.createUnit(UniverInstanceType.UNIVER_SHEET, {});
    return { univer, api: FUniver.newAPI(univer) };
  }

  function initUniver() {
    if (S.u) return;
    S.u = makeUniverApp('univerMount');
    // 单元格点击 → 行选中（数据管理页的行操作依赖它）
    try {
      const api = S.u.api;
      const sh0 = api.getActiveWorkbook().getSheets()[0];
      let ev = null;
      try { const hooks = sh0.getSheetHooks(); if (hooks && hooks.CellClicked) ev = hooks.CellClicked; } catch (e) { /* ignore */ }
      if (!ev && window.UniverSheetsUi) {
        for (const k of Object.keys(window.UniverSheetsUi)) {
          const o = window.UniverSheetsUi[k];
          if (o && typeof o === 'object' && o.CellClicked) { ev = o.CellClicked; break; }
        }
      }
      if (!ev) ev = 'CellClicked';
      api.addEvent(ev, (params) => onSheetCell(params && typeof params.row === 'number' ? params.row : null));
    } catch (e) { /* ignore */ }
    // 兜底行选中：body 捕获阶段监听 mouseup（画布会吞掉冒泡），从选区读取行号
    try {
      const onMountMouseup = (e) => {
        if (e.target && e.target.closest && e.target.closest('#annualGroupGutter')) return;
        if (!e.target || !e.target.closest || !e.target.closest('#univerMount')) return;
        setTimeout(() => {
          try {
            const wb2 = S.u.api.getActiveWorkbook();
            // 多表时代：从「最近渲染的表」读选区（getSheets()[0] 会读到别的功能表）
            const nm = (S.yearToggle && S.nav === 'r:annual' && S.yearToggle.sheet) || S.lastRenderedSheet;
            const sh2 = wb2.getSheets().find((x) => { try { return nm && x.getSheetName() === nm; } catch (e2) { return false; } }) || wb2.getSheets()[0];
            let row = null;
            const ar = sh2.getActiveRange && sh2.getActiveRange();
            if (ar && typeof ar.getRow === 'function') row = ar.getRow();
            onSheetCell(row);
          } catch (e) { /* ignore */ }
        }, 0);
      };
      document.body.addEventListener('mouseup', onMountMouseup, true);
    } catch (e) { /* ignore */ }
    placeMount();
  }

  // Univer 0.6.9 挂载后会把容器挪到 body 下：拉回工作台容器并触发画布重算
  // （客户动态不用主画布，它有两块自己的挂载：placeDashMounts）
  function placeMount() {
    const host = document.getElementById('page-sheet');
    const m = document.getElementById('univerMount');
    if (!host || !m) return;
    if (m.parentElement !== host) {
      host.appendChild(m);
      window.dispatchEvent(new Event('resize'));
    }
  }

  // 抓取工作表的单元格样式（稀疏表 "行:列" → 样式对象），用于用户格式差异记忆
  function captureCellStyles(sh) {
    const map = {};
    try {
      const maxR = Math.min(sh.getMaxRows(), 300);
      const maxC = Math.min(sh.getMaxColumns(), 60);
      const wraps = sh.getRange(0, 0, maxR, maxC).getCellStyles();
      for (let ri = 0; ri < wraps.length; ri++) {
        const row = wraps[ri] || [];
        for (let ci = 0; ci < row.length; ci++) {
          const w = row[ci];
          const st = w && w._style ? w._style : null;
          if (!st) continue;
          const j = JSON.stringify(st);
          if (j === '{}' || j === '') continue;
          map[ri + ':' + ci] = st;
        }
      }
    } catch (e) { /* ignore */ }
    return map;
  }

  // 表内容未变时仅激活（毫秒级切换，避免大表整张重渲染卡顿）。token 由调用方按 数据版本+查询条件 构造。
  async function activateIfFresh(sheetName, token) {
    if (S.sheetToken[sheetName] !== token || S.isRendering) return false;
    const wb = S.u.api.getActiveWorkbook();
    if (!wb) return false;
    const sh = wb.getSheets().find((x) => { try { return x.getSheetName() === sheetName; } catch (e) { return false; } });
    if (!sh) return false;
    try { sh.activate(); } catch (e) { /* ignore */ }
    placeMount();   // 短路路径也要把画布拉回当前页面容器（如从客户动态切回工作台）
    try {
      const um = document.getElementById('univerMount');
      if (um) um.dispatchEvent(new Event('resize'));
      window.dispatchEvent(new Event('resize'));
    } catch (e) { /* ignore */ }
    S.lastRenderedSheet = sheetName;
    return true;
  }

  // 收集进行中的异步保存：平时 fire-and-forget，关闭软件前 __flushMemory 等待全部落库
  function trackSave(p) {
    if (!p || !p.finally) return;
    S.pendingSaves.push(p);
    p.finally(() => { const i = S.pendingSaves.indexOf(p); if (i >= 0) S.pendingSaves.splice(i, 1); });
  }

  // 把某张表的当前用户调整（布局+单元格样式）与基线差分后落库。
  // 调用时机：渲染前 + 每 2.5s 定时（用户调整后直接关闭软件也不丢）。
  function syncSheetMemory(sh) {
    try {
      const layName = String(sh.getSheetName());
      const base = S.appliedLayouts[layName];
      if (!base) return;   // 本会话未渲染过的表不快照（避免默认值污染记忆）
      // ---- 布局（列宽/行高/冻结） ----
      const widths = [];
      const nCols = Math.min(sh.getMaxColumns(), 60);
      for (let ci = 0; ci < nCols; ci++) {
        let w = null;
        try { w = sh.getColumnWidth(ci); } catch (e) { /* ignore */ }
        widths.push(typeof w === 'number' && w > 0 ? Math.round(w) : null);
      }
      const heights = {};
      const nRows = Math.min(sh.getMaxRows(), 300);
      for (let ri = 0; ri < nRows; ri++) {
        let h = null;
        try { h = sh.getRowHeight(ri); } catch (e) { /* ignore */ }
        // 26=数据行默认高、24=Univer 空表默认高、30=标题行批量高：都视为程序默认值
        if (typeof h === 'number' && h > 0 && h !== 26 && h !== 24 && h !== 30) heights[ri] = Math.round(h);
      }
      let fz = null;
      try {
        const fr = sh.getFreeze();
        if (fr && (fr.ySplit || fr.xSplit)) fz = { y: fr.ySplit || 0, x: fr.xSplit || 0 };
      } catch (e) { /* ignore */ }
      const wChanged = (base.w || []).some((bw, ci) => (widths[ci] || null) !== (bw || null));
      const hKeys = new Set([...Object.keys(base.h || {}), ...Object.keys(heights)]);
      const hChanged = [...hKeys].some((k) => {
        const bv = base.h && base.h[k] === 24 ? null : ((base.h || {})[k] || null);
        return (heights[k] || null) !== bv;
      });
      const fzChanged = JSON.stringify(fz) !== JSON.stringify(base.fz || null);
      // ---- 单元格样式差分（与渲染完成基线比） ----
      let stylesChanged = false;
      const baseStyles = S.appliedStyles[layName];
      if (baseStyles) {
        const live = captureCellStyles(sh);
        const user = S.userStyles[layKeySafe(layName)] || {};
        const merged = Object.assign({}, user);
        const keys = new Set([...Object.keys(baseStyles), ...Object.keys(live)]);
        keys.forEach((k) => {
          const lv = live[k] ? JSON.stringify(live[k]) : '';
          const bv = baseStyles[k] ? JSON.stringify(baseStyles[k]) : '';
          if (lv !== bv) {
            if (lv && lv !== '{}') { merged[k] = live[k]; stylesChanged = true; }
            else if (merged[k]) { delete merged[k]; stylesChanged = true; }
          }
        });
        if (stylesChanged) S.userStyles[layKeySafe(layName)] = merged;
        // 基线同步为当前（避免重复差分），下次用户再改仍能捕获
        S.appliedStyles[layName] = live;
      }
      if (wChanged || hChanged || fzChanged || stylesChanged) {
        S.sheetLayouts[layName] = { w: widths, h: heights, fz: fz || null };
        trackSave(api('set_pref', 'sheetLayouts', JSON.stringify(S.sheetLayouts)));
        if (stylesChanged) {
          trackSave(api('set_pref', 'userStyles', JSON.stringify(S.userStyles)));
        }
      }
    } catch (e) { /* ignore */ }
  }
  function layKeySafe(n) { return String(n || ''); }

  function clearSheet(sh, opts) {
    // 注意：FFacade 数字坐标的 getRange 是 0 基；FRange 只有 clearFormat（单数），
    // clearFormats 是 FWorksheet 的方法——旧代码调错名字异常被吞，样式从未清过（残留根源）。
    // 大表整表 clearContent+clearFormat 是分钟级慢路径（5.3 万行 ≈ 6s）。
    // 传入 keepRows/keepCols（本次写入范围）且有上一次写入记录（S.sheetExtent）时，
    // 只清「新范围之外的尾巴」——重叠区会被本次 setValues 全量覆盖，无需清。
    const ext = S.sheetExtent ? S.sheetExtent[String(sh.getSheetName() || '')] : null;
    try { sh.getMergeData().forEach(r=>r.breakApart()); } catch(e) {}
    try { sh.cancelFreeze(); } catch(e) {}
    try { sh.showColumns(0,sh.getMaxColumns()); } catch(e) {}
    try { sh.showRows(0,Math.min(sh.getMaxRows(),Math.max(opts?.keepRows || 0,ext?.rows || 0))); } catch(e) {}
    if (opts && opts.keepRows != null && ext) {
      const nrows = sh.getMaxRows(), ncols = sh.getMaxColumns();
      const oldR = Math.min(ext.rows, nrows), oldC = Math.min(ext.cols, ncols);
      const nr = opts.keepRows, nc = opts.keepCols;
      try {
        if (oldR > nr) {
          sh.getRange(nr, 0, oldR - nr, oldC).clearContent();
          sh.getRange(nr, 0, oldR - nr, oldC).clearFormat();
        }
        if (oldC > nc) {
          sh.getRange(0, nc, Math.min(nr, oldR), oldC - nc).clearContent();
          sh.getRange(0, nc, Math.min(nr, oldR), oldC - nc).clearFormat();
        }
      } catch (e0) { /* ignore */ }
      return;
    }
    try { sh.getRange(0, 0, sh.getMaxRows(), sh.getMaxColumns()).clearContent(); }
    catch (e) { try { sh.clearContents(); } catch (e2) { /* ignore */ } }
    try { sh.getRange(0, 0, sh.getMaxRows(), sh.getMaxColumns()).clearFormat(); }
    catch (e) { try { sh.clearFormats(); } catch (e2) { /* ignore */ } }
    // 表头类结构残留一并还原：合并的表头块（getMergeData→breakApart）、冻结的表头行/列。
    // 不清的话空表上仍能看到合并格子轮廓和冻结线，新报表再 merge 还可能因旧合并区重叠出错。
    try { sh.getMergeData().forEach((r) => r.breakApart()); } catch (e) { /* ignore */ }
    try { sh.cancelFreeze(); }
    catch (e) { try { sh.setFrozenRows(0); sh.setFrozenColumns(0); } catch (e2) { /* ignore */ } }
  }

  async function clearAllStatSheets() {
    // 兼容保留：清掉除第一张外的所有表并把首表重置为空 Home（各功能表会在导航时重建）
    if (!S.u) return;
    const wb = S.u.api.getActiveWorkbook();
    if (!wb) return;
    const sheets = wb.getSheets();
    const keepId = sheets[0].getSheetId();
    for (const sh of sheets.slice(1)) {
      try {
        await S.u.api.executeCommand('sheet.command.remove-sheet', { unitId: wb.getId(), subUnitId: sh.getSheetId() });
      } catch (e) { /* ignore */ }
    }
    const keep = wb.getSheets().find((s) => s.getSheetId() === keepId) || wb.getSheets()[0];
    clearSheet(keep);
    try { keep.setName('Home'); } catch (e) { /* ignore */ }
    try { keep.activate(); } catch (e) { /* ignore */ }
  }

  // 销售数据变更（导入/删除/清空/还原）后统一调用：已生成报表的结果缓存清掉，
  // 并把对应的报表工作表整张移除（各功能的数据工作表在导航时总会重新渲染，无需处理）
  async function invalidateStatOutputs() {
    invalidateCtx();
  }

  // 数字格式补隐藏零段（#,##0.00 → #,##0.00;-#,##0.00;），与「零值不显示」口径一致
  function zhFmt(fmt) {
    const f = String(fmt || '');
    if (!f || f === '@' || f.includes(';')) return f;
    return f + ';-' + f + ';';
  }
  function shiftA1(rng, n) {
    if (!n) return rng;
    const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(String(rng || ''));
    if (!m) return rng;
    return m[1] + (Number(m[2]) + n) + ':' + m[3] + (Number(m[4]) + n);
  }

  let renderChain = Promise.resolve();
  function renderSpec(spec) {
    // 串行队列：桥接切换/快速切换菜单时的并发渲染排队执行，避免交叉写同一张表卡死主线程
    const p = renderChain.then(() => renderSpecInner(spec));
    renderChain = p.catch(() => { /* ignore */ });
    return p;
  }

  async function renderSpecInner(spec) {
    const mySeq = (window.__rs = (window.__rs || 0) + 1);
    S.isRendering = true;
    try {
      await renderSpecInnerBody(spec);
      try { if (window.__renderProfile) window.__renderProfile.last.__done = true; } catch (e) { /* ignore */ }
    } finally {
      S.isRendering = false;
      window.__rsDone = mySeq;
    }
  }

  // Every generated cell receives a complete baseline, including empty cells.
  // Coordinates are not record identities: old per-row formatting must not be replayed.
  function cleanGeneratedStyle(style) {
    const edge={s:1,cl:{rgb:'#b8c4d2'}};
    return Object.assign(style, Object.assign({bl:0,it:0,ul:null,st:null,
      bg:{rgb:'#ffffff'},cl:{rgb:'#222222'},n:null,ht:0,vt:2,
      bd:{t:edge,b:edge,l:edge,r:edge}},style));
  }

  // 年度固定模板：样式随数据一次写入，避免逐格百分比修复和格式命令。
  function renderAnnualTemplate(spec, sh) {
    const started = performance.now();
    const nr = spec.values.length;
    const nc = spec.values[0].length;
    const key = String(spec.sheet);
    const summaryRows = new Set([nr - 2, nr - 1, ...spec.yearGroups.map((g) => g.row - 1)]);
    const user = {};
    const baseline = {};
    const matrix = spec.values.map((row, ri) => Array.from({ length: nc }, (_, ci) => {
      const percent = ri > 0 && ci >= 5 && ci % 2 === 1;
      const style = Object.assign({ bl: ri === 0 || summaryRows.has(ri) ? 1 : 0,
        ht: ci === 0 ? 1 : 3 }, ri === 0 ? { bg: { rgb: '#1f4e78' }, cl: { rgb: '#ffffff' } } : {}, user[ri + ':' + ci] || {});
      style.ht = ri === 0 ? 2 : ci === 0 ? (summaryRows.has(ri) ? 2 : 1) : percent ? 2 : 3; // 表头/比例居中，A列左，金额右
      if (ri > 0 && ci > 0) style.n = { pattern: percent ? '0.00%;-0.00%;' : '#,##0.00;-#,##0.00;' };
      cleanGeneratedStyle(style);
      baseline[ri + ':' + ci] = style;
      const v = row[ci] === undefined || row[ci] === 0 ? '' : row[ci];
      return { v, s: style };
    }));
    sh.getRange(0, 0, nr, nc).setValues(matrix);
    const raw = S.sheetLayouts[key];
    const saved = Array.isArray(raw) ? { w: raw, h: {} } : (raw || {});
    // 年度报表固定尺寸，覆盖旧的列宽/行高记忆。
    const widths = Array.from({length: nc}, (_, ci) =>
      ci === 0 ? 120 : ci >= 5 && ci % 2 === 1 ? 100 : 150);
    widths.forEach((w, ci) => sh.setColumnWidth(ci, w));
    sh.setRowHeights(0, nr, 30);
    const heights = {};
    const fz = saved.fz || { y: 1, x: 0 };
    sh.setFrozenRows(fz.y || 0); sh.setFrozenColumns(fz.x || 0);
    applyAnnualGroups(spec, sh);
    S.appliedLayouts[key] = { w: widths, h: heights, fz };
    S.appliedStyles[key] = baseline;
    S.sheetExtent = S.sheetExtent || {};
    S.sheetExtent[key] = { rows: nr, cols: nc };
    S.lastRenderedSheet = key;
    sh.activate();
    try { window.dispatchEvent(new Event('resize')); } catch (e) { /* ignore */ }
    window.__annualTemplateProfile = { rows: nr, columns: nc,
      milliseconds: Math.round(performance.now() - started), bulkWrites: 1 };
  }

  // 销售业绩固定模板：每个格子的样式一次合并，金额与末尾比例行分别处理。
  function renderPerformanceTemplate(spec, sh) {
    const nr = spec.values.length, nc = spec.values[0].length;
    const key = String(spec.sheet);
    const isOrders = key === '客户订单统计';
    const user = {};
    const styles = Array.from({ length: nr }, () => Array.from({ length: nc }, () => ({})));
    const eachRange = (range, fn) => {
      const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(range);
      if (!m) return;
      const col = (v) => [...v].reduce((a,c) => a*26+c.charCodeAt(0)-64,0)-1;
      for (let r=+m[2]-1; r<Math.min(+m[4],nr); r++)
        for (let c=col(m[1]); c<=Math.min(col(m[3]),nc-1); c++) fn(styles[r][c]);
    };
    (spec.bold || []).forEach((range) => eachRange(range, (st) => { st.bl=1; }));
    (spec.bg || []).forEach((b) => eachRange(b.range, (st) => { st.bg={rgb:b.color}; }));
    (spec.numFmt || []).forEach((f) => eachRange(f.range, (st) => { st.n={pattern:zhFmt(f.fmt)}; }));
    const baseline = {};
    const matrix = spec.values.map((row,r) => Array.from({length:nc},(_,c) => {
      const fixed=styles[r][c];
      const st=Object.assign({bl:0},fixed,user[r+':'+c] || {});
      if(!isOrders) {
        // Row positions change with report filters. Never replay old summary
        // fills/bold/text colors onto a different salesperson at the same row.
        st.bg=fixed.bg || {rgb:'#ffffff'};
        st.bl=fixed.bl || 0;
        st.cl=fixed.cl || {rgb:'#222222'};
        const edge={s:1,cl:{rgb:'#b8c4d2'}};
        st.bd={t:edge,b:edge,l:edge,r:edge};
      }
      if (fixed.n) st.n=fixed.n;
      const percent=!!(fixed.n && fixed.n.pattern.includes('%'));
      const isHeader = r===0 || (isOrders && row[0] === 'Sales 2' && row[1] === 'Sales Order');
      st.ht=isHeader ? 2 : isOrders ? (c===0 ? 1 : [1,2,4,6,8].includes(c) ? 2 : 3) : [0,1,5].includes(c) ? 1 : percent ? 2 : 3;
      if(isHeader) st.cl={rgb:'#ffffff'};
      cleanGeneratedStyle(st);
      baseline[r+':'+c]=st;
      return {v:row[c]===0 || row[c]===undefined ? '' : row[c],s:st};
    }));
    sh.getRange(0,0,nr,nc).setValues(matrix);
    const widths=Array.from({length:nc},(_,c)=>(isOrders ? [150,100,100,180,100,150,100,150,100] : [150,150,180,100,100,150])[c] || 150);
    widths.forEach((w,c)=>sh.setColumnWidth(c,w));
    sh.showRows(0,nr);
    sh.setRowHeights(0,nr,30);
    sh.setFrozenRows(1); sh.setFrozenColumns(isOrders ? 0 : 2);
    S.appliedLayouts[key]={w:widths,h:{},fz:{y:1,x:isOrders ? 0 : 2}};
    S.appliedStyles[key]=baseline;
    S.sheetExtent=S.sheetExtent || {};
    S.sheetExtent[key]={rows:nr,cols:nc};
    S.lastRenderedSheet=key;
    sh.activate();
    try { window.dispatchEvent(new Event('resize')); } catch(e) { /* ignore */ }
  }

  // Default report conventions: identifiers/counts/ratios 100 centered,
  // customer names 300 left, money 150 right, headers centered, row height 30.
  const REPORT_DEFAULT_LAYOUT = Object.freeze({compact:100,customerName:300,money:150,rowHeight:30});
  const CUSTOMER_BATCH_IDS=['customerType','salesAccount','actDetail','actBySales1','actByCity','actByType','actSleeping','actReact','actNever','actNew'];
  const CUSTOMER_BATCH_SHEETS=['客户类型统计','客户开户统计','客户活跃统计','活跃-按销售1','活跃-按城市','活跃-按客户类型','沉睡客户','沉睡客户唤醒统计','从未购买','新开户'];
  async function customerBatchContext(opts, knownRevision) {
    const key=JSON.stringify([opts.start || '',opts.end || '']);
    const cached=S.customerBatchSource && S.customerBatchSource.key===key ? S.customerBatchSource.ctx : null;
    const result=await normalizeCtx(await api('get_customer_reports_context',opts.start || '',opts.end || '',knownRevision || (cached && cached.revision) || ''));
    if(result.unchanged) {
      if(knownRevision && result.revision===knownRevision) return result;
      if(cached && cached.revision===result.revision) return cached;
      throw new Error('Customer source cache is unavailable');
    }
    S.customerBatchSource={key,ctx:result};
    return result;
  }
  function renderCustomerBatchTemplate(spec,sh,values,off) {
    const key=String(spec.sheet),nr=values.length,nc=spec.values[0].length;
    const grouped=key.startsWith('活跃-'),type=key==='客户类型统计',account=key==='客户开户统计';
    const sleeping=key==='沉睡客户',react=key==='沉睡客户唤醒统计',never=key==='从未购买',fresh=key==='新开户';
    const widths=Array.from({length:nc},(_,c)=>type ? (c===0?250:200) : account ? (c===0?200:100) : grouped ? (c===0?200:c>=9?150:100) : never ? (c===0?100:c===1?300:120) : react ? (c===1?250:c===8?150:100) : sleeping ? (c===1?250:[2,3,5].includes(c)?150:100) : fresh ? (c===1?300:c===5?150:100) : (c===1?300:[2,5].includes(c)?180:c===0?100:150));
    const left=type || account || grouped ? [0] : sleeping ? [1,2,3] : [1];
    const right=type ? [1] : grouped ? [9,10] : account || never ? [] : react ? [8] : [5];
    // Version the defaults once; subsequent manual edits remain editable and remembered.
    S.customerBatchLayouts=S.customerBatchLayouts || {};
    if(!S.customerBatchLayouts[key]) {
      delete S.userStyles[key];delete S.sheetLayouts[key];S.customerBatchLayouts[key]=1;
      trackSave(api('set_pref','customerBatchLayouts',JSON.stringify(S.customerBatchLayouts)));
      trackSave(api('set_pref','userStyles',JSON.stringify(S.userStyles)));
      trackSave(api('set_pref','sheetLayouts',JSON.stringify(S.sheetLayouts)));
    }
    const user={},lay=S.sheetLayouts[key] || {},baseline={},formats=[];
    const col=v=>[...v].reduce((a,c)=>a*26+c.charCodeAt(0)-64,0)-1;
    (spec.numFmt || []).forEach(f=>{
      const m=/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(f.range);
      if(m) formats.push({r0:+m[2]-1+off,r1:+m[4]-1+off,c0:col(m[1]),c1:col(m[3]),pattern:zhFmt(f.fmt.includes('%')?'0.00%':f.fmt)});
    });
    const matrix=values.map((row,r)=>Array.from({length:nc},(_,c)=>{
      const header=r===off,total=/^(Grand Total|Total|合计)$/.test(String(row[0] || ''));
      const st={bl:r===0 || header || total?1:0,vt:2,ht:r<off?1:header?2:left.includes(c)?1:right.includes(c)?3:2};
      if(header || total) st.bg={rgb:'#d9eaf7'};
      const f=formats.find(f=>r>=f.r0 && r<=f.r1 && c>=f.c0 && c<=f.c1);if(f) st.n={pattern:f.pattern};
      Object.assign(st,user[r+':'+c] || {});
      cleanGeneratedStyle(st);
      if(r<300) baseline[r+':'+c]=st;
      return {v:row[c]===0 || row[c]===undefined?'':row[c],s:st};
    }));
    sh.getMergeData().forEach(r=>r.breakApart());
    sh.getRange(0,0,nr,nc).setValues(matrix);
    if(off) {sh.getRange(0,0,1,nc).merge();sh.getRange(1,0,1,nc).merge();}
    widths.forEach((w,c)=>{if(lay.w && lay.w[c]>0) widths[c]=lay.w[c];sh.setColumnWidth(c,widths[c]);});
    sh.showRows(0,nr);sh.setRowHeights(0,nr,30);
    const heights={};Object.entries(lay.h || {}).forEach(([r,h])=>{if(+r<nr && h>=10){sh.setRowHeight(+r,h);heights[r]=h;}});
    const fz=lay.fz || {y:off+1,x:grouped?1:0};sh.setFrozenRows(fz.y || 0);sh.setFrozenColumns(fz.x || 0);
    S.appliedLayouts[key]={w:widths,h:heights,fz};S.appliedStyles[key]=baseline;
    S.sheetExtent=S.sheetExtent || {};S.sheetExtent[key]={rows:nr,cols:nc};S.lastRenderedSheet=key;
    sh.activate();try{window.dispatchEvent(new Event('resize'));}catch(e){}
  }

  function renderSalesQueryTemplate(spec,sh) {
    const key=String(spec.sheet),nr=spec.values.length,nc=spec.values[0].length,baseline={};
    const widths=[100,100,300,120,120,120,120,100,300,100,200,100,200,100,150,150];
    const matrix=spec.values.map((row,r)=>Array.from({length:nc},(_,c)=>{
      const st={bl:r===0?1:0,vt:2,ht:r===0?2:[2,3,4,8,10,12].includes(c)?1:c>=14?3:2};
      if(r===0) st.bg={rgb:'#d9eaf7'};
      if(r>0 && c>=13) st.n={pattern:zhFmt(c===13?'#,##0':'#,##0.00')};
      cleanGeneratedStyle(st);
      if(r<300) baseline[r+':'+c]=st;
      return {v:row[c]===undefined || row[c]===null?'':row[c],s:st};
    }));
    sh.getMergeData().forEach(r=>r.breakApart());
    sh.getRange(0,0,nr,nc).setValues(matrix);
    widths.forEach((w,c)=>sh.setColumnWidth(c,w));
    sh.showRows(0,nr);sh.setRowHeights(0,nr,30);
    sh.setFrozenRows(1);sh.setFrozenColumns(0);
    S.appliedLayouts[key]={w:widths,h:{},fz:{y:1,x:0}};S.appliedStyles[key]=baseline;
    S.sheetExtent=S.sheetExtent || {};S.sheetExtent[key]={rows:nr,cols:nc};
    S.lastRenderedSheet=key;sh.activate();
    try{window.dispatchEvent(new Event('resize'));}catch(e){}
  }

  async function productContext(opts,known) {
    const key=JSON.stringify([opts.start || '',opts.end || '']);
    const cache=S.productSource && S.productSource.key===key ? S.productSource.ctx : null;
    const ctx=await normalizeCtx(await api('get_product_context',opts.start || '',opts.end || '',known || (cache && cache.revision) || ''));
    if(ctx.unchanged) {if(known) return ctx;if(cache && cache.revision===ctx.revision)return cache;throw new Error('Product cache unavailable');}
    S.productSource={key,ctx};return ctx;
  }
  function stoneTypeOptions(types) {
    const opts=[...new Set((types || []).map(v=>String(v || '').trim().toUpperCase()).filter(v=>v && !STATS.helpers.isCabinetType(v)))].sort();
    if(!opts.includes('SLAB'))opts.unshift('SLAB');
    return opts;
  }
  function renderProductCustomerTemplate(spec,sh,values,off) {
    const customer=!!spec.customerTemplate,cab=spec.sheet==='橱柜颜色统计';
    const key=String(spec.sheet),nr=values.length,nc=spec.values[0].length,baseline={};
    const widths=customer?[100,300,100,150,150,100,150,250,250,400]:cab?[200,120,120,180,120]:[100,100,300,100,100,180,100,100];
    const left=customer?[1,3,4,7,8,9]:cab?[0]:[2],right=customer?[]:cab?[3]:[5];
    const formats=[];const col=v=>[...v].reduce((a,c)=>a*26+c.charCodeAt(0)-64,0)-1;
    (spec.numFmt || []).forEach(f=>{const m=/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(f.range);if(m)formats.push({r0:+m[2]-1+off,r1:+m[4]-1+off,c0:col(m[1]),c1:col(m[3]),pattern:zhFmt(f.fmt)});});
    const matrix=values.map((row,r)=>Array.from({length:nc},(_,c)=>{
      const header=r===off,total=!customer && r===nr-1;
      const st={bl:r===0 || header || total?1:0,vt:2,ht:r<off?1:header?2:left.includes(c)?1:right.includes(c)?3:2};
      if(total) st.bg={rgb:'#d9bd8b'};else if(header && !cab)st.bg={rgb:'#d9e6f3'};
      const f=formats.find(f=>r>=f.r0 && r<=f.r1 && c>=f.c0 && c<=f.c1);if(f)st.n={pattern:f.pattern};
      cleanGeneratedStyle(st);
      if(r<300)baseline[r+':'+c]=st;
      return {v:row[c]===undefined || row[c]===null?'':row[c],s:st};
    }));
    sh.getMergeData().forEach(r=>r.breakApart());sh.getRange(0,0,nr,nc).setValues(matrix);
    if(off){sh.getRange(0,0,1,nc).merge();sh.getRange(1,0,1,nc).merge();}
    widths.forEach((w,c)=>sh.setColumnWidth(c,w));sh.showRows(0,nr);sh.setRowHeights(0,nr,30);
    sh.setFrozenRows(off+1);sh.setFrozenColumns(0);
    S.appliedLayouts[key]={w:widths,h:{},fz:{y:off+1,x:0}};S.appliedStyles[key]=baseline;
    S.sheetExtent=S.sheetExtent || {};S.sheetExtent[key]={rows:nr,cols:nc};S.lastRenderedSheet=key;sh.activate();
    try{window.dispatchEvent(new Event('resize'));}catch(e){}
  }

  function renderCustomerStatsTemplate(spec, sh) {
    const key=String(spec.sheet),nr=spec.values.length,nc=spec.values[0].length;
    const money=[4,7,12],pct=[5,8,10,13,15],baseline={};
    const widths=Array.from({length:nc},(_,c)=>c===1 ? REPORT_DEFAULT_LAYOUT.customerName : money.includes(c) ? REPORT_DEFAULT_LAYOUT.money : REPORT_DEFAULT_LAYOUT.compact);
    const matrix=spec.values.map((row,r)=>Array.from({length:nc},(_,c)=>{
      const header=r<2,total=r===2;
      const st={bl:header || total ? 1:0,vt:2,ht:header ? 2 : c===1 ? 1 : money.includes(c) ? 3 : 2};
      if(header || total) st.bg={rgb:header ? '#d9e6f3':'#d9bd8b'};
      if(!header && c>=4) st.n={pattern:zhFmt(money.includes(c) ? '#,##0.00' : pct.includes(c) ? '0.00%' : '#,##0')};
      cleanGeneratedStyle(st);
      if(r<300) baseline[r+':'+c]=st;
      return {v:row[c]===0 || row[c]===undefined ? '' : row[c],s:st};
    }));
    sh.getMergeData().forEach(r=>r.breakApart());
    sh.getRange(0,0,nr,nc).setValues(matrix);
    (spec.merges || []).forEach(r=>sh.getRange(r).merge());
    widths.forEach((w,c)=>sh.setColumnWidth(c,w));
    sh.showRows(0,nr);sh.setRowHeights(0,nr,REPORT_DEFAULT_LAYOUT.rowHeight);
    sh.setFrozenRows(3);sh.setFrozenColumns(2);
    S.appliedLayouts[key]={w:widths,h:{},fz:{y:3,x:2}};
    S.appliedStyles[key]=baseline;
    S.sheetExtent=S.sheetExtent || {};S.sheetExtent[key]={rows:nr,cols:nc};
    S.lastRenderedSheet=key;sh.activate();
    try {window.dispatchEvent(new Event('resize'));} catch(e) { /* ignore */ }
  }

  function renderHistoryTemplate(spec, sh, values) {
    const key=String(spec.sheet), nr=values.length, nc=spec.values[0].length;
    // Infer old saved report layouts as well as newly generated ones.
    const view=spec.historyView || (nc===7 ? 'customer' : nc===8 ? 'sales' : nc===9 ? 'invoice' : spec.values[0][0]==='Item No.' ? 'item' : 'all');
    const layouts={
      all:{w:[100,100,100,100,120,100,250,100,100,120],left:[6],right:[4,9]},
      customer:{w:[100,100,100,100,120,100,120],left:[],right:[4,6]},
      item:{w:[100,300,100,100,100,120,100,300,100,120],left:[1,7],right:[5,9]},
      invoice:{w:[100,100,300,100,100,100,100,120,120],left:[2],right:[7,8]},
      sales:{w:[100,120,100,100,120,100,300,120],left:[6],right:[4,7]}
    };
    const layout=layouts[view], baseline={}, formats=[];
    const col=(v)=>[...v].reduce((a,c)=>a*26+c.charCodeAt(0)-64,0)-1;
    (spec.numFmt || []).forEach(f=>{
      const m=/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(f.range);
      if(m) formats.push({r0:+m[2]+1,r1:+m[4]+1,c0:col(m[1]),c1:col(m[3]),pattern:zhFmt(f.fmt)});
    });
    const matrix=values.map((row,r)=>Array.from({length:nc},(_,c)=>{
      const header=r===2, total=r===nr-1;
      const st={bl:r===0 || header || total ? 1:0,vt:2,
        ht:r<2 ? 1 : header ? 2 : layout.left.includes(c) ? 1 : layout.right.includes(c) ? 3 : 2};
      if(header || total) st.bg={rgb:header ? '#d9eaf7':'#d9e6f3'};
      const f=formats.find(f=>r>=f.r0 && r<=f.r1 && c>=f.c0 && c<=f.c1);
      if(f) st.n={pattern:f.pattern};
      cleanGeneratedStyle(st);
      if(r<300) baseline[r+':'+c]=st;
      return {v:row[c]===0 || row[c]===undefined ? '' : (r===1 ? String(row[c]).replace(/\n/g,' | ') : row[c]),s:st};
    }));
    sh.getMergeData().forEach(r=>r.breakApart());
    sh.getRange(0,0,nr,nc).setValues(matrix);
    sh.getRange(0,0,1,nc).merge();sh.getRange(1,0,1,nc).merge();
    layout.w.forEach((w,c)=>sh.setColumnWidth(c,w));
    sh.showRows(0,nr);sh.setRowHeights(0,nr,30);
    sh.setFrozenRows(3);sh.setFrozenColumns(0);
    S.appliedLayouts[key]={w:layout.w,h:{},fz:{y:3,x:0}};
    S.appliedStyles[key]=baseline;
    S.sheetExtent=S.sheetExtent || {};S.sheetExtent[key]={rows:nr,cols:nc};
    S.lastRenderedSheet=key;sh.activate();
    try {window.dispatchEvent(new Event('resize'));} catch(e) { /* ignore */ }
  }

  function renderCategoryTemplate(spec, sh) {
    const nr=spec.values.length, nc=spec.values[0].length, key=String(spec.sheet);
    const user={}, baseline={};
    const qty=(spec.numFmt[0] || {}).fmt === '#,##0';
    const pattern=zhFmt(qty ? '#,##0' : '#,##0.00');
    const matrix=spec.values.map((row,r)=>Array.from({length:nc},(_,c)=>{
      const header=r<2, total=r===nr-1;
      const st=Object.assign({bl:header || total ? 1 : 0},
        header ? {bg:{rgb:'#D9E8F5'}} : total ? {bg:{rgb:'#d9e6f3'}} : {},user[r+':'+c] || {});
      st.ht=header ? 2 : c<3 ? 1 : qty ? 2 : 3;
      st.vt=2;
      if(!header && c>=3) st.n={pattern};
      cleanGeneratedStyle(st);
      if(r<300 && c<60) baseline[r+':'+c]=st;
      return {v:row[c]===0 || row[c]===undefined ? '' : row[c],s:st};
    }));
    // 年度/月数变化时先解除旧合并，避免原表头范围残留。
    sh.getMergeData().forEach((r)=>r.breakApart());
    sh.getRange(0,0,nr,nc).setValues(matrix);
    (spec.merges || []).forEach((r)=>sh.getRange(r).merge());
    const monthStart=spec.values[1].findIndex((v)=>v!=='' && v!=null);
    const widths=Array.from({length:nc},(_,c)=>c===0 || c===1 ? 100 : c===2 ? 200 : c<monthStart ? 120 : 100);
    widths.forEach((w,c)=>sh.setColumnWidth(c,w));
    sh.showRows(0,nr);sh.setRowHeights(0,nr,30);
    sh.setFrozenRows(2);sh.setFrozenColumns(3);
    S.appliedLayouts[key]={w:widths,h:{},fz:{y:2,x:3}};
    S.appliedStyles[key]=baseline;
    S.sheetExtent=S.sheetExtent || {};S.sheetExtent[key]={rows:nr,cols:nc};
    S.lastRenderedSheet=key;sh.activate();
    try {window.dispatchEvent(new Event('resize'));} catch(e) { /* ignore */ }
  }

  async function renderSpecInnerBody(spec) {
    hideCostColumn(spec);
    placeMount();
    // 每个功能拥有自己的工作表：按 spec.sheet 找表（无则新建），原地清空重写，不影响其他功能的表
    const wb = S.u.api.getActiveWorkbook();
    if (!wb) return;
    const name = String(spec.sheet || 'Home');
    const ncPreview = (sp) => Math.max(...((sp && sp.values) || [['']]).map((r) => r.length), 1);
    const findSheet = () => {
      try { const hit = wb.getSheetByName(name); if (hit) return hit; } catch (e) { /* 老 facade 无此方法 */ }
      return wb.getSheets().find((s) => { try { return s.getSheetName() === name; } catch (e2) { return false; } }) || null;
    };
    const need = (spec.values || []).length + 2;
    let freshSheet = false;
    let sh = findSheet();
    // 行容量不足或严重过剩时删旧重建（快照建表毫秒级；而大表整表 clearContent+clearFormat 是
    // 分钟级慢路径——如销售记录从"全部"(5.3万行)切回 100，复用旧表会整表清空卡死）。
    // 建表走「裸 {id,name} 的 insert-sheet」：历史验证 numfmt 等命令健康
    // （带自定义 snapshot 的 insert-sheet / wb.create 均会导致 "numfmt.set is not registered"，大范围卡死）。
    if (sh && (sh.getMaxRows() < need || sh.getMaxRows() > need + 1000)) {
      try {
        // 不能删当前激活表（命令会被静默拒绝）——先切到另一张表再删
        const other = wb.getSheets().find((x) => { try { return x.getSheetId() !== sh.getSheetId(); } catch (e1) { return false; } });
        if (other) { try { other.activate(); } catch (e0) { /* ignore */ } }
        await S.u.api.executeCommand('sheet.command.remove-sheet', { unitId: wb.getId(), subUnitId: sh.getSheetId() });
      } catch (e) { /* ignore */ }
      sh = null;
    }
    if (!sh) {
      const newId = 'sh_' + Date.now() + '_' + Math.floor(Math.random() * 10000);
      try {
        // 快照建表带 rowCount/columnCount：数万行表 50ms 级建好（insertRows 扩行是分钟级，已弃用）
        await S.u.api.executeCommand('sheet.command.insert-sheet', {
          unitId: wb.getId(),
          sheet: { id: newId, name,
                   rowCount: Math.max(need + 50, 1000), columnCount: Math.max(ncPreview(spec) + 4, 26) },
        });
      } catch (e) { /* ignore */ }
      // 删旧建同名的场景 Univer 会自动改名（"xxx (1)"），按名字找会抓错表——按我们指定的 id 找回
      sh = wb.getSheets().find((x) => { try { return x.getSheetId() === newId; } catch (e1) { return false; } })
        || findSheet() || wb.getSheets()[0];
      freshSheet = true;   // 新建空表：跳过 clearSheet（大范围 clear 是隐藏的极慢路径）
      try { sh.activate(); } catch (e2) { /* ignore */ }
      // 旧表若因「是激活表且无其他表可激活」删除被拒（如全新工作簿只有一张表），现在新表已激活，补删同名残表
      const leftover = wb.getSheets().find((x) => {
        try { return x.getSheetId() !== sh.getSheetId() && x.getSheetName() === name; } catch (e3) { return false; }
      });
      if (leftover) {
        try { await S.u.api.executeCommand('sheet.command.remove-sheet', { unitId: wb.getId(), subUnitId: leftover.getSheetId() }); } catch (e4) { /* ignore */ }
      }
    }
    try { sh.setName(name); } catch (e) { /* ignore */ }
    // 清理初始空表（Sheet1/Home）：已有其他表且它仍为空时移除。
    // 若空表正是当前激活表，先把目标表激活再删，否则视图会停留在被删表上变成空画布。
    try {
      const all = wb.getSheets();
      if (all.length > 1) {
        const blank = all.find((s) => {
          try { return s !== sh && /^(Sheet1|Home)$/.test(s.getSheetName()) && !s.getRange(0, 0, 1, 1).getValue(); } catch (e2) { return false; }
        });
        if (blank) {
          try { sh.activate(); } catch (e3) { /* ignore */ }
          await S.u.api.executeCommand('sheet.command.remove-sheet', { unitId: wb.getId(), subUnitId: blank.getSheetId() });
        }
      }
    } catch (e) { /* ignore */ }
    // 渲染前只同步"上次渲染的表"（用户刚编辑过的那张；含目标表自身=重新生成场景），
    // 避免对全簿所有表做布局/样式快照导致切换卡顿。
    try {
      const wb0 = S.u.api.getActiveWorkbook();
      if (S.lastRenderedSheet) {
        const prev = wb0.getSheets().find((x) => { try { return x.getSheetName() === S.lastRenderedSheet; } catch (e) { return false; } });
        if (prev) syncSheetMemory(prev);
      }
    } catch (e) { /* ignore */ }
    // 写入范围先算（clearSheet 尾巴清理要用；勿在使用后声明——TDZ）
    let values = spec.values;
    const off = spec.info ? 2 : 0;
    if (off) {
      // 报表头两行：标题（表名）+ 口径说明，其下才是表格；格式区域整体下移
      const width = Math.max(...spec.values.map((r) => r.length));
      const title = new Array(width).fill(''); title[0] = spec.sheet;
      const info = new Array(width).fill(''); info[0] = spec.info;
      values = [title, info].concat(spec.values);
    }
    const nr = values.length;
    const nc = Math.max(...values.map((r) => r.length));
    if (!freshSheet) clearSheet(sh, { keepRows: nr, keepCols: nc });   // 复用表只清尾巴（整表 clear 是分钟级慢路径）
    // 防御性小步扩容（建表已按需给行数；此分支仅覆盖建表 fallback 路径的小缺口）
    try {
      if (sh.getMaxRows() < need) sh.insertRows(sh.getMaxRows(), need - sh.getMaxRows());
    } catch (e) { /* ignore */ }
    if(spec.customerTemplate || ['橱柜颜色统计','石材销量统计'].includes(spec.sheet)){renderProductCustomerTemplate(spec,sh,values,off);return;}
    if(spec.salesQueryTemplate) {renderSalesQueryTemplate(spec,sh);return;}
    if (CUSTOMER_BATCH_SHEETS.includes(spec.sheet)) {
      renderCustomerBatchTemplate(spec,sh,values,off);return;
    }
    if (spec.sheet === '客户统计' && !off) {
      renderCustomerStatsTemplate(spec, sh);
      return;
    }
    if (spec.sheet === '销售历史查询' && off===2) {
      renderHistoryTemplate(spec, sh, values);
      return;
    }
    if (spec.sheet === '销售查询' && !off) {
      renderCategoryTemplate(spec, sh);
      return;
    }
    if (['销售业绩', '客户订单统计'].includes(spec.sheet) && !off) {
      renderPerformanceTemplate(spec, sh);
      return;
    }
    if (spec.yearGroups && !off) {
      renderAnnualTemplate(spec, sh);
      return;
    }
    // 数字格式随值一次写入（{v, s:{n:{pattern}}}）：绕开 numfmt 命令、万行级 250ms 内完成；
    // 同时 0 值渲染为空单元格（客户口径）。zhFmt 已补隐藏零段。
    const colIdxOf = (letters) => letters.split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1;
    const fmtRangeCols = {};
    (spec.numFmt || []).forEach((f) => {
      const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(String(f.range || ''));
      if (!m) return;
      const pat = zhFmt(f.fmt);
      if (!pat) return;
      const c1 = colIdxOf(m[1]);
      const r1 = Number(m[2]);
      const c2 = colIdxOf(m[3]);
      const r2 = Number(m[4]);
      for (let c = c1; c <= c2; c++) fmtRangeCols[c] = { startRow0: r1 - 1, pat };
    });
    const matrix = values.map((r,ri)=>Array.from({length:nc},(_,ci)=>{
      const v=r[ci],fr=fmtRangeCols[ci];
      const st=cleanGeneratedStyle({});
      if(fr && ri>=fr.startRow0+off && typeof v==='number')st.n={pattern:fr.pat};
      return {v:v===0 || v==null?'':v,s:st};
    }));
    const __pf = (window.__renderProfile = window.__renderProfile || { last: {} });
    const __t = (k) => { __pf.last[k] = Math.round(performance.now() - __pf._t0); __pf._t0 = performance.now(); };
    __pf._t0 = performance.now(); __pf.last = {};
    __t('expandRows');
    sh.getRange(0, 0, nr, nc).setValues(matrix);
    __t('setValues');
    __t('numFmt');   // 数字格式已随 setValues 写入（见矩阵构造），不再走命令
    (spec.bold || []).forEach((r) => { try { sh.getRange(shiftA1(r, off)).setFontWeight('bold'); } catch (e) { /* ignore */ } });
    (spec.bg || []).forEach((b) => { try { sh.getRange(shiftA1(b.range, off)).setBackground(b.color); } catch (e) { /* ignore */ } });
    (spec.align || []).forEach((a) => {
      try {
        if (a.horiz === 'left') {
          const m2 = /^(?:[A-Z]+)(\d+):(?:[A-Z]+)(\d+)$/.exec(String(a.range));
          if (m2 && (Number(m2[2]) - 1) > 3000) return;   // 3千行以上的 left 对齐=文本默认对齐，跳过大范围样式写入（约省1.5s）
        }
        sh.getRange(shiftA1(a.range, off)).setHorizontalAlignment(a.horiz);
      } catch (e) { /* ignore */ }
    });
    (spec.merges || []).forEach((m) => { try { sh.getRange(m).merge(); } catch (e) { /* ignore */ } });
    // 自适应列宽：按内容计算（CJK 记 2 个单位宽，数字格式预留 25% 余量），钳制在 56~420px
    const unitW = (v) => { let w = 0; for (const ch of String(v === null || v === undefined ? '' : v)) w += ch.charCodeAt(0) > 255 ? 2 : 1; return w; };
    const colUnits = new Array(nc).fill(0);
    // 大表采样：只统计前 500 行内容宽（5 万行逐格计算会显著拖慢渲染）
    const widthSample = spec.values.slice(0, 500);
    for (const row of widthSample) for (let c = 0; c < row.length; c++) {
      const u = unitW(row[c]);
      if (u > colUnits[c]) colUnits[c] = u;
    }
    const finalWidths = [];
    for (let c = 0; c < nc; c++) {
      const w = Math.max(56, Math.min(420, Math.round(colUnits[c] * 8.4 * 1.25) + 26));
      try { sh.setColumnWidth(c, w); } catch (e) { /* ignore */ }
      finalWidths.push(w);
    }
    // 应用记忆布局（列宽+行高）：该表被用户调整过就优先用用户的设定。兼容旧数组格式（仅列宽）。
    const layKey = String(spec.sheet || '');
    const savedRaw = S.sheetLayouts[layKey];
    const savedLay = Array.isArray(savedRaw) ? { w: savedRaw, h: {} } : (savedRaw || null);
    if (savedLay && savedLay.w && savedLay.w.length) {
      const lim = Math.min(savedLay.w.length, sh.getMaxColumns());
      for (let c = 0; c < lim; c++) {
        if (typeof savedLay.w[c] === 'number' && savedLay.w[c] >= 40) {
          try { sh.setColumnWidth(c, savedLay.w[c]); } catch (e) { /* ignore */ }
          if (c < finalWidths.length) finalWidths[c] = savedLay.w[c];
        }
      }
    }
    __t('stylesBoldBgAlign');   // 粗体/背景/对齐/合并/列宽合计
    // 自适应行高：批量设置（标题行 30、口径行 24、数据行 26）
    try { sh.setRowHeights(0, 1, 30); } catch (e) { /* ignore */ }
    if (off) { try { sh.setRowHeights(1, 1, 24); } catch (e) { /* ignore */ } }
    try { sh.setRowHeights(off, Math.max(0, nr - off), 26); } catch (e) { /* ignore */ }
    __t('rowHeights');
    if (off) { try { sh.getRange(0, 0, 1, nc).setFontWeight('bold'); } catch (e) { /* ignore */ } }
    // 应用记忆行高（用户调整过的行）
    const appliedH = {};
    const maxR = sh.getMaxRows();
    if (savedLay && savedLay.h) {
      Object.entries(savedLay.h).forEach(([ri, h]) => {
        const idx = Number(ri);
        // 注意：不能用 nr（占位表 nr=1 会跳过），按 sheet 实际行数应用
        if (idx < maxR && typeof h === 'number' && h >= 10) {
          try { sh.setRowHeight(idx, h); } catch (e) { /* ignore */ }
          appliedH[idx] = h;
        }
      });
    }
    if (spec.freeze) {
      try { sh.setFrozenRows((spec.freeze.rows || 0) + off); } catch (e) { /* ignore */ }
      try { sh.setFrozenColumns(spec.freeze.cols || 0); } catch (e) { /* ignore */ }
    }
    // 应用记忆冻结窗格（用户调整过就覆盖报表默认）
    let appliedFz = spec.freeze ? { y: (spec.freeze.rows || 0) + off, x: spec.freeze.cols || 0 } : null;
    if (savedLay && savedLay.fz && (savedLay.fz.y || savedLay.fz.x)) {
      try { sh.setFrozenRows(savedLay.fz.y || 0); } catch (e) { /* ignore */ }
      try { sh.setFrozenColumns(savedLay.fz.x || 0); } catch (e) { /* ignore */ }
      appliedFz = { y: savedLay.fz.y || 0, x: savedLay.fz.x || 0 };
    }
    // 布局基线 = 本次实际应用的列宽/行高/冻结
    S.appliedLayouts[layKey] = { w: finalWidths, h: appliedH, fz: appliedFz };
    // 回放用户修饰（只写样式，不动数值），随后记录样式基线（供定时差分捕获后续修饰）
    const userSt = {}; // Do not restore formatting by stale row coordinates.
    Object.entries(userSt).forEach(([k, styleObj]) => {
      const parts = k.split(':');
      const ri = Number(parts[0]), ci = Number(parts[1]);
      // 年度百分比是报表固定格式，不能被旧的用户样式（常规/小数）覆盖。
      const annualPercent = spec.yearGroups && ri >= 1 && ci >= 5 && ci % 2 === 1;
      const applied = annualPercent ? Object.assign({}, styleObj, { n: { pattern: '0.00%;-0.00%;' } }) : styleObj;
      try { sh.getRange(ri, ci, 1, 1).setValues([[{ s: applied }]]); } catch (e) { /* ignore */ }
    });
    if (spec.yearGroups) {
      const styles = captureCellStyles(sh);
      for (let ri = 1; ri < spec.values.length; ri++) {
        for (let ci = 5; ci < spec.values[ri].length; ci += 2) {
          const v = spec.values[ri][ci];
          if (typeof v !== 'number') continue;
          const style = Object.assign({}, styles[ri + ':' + ci] || {}, { n: { pattern: '0.00%;-0.00%;' } });
          sh.getRange(ri, ci, 1, 1).setValues([[{ v: v === 0 ? '' : v, s: style }]]);
        }
      }
    }
    S.appliedStyles[layKey] = captureCellStyles(sh);
    __t('replayUserFmt');   // 用户样式回放+基线+冻结合计
    // 激活目标表并强制视图重绑：页签栏已隐藏，切换渲染面只能靠程序保证（activate + resize）
    if (spec.yearGroups) applyAnnualGroups(spec, sh);
    S.lastRenderedSheet = name;
    try { sh.activate(); } catch (e) { /* ignore */ }
    __t('activate');
    S.sheetExtent = S.sheetExtent || {};
    S.sheetExtent[String(spec.sheet || 'Home')] = { rows: nr, cols: nc };   // 供下次渲染做尾巴清理
    try {
      const um = document.getElementById('univerMount');
      if (um) um.dispatchEvent(new Event('resize'));
      window.dispatchEvent(new Event('resize'));
    } catch (e) { /* ignore */ }
    __t('resizeDispatch');
  }

  // ================= statistics（列表式结果，不再依赖 Univer 工作表） =================
  // 三大板块与原 Apps Script Sidebar 对应：
  // 销售统计=销售业绩/客户订单；销售查询=分类统计/销售历史查询；
  // 客户统计=客户销售/客户类型/客户开户/客户活跃。actions 即原 Sidebar 里的多个生成按钮。
  const STAT_HUBS = [
    { id: 'g-sales', labelKey: 'navSalesHub', icon: '🧾', reports: [
      { id: 'salesPerformance', labelKey: 'statSales', fields: ['range'] },
      { id: 'customerOrders', labelKey: 'statOrders', fields: ['range'] },
    ] },
    { id: 'g-query', labelKey: 'navQueryHub', icon: '🔎', reports: [
      { id: 'categoryQuery', labelKey: 'statCategory', fields: ['range', 'category'], actions: [
        { labelKey: 'genAmount', opts: { mode: 'amount' } },
        { labelKey: 'genQty', opts: { mode: 'qty' } },
      ] },
      { id: 'salesHistory', labelKey: 'statSalesHistory', fields: ['range', 'salesHistory'], actions: [
        { labelKey: 'btnQuery', opts: {} },
      ] },
    ] },
    { id: 'g-cust', labelKey: 'navCustHub', icon: '👥', reports: [
      { id: 'customerStats', labelKey: 'statCustomers', fields: ['range'], actions: [
        { labelKey: 'genAll', opts: { scope: 'all' } },
        { labelKey: 'genCabinet', opts: { scope: 'cabinet' } },
        { labelKey: 'genStone', opts: { scope: 'stone' } },
      ] },
      { id: 'customerType', labelKey: 'statCustType', fields: ['range'] },
      { id: 'salesAccount', labelKey: 'statSalesAccount', fields: ['range'] },
      // 客户活跃统计：一个维度一张独立 Excel 表（完全隔离，便于给客户配不同样式）；
      // 不受起始时间限制（始终统计全部客户），只提供 客户/销售/城市 模糊查询
      { id: 'actDetail', labelKey: 'actDetail', fields: ['actf'] },
      { id: 'actBySales1', labelKey: 'actSales1', fields: ['actf'] },
      { id: 'actByCity', labelKey: 'actCity', fields: ['actf'] },
      { id: 'actByType', labelKey: 'actCustType', fields: ['actf'] },
      { id: 'actSleeping', labelKey: 'actSleeping', fields: ['actf'] },
      // 沉睡客户唤醒统计（actReact）已恢复显示
      { id: 'actReact', labelKey: 'actReact', fields: [] },
      { id: 'actNever', labelKey: 'actNever', fields: ['actf'] },
      { id: 'actNew', labelKey: 'actNew', fields: ['actf'] },
    ] },
    { id: 'g-stats', labelKey: 'navStats', icon: '📊', reports: [
      { id: 'annual', labelKey: 'statAnnual', fields: [] },   // 年度统计固定全量（无时间段）
      { id: 'cabinetColors', labelKey: 'statColors', fields: ['range'], actions: [
        { labelKey: 'genNoUni', opts: { includeUniversal: false } },
        { labelKey: 'genWithUni', opts: { includeUniversal: true } },
      ] },
      { id: 'stoneSales', labelKey: 'statStone', fields: ['range', 'type'] },
    ] },
  ];
  const STAT_DEFS = [];
  STAT_HUBS.forEach((h) => h.reports.forEach((r) => STAT_DEFS.push(r)));

  function distinctTypes(sales) {
    const s = new Set();
    (sales || []).forEach((r) => { const v = String(r.type || '').trim().toUpperCase(); if (v) s.add(v); });
    return Array.from(s).sort();
  }

  async function statOptionFields(def) {
    const fields = [];
    if (def.fields.includes('range')) {
      // 销售历史查询默认全部时间（Start/End 留空）；其他报表默认 TTM（滚动12个月）
      const allTime = def.fields.includes('salesHistory');
      const r = allTime ? { start: '', end: '' } : ttmRange();
      fields.push({ key: 'start', label: t('dlgStart'), type: 'date', value: r.start });
      fields.push({ key: 'end', label: t('dlgEnd'), type: 'date', value: r.end });
    }
    if (def.fields.includes('salesHistory')) {
      // 销售历史查询：三个精确条件（模糊下拉选择后回填编码），至少填一个
      fields.push({ key: 'custNo', label: t('custNo'), type: 'combo', fuzzy: 'customers', value: '' });
      fields.push({ key: 'itemNo', label: t('colItemNo'), type: 'combo', fuzzy: 'items', value: '' });
      fields.push({ key: 'invoice', label: t('colInvoice'), type: 'combo', fuzzy: 'invoices', value: '' });
      fields.push({ key: 'sales', label: t('colRep1'), type: 'combo', fuzzy: 'rep1', value: '' });
    }
    if (def.fields.includes('actf')) {
      // 客户活跃统计系列：客户/销售/城市模糊查询（可留空=全部）；无时间段（始终统计全部客户）
      if(!['actBySales1','actByCity','actByType'].includes(def.id)) fields.push({ key: 'fCust', label: t('custNo'), type: 'combo', fuzzy: 'customers', value: '' });
      fields.push({ key: 'fSales', label: t('colRep1'), type: 'combo', fuzzy: 'rep1', value: '' });
      fields.push({ key: 'fCity', label: t('dlgCity'), type: 'combo', fuzzy: 'cities', value: '' });
    }
    // 下拉类别/型号用轻量聚合接口（SQL DISTINCT），不再为下拉拉全表明细
    const ensureTypes = async () => {
      try { return (await api('get_lookup_data')).types || []; } catch (e) { return []; }
    };
    if (def.fields.includes('category')) {
      const opts = ['ALL', ...(await ensureTypes())];
      fields.push({ key: 'category', label: t('dlgCategory'), type: 'combo', options: opts, value: 'ALL' });
    }
    if (def.fields.includes('type')) {
      const opts=stoneTypeOptions(await api('get_product_types'));
      const saved=String((S.statOpts[def.id] || {}).type || '').trim().toUpperCase();
      S.statOpts[def.id]=Object.assign({},S.statOpts[def.id],{type:opts.includes(saved)?saved:'SLAB'});
      if(opts.length>1)fields.push({key:'type',label:t('dlgType'),type:'select',options:opts,value:'SLAB'});
    }
    return fields;
  }

  async function runStat(def, opts, btn) {
    const btns = Array.from(document.querySelectorAll('#stat-actions button'));
    btns.forEach((b) => (b.disabled = true));
    statOverlay(true, t('loadingGen'));
    loadBar(true);
    await nextPaint();
    try {
      await runStatInner(def, opts, btn);
    } finally {
      statOverlay(false);
      btns.forEach((b) => (b.disabled = false));
      loadBar(false);
    }
  }

  async function runStatInner(def, opts, btn) {
    // 计时埋点：取数/计算/渲染 三段耗时写进状态栏（排查“统计要一分钟以上”用）
    if(def.id.startsWith('act')) {opts=Object.assign({},opts,{start:'',end:''});
      if(['actBySales1','actByCity','actByType'].includes(def.id)) opts.fCust='';
    }
    const pf0 = performance.now();
    const annualNow = new Date();
    const annualWindow = annualNow.getFullYear() + '-' + annualNow.getMonth();
    const priorAnnual = def.id === 'annual' ? S.statResults[def.id] : null;
    const reusableAnnual = priorAnnual && !priorAnnual.saveFailed &&
      priorAnnual.annualAlgorithm === 2 && priorAnnual.annualWindow === annualWindow;
    const knownAnnualRevision = reusableAnnual ? priorAnnual.dataRevision || '' : '';
    const performanceRange = JSON.stringify([opts.start || '', opts.end || '']);
    const priorPerformance = def.id === 'salesPerformance' ? S.statResults[def.id] : null;
    const reusablePerformance = priorPerformance && !priorPerformance.saveFailed &&
      priorPerformance.performanceAlgorithm === 1 && priorPerformance.performanceRange === performanceRange;
    const knownPerformanceRevision = reusablePerformance ? priorPerformance.dataRevision || '' : '';
    const priorOrders = def.id === 'customerOrders' ? S.statResults[def.id] : null;
    const reusableOrders = priorOrders && !priorOrders.saveFailed &&
      priorOrders.ordersAlgorithm === 1 && priorOrders.ordersRange === performanceRange;
    const knownOrdersRevision = reusableOrders ? priorOrders.dataRevision || '' : '';
    const categoryRange = JSON.stringify([opts.start || '', opts.end || '', opts.category || 'ALL', opts.mode || 'amount', annualWindow]);
    const priorCategory = def.id === 'categoryQuery' ? S.statResults[def.id] : null;
    const reusableCategory = priorCategory && !priorCategory.saveFailed &&
      priorCategory.categoryAlgorithm === 1 && priorCategory.categoryRange === categoryRange;
    const knownCategoryRevision = reusableCategory ? priorCategory.dataRevision || '' : '';
    const historyRange = JSON.stringify([opts.start || '',opts.end || '',opts.custNo || '',opts.itemNo || '',opts.invoice || '',opts.sales || '']);
    const priorHistory = def.id === 'salesHistory' ? S.statResults[def.id] : null;
    const reusableHistory = priorHistory && !priorHistory.saveFailed && priorHistory.historyAlgorithm===1 && priorHistory.historyRange===historyRange;
    const knownHistoryRevision = reusableHistory ? priorHistory.dataRevision || '' : '';
    const customerStatsRange=JSON.stringify([opts.start || '',opts.end || '',opts.scope || 'all']);
    const priorCustomerStats=def.id==='customerStats' ? S.statResults[def.id] : null;
    const reusableCustomerStats=priorCustomerStats && !priorCustomerStats.saveFailed && priorCustomerStats.customerStatsAlgorithm===1 && priorCustomerStats.customerStatsRange===customerStatsRange;
    const knownCustomerStatsRevision=reusableCustomerStats ? priorCustomerStats.dataRevision || '' : '';
    const batchRange=JSON.stringify([opts.start || '',opts.end || '',opts.fCust || '',opts.fSales || '',opts.fCity || '']);
    const priorBatch=CUSTOMER_BATCH_IDS.includes(def.id) ? S.statResults[def.id] : null;
    const reusableBatch=priorBatch && !priorBatch.saveFailed && priorBatch.customerBatchAlgorithm===1 && priorBatch.customerBatchRange===batchRange;
    const knownBatchRevision=reusableBatch ? priorBatch.dataRevision || '' : '';
    const isProduct=['cabinetColors','stoneSales'].includes(def.id);
    const productRange=JSON.stringify([opts.start || '',opts.end || '',opts.type || 'SLAB',opts.includeUniversal===true]);
    const priorProduct=isProduct ? S.statResults[def.id]:null;
    const reusableProduct=priorProduct && !priorProduct.saveFailed && priorProduct.productAlgorithm===1 && priorProduct.productRange===productRange;
    const knownProductRevision=reusableProduct ? priorProduct.dataRevision || '':'';
    const priorFast = reusableAnnual ? priorAnnual : reusablePerformance ? priorPerformance : reusableOrders ? priorOrders : reusableCategory ? priorCategory : reusableHistory ? priorHistory : reusableCustomerStats ? priorCustomerStats : reusableBatch ? priorBatch : reusableProduct ? priorProduct : null;
    S.ctxCache = null; // 每次重新生成都读取最新数据，避免把旧缓存保存成新报表。
    const ctx = getCtx(opts.start, opts.end) || await (async () => {
      const c = await normalizeCtx(def.id === 'annual' ? await api('get_annual_context', knownAnnualRevision) :
        def.id === 'salesPerformance' ? await api('get_performance_context', opts.start || '', opts.end || '', knownPerformanceRevision) :
        def.id === 'customerOrders' ? await api('get_orders_context', opts.start || '', opts.end || '', knownOrdersRevision) :
        def.id === 'categoryQuery' ? await api('get_category_context', opts.start || '', opts.end || '', knownCategoryRevision) :
        def.id === 'salesHistory' ? await api('get_history_context', opts.start || '', opts.end || '', {custNo:opts.custNo,itemNo:opts.itemNo,invoice:opts.invoice,sales:opts.sales}, knownHistoryRevision) :
        def.id === 'customerStats' ? await api('get_customer_stats_context', opts.start || '', opts.end || '', knownCustomerStatsRevision) :
        CUSTOMER_BATCH_IDS.includes(def.id) ? await customerBatchContext(opts,knownBatchRevision) :
        isProduct ? await productContext(opts,knownProductRevision) :
        await api('get_stats_context', opts.start || '', opts.end || ''));
      if (!isProduct && !CUSTOMER_BATCH_IDS.includes(def.id) && !['annual', 'salesPerformance', 'customerOrders', 'categoryQuery', 'salesHistory', 'customerStats'].includes(def.id)) putCtx(opts.start || '', opts.end || '', c);
      return c;
    })();
    const pf1 = performance.now();
    const revision = ctx && ctx.revision;
    if (ctx && ctx.unchanged && priorFast) {
      priorFast.stale = false;
      if (!await activateIfFresh(priorFast.sheet, 'stat:' + S.dataToken)) {
        await renderSpec(priorFast);
        S.sheetToken[priorFast.sheet] = 'stat:' + S.dataToken;
      }
      S.yearToggle = priorFast.yearGroups ? { sheet: priorFast.sheet, groups: priorFast.yearGroups } : null;
      renderStatExtra(def);
      const elapsed = ((performance.now() - pf0) / 1000).toFixed(1);
      const st = document.getElementById('stat-state');
      if (st) st.textContent = reportState(priorFast) + (LANG === 'zh' ? ' · 数据未变，复用结果 · 总计 ' : ' · Data unchanged, reused · Total ') + elapsed + 's';
      toast(LANG === 'zh' ? '数据未变化，已显示保存的报表。' : 'Data unchanged. Showing saved report.');
      return;
    }
    if (!ctx || !ctx.sales || !ctx.sales.length) {
      if (def.id === 'salesHistory' && [opts.custNo,opts.itemNo,opts.invoice,opts.sales].some(v=>String(v || '').trim())) {
        toast(LANG === 'zh' ? '没有符合这些条件的销售记录。' : 'No sales records match these filters.');
        return;
      }
      // 区分「库里完全没有销售」和「所选时间段内没有销售」
      if (opts.start || opts.end) {
        let months = [];
        try { months = (await api('list_batches')).map((b) => b.month); } catch (e) { /* ignore */ }
        toast(tf('periodEmpty', { m: months.join(', ') || '-' }));
      } else {
        toast(t('statsNoData'));
      }
      return;
    }
    const spec = STATS[def.id](ctx.sales, Object.assign({}, opts, { reps: ctx.reps, stock: ctx.stock, customers: ctx.customers }), t);
    hideCostColumn(spec);
    const pf2 = performance.now();
    if (!spec || spec.error) {
      const key = 'err' + String((spec && spec.error) || '');
      toast(t(key) !== key ? t(key) : tf('statEmpty', { msg: (spec && spec.error) || '' }));
      return;
    }
    // 渲染行数封顶：Univer 对万行级表格无稳定路径（扩行悬死），超限截断并提示缩小条件。
    // 数据缓存仍保存完整 spec（导出/切回重铺同受此限，保证不再卡死）。
    const RENDER_CAP = 60000;   // 快照建表+格式随值写入后，万行级渲染已稳定（5 万行 ~1s）
    const totalRows = (spec.values || []).length - 1;
    if (totalRows > RENDER_CAP) {
      const v = spec.values;
      const keep = [v[0], ...v.slice(1, RENDER_CAP + 1)];
      // 超限截断时保留 Total/合计行（其数值按全量计算，仍然正确），不能把合计裁掉
      const lastRow = v[v.length - 1];
      if (v.length - 1 > RENDER_CAP && /^(total|grand total|.*总计|.*合计)/i.test(String(lastRow && lastRow[0] || ''))) {
        keep.push(lastRow);
      }
      spec.values = keep;
      spec.truncated = { shown: RENDER_CAP, total: totalRows };
      try { toast(tf('statTruncated', { n: RENDER_CAP, t: totalRows })); } catch (e) { /* ignore */ }
    }
    if (spec.yearGroups) { try { toast(t('annualCollapseHint')); } catch (e) { /* ignore */ } }
    initUniver();
    try {
      await renderSpec(spec);
    } catch (e) {
      toast('Render error: ' + (e && e.message ? e.message : e));
      return;
    }
    const pf3 = performance.now();
    spec.perf = { ctx: (pf1 - pf0) / 1000, calc: (pf2 - pf1) / 1000, render: (pf3 - pf2) / 1000 };
    try { console.info('[perf]', def.id, 'ctx=' + spec.perf.ctx.toFixed(2) + 's calc=' + spec.perf.calc.toFixed(2) + 's render=' + spec.perf.render.toFixed(2) + 's'); } catch (e) { /* ignore */ }
    spec.doneAt = now2();
    if (def.id === 'annual') {
      spec.dataRevision = revision;
      spec.annualWindow = annualWindow;
      spec.annualAlgorithm = 2;
    }
    if(isProduct){spec.dataRevision=revision;spec.productRange=productRange;spec.productAlgorithm=1;}
    if (CUSTOMER_BATCH_IDS.includes(def.id)) {
      spec.dataRevision=revision;spec.customerBatchRange=batchRange;spec.customerBatchAlgorithm=1;
    }
    if (def.id === 'customerStats') {
      spec.dataRevision=revision;spec.customerStatsRange=customerStatsRange;spec.customerStatsAlgorithm=1;
    }
    if (def.id === 'salesHistory') {
      spec.dataRevision=revision;spec.historyRange=historyRange;spec.historyAlgorithm=1;
    }
    if (def.id === 'categoryQuery') {
      spec.dataRevision = revision;
      spec.categoryRange = categoryRange;
      spec.categoryAlgorithm = 1;
    }
    if (def.id === 'customerOrders') {
      spec.dataRevision = revision;
      spec.ordersRange = performanceRange;
      spec.ordersAlgorithm = 1;
    }
    if (def.id === 'salesPerformance') {
      spec.dataRevision = revision;
      spec.performanceRange = performanceRange;
      spec.performanceAlgorithm = 1;
    }
    S.yearToggle = spec.yearGroups ? { sheet: spec.sheet, groups: spec.yearGroups } : null;
    spec.stale = false;
    const saveStarted = performance.now();
    try {
      const saved = await api('save_report', def.id, spec, opts, revision);
      if (!saved || !saved.ok) throw new Error('Save failed');
      spec.stale = !!saved.stale;
    } catch (e) {
      spec.saveFailed = true;
      toast(LANG === 'zh' ? '报表已生成，但数据库保存失败，请重新生成重试。' : 'Report generated, but saving failed. Generate again to retry.');
    }
    spec.perf.save = (performance.now() - saveStarted) / 1000;
    spec.perf.total = (performance.now() - pf0) / 1000;
    S.statResults[def.id] = spec;
    S.statSheetName[def.id] = spec.sheet;
    S.sheetToken[spec.sheet] = 'stat:' + S.dataToken;   // 数据未变时切回该报表=纯激活
    renderStatExtra(def);
    const st = document.getElementById('stat-state');
    if (st) {
      let txt = reportState(spec);
      if (spec.perf) txt += '（' + tf('statTiming', { a: spec.perf.ctx.toFixed(1), b: spec.perf.calc.toFixed(1), c: spec.perf.render.toFixed(1) }) + '）';
      txt += LANG === 'zh' ? ' · 保存 ' + spec.perf.save.toFixed(1) + 's · 总计 ' + spec.perf.total.toFixed(1) + 's' : ' · Save ' + spec.perf.save.toFixed(1) + 's · Total ' + spec.perf.total.toFixed(1) + 's';
      st.textContent = txt;
    }
    if (btn) {
      document.querySelectorAll('#stat-actions button').forEach((b) => b.classList.remove('triggered'));
      btn.classList.add('triggered');
      S.statTriggered[def.id] = btn.textContent;
    }
    toast(tf('statDone', { s: spec.sheet }));
  }

  // 汇总条（客户活跃统计等带 summary 的报表显示在表格上方）
  function renderStatExtra(def) {
    const wrap = $('#stat-extra');
    if (!wrap) return;
    const spec = S.statResults[def.id];
    if (!spec || !spec.summary || !spec.summary.length) { wrap.innerHTML = ''; return; }
    wrap.innerHTML = '<div class="stat-summary">' + spec.summary.map(([k, v]) =>
      '<span class="chip"><b>' + esc(k) + '</b> ' + (typeof v === 'number' && !Number.isInteger(v) && v <= 1 ? (v * 100).toFixed(1) + '%' : esc(v)) + '</span>').join('') + '</div>';
  }

  async function exportStatResult(def) {
    if(costToken && !await checkCostAccess())return;
    const spec = hideCostColumn(S.statResults[def.id]);
    if (!spec || !spec.values || !spec.values.length) { toast(t('exportEmpty')); return; }
    // 先走后端接口（本地 exe 版桥接实现）；远程版后端无此接口，失败后走浏览器本地导出
    try {
      const res = await api('export_sheet_xlsx', JSON.stringify({ name: spec.sheet, values: spec.values }));
      if (res && res.ok) { toast(tf('exportDone', { p: res.path })); return; }
    } catch (e) { /* 忽略，走下面本地导出 */ }
    // 浏览器本地导出（远程版）：SheetJS 在页面内生成 xlsx 并下载
    try {
      await ensureXlsxLib();
      const ws = XLSX.utils.aoa_to_sheet(spec.values);
      const wb = XLSX.utils.book_new();
      const sheetName = String(spec.sheet || 'Sheet').replace(/[\\/?*[\]:]/g, '').slice(0, 31) || 'Sheet';
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
      const fileName = sheetName + '.xlsx';
      XLSX.writeFile(wb, fileName);
      toast(tf('exportDone', { p: fileName }));
    } catch (e) {
      toast(tf('exportFail', { e: e.message || e }));
    }
  }

  // 懒加载 SheetJS（xlsx.full.min.js 与 app.js 同目录），只在导出时加载一次
  let _xlsxLibPromise = null;
  function ensureXlsxLib() {
    if (typeof XLSX !== 'undefined') return Promise.resolve();
    if (_xlsxLibPromise) return _xlsxLibPromise;
    _xlsxLibPromise = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = 'xlsx.full.min.js';
      s.onload = () => resolve();
      s.onerror = () => reject(new Error('xlsx lib load failed'));
      document.head.appendChild(s);
    });
    return _xlsxLibPromise;
  }

  let statCardToken = 0;   // 快速连续切换菜单时，只让最后一次的渲染生效
  async function renderStatCard() {
    const def = S.statDef;
    if (!def) return;
    const tk = ++statCardToken;
    loadBar(true);
    try {
      await renderStatCardInner(def, tk);
    } finally {
      loadBar(false);
    }
  }

  async function renderStatCardInner(def, tk) {
    const def2 = def;
    if (!def2) return;
    S.reportsLoaded = S.reportsLoaded || {};
    if (!S.reportsLoaded[def.id] && !S.statResults[def.id]) {
      try {
        const result = await api('load_report', def.id);
        const record = result.report;
        if (record && !S.statResults[def.id]) {
          S.statResults[def.id] = record.spec;
          S.statOpts[def.id] = record.options || {};
          S.statSheetName[def.id] = record.spec.sheet;
        }
        S.reportsLoaded[def.id] = true;
      } catch (e) {
        toast(LANG === 'zh' ? '已保存报表读取失败，请重新打开重试。' : 'Could not load saved report. Reopen to retry.');
      }
    }
    if (tk !== undefined && tk !== statCardToken) return;
    $('#stat-cond-title').textContent = '▶ ' + t(def.labelKey);
    const box = $('#stat-cond-fields');
    box.innerHTML = '';
    if (!S.statOpts[def.id]) S.statOpts[def.id] = {};
    // 快速时间段：上月 / 今年 / TTM / 重置（重置=不限时段即全部数据）
    if ((def.fields || []).includes('range')) {
      const now = new Date();
      const quick = document.createElement('div');
      quick.className = 'stat-quick';
      const setRange = (s, e) => {
        S.statOpts[def.id].start = s || '';
        S.statOpts[def.id].end = e || '';
        const i1 = box.querySelector('input[data-k="start"]');
        const i2 = box.querySelector('input[data-k="end"]');
        if (i1) i1.value = s || '';
        if (i2) i2.value = e || '';
      };
      const quicks = [
        ['quickLM', () => [new Date(now.getFullYear(), now.getMonth() - 1, 1), new Date(now.getFullYear(), now.getMonth(), 0)]],
        ['quickTY', () => [new Date(now.getFullYear(), 0, 1), now]],
        ['quickTTM', () => { const r = ttmRange(); return [r.start, r.end]; }],
        ['quickReset', () => ['', '']],
      ];
      quicks.forEach(([key, fn]) => {
        const b = document.createElement('button');
        b.type = 'button';
        b.className = 'btn sm';
        b.textContent = key === 'quickTTM' ? 'TTM' : t(key);
        b.onclick = () => {
          const pair = fn();
          // 快速段回调可返回 Date 或已格式化的日期字符串（ttmRange 返回字符串），统一归一
          const fmt = (v) => (v instanceof Date ? isoDate(v) : String(v || ''));
          setRange(fmt(pair[0]), fmt(pair[1]));
        };
        quick.appendChild(b);
      });
      box.appendChild(quick);
    }
    const fields = await statOptionFields(def);
    if (tk !== undefined && tk !== statCardToken) return;   // 期间用户已切走，放弃本次渲染
    fields.forEach((f) => {
      if (S.statOpts[def.id][f.key] === undefined && f.value !== undefined) S.statOpts[def.id][f.key] = f.value;
      const row = document.createElement('label');
      row.className = 'cond-row';
      const cur = S.statOpts[def.id][f.key];
      row.innerHTML = '<span>' + esc(f.label) + '</span><input type="' + (f.type === 'combo' ? 'text' : (f.type || 'text')) + '" data-k="' + f.key + '" value="' + esc(cur === undefined ? '' : cur) + '">';
      if(f.type==='select') row.innerHTML='<span>'+esc(f.label)+'</span><select data-k="'+f.key+'">'+f.options.map(v=>'<option value="'+esc(v)+'"'+(v===cur?' selected':'')+'>'+esc(v)+'</option>').join('')+'</select>';
      row.querySelectorAll('[data-k]').forEach((el) => {
        el.addEventListener('change', () => {
          S.statOpts[def.id][el.dataset.k] = el.value;
        });
      });
      box.appendChild(row);
      if (f.type === 'combo') {
        const input = row.querySelector('input[data-k]');
        if (f.options) {
          attachFuzzy(input, { fetch: (text) => fuzzySimple(f.options, text) });
        } else if (f.fuzzy === 'customers') {
          attachFuzzy(input, { fetch: fuzzyCustomersCode });
        } else if (f.fuzzy === 'items') {
          attachFuzzy(input, { fetch: fuzzyItems });
        } else if (f.fuzzy === 'invoices') {
          attachFuzzy(input, { fetch: fuzzyInvoices });
        } else if (f.fuzzy === 'rep1') {
          attachFuzzy(input, { fetch: fuzzyRep1 });
        } else if (f.fuzzy === 'cities') {
          attachFuzzy(input, { fetch: (text) => getLookup().then((lk) => fuzzySimple(lk.cities || [], text)) });
        }
      }
    });
    // 功能按钮区：原 Sidebar 里每个板块的多个生成按钮 + 导出 Excel
    const actions = $('#stat-actions');
    actions.innerHTML = '';
    const btns = (def.actions || [{ labelKey: 'btnGenerate', opts: {} }]).map((ba) => ({ label: t(ba.labelKey), opts: ba.opts }));
    btns.push({ label: t('btnExport'), opts: null, export: true });
    btns.forEach((b) => {
      const el = document.createElement('button');
      el.className = 'btn' + (b.export ? '' : ' primary');
      el.textContent = b.label;
      el.onclick = () => {
        if (b.export) { exportStatResult(def); return; }
        runStat(def, Object.assign({}, collectStatOpts(def), b.opts), el);
      };
      actions.appendChild(el);
    });
    const st = document.createElement('span');
    st.className = 'info';
    st.id = 'stat-state';
    actions.appendChild(st);
    // 恢复置灰标记：当前工作表由哪个功能按钮触发渲染（灰色=当前表格来源）
    const trig = S.statTriggered[def.id];
    if (trig) {
      const hit = [...actions.querySelectorAll('button')].find((b) => b.textContent === trig);
      if (hit) hit.classList.add('triggered');
    }
    if(def.id==='salesHistory' && S.costRefreshHistory){
      S.costRefreshHistory=false;
      await runStat(def,Object.assign({},collectStatOpts(def)),null);return;
    }
    const cached = S.statResults[def.id];
    $('#stat-state').textContent = reportState(cached);
    renderStatExtra(def);
    // 切换报表必须切到它自己的表格：
    //  - 有最近一次结果 → 重铺（重绘矩阵，不重算）
    //  - 没有结果（从未生成/数据变更后被清）→ 铺该报表自己的空表，绝不停留在别的功能的表上
    S.yearToggle = cached && cached.yearGroups ? { sheet: cached.sheet, groups: cached.yearGroups } : null;
    if (cached) {
      // 数据未变（token 命中）→ 纯激活，避免整张重铺卡顿；否则重铺缓存矩阵
      if (!await activateIfFresh(cached.sheet, 'stat:' + S.dataToken)) {
        try { await renderSpec(cached); } catch (e) { /* ignore */ }
        S.sheetToken[cached.sheet] = 'stat:' + S.dataToken;
      }
    } else {
      const nm = S.statSheetName[def.id] || STATS.sheetNames[def.id] || '';
      if (nm) {
        if (!await activateIfFresh(nm, 'stat:' + S.dataToken)) {
          try { await renderSpec({ sheet: nm, values: [['']] }); } catch (e) { /* ignore */ }
          S.sheetToken[nm] = 'stat:' + S.dataToken;
        }
      }
    }
  }

  // 初始化（对应原 Apps Script「初始化」）：补全销售2 / 补齐销售分组 / 刷新查询数据源
  async function initializeFlow() {
    const ok = await confirmModal(t('initConfirm'), t('confirm'));
    if (!ok) return;
    let res;
    try { res = await api('initialize_data'); } catch (e) { toast('Init error: ' + e); return; }
    LOOKUP = null;
    REPS_CACHE = null;
    renderTab(S.nav);
    toast(tf('initDone', { s2: res.sales2Filled || 0, s1: (res.newS1 || []).length, s2n: (res.newS2 || []).length }));
    saved();
  }

  // ================= menus & tabs =================
  function buildMenus() {
    const mb = $('#menubar');
    mb.innerHTML = '';
    const mk = (label) => {
      const item = document.createElement('div');
      item.className = 'menu-item';
      item.textContent = label;
      return item;
    };
    const file = mk(t('menuFile'));
    const fileDrop = document.createElement('div');
    fileDrop.className = 'dropdown';
    const fileItems = [
      { label: LANG==='zh'?'解锁管理权限…':'Unlock management…', fn: unlockCost },
      { label: t('miInitialize'), fn: initializeFlow, admin:true },
      { sep: true },
      { sep: true },
      { label: t('miExportBackup'), fn: exportBackup, admin:true },
      { label: t('miRestoreBackup'), fn: restoreBackup, admin:true },
    ];
    // 统计功能已平铺在「统计报表」页顶部按钮栏，顶部不再设「统计」下拉菜单
    fileItems.forEach((f) => {
      if (f.sep) {
        const sep = document.createElement('div');
        sep.className = 'dd-sep';
        fileDrop.appendChild(sep);
        return;
      }
      const it = document.createElement('div');
      it.className = 'dd-item';
      it.textContent = f.label;
      const locked=f.admin && !costVisible();
      if(locked){it.style.opacity='.4';it.style.cursor='not-allowed';it.setAttribute('aria-disabled','true');}
      it.onclick = (e) => { e.stopPropagation(); if(locked)return; closeDropdowns(); f.fn(); };
      fileDrop.appendChild(it);
    });
    file.appendChild(fileDrop);
    [file].forEach((m) => {
      m.onclick = (e) => {
        e.stopPropagation();
        const open = m.classList.contains('open');
        closeDropdowns();
        if (!open) m.classList.add('open');
      };
    });
    const brand = document.createElement('div');
    brand.className = 'brand';
    brand.textContent = '📊 ' + t('appName');
    const spacer = document.createElement('div');
    spacer.className = 'menu-spacer';
    const langBtn = document.createElement('button');
    langBtn.className = 'btn ghost';
    langBtn.textContent = t('langName');
    langBtn.onclick = async () => {
      LANG = LANG === 'en' ? 'zh' : 'en';
      try { await api('set_lang', LANG); } catch (e) { /* ignore */ }
      applyI18n();
    };
    const fsBtn = document.createElement('button');
    fsBtn.className = 'btn ghost';
    fsBtn.textContent = '⛶ ' + t('fullscreen');
    fsBtn.onclick = toggleFullscreen;
    const saveBtn = document.createElement('button');
    saveBtn.className = 'btn ghost';
    saveBtn.textContent = '💾 ' + t('miSave') + ' (Ctrl+S)';   // 快捷键直接标在按钮上，使用者不用查
    saveBtn.title = t('manualSavedTip');
    saveBtn.onclick = manualSave;
    mb.appendChild(brand);
    mb.appendChild(file);
    mb.appendChild(spacer);
    mb.appendChild(saveBtn);
    mb.appendChild(langBtn);
    mb.appendChild(fsBtn);
  }

  // 手动保存：与切换/退出时同一套落库链（分组直接编辑 + 客户动态两块面板 + 主画布最近表 + 待写队列）。
  // 平时不动它也行——切换/关闭仍会自动保存；这个按钮用于“改完样式当场保存”。
  async function manualSave() {
    try {
      if (S.nav === 'reps' && S.repGrid) {
        const n = syncRepsFromGrid();
        if (n) toast(tf('repsAutoSaved', { n }));
      }
      syncDashMemory();
      if (S.u && S.u.api && S.lastRenderedSheet && !S.isRendering) {
        const wb = S.u.api.getActiveWorkbook();
        const prev = wb.getSheets().find((x) => { try { return x.getSheetName() === S.lastRenderedSheet; } catch (e) { return false; } });
        if (prev) syncSheetMemory(prev);
      }
      await Promise.allSettled(S.pendingSaves.slice());
      toast(t('manualSaved'));
    } catch (e) {
      toast(t('manualSavedFail'));
    }
  }

  // Ctrl+S 手动保存（拦截浏览器默认）
  document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && String(e.key).toLowerCase() === 's') {
      e.preventDefault();
      manualSave();
    }
  });

  function closeDropdowns() {
    document.querySelectorAll('.menu-item.open').forEach((m) => m.classList.remove('open'));
  }
  document.addEventListener('click', closeDropdowns);

  function toggleFullscreen() {
    if (window.pywebview && window.pywebview.api) { window.pywebview.api.toggle_fullscreen(); return; }
    if (document.fullscreenElement) document.exitFullscreen();
    else document.documentElement.requestFullscreen();
  }
  window.addEventListener('keydown', (e) => {
    if (e.key === 'F11') { e.preventDefault(); toggleFullscreen(); }
  });
  window.addEventListener('resize', () => { if (S.dash.chart) S.dash.chart.resize(); });

  // ================= sidebar navigation（多级可折叠） =================
  const NAV_ICONS = {
    sales: '📋', customers: '👥', inventory: '📦', reps: '🧑‍💼', history: '🕘',
    dashboard: '📈', annual: '📅', salesPerformance: '🏆', customerOrders: '🧾',
    categoryQuery: '🔎', customerStats: '👤', cabinetColors: '🎨', stoneSales: '🪨',
    salesHistory: '🕘', customerType: '🏷', salesAccount: '🆕', customerActivity: '💓',
    actDetail: '💓', actBySales1: '🧑‍💼', actByCity: '🏙', actByType: '🏷',
    actSleeping: '😴', actNever: '🚫', actNew: '🌟', actReact: '🔔',
  };
  // children 里还可以再嵌套 children，即可支持三级及以上菜单
  // 报表类功能全部收进左侧多级菜单（r:报表ID 直达该报表），不再使用顶部横向选项卡
  const NAV = [
    { id: 'g-data', labelKey: 'navData', icon: '🗂', children: [
      { id: 'sales', labelKey: 'tabSales' },
      { id: 'customers', labelKey: 'tabCustomers' },
      { id: 'inventory', labelKey: 'tabInventory' },
      { id: 'reps', labelKey: 'tabReps' },
      { id: 'history', labelKey: 'tabHistory' },
    ] },
    { id: 'g-report', labelKey: 'navReports', icon: '📊', children: STAT_HUBS.map((h) => ({
      id: 'g-' + h.id, labelKey: h.labelKey, children: h.reports.map((r) => ({ id: 'r:' + r.id, labelKey: r.labelKey })),
    })) },
    { id: 'dashboard', labelKey: 'tabDashboard' },
  ];
  function buildSidebar() {
    const sb = $('#sidebar');
    sb.classList.toggle('collapsed', S.sidebarCollapsed);
    sb.innerHTML = '';
    const toggle = document.createElement('button');
    toggle.className = 'nav-toggle';
    toggle.textContent = S.sidebarCollapsed ? '»' : '«';
    toggle.title = S.sidebarCollapsed ? (LANG === 'zh' ? '展开侧边栏' : 'Expand sidebar') : (LANG === 'zh' ? '收起侧边栏' : 'Collapse sidebar');
    toggle.onclick = async () => {
      S.sidebarCollapsed = !S.sidebarCollapsed;
      try { await api('set_pref', 'sidebarCollapsed', S.sidebarCollapsed ? '1' : '0'); } catch (e) { /* ignore */ }
      buildSidebar();
    };
    sb.appendChild(toggle);
    renderNavNodes(NAV, sb, 0);
  }
  // 一级菜单：所有功能项平铺直达（与客户动态同级），分组只作静态小标题不可折叠；
  // 纯分组的中间层（如「报表」下全是板块）不再显示标题，避免两层标题堆叠。
  function renderNavNodes(nodes, sb, depth) {
    nodes.forEach((n) => {
      if (n.children) {
        const allGroups = n.children.every((c) => c.children);
        if (!allGroups) {
          const cap = document.createElement('div');
          cap.className = 'nav-cap';
          cap.textContent = (n.icon ? n.icon + ' ' : '') + t(n.labelKey);
          sb.appendChild(cap);
        }
        renderNavNodes(n.children, sb, depth + 1);
      } else {
        const a = document.createElement('div');
        a.className = 'nav-item' + (S.nav === n.id ? ' active' : '');
        if (S.sidebarCollapsed) {
          a.textContent = NAV_ICONS[n.id.replace('hub:', '').replace('r:', '')] || '·';
          a.title = t(n.labelKey);
        } else {
          a.textContent = (n.icon ? n.icon + ' ' : '') + t(n.labelKey);
          a.title = '';
        }
        const locked=['customers','inventory','reps','history'].includes(n.id) && !costVisible();
        if(locked){a.style.opacity='.4';a.style.cursor='not-allowed';a.setAttribute('aria-disabled','true');}
        a.onclick = () => {if(!locked)navigate(n.id);};
        sb.appendChild(a);
      }
    });
  }
  function navigate(id) {
    if(['customers','inventory','reps','history'].includes(id) && !costVisible()){toast('请先解锁管理权限 / Unlock management first');return;}
    ['sales-import','cust-add','cust-import','cust-sync'].forEach(id=>{const b=document.getElementById(id);if(b)b.disabled=!costVisible();});
    // 离开当前页面时立即落库（不等定时器）：分组直接编辑 + 上一张渲染表的布局/样式调整。
    // 覆盖切到 history/dashboard 等不触发渲染的页面；关闭软件无可靠钩子，仍由 2.5s 定时器兜底最后几秒。
    if (id !== S.nav) {
      if (S.nav === 'reps') {
        const n = syncRepsFromGrid();
        if (n) toast(tf('repsAutoSaved', { n }));
      }
      if (S.nav === 'dashboard') syncDashMemory();   // 离开客户动态：保存两块面板的布局/样式调整
      try {
        if (S.u && S.u.api && S.lastRenderedSheet) {
          const wb0 = S.u.api.getActiveWorkbook();
          const prev = wb0.getSheets().find((x) => { try { return x.getSheetName() === S.lastRenderedSheet; } catch (e) { return false; } });
          if (prev) syncSheetMemory(prev);
        }
      } catch (e) { /* ignore */ }
    }
    // 统一直达报表：r:报表ID；兼容旧 stat:报表ID / hub:板块ID（折算成板块第一个报表）
    if (id.startsWith('stat:')) id = 'r:' + id.slice(5);
    if (id.startsWith('hub:')) {
      const hub = STAT_HUBS.find((h) => h.id === id.slice(4));
      id = 'r:' + (hub && hub.reports[0] ? hub.reports[0].id : 'salesPerformance');
    }
    S.nav = id;
    const isReport = id.startsWith('r:');
    const isData = ['sales', 'customers', 'inventory', 'reps', 'history'].includes(id);
    const isDash = id === 'dashboard';
    document.querySelectorAll('.page').forEach((p) => (p.style.display = 'none'));
    const page = document.getElementById(isDash ? 'page-dashboard' : 'page-sheet');
    if (page) page.style.display = 'flex';
    ['wb-sales', 'wb-customers', 'wb-inventory', 'wb-reps', 'wb-history', 'wb-stat'].forEach((wb) => {
      const el = document.getElementById(wb);
      if (el) el.style.display = 'none';
    });
    const um = document.getElementById('univerMount');
    if (um) um.style.display = 'block';   // 导入记录页会隐藏工作表，离开时务必恢复
    if (isDash) {
      renderDashboard();   // 客户动态用两块独立 Excel 面板（懒加载引擎），不用主画布
      buildSidebar();
      return;
    }
    initUniver();
    try{S.u.api.getActiveWorkbook().setEditable(true);}catch(e){}
    // 工作簿常驻，进页面时确保已初始化
    if (isReport) {
      const ws = document.getElementById('wb-stat');
      if (ws) ws.style.display = 'block';
      S.statDef = STAT_DEFS.find((d) => d.id === id.slice(2)) || STAT_DEFS[0] || null;
      renderStatCard();
    } else if (isData) {
      const wb = document.getElementById('wb-' + id);
      if (wb) wb.style.display = 'block';
      const um = document.getElementById('univerMount');
      if (um) um.style.display = (id === 'history') ? 'none' : 'block';
      S.sheetCtx = null;
      S.sheetSel = -1;
      // 切页面时把各数据页的 Edit/Delete 复位为禁用，避免上一页残留的启用状态 + 旧选中导致点删除没反应或删错行；点行后 onSheetCell 再按当前页启用
      ['sales-edit', 'cust-edit', 'cust-del', 'inv-edit', 'inv-del', 'reps-edit', 'reps-del'].forEach((bid) => { const b = document.getElementById(bid); if (b) b.disabled = true; });
      renderTab(id);
    }
    buildSidebar();
  }
  function renderTab(id) {
    try{S.u.api.getActiveWorkbook().setEditable(true);}catch(e){}
    loadBar(true);
    let p = null;
    if (id === 'sales') p = renderSales();
    else if (id === 'customers') p = renderCustomers();
    else if (id === 'inventory') p = renderInventory();
    else if (id === 'reps') p = renderReps();
    else if (id === 'history') p = renderHistory();
    else if (id === 'dashboard') p = renderDashboard();
    Promise.resolve(p).finally(() => {
      try{S.u.api.getActiveWorkbook().setEditable(!(['sales','customers'].includes(S.nav) && !costVisible()));}catch(e){}
      loadBar(false);
    });
  }

  function collectStatOpts(def) {
    const vals = {};
    $('#stat-cond-fields').querySelectorAll('[data-k]').forEach((el) => {
      vals[el.dataset.k] = el.type === 'checkbox' ? el.checked : el.value;
    });
    S.statOpts[def.id] = vals;
    return vals;
  }

  // ================= customer trends dashboard（1竖2横） =================
  // 左=客户列表（竖，带选择勾选框 + 时间范围）；右上=客户销售动态分析（年份表+月份表，横）；
  // 右下=最近12个完整月份销售趋势折线图（横）。对应原 Apps Script「客户动态分析」表布局。
  const PERIOD_KEYS = [['ALL', 'periodAll'], ['LAST_MONTH', 'periodLastMonth'], ['THIS_YEAR', 'periodThisYear'], ['YTD', 'periodYtd']];
  let dashboardRenderToken=0;
  const DASH_SNAPSHOT_KEY='bcs-dashboard-v2';
  function dashStatus(message) {
    let el=document.getElementById('dash-status');
    if(!el){el=document.createElement('span');el.id='dash-status';el.style.cssText='font-size:12px;color:#65758b;margin-left:8px';$('#dash-periods').appendChild(el);}
    el.textContent=message;
  }
  async function drawDashboard() {
    const dd=S.dash.cache;
    if(!dd)return;
    dd.period=S.dash.period;
      // 时间段按钮（对应原表 D1「时间范围」下拉）
      const seg = $('#dash-periods');
      seg.innerHTML = '';
      PERIOD_KEYS.forEach(([k, lk]) => {
        const b = document.createElement('button');
        b.className = 'btn sm seg-btn' + (S.dash.period === k ? ' primary' : '');
        b.textContent = t(lk);
        b.onclick = () => { S.dash.period = k; renderDashboard(true); };
        seg.appendChild(b);
      });
      $('#dash-list-title').textContent = t('dashList') + '（' + t(PERIOD_KEYS.find(([k]) => k === S.dash.period)[1]) + '）';
      await initDashEngines();
      if (!dd.customers || !Object.keys(dd.customers).length) {
        $('#dash-cust-title').textContent = '';
        if (S.dash.chart) S.dash.chart.clear();
        await renderDashList();
        await renderDashAna();
        return;
      }
      // 选中客户：记住选择；不存在则取当前榜第一
      const ranking = dd.rankings[S.dash.period] || [];
      if (!ranking.some((r) => r[0] === S.dash.customerNo)) {
        S.dash.customerNo = ranking.length ? ranking[0][0] : '';
      }
      await renderDashList();
      await renderDashAna();
      renderDashChart();
  }
  async function renderDashboard(localOnly=false) {
    const renderToken=++dashboardRenderToken;
    const windowKey=isoDate(new Date());
    loadBar(true);
    try {
      if(!S.dash.cache && !S.dash.snapshotLoaded) {
        S.dash.snapshotLoaded=true;
        let saved=null;
        try{saved=JSON.parse(localStorage.getItem(DASH_SNAPSHOT_KEY) || 'null');}catch(e){}
        if(!saved){try{const r=await api('load_report','customerDashboard');saved=r && r.report && r.report.spec;}catch(e){}}
        if(renderToken!==dashboardRenderToken)return;
        if(saved && saved.dashboardAlgorithm===2 && saved.dashboard && saved.dashboard.customers && saved.dashboard.rankings){
          S.dash.cache=saved.dashboard;S.dash.cacheWindow=saved.dashboardWindow;
          S.dash.cacheRevision=saved.dataRevision;S.dash.cacheToken=S.dataToken;
        }
      }
      const valid=S.dash.cache && S.dash.cacheToken===S.dataToken && S.dash.cacheWindow===windowKey;
      if(S.dash.cache) {
        await drawDashboard();
        if(localOnly && valid){dashStatus(LANG==='zh'?'已复用当前数据':'Using current data');return;}
        dashStatus(LANG==='zh'?'已显示保存结果，正在检查更新…':'Showing saved results; checking for updates…');
        loadBar(false);
      }
      const ctx=await normalizeCtx(await api('get_customer_stats_context','','',valid ? S.dash.cacheRevision || '' : ''));
      if(renderToken!==dashboardRenderToken)return;
      if(!ctx.unchanged){
        S.dash.cache=STATS.customerDashboardData(ctx.sales,{period:S.dash.period,customers:ctx.customers || []},ctx.reps);
        S.dash.cacheRevision=ctx.revision;S.dash.cacheToken=S.dataToken;S.dash.cacheWindow=windowKey;
        const snapshot={values:[],dashboardAlgorithm:2,dashboard:S.dash.cache,dataRevision:ctx.revision,dashboardWindow:windowKey};
        try{localStorage.setItem(DASH_SNAPSHOT_KEY,JSON.stringify(snapshot));}catch(e){}
        trackSave(api('save_report','customerDashboard',snapshot,{},ctx.revision).catch(()=>{}));
        await drawDashboard();
      }
      dashStatus(LANG==='zh'?'数据已更新':'Data is current');
    } catch(e) {
      if(S.dash.cache)dashStatus(LANG==='zh'?'当前显示保存结果；检查更新失败，请重新进入重试':'Showing saved results; update check failed');
      else toast('Load error: '+(e.message || e));
    } finally {if(renderToken===dashboardRenderToken)loadBar(false);}
  }

  // 客户动态两块面板各挂一个独立 Univer 实例（一个画布只能显示一张表，左右面板=两台引擎）。
  // 引擎懒加载：首次进入客户动态才创建；主画布（page-sheet）与此互不影响。
  async function initDashEngines() {
    if (S.dashEngines) { placeDashMounts(); return; }
    S.dashEngines = {
      list: makeUniverApp('dashListMount'),
      ana: makeUniverApp('dashAnaMount'),
    };
    placeDashMounts();
    // 左侧列表：点击行选中客户。挂在 mousedown 上（实测 mousedown 后选区立即正确，
    // mouseup 后 getActiveRange 反而读到旧值/默认值——挂 mouseup 会固定选中上一行）；
    // 轮询几次拿到新选区（Univer 应用选区可能晚于监听回调），拿到即选。
    try {
      document.body.addEventListener('mousedown', (e) => {
        if (!S.dashEngines || !e.target || !e.target.closest || !e.target.closest('#dashListMount')) return;
        const readRow = () => {
          try {
            const wb = S.dashEngines.list.api.getActiveWorkbook();
            const sh = wb.getSheets().find((x) => { try { return x.getSheetName() === '客户列表'; } catch (e2) { return false; } });
            const ar = sh && sh.getActiveRange ? sh.getActiveRange() : null;
            return ar && typeof ar.getRow === 'function' ? ar.getRow() : null;
          } catch (e3) { return null; }
        };
        let tries = 0;
        const tick = () => {
          const row = readRow();
          const no = row != null ? (S.dash.rowMap || {})[row] : null;
          tries++;
          if ((no && no !== S.dash.customerNo) || tries >= 6) {
            if (no && no !== S.dash.customerNo) {
              S.dash.customerNo = no;
              updateDashSelection();
              renderDashAna();
              renderDashChart();
            }
            return;
          }
          setTimeout(tick, 80);
        };
        setTimeout(tick, 0);
      }, true);
    } catch (e) { /* ignore */ }
  }

  function updateDashSelection() {
    try {
      const wb=S.dashEngines.list.api.getActiveWorkbook();
      const sh=wb.getSheets().find(x=>x.getSheetName()==='客户列表');
      if(!sh)return;
      if(S.dash.selectedRow!=null)sh.getRange(S.dash.selectedRow,0,1,5).setBackground(null);
      const pair=Object.entries(S.dash.rowMap || {}).find(([,no])=>no===S.dash.customerNo);
      S.dash.selectedRow=pair ? Number(pair[0]):null;
      if(S.dash.selectedRow!=null)sh.getRange(S.dash.selectedRow,0,1,5).setBackground('#dbe9ff');
    }catch(e){ /* native selection remains available */ }
  }

  // Univer UI 会把挂载元素挪到 body 下：拉回面板容器并触发画布重算
  function placeDashMounts() {
    [['dashListMount', 'dash-list-box'], ['dashAnaMount', 'dash-ana-box']].forEach(([mid, hid]) => {
      const m = document.getElementById(mid);
      const h = document.getElementById(hid);
      if (m && h && m.parentElement !== h) h.appendChild(m);
    });
    try {
      window.dispatchEvent(new Event('resize'));
      ['dashListMount', 'dashAnaMount'].forEach((mid) => {
        const el = document.getElementById(mid);
        if (el) el.dispatchEvent(new Event('resize'));
      });
    } catch (e) { /* ignore */ }
  }

  // 离开客户动态/关闭软件时保存两块面板的布局与样式调整（与主画布同一套记忆键：按表名）
  function syncDashMemory() {
    if (!S.dashEngines || S.isRendering) return;
    [['list', '客户列表'], ['ana', '客户动态分析']].forEach(([k, nm]) => {
      try {
        const wb = S.dashEngines[k].api.getActiveWorkbook();
        const sh = wb.getSheets().find((x) => { try { return x.getSheetName() === nm; } catch (e) { return false; } });
        if (sh) syncSheetMemory(sh);
      } catch (e) { /* ignore */ }
    });
  }

  // 小表直渲染：临时表重建（容量充足、无残留），格式随值写入；应用并记录布局/样式记忆（表名=记忆键）
  async function renderDashSheet(engKey, spec) {
    const eng = S.dashEngines && S.dashEngines[engKey];
    if (!eng) return;
    const wb = eng.api.getActiveWorkbook();
    if (!wb) return;
    const name = String(spec.sheet);
    const nc = Math.max(...spec.values.map((r) => r.length), 1);
    const old = wb.getSheets().find((x) => { try { return x.getSheetName() === name; } catch (e) { return false; } }) || null;
    let sh=old;
    // Reuse the fixed dashboard template without rescanning manual styles on each customer click.
    if(!old || old.getMaxRows()<spec.values.length || old.getMaxColumns()<nc || (engKey==='ana' && old.getMaxRows()!==14)) {
    // 先建 ~tmp 表、激活、删旧表、再改回正名：避开「删激活表被静默拒绝」「同名自动改名(1)」两个坑
    const tmpId = 'dsh_' + Date.now() + '_' + Math.floor(Math.random() * 10000);
    await eng.api.executeCommand('sheet.command.insert-sheet', {
      unitId: wb.getId(),
      sheet: { id: tmpId, name: name + '~tmp', rowCount: engKey==='ana'?14:Math.max(spec.values.length + 20, 60), columnCount: engKey==='ana'?9:Math.max(nc + 4, 10) },
    });
    sh = wb.getSheets().find((x) => x.getSheetId() === tmpId);
    if (!sh) return;
    try { sh.activate(); } catch (e) { /* ignore */ }
    if (old) {
      try { await eng.api.executeCommand('sheet.command.remove-sheet', { unitId: wb.getId(), subUnitId: old.getSheetId() }); } catch (e) { /* ignore */ }
    }
    try { sh.setName(name); } catch (e) { /* ignore */ }
    // 清理 createUnit 自带的默认空表（保持引擎里只有业务表）
    const blank = wb.getSheets().find((x) => {
      let n2 = '';
      try { n2 = x.getSheetName(); } catch (e) { /* ignore */ }
      return (n2 === 'Sheet1' || n2 === '工作表1') && x.getSheetId() !== sh.getSheetId();
    });
    if (blank) {
      try { await eng.api.executeCommand('sheet.command.remove-sheet', { unitId: wb.getId(), subUnitId: blank.getSheetId() }); } catch (e) { /* ignore */ }
    }
    } else {
      const extent=(S.dash.sheetExtents || {})[engKey];
      if(extent && extent.rows>spec.values.length) {
        const tail=sh.getRange(spec.values.length,0,extent.rows-spec.values.length,extent.cols);
        tail.clearContent();tail.clearFormat();
      }
    }
    S.dash.sheetExtents=S.dash.sheetExtents || {};
    S.dash.sheetExtents[engKey]={rows:spec.values.length,cols:nc};
    // 矩阵构造：0→空串（全套系统口径），numFmt 列金额格式随值写入
    const colIdx = (s) => { let n2 = 0; for (const ch of s) n2 = n2 * 26 + (ch.charCodeAt(0) - 64); return n2 - 1; };
    const fr = {};
    (spec.numFmt || []).forEach((f) => {
      const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(f.range);
      if (!m) return;
      const c1 = colIdx(m[1]), c2 = colIdx(m[3]);
      for (let c = c1; c <= c2; c++) fr[c] = { startRow0: Number(m[2]) - 1, pat: f.fmt };
    });
    const matrix = spec.values.map((row, ri) => row.map((v, ci) => {
      const st={bl:ri===0?1:0,vt:2,bg:{rgb:'#ffffff'},ht:ri===0?2:engKey==='list'?(ci===1?1:ci===2?3:2):([0,4,5].includes(ci)?2:3)};
      // Explicit borders remain visible on white fill and selected-row fill.
      const edge={s:1,cl:{rgb:'#b8c4d2'}};
      st.bd={t:edge,b:edge,l:edge,r:edge};
      const f=fr[ci];if(f && ri>=f.startRow0 && typeof v==='number')st.n={pattern:zhFmt(f.pat)};
      cleanGeneratedStyle(st);
      return {v:v===0?'':v,s:st};
    }));
    try { sh.getRange(0, 0, matrix.length, nc).setValues(matrix); } catch (e) { /* ignore */ }
    try { sh.getRange(0, 0, 1, nc).setFontWeight('bold'); } catch (e) { /* ignore */ }
    (spec.bg || []).forEach((b) => { try { sh.getRange(b.range).setBackground(b.color); } catch (e) { /* ignore */ } });
    (spec.bold || []).forEach((rg) => { try { sh.getRange(rg).setFontWeight('bold'); } catch (e) { /* ignore */ } });
    const layKey=name;
    const savedLay=null;
    const finalWidths=engKey==='list'?[100,400,150,80,80]:[100,200,200,200,100,100,200,200,200];
    finalWidths.forEach((w,c)=>sh.setColumnWidth(c,w));
    sh.setRowHeights(0,matrix.length,30);
    if(engKey==='ana')sh.setRowHeight(13,30);
    const appliedH = {};
    if (savedLay && savedLay.h) {
      Object.entries(savedLay.h).forEach(([ri, h]) => {
        const idx = Number(ri);
        if (idx < sh.getMaxRows() && typeof h === 'number' && h >= 10) {
          try { sh.setRowHeight(idx, h); } catch (e) { /* ignore */ }
          appliedH[idx] = h;
        }
      });
    }
    // 冻结：报表默认 + 用户记忆覆盖
    let fz = spec.freeze ? { y: spec.freeze.rows || 0, x: spec.freeze.cols || 0 } : null;
    if (savedLay && savedLay.fz && (savedLay.fz.y || savedLay.fz.x)) fz = { y: savedLay.fz.y || 0, x: savedLay.fz.x || 0 };
    if (fz && (fz.y || fz.x)) { try { sh.setFrozenRows(fz.y || 0); } catch (e) { /* ignore */ } try { sh.setFrozenColumns(fz.x || 0); } catch (e) { /* ignore */ } }
    S.appliedLayouts[layKey] = { w: finalWidths, h: appliedH, fz };
    S.appliedStyles[layKey] = captureCellStyles(sh);
    try { sh.activate(); } catch (e) { /* ignore */ }
    placeDashMounts();
  }

  // 左面板：客户列表（点行选中、搜索框过滤、选中行高亮）
  async function renderDashList() {
    const dd = S.dash.cache || {};
    const ranking0 = (dd.rankings && dd.rankings[S.dash.period]) || [];
    const q = String(S.dash.query || '').trim().toLowerCase();
    const ranking = q ? ranking0.filter((r) => String(r[0]).toLowerCase().includes(q) || String(r[1] || '').toLowerCase().includes(q)) : ranking0;
    S.dash.rowMap = {};
    const values = [[t('custNo'), t('custName'), t('dashTotalAmt'), t('colRep1'), t('colRep2')]];
    let selRow = 0;
    ranking.forEach((r, i) => {
      S.dash.rowMap[i + 1] = r[0];   // mouseup 读 0 基行号：第 i+1 行对应第 i 个客户
      values.push([r[0], r[1], r[2], r[3], r[4]]);
      if (r[0] === S.dash.customerNo) selRow = i + 2;   // 1 基工作表行号（第 1 行是表头，客户从第 2 行起）
    });
    const spec = { sheet: '客户列表', values, numFmt: [], bold: [], bg: [], freeze: { rows: 1, cols: 0 } };
    const n = values.length;
    if (n > 1) spec.numFmt.push({ range: 'C2:C' + n, fmt: '#,##0.00;-#,##0.00;' });
    spec.bg.push({ range: 'A1:E1', color: '#d9eaf7' });
    if (selRow) spec.bg.push({ range: 'A' + selRow + ':E' + selRow, color: '#dbe9ff' });
    await renderDashSheet('list', spec);
    S.dash.selectedRow=selRow ? selRow-1:null;
  }

  // 右上面板：年份表（全部历史，A:D）+ 最近12个完整月表（F:I），各带合计行
  async function renderDashAna() {
    const dd = S.dash.cache || {};
    const cust = dd.customers ? dd.customers[S.dash.customerNo] : null;
    const titleEl = $('#dash-cust-title');
    if (titleEl) {
      titleEl.textContent = cust ? cust.no + '  ' + cust.name : '';
      if(cust){
        const badge=document.createElement('span');
        badge.textContent=(cust.isNew?'New / ':'')+(cust.status || 'Never');
        badge.style.cssText='display:inline-block;margin-left:12px;padding:3px 10px;border-radius:5px;font-size:15px;font-weight:700;';
        badge.style.color=cust.status==='Active'?'#166534':cust.status==='Sleeping'?'#475569':'#854d0e';
        badge.style.background=cust.status==='Active'?'#dcfce7':cust.status==='Sleeping'?'#e2e8f0':'#fef9c3';
        titleEl.appendChild(badge);
      }
    }
    const moneyHead = (first) => [first, t('dashTotalAmt'), t('dashCabinetAmt'), t('dashStoneAmt')];
    S.dash.yearTotal=[t('total'),'','',''];S.dash.monthTotal=[t('total'),'','',''];
    const yearRows = [moneyHead(t('dashYear'))];
    const monthRows = [moneyHead(t('dashMonth'))];
    if (cust) {
      let yt = 0, yc = 0, ys = 0;
      const currentYear=new Date().getFullYear();
      Object.keys(cust.years || {}).filter(yr=>Number(yr)>=currentYear-11 && Number(yr)<=currentYear).sort().forEach((yr) => {
        const m = cust.years[yr]; yt += m.total; yc += m.cabinet; ys += m.stone;
        yearRows.push([yr, m.total, m.cabinet, m.stone]);
      });
      S.dash.yearTotal=[t('total'),yt,yc,ys];
      let tt = 0, tc = 0, ts = 0;
      dd.monthKeys.slice(-12).forEach((mk) => {
        const m = cust.months[mk]; tt += m.total; tc += m.cabinet; ts += m.stone;
        monthRows.push([mk, m.total, m.cabinet, m.stone]);
      });
      S.dash.monthTotal=[t('total'),tt,tc,ts];
    }
    const nR = 13; // header plus twelve detail rows; append a normal summary row below
    const values = [];
    for (let i = 0; i < nR; i++) values.push(['', '', '', '', '', '', '', '', '']);   // A..I（E 列空档）
    yearRows.forEach((r, i) => { for (let c = 0; c < 4; c++) values[i][c] = r[c]; });
    monthRows.forEach((r, i) => { for (let c = 0; c < 4; c++) values[i][5 + c] = r[c]; });
    values.push([...S.dash.yearTotal,'',...S.dash.monthTotal]);
    const spec = { sheet: '客户动态分析', values, numFmt: [], bold: [], bg: [], freeze: { rows: 1, cols: 0 } };
    const money = '#,##0.00;-#,##0.00;';
    spec.numFmt.push({ range: 'B2:D14', fmt: money });
    spec.numFmt.push({ range: 'G2:I14', fmt: money });
    spec.bg.push({ range: 'A1:D1', color: '#d9eaf7' }, { range: 'F1:I1', color: '#d9eaf7' });
    spec.bg.push({range:'A14:I14',color:'#d9e6f3'});
    spec.bold.push('A14:I14');

    renderDashAnalysisTable(spec);
    document.getElementById('dash-fixed-totals')?.remove();
  }

  // One native table owns all column widths and both scroll axes. Sticky cells
  // pin the header/footer vertically without independent horizontal positioning.
  function renderDashAnalysisTable(spec) {
    const box=document.getElementById('dash-ana-box');
    const mount=document.getElementById('dashAnaMount');
    if(mount)mount.style.display='none';
    let view=document.getElementById('dash-analysis-table');
    if(!view){view=document.createElement('div');view.id='dash-analysis-table';box.appendChild(view);}
    const x=view.scrollLeft,y=view.scrollTop;
    const widths=[100,200,200,200,100,100,200,200,200];
    const row=(values,header=false)=>'<tr>'+values.map((v,c)=>{
      const tag=header?'th':'td',center=header || [0,4,5].includes(c);
      const value=typeof v==='number'?(v===0?'':fmtMoney(v)):v;
      return '<'+tag+' style="text-align:'+(center?'center':'right')+'">'+esc(value==null?'':value)+'</'+tag+'>';
    }).join('')+'</tr>';
    view.innerHTML='<table aria-label="'+esc(t('dashAnalysis'))+'" style="width:'+widths.reduce((a,b)=>a+b,0)+'px"><colgroup>'+widths.map(w=>'<col style="width:'+w+'px">').join('')+'</colgroup><thead>'+row(spec.values[0],true)+'</thead><tbody>'+spec.values.slice(1,-1).map(r=>row(r)).join('')+'</tbody><tfoot>'+row(spec.values[spec.values.length-1])+'</tfoot></table>';
    view.scrollLeft=x;view.scrollTop=y;
  }

  function renderDashChart() {
    const dd = S.dash.cache;
    const cust = dd.customers[S.dash.customerNo];
    const el = document.getElementById('dash-chart');
    if (!cust || !el) {if(S.dash.chart)S.dash.chart.clear();return;}
    if (!window.echarts) return;
    if (!S.dash.chart || S.dash.chart._dom !== el) S.dash.chart = window.echarts.init(el);
    const totals = dd.monthKeys.map((mk) => cust.months[mk].total);
    const cabs = dd.monthKeys.map((mk) => cust.months[mk].cabinet);
    const stones = dd.monthKeys.map((mk) => cust.months[mk].stone);
    const totalName = LANG === 'zh' ? '总金额' : 'Total';
    const cabName = LANG === 'zh' ? '橱柜金额' : 'Cabinet';
    const stoneName = LANG === 'zh' ? '石材金额' : 'Stone';
    // 原表图表为折线图：总金额/橱柜金额/石材金额 三条线，图例在顶部，月份斜排
    S.dash.chart.setOption({
      tooltip: { trigger: 'axis' },
      legend: { data: [totalName, cabName, stoneName], top: 0 },
      grid: { left: 10, right: 16, top: 34, bottom: 10, containLabel: true },
      xAxis: { type: 'category', boundaryGap: false, data: dd.monthKeys, axisLabel: { rotate: 40 } },
      yAxis: { type: 'value', axisLabel: { formatter: (v) => Number(v).toLocaleString('en-US') }, name: LANG === 'zh' ? '销售金额' : 'Amount', nameLocation: 'middle', nameGap: 46 },
      series: [
        { name: totalName, type: 'line', smooth: false, symbol: 'circle', symbolSize: 6, lineStyle: { width: 2 }, data: totals, itemStyle: { color: '#2456a6' } },
        { name: cabName, type: 'line', smooth: false, symbol: 'circle', symbolSize: 6, lineStyle: { width: 2 }, data: cabs, itemStyle: { color: '#e6a23c' } },
        { name: stoneName, type: 'line', smooth: false, symbol: 'circle', symbolSize: 6, lineStyle: { width: 2 }, data: stones, itemStyle: { color: '#67c23a' } },
      ],
    }, true);
    S.dash.chart.resize();
  }

  // ================= 数据管理页：列表全部呈现在单一 Excel 工作表 =================
  const SALES_COLS = [
    ['month', 'colMonth'], ['itemNo', 'colItemNo'], ['description', 'colDesc'], ['color', 'colColor'],
    ['type', 'colType'], ['soldDate', 'colDate'], ['invoice', 'colInvoice'], ['custNo', 'colCustNo'],
    ['custName', 'colCustName'], ['rep1', 'colRep1'], ['rep1Name', 'colRep1Name'], ['rep2', 'colRep2'],
    ['rep2Name', 'colRep2Name'], ['qty', 'colQty'], ['amount', 'colAmount'], ['cost', 'colCost'],
  ];

  function salesColumns(){return SALES_COLS.filter(([key])=>key!=='cost' || costVisible());}
  function colLetter(n) { let t = ''; while (n > 0) { const m2 = (n - 1) % 26; t = String.fromCharCode(64 + m2 + 1) + t; n = Math.floor((n - 1) / 26); } return t; }
  function rngA(r1, c1, r2, c2) { return colLetter(c1) + r1 + ':' + colLetter(c2) + r2; }
  function sheetSelEntry() {
    const ctx = S.sheetCtx;
    if (!ctx || S.sheetSel == null || S.sheetSel < 0) return null;
    return ctx.rows[S.sheetSel] || null;
  }
  function sheetRowBg(row, color) {
    try {
      const ctx = S.sheetCtx;
      const wb = S.u.api.getActiveWorkbook();
      const sh = wb.getSheets()[0];
      sh.getRange(row, 0, 1, ctx ? ctx.colCount : 20).setBackground(color || '#ffffff');
    } catch (e) { /* ignore */ }
  }
  function highlightSheetRow(row) {
    // Native cell/row selection is sufficient; do not persist selection as cell formatting.
    if(S.sheetCtx && ['sales','customers'].includes(S.sheetCtx.kind)) {S.lastHLRow=null;return;}
    if (S.lastHLRow != null && S.lastHLRow !== row) sheetRowBg(S.lastHLRow, null);
    S.lastHLRow = row;
    if (row != null && row >= 0) sheetRowBg(row, '#dbe9ff');
  }
  // Keep display arrows out of spec.values so Excel export stays unchanged.
  function applyAnnualGroups(spec, sh) {
    // 旧报表的隐藏行不能沿用到新行号；年份、TTM 和合计必须始终可见。
    sh.showRows(0, spec.values.length);
    spec.yearGroups.forEach((g) => {
      if (!g.months || !g.months.length) return;
      if (typeof g.collapsed !== 'boolean') g.collapsed = true;
      const first = g.months[0] - 1;
      if (g.collapsed) sh.hideRows(first, g.months.length);
      else sh.showRows(first, g.months.length);
      g.label = String(spec.values[g.row - 1][0]);
      sh.getRange(g.row - 1, 0, 1, 1).setValues([[g.label]]);
    });
    spec.yearGroups.forEach((g) => sh.showRows(g.row - 1, 1));
    S.yearToggle = { sheet: spec.sheet, groups: spec.yearGroups };
    ensureAnnualGutter();
  }

  let annualGutter = null;
  let annualGutterTimer = null;
  function ensureAnnualGutter() {
    if (annualGutterTimer) return;
    const mount = document.getElementById('univerMount');
    if (!mount) return;
    annualGutter = document.createElement('div');
    annualGutter.id = 'annualGroupGutter';
    annualGutter.setAttribute('aria-label', 'Annual month groups');
    annualGutter.style.cssText = 'position:absolute;inset:0;pointer-events:none;z-index:20;overflow:hidden;';
    mount.appendChild(annualGutter);
    const update = () => {
      const state = S.yearToggle;
      if (S.nav !== 'r:annual' || !state || !S.u || S.isRendering) {
        annualGutter.style.display = 'none'; return;
      }
      try {
        const wb = S.u.api.getActiveWorkbook();
        const sh = wb.getSheets().find((x) => x.getSheetName() === state.sheet);
        const canvases = [...mount.querySelectorAll('canvas')];
        const canvas = canvases.sort((a,b) => b.clientHeight - a.clientHeight)[0];
        if (!sh || !canvas) { annualGutter.style.display = 'none'; return; }
        const mr = mount.getBoundingClientRect(), cr = canvas.getBoundingClientRect();
        const zoom = sh.getSheet().getZoomRatio();
        const scroll = wb.getScrollStateBySheetId(sh.getSheetId()) || {};
        const sy = Number(scroll.viewportScrollY) || 0;
        const header = sh.getRange(0,0,1,1).getCellRect();
        const freeze = sh.getFreeze();
        const frozenRows = freeze && freeze.ySplit || 0;
        const boundary = frozenRows ? sh.getRange(frozenRows-1,0,1,1).getCellRect().bottom : header.top;
        const clipTop = cr.top - mr.top + boundary * zoom;
        const bottom = Math.min(mr.height, cr.bottom - mr.top) - 18;
        annualGutter.style.display = '';
        const ids = new Set();
        for (const g of state.groups || []) {
          const id = String(g.row); ids.add(id);
          let btn = annualGutter.querySelector('[data-year-row="' + id + '"]');
          if (!btn) {
            btn = document.createElement('button'); btn.type = 'button'; btn.dataset.yearRow = id;
            btn.style.cssText = 'position:absolute;width:18px;height:20px;padding:0;border:1px solid #9aa8b6;border-radius:3px;background:#fff;color:#254b70;font:bold 15px Arial;pointer-events:auto;cursor:pointer;';
            btn.addEventListener('mousedown', (e) => e.stopPropagation());
            btn.addEventListener('mouseup', (e) => e.stopPropagation());
            btn.addEventListener('click', (e) => { e.stopPropagation(); onSheetCell(Number(btn.dataset.yearRow)-1); update(); });
            annualGutter.appendChild(btn);
          }
          const rect = sh.getRange(g.row-1,0,1,1).getCellRect();
          const y = cr.top - mr.top + (rect.top - (g.row-1 >= frozenRows ? sy : 0)) * zoom + rect.height * zoom / 2 - 10;
          btn.style.left = Math.max(0, cr.left-mr.left+2) + 'px';
          btn.style.top = y + 'px';
          btn.style.display = y < clipTop || y+20 > bottom ? 'none' : '';
          btn.textContent = g.collapsed ? '+' : '−';
          btn.setAttribute('aria-expanded', String(!g.collapsed));
          btn.title = (g.collapsed ? 'Expand ' : 'Collapse ') + (g.label || '');
        }
        for (const btn of [...annualGutter.children]) if (!ids.has(btn.dataset.yearRow)) btn.remove();
      } catch (e) { annualGutter.style.display = 'none'; }
    };
    annualGutterTimer = setInterval(update, 100);
    update();
  }

  let lastAnnualClick = { sheet: '', row: -1, at: -Infinity };
  function onSheetCell(row) {
    if (S.nav === 'r:annual' && S.yearToggle && row != null) {
      const state = S.yearToggle;
      const g = (state.groups || []).find((x) => x.row === row + 1);
      if (g && g.months && g.months.length) {
        const now = performance.now();
        // CellClicked and the mouseup fallback can report the same click.
        if (lastAnnualClick.sheet === state.sheet && lastAnnualClick.row === row && now - lastAnnualClick.at < 180) return;
        lastAnnualClick = { sheet: state.sheet, row, at: now };
        try {
          const sh = S.u.api.getActiveWorkbook().getSheets().find((x) => x.getSheetName() === state.sheet);
          if (!sh) return;
          const collapsed = !g.collapsed;
          const ok = collapsed ? sh.hideRows(g.months[0] - 1, g.months.length) : sh.showRows(g.months[0] - 1, g.months.length);
          if (ok === false) throw new Error('Unable to change month visibility');
          g.collapsed = collapsed;
          if (g.label != null) sh.getRange(g.row - 1, 0, 1, 1).setValues([[g.label]]);
        } catch (e) {
          toast(LANG === 'zh' ? '月份折叠失败，请重新生成报表。' : 'Could not toggle months. Please regenerate the report.');
        }
        return;
      }
    }
    if(S.nav && S.nav.startsWith('r:')) return;
    const ctx = S.sheetCtx;
    if (!ctx || row === null || row === undefined || row < ctx.startRow) return;
    const idx = row - ctx.startRow;
    if (!ctx.rows[idx]) return;
    S.sheetSel = idx;
    highlightSheetRow(row);
    const enable = (id) => { const b = document.getElementById(id); if (b) b.disabled = false; };
    if (ctx.kind === 'sales' && costVisible()) enable('sales-edit');
    else if (ctx.kind === 'customers' && costVisible()) { enable('cust-edit'); enable('cust-del'); }
    else if (ctx.kind === 'inventory') { enable('inv-edit'); enable('inv-del'); }
    else if (ctx.kind === 'reps') { enable('reps-edit'); enable('reps-del'); }
  }

  let salesToken = 0;
  async function renderSales(force=false) {
    try{S.u.api.getActiveWorkbook().setEditable(true);}catch(e){}
    try {
    const token = ++salesToken;
    const all = !S.pageSize;   // pageSize=0 = 加载全部
    const tk = 'sales:' + S.dataToken + ':' + JSON.stringify(S.salesFilters) + ':' + S.salesPage + ':' + (all ? 'all' : S.pageSize);
    if (!force && await activateIfFresh(t('tabSales'), tk)) { S.sheetCtx = S.sheetCtxCache[t('tabSales')] || null; return; }
    let res;
    const ALL_CAP = 60000;   // 全部档上限（真实数据 5.35 万行全量可渲染）
    try {
      res = await normalizeQuery(await api('query_sales', all ? 0 : S.salesPage * S.pageSize, all ? ALL_CAP : S.pageSize, JSON.stringify(S.salesFilters)));
    } catch (e) { toast('Load error: ' + (e && e.message ? e.message : e)); return; }
    if (token !== salesToken) return;
    if (!res || !Array.isArray(res.rows)) { toast('Load error: bad response'); return; }
    const pages = all ? 1 : Math.max(1, Math.ceil(res.total / S.pageSize));
    if (!all && S.salesPage >= pages) {S.salesPage=pages-1;return renderSales(true);}
    const start = all ? 0 : S.salesPage * S.pageSize;
    S.lastSalesTotal = res.total;
    $('#sales-info').textContent = res.total
      ? (all
        ? (res.rows.length < res.total ? res.rows.length + ' / ' + res.total + ' (' + t('allCapHint') + ')' : res.total + ' / ' + res.total)
        : (start + 1) + '-' + Math.min(start + S.pageSize, res.total) + ' / ' + res.total)
      : '';
    $('#sales-page').textContent = (S.salesPage + 1) + ' / ' + pages;
    const values = [salesColumns().map(([, lk]) => t(lk))];
    res.rows.forEach((r) => values.push(salesColumns().map(([k]) => (r[k] === null || r[k] === undefined ? '' : r[k]))));
    const n = values.length;
    S.sheetCtx = { kind: 'sales', rows: res.rows, startRow: 1, colCount: salesColumns().length };
    S.sheetCtxCache[t('tabSales')] = S.sheetCtx;
    await renderSpec({
      sheet: t('tabSales'), values, salesQueryTemplate:true,
      numFmt: [
        { range: 'N2:N' + n, fmt: '#,##0' },
        { range: 'O2:P' + n, fmt: '#,##0.00' },
      ],
      freeze: { rows: 1, cols: 0 },
    });
    S.sheetToken[t('tabSales')] = tk;
    S.lastHLRow = null;
    } finally {
      try{S.u.api.getActiveWorkbook().setEditable(S.nav!=='sales' || costVisible());}catch(e){}
    }
  }

  function editSalesRow(r) {
    if(!costVisible()){toast('请先解锁管理权限');return;}
    const fields = [
      { key: 'itemNo', label: t('colItemNo'), value: r.itemNo },
      { key: 'description', label: t('colDesc'), value: r.description },
      { key: 'color', label: t('colColor'), value: r.color },
      { key: 'type', label: t('colType'), value: r.type },
      { key: 'soldDate', label: t('colDate'), type: 'date', value: r.soldDate },
      { key: 'invoice', label: t('colInvoice'), value: r.invoice },
      { key: 'custNo', label: t('colCustNo'), value: r.custNo },
      { key: 'custName', label: t('colCustName'), value: r.custName },
      { key: 'rep1', label: t('colRep1'), value: r.rep1 },
      { key: 'rep2', label: t('colRep2'), value: r.rep2 },
      { key: 'qty', label: t('colQty'), type: 'number', value: r.qty },
      { key: 'amount', label: t('colAmount'), type: 'number', value: r.amount },
      { key: 'cost', label: t('colCost'), type: 'number', value: r.cost },
    ];
    formModal(t('edit') + ' — ' + r.month, fields.filter(f=>f.key!=='cost' || costVisible())).then(async (values) => {
      if (!values) return;
      values.qty = Number(values.qty) || 0;
      values.amount = Number(values.amount) || 0;
      if(costVisible())values.cost = Number(values.cost) || 0;else delete values.cost;
      await api('update_sale', r.id, JSON.stringify(values));
      invalidateCtx();
      renderSales();
      saved();
    });
  }

  // ================= customers page =================
  const CUST_COLS = [['no', 'custNo'], ['name', 'custName'], ['sales1', 'colRep1'], ['created', 'custCreated'], ['city', 'colCity'], ['state', 'colState'], ['phone', 'colPhone'], ['contact', 'colContact'], ['type', 'colCustType'], ['email', 'colEmail']];
  let customerQueryToken=0;
  async function renderCustomers() {
    const queryToken=++customerQueryToken;
    const tk = 'cust:' + S.dataToken + ':' + (S.custQuery || '');
    if (await activateIfFresh(t('tabCustomers'), tk)) { S.sheetCtx = S.sheetCtxCache[t('tabCustomers')] || null; return; }
    const list = await api('get_all_customers', S.custQuery);
    if(queryToken!==customerQueryToken)return;
    $('#cust-count').textContent = list.length ? String(list.length) : '';
    const values = [CUST_COLS.map(([, lk]) => t(lk))];
    list.slice(0, 5000).forEach((r) => values.push(CUST_COLS.map(([k]) => (r[k] === null || r[k] === undefined ? '' : r[k]))));
    S.sheetCtx = { kind: 'customers', rows: list.slice(0, 5000), startRow: 1, colCount: CUST_COLS.length };
    S.sheetCtxCache[t('tabCustomers')] = S.sheetCtx;
    await renderSpec({ sheet: t('tabCustomers'), values, customerTemplate:true, freeze: { rows: 1, cols: 0 } });
    S.sheetToken[t('tabCustomers')] = tk;
    S.lastHLRow = null;
  }

  function editCustomer(row) {
    if(!costVisible()){toast('请先解锁管理权限');return;}
    const isNew = !row;
    const base = row || { no: '', name: '', city: '', state: '', phone: '', contact: '', type: '', email: '', sales1: '', created: '' };
    const fields = [
      { key: 'no', label: t('custNo'), value: base.no },
      { key: 'name', label: t('custName'), value: base.name },
      { key: 'sales1', label: t('colRep1'), value: base.sales1 },
      { key: 'created', label: t('custCreated'), type: 'date', value: base.created },
      { key: 'city', label: t('colCity'), value: base.city },
      { key: 'state', label: t('colState'), value: base.state },
      { key: 'phone', label: t('colPhone'), value: base.phone },
      { key: 'contact', label: t('colContact'), value: base.contact },
      { key: 'type', label: t('colCustType'), value: base.type },
      { key: 'email', label: t('colEmail'), value: base.email },
    ];
    formModal(isNew ? t('add') : t('edit'), fields).then(async (values) => {
      if (!values || !values.no) return;
      if (isNew) {
        const existing = await api('get_all_customers', values.no);
        if (existing.some((x) => x.no === values.no)) { toast(tf('dupNo', { no: values.no })); return; }
      }
      invalidateCtx();
      await api('upsert_customer', JSON.stringify(values));
      renderCustomers();
      saved();
    });
  }

  // ================= inventory page =================
  const INV_COLS = [['itemNo', 'invItemNo'], ['description', 'invDesc'], ['available', 'invAvailable'], ['inTransit', 'invInTransit'], ['stockCalc', 'invStock']];
  async function renderInventory() {
    const tk = 'inv:' + S.dataToken + ':' + (S.invQuery || '');
    if (await activateIfFresh(t('tabInventory'), tk)) { S.sheetCtx = S.sheetCtxCache[t('tabInventory')] || null; return; }
    const list = await api('get_all_stock', S.invQuery);
    $('#inv-count').textContent = list.length ? String(list.length) : '';
    const values = [INV_COLS.map(([, lk]) => t(lk))];
    list.forEach((r) => values.push([
      r.itemNo, r.description,
      Number(r.available) || 0, Number(r.inTransit) || 0,
      (Number(r.available) || 0) + (Number(r.inTransit) || 0),
    ]));
    const n = values.length;
    S.sheetCtx = { kind: 'inventory', rows: list, startRow: 1, colCount: INV_COLS.length };
    S.sheetCtxCache[t('tabInventory')] = S.sheetCtx;
    await renderSpec({
      sheet: t('tabInventory'), values,
      numFmt: [{ range: 'C2:E' + n, fmt: '#,##0' }],
      freeze: { rows: 1, cols: 0 },
    });
    S.sheetToken[t('tabInventory')] = tk;
    S.lastHLRow = null;
  }

  function editStock(row) {
    const isNew = !row;
    const base = row || { itemNo: '', description: '', available: 0, inTransit: 0 };
    const fields = [
      { key: 'itemNo', label: t('invItemNo'), value: base.itemNo },
      { key: 'description', label: t('invDesc'), value: base.description },
      { key: 'available', label: t('invAvailable'), type: 'number', value: base.available },
      { key: 'inTransit', label: t('invInTransit'), type: 'number', value: base.inTransit },
    ];
    formModal(isNew ? t('add') : t('edit'), fields).then(async (values) => {
      if (!values || !values.itemNo) return;
      values.available = Number(values.available) || 0;
      values.inTransit = Number(values.inTransit) || 0;
      if (isNew) {
        const existing = await api('get_all_stock', values.itemNo);
        if (existing.some((x) => x.itemNo === values.itemNo)) { toast(tf('dupNo', { no: values.itemNo })); return; }
      }
      invalidateCtx();
      await api('upsert_stock', JSON.stringify(values));
      renderInventory();
      saved();
    });
  }

  // ================= reps page =================
  async function renderReps() {
    const tk = 'reps:' + S.dataToken;
    if (await activateIfFresh(t('tabReps'), tk)) { S.sheetCtx = S.sheetCtxCache[t('tabReps')] || null; return; }
    const reps = await api('get_all_reps');
    const s1 = reps.s1 || [], s2 = reps.s2 || [];
    const values = [[t('repsS1'), '', '', '']];
    const ctxRows = [];
    values.push(['', t('repCode'), t('repName'), t('repGroup')]);
    s1.forEach((r) => { ctxRows.push({ kind: 's1', rec: r }); values.push(['', r.code, r.name, r.group || '']); });
    values.push(['', '', '', '']);
    values.push([t('repsS2'), '', '', '']);
    values.push(['', t('repCode'), t('repName'), t('repGroup')]);
    s2.forEach((r) => { ctxRows.push({ kind: 's2', rec: r }); values.push(['', r.code, r.name, r.group || '']); });
    const n = values.length;
    S.sheetCtx = { kind: 'reps', rows: ctxRows, startRow: 2, colCount: 4 };
    S.sheetCtxCache[t('tabReps')] = S.sheetCtx;
    const bold = [rngA(1, 1, 1, 4), rngA(2, 1, 2, 4), rngA(n - s2.length - 1, 1, n - s2.length - 1, 4), rngA(n - s2.length, 1, n - s2.length, 4)];
    const bg = [{ range: rngA(1, 1, 1, 4), color: '#d9e6f3' }, { range: rngA(n - s2.length - 1, 1, n - s2.length - 1, 4), color: '#d9e6f3' }];
    await renderSpec({ sheet: t('tabReps'), values, bold, bg, freeze: { rows: 0, cols: 0 } });
    S.sheetToken[t('tabReps')] = tk;
    S.lastHLRow = null;
    // 网格映射：行号 → {kind, code, name, group}（直接编辑同步的基线；B=code C=name D=group）
    S.repSheetName = t('tabReps');
    S.repGrid = {};
    let ri = 0;
    ctxRows.forEach((e) => {
      while (ri < values.length && String(values[ri] ? values[ri][1] : '') !== String(e.rec.code)) ri++;
      S.repGrid[ri] = { kind: e.kind, code: e.rec.code, name: e.rec.name || '', group: e.rec.group || '' };
      ri++;
    });
  }

  // 分组表直接编辑同步：定时对比表格 C/D 列与基线，差异写回数据库（upsert=锁定行）
  // 用户在 Excel 格子里直接改 姓名/组别 后约 2.5s 自动保存；改 code 列不生效（key 不可变，用编辑按钮改）
  function syncRepsFromGrid() {
    let saved = 0;
    if (!costVisible() || S.nav !== 'reps' || S.isRendering || !S.u || !S.repGrid) return saved;
    let wb = null;
    try { wb = S.u.api.getActiveWorkbook(); } catch (e) { return; }
    const sh = wb.getSheets().find((x) => { try { return x.getSheetName() === S.repSheetName; } catch (e2) { return false; } });
    if (!sh) return;
    Object.keys(S.repGrid).forEach((k) => {
      const ri = Number(k);
      const e = S.repGrid[k];
      let row = null;
      try { row = sh.getRange(ri, 1, 1, 3).getValues()[0]; } catch (err) { return; }
      const code = String(row[0] === null || row[0] === undefined ? '' : row[0]).trim();
      if (code !== String(e.code)) return;   // 行被用户挪动/清空 code：不同步，避免写错行
      const name = String(row[1] === null || row[1] === undefined ? '' : row[1]).trim();
      const group = String(row[2] === null || row[2] === undefined ? '' : row[2]).trim();
      if (name !== e.name || group !== e.group) {
        trackSave(api('upsert_rep', e.kind, JSON.stringify({ code: e.code, name, group })));
        e.name = name;
        e.group = group;   // 更新基线，避免重复写库
        saved++;
      }
    });
    return saved;
  }

  function editRep(kind, row) {
    const isNew = !row;
    const base = row || { code: '', name: '', group: '' };
    const groupHint = kind === 's1' ? 'House / OutSide / Other Branch / IC' : 'inside / (blank)';
    const fields = [
      { key: 'code', label: t('repCode'), value: base.code },
      { key: 'name', label: t('repName'), value: base.name },
      { key: 'group', label: t('repGroup'), value: base.group, placeholder: groupHint },
    ];
    formModal((isNew ? t('add') : t('edit')) + ' — ' + (kind === 's1' ? t('repsS1') : t('repsS2')), fields).then(async (values) => {
      if (!values || !values.code) return;
      values.code = values.code.trim().toUpperCase();
      if (isNew) {
        const reps = await api('get_all_reps');
        if (reps[kind].some((x) => x.code.toUpperCase() === values.code)) { toast(tf('dupNo', { no: values.code })); return; }
      }
      invalidateCtx();
      await api('upsert_rep', kind, JSON.stringify(values));
      renderReps();
      saved();
    });
  }

  // ================= import history page（HTML 多选列表，不进工作表） =================
  async function renderHistory() {
    const batches = await api('list_batches');
    const wrap = document.getElementById('hist-table');
    if (!wrap) return;
    let html = '<table class="grid"><thead><tr>' +
      '<th style="width:36px"><input type="checkbox" id="hist-all"></th>' +
      '<th>' + esc(t('histMonth')) + '</th><th>' + esc(t('histFile')) + '</th><th>' + esc(t('histRows')) + '</th><th>' + esc(t('histTime')) + '</th></tr></thead><tbody>';
    if (!batches.length) html += '<tr><td colspan="5" class="empty">' + esc(t('histEmpty')) + '</td></tr>';
    batches.forEach((b) => {
      html += '<tr><td><input type="checkbox" data-hsel="' + esc(b.id) + '"' + (S.histSel.has(b.id) ? ' checked' : '') + '></td>' +
        '<td><b>' + esc(b.month) + '</b></td><td>' + esc(b.file) + '</td><td class="num">' + esc(b.rows) + '</td><td>' + esc(b.importedAt) + '</td></tr>';
    });
    html += '</tbody></table>';
    wrap.innerHTML = html;
    wrap.querySelectorAll('[data-hsel]').forEach((cb) => (cb.onchange = () => {
      if (cb.checked) S.histSel.add(cb.dataset.hsel); else S.histSel.delete(cb.dataset.hsel);
      updateHistDelBtn();
    }));
    const all = wrap.querySelector('#hist-all');
    if (all) all.onchange = () => {
      wrap.querySelectorAll('[data-hsel]').forEach((cb) => {
        cb.checked = all.checked;
        if (all.checked) S.histSel.add(cb.dataset.hsel); else S.histSel.delete(cb.dataset.hsel);
      });
      updateHistDelBtn();
    };
    updateHistDelBtn();
  }

  function updateHistDelBtn() {
    const btn = document.getElementById('hist-del-sel');
    if (!btn) return;
    const n = S.histSel.size;
    btn.textContent = tf('delSelected') + (n ? ' (' + n + ')' : '');
    btn.disabled = n === 0;
  }

  async function deleteSelectedBatches() {
    const batches = (await api('list_batches')).filter((b) => S.histSel.has(b.id));
    if (!batches.length) return;
    const rows = batches.reduce((sum, b) => sum + b.rows, 0);
    const ok = await confirmModal(tf('delSelectedConfirm', { k: batches.length, n: rows }));
    if (!ok) return;
    for (const b of batches) await api('delete_batch', b.id);
    invalidateCtx();
    await invalidateStatOutputs();
    S.histSel.clear();
    renderHistory();
    if (S.nav === 'sales') renderSales();
    toast(tf('delMonthDone', { m: batches.map((b) => b.month).join(', ') }));
    saved();
  }

  // ================= import flows =================
  function normHeader(s) {
    return String(s || '').toLowerCase().replace(/[\s._#]/g, '').trim();
  }
  const HEADER_MAP = {
    itemno: 'itemNo', item: 'itemNo', 品名: 'itemNo',
    description: 'description', 描述: 'description',
    color: 'color', 颜色: 'color',
    name: 'name',
    type: 'type', 类别: 'type',
    category: 'category',
    producttype: 'productType',
    solddate: 'soldDate', 销售日期: 'soldDate', 日期: 'soldDate',
    invoiceno: 'invoice', invoice: 'invoice', 发票号码: 'invoice', 发票号: 'invoice',
    itemledgerno: 'ledger',
    customerid: 'custNo', customerno: 'custNo', 客户编号: 'custNo', 客户号: 'custNo',
    customername: 'custName', 客户名称: 'custName',
    repcode: 'rep1', rep1code: 'rep1', rep1: 'rep1', 销售1: 'rep1', 销售1代码: 'rep1', sales1: 'rep1',
    salesperson: 'rep1', salespersoncode: 'rep1', salesperson1: 'rep1',
    repname: 'rep1Name', 销售1名称: 'rep1Name',
    rep2: 'rep2', rep2code: 'rep2', 销售2: 'rep2', 销售2代码: 'rep2', sales2: 'rep2',
    部门1: 'rep1Group', department1: 'rep1Group', dept1: 'rep1Group',
    部门2: 'rep2Group', department2: 'rep2Group', dept2: 'rep2Group',
    rep2name: 'rep2Name',
    soldqty: 'qty', qty: 'qty', 数量: 'qty',
    salesamount: 'amount', amount: 'amount', 金额: 'amount',
    costamount: 'cost', cost: 'cost', 成本: 'cost', 成本金额: 'cost', 成本额: 'cost', unitcost: 'cost', 单位成本: 'cost',
  };
  // 中文 14 列模板：部门1/部门2 即销售1/销售2 的分组（Group 信息随销售文件走）
  const ZH14_POSITIONAL = ['type', 'itemNo', 'description', 'color', 'soldDate', 'invoice', 'custNo', 'custName', 'rep1', 'rep1Group', 'rep2', 'rep2Group', 'qty', 'amount'];

  function parseSalesRows(headers, rows) {
    const zh = (headers || []).some((h) => String(h).trim() === '类别');
    let keys = (headers || []).map((h) => {
      const n = normHeader(h);
      if (HEADER_MAP[n]) return HEADER_MAP[n];
      const zhKey = String(h || '').trim();
      if (HEADER_MAP[zhKey]) return HEADER_MAP[zhKey];
      return null;
    });
    if (zh && keys.filter(Boolean).length < 8) keys = ZH14_POSITIONAL;
    if (keys.filter(Boolean).length < 8) {
      // 兜底：英文 19 列格式位置映射（含销售姓名与第 19 列 Cost Amount 成本）
      keys = ['itemNo', 'description', 'color', 'name', 'type', 'category', 'productType', 'soldDate', 'invoice', 'ledger', 'custNo', 'custName', 'rep1', 'rep1Name', 'rep2', 'rep2Name', 'qty', 'amount', 'cost'];
    }
    const recs = [];
    (rows || []).forEach((arr, idx) => {
      if (!arr || arr.every((v) => v === null || v === undefined || v === '')) return;
      const rec = { name: '', category: '', productType: '', ledger: '', rep1Name: '', rep2Name: '', rep1Group: '', rep2Group: '', cost: 0 };
      arr.forEach((v, ci) => {
        const k = keys[ci];
        if (!k || k[0] === '_') return;
        if (k === 'qty' || k === 'amount' || k === 'cost') rec[k] = Number(STATS.helpers.amount(v)) || 0;
        else if (k === 'soldDate') rec[k] = STATS.helpers.date(v);
        else rec[k] = String(v === null || v === undefined ? '' : v).trim();
      });
      if (rec.itemNo || rec.invoice || rec.amount) recs.push(rec);
    });
    // 表头映射路径下，记录没识别、但列里有实际数据的表头（位置兜底路径每列都有归属，自然为空）——
    // 导入确认框里提示出来，避免新表头对不上时静默丢列
    recs.unmapped = (headers || [])
      .map((h, ci) => ({ h: String(h || '').trim(), ci }))
      .filter((x) => x.h && !keys[x.ci] && (rows || []).some((r) => r && r[x.ci] !== null && r[x.ci] !== undefined && r[x.ci] !== ''))
      .map((x) => x.h);
    return recs;
  }

  // ================= 网页版文件导入/备份（远程版） =================
  // exe 版后端提供原生文件对话框（import_*_file / export_backup / backup_restore），
  // 远程版 HTTP 后端没有这些方法；先调 api()（exe 版 / 演示模式走原生对话框或 MOCK 数据），
  // 后端报 unknown method 时再走浏览器本地选文件 + SheetJS 解析，返回与后端一致的形状，
  // 下游流程（parseSalesRows / apply_*_import）不用改。
  // 解析逻辑与 ui/webimport.js 同源；内联于此，避免依赖 index.html 另行加载它。
  async function importFileWithFallback(apiMethod, pickAndParse) {
    let res = null;
    try { res = await api(apiMethod); } catch (e) { res = null; }
    if (res && res.ok) return res;
    // 取消 / 读取错误等：按原流程提示，不再弹第二次选文件框
    if (res && res.error && !/unknown method/i.test(String(res.error))) return res;
    return pickAndParse();
  }

  function webImportNormHeader(h) {
    return String(h == null ? '' : h).toLowerCase().replace(/[\s_\-]+/g, '');
  }

  function webImportPickFiles(accept, multiple) {
    return new Promise((resolve) => {
      const inp = document.createElement('input');
      inp.type = 'file';
      inp.accept = accept || '.xlsx,.xls,.csv';
      inp.multiple = !!multiple;
      inp.onchange = () => resolve(Array.from(inp.files || []));
      // 用户取消选择时也 resolve，避免永远挂起
      inp.addEventListener('cancel', () => resolve([]));
      setTimeout(() => resolve([]), 120000); // 2 分钟兜底
      inp.click();
    });
  }

  function webImportReadWorkbook(file) {
    return new Promise((resolve, reject) => {
      const rd = new FileReader();
      rd.onload = (e) => {
        try {
          const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
          resolve(wb);
        } catch (err) { reject(err); }
      };
      rd.onerror = () => reject(rd.error || new Error('read error'));
      rd.readAsArrayBuffer(file);
    });
  }

  function webImportSheetToHeadersRows(ws) {
    // header:1 取原始二维数组；第一行非空的作为表头
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
    if (!aoa.length) return { headers: [], rows: [] };
    let hi = 0;
    while (hi < aoa.length && aoa[hi].every((v) => v === '' || v === null || v === undefined)) hi++;
    if (hi >= aoa.length) return { headers: [], rows: [] };
    const headers = aoa[hi].map((v) => String(v == null ? '' : v).trim());
    const rows = aoa.slice(hi + 1).map((r) =>
      headers.map((_, ci) => {
        const v = r[ci];
        if (v instanceof Date) {
          // 日期转 YYYY-MM-DD 文本，保持与原 Python 行为一致
          const m = String(v.getMonth() + 1).padStart(2, '0'), d = String(v.getDate()).padStart(2, '0');
          return v.getFullYear() + '-' + m + '-' + d;
        }
        return v === null || v === undefined ? '' : v;
      })
    );
    return { headers, rows };
  }

  async function webImportReadSalesFiles(files) {
    const out = { ok: true, files: [] };
    for (const f of files) {
      try {
        const wb = await webImportReadWorkbook(f);
        // 销售文件取第一张表
        const ws = wb.Sheets[wb.SheetNames[0]];
        const hr = webImportSheetToHeadersRows(ws);
        out.files.push({ fileName: f.name, headers: hr.headers, rows: hr.rows, ok: true });
      } catch (e) {
        out.files.push({ fileName: f.name, ok: false, error: String((e && e.message) || e) });
      }
    }
    return out;
  }

  // 表头映射：客户 / 库存 / 销售员
  const WEBIMPORT_CUST_MAP = {
    no: 'no', customerno: 'no', customerid: 'no', 客户编号: 'no', 客户号: 'no',
    name: 'name', customername: 'name', 客户名称: 'name',
    city: 'city', 城市: 'city',
    state: 'state', 州: 'state', 省: 'state',
    phone: 'phone', 电话: 'phone',
    contact: 'contact', 联系人: 'contact',
    type: 'type', customertype: 'type', 客户类型: 'type',
    email: 'email', 邮箱: 'email',
    sales1: 'sales1', rep1: 'sales1', 销售1: 'sales1', 销售员: 'sales1',
    created: 'created', creationdate: 'created', 开户日期: 'created', 创建日期: 'created',
  };
  const WEBIMPORT_STOCK_MAP = {
    itemno: 'itemNo', item: 'itemNo', 品名: 'itemNo',
    description: 'description', 描述: 'description',
    available: 'available', availableqty: 'available', 库存: 'available',
    intransit: 'inTransit', intransitqty: 'inTransit', 在途: 'inTransit',
  };
  const WEBIMPORT_REP_MAP = {
    code: 'code', repcode: 'code', 销售代码: 'code', 代码: 'code',
    name: 'name', repname: 'name', 姓名: 'name', 名称: 'name',
    group: 'group', 部门: 'group', 分组: 'group',
  };

  function webImportMapRows(headers, rows, hmap, numKeys) {
    const keys = headers.map((h) => hmap[webImportNormHeader(h)] || hmap[String(h).trim()] || null);
    const recs = [];
    rows.forEach((arr) => {
      if (!arr || arr.every((v) => v === '' || v === null || v === undefined)) return;
      const rec = {};
      arr.forEach((v, ci) => {
        const k = keys[ci];
        if (!k) return;
        if ((numKeys || []).includes(k)) {
          const n = Number(String(v).replace(/[$,]/g, ''));
          rec[k] = isNaN(n) ? 0 : n;
        } else {
          rec[k] = String(v === null || v === undefined ? '' : v).trim();
        }
      });
      recs.push(rec);
    });
    return recs;
  }

  function webImportFindSheet(wb, names) {
    const lower = wb.SheetNames.map((s) => String(s).toLowerCase());
    for (const n of names) {
      const i = lower.findIndex((s) => s.indexOf(n) >= 0);
      if (i >= 0) return wb.Sheets[wb.SheetNames[i]];
    }
    return null;
  }

  async function webImportReadMasterFile(file) {
    if (!file) return { ok: false, error: 'cancelled' };
    const wb = await webImportReadWorkbook(file);
    const out = { ok: true, customers: [], stock: [], reps: { s1: [], s2: [] } };
    // 客户表
    let ws = webImportFindSheet(wb, ['customer', '客户']) || wb.Sheets[wb.SheetNames[0]];
    let hr = webImportSheetToHeadersRows(ws);
    out.customers = webImportMapRows(hr.headers, hr.rows, WEBIMPORT_CUST_MAP);
    // 库存表
    ws = webImportFindSheet(wb, ['stock', 'stone', '库存']);
    if (ws) { hr = webImportSheetToHeadersRows(ws); out.stock = webImportMapRows(hr.headers, hr.rows, WEBIMPORT_STOCK_MAP, ['available', 'inTransit']); }
    // 销售员表(s1/s2)
    const s1ws = webImportFindSheet(wb, ['sales1', 'rep1', '销售1']);
    const s2ws = webImportFindSheet(wb, ['sales2', 'rep2', '销售2']);
    if (s1ws) { hr = webImportSheetToHeadersRows(s1ws); out.reps.s1 = webImportMapRows(hr.headers, hr.rows, WEBIMPORT_REP_MAP).map((r) => ({ kind: 's1', ...r })); }
    if (s2ws) { hr = webImportSheetToHeadersRows(s2ws); out.reps.s2 = webImportMapRows(hr.headers, hr.rows, WEBIMPORT_REP_MAP).map((r) => ({ kind: 's2', ...r })); }
    return out;
  }

  async function webImportReadCustomersFile(file) {
    if (!file) return { ok: false, error: 'cancelled' };
    const wb = await webImportReadWorkbook(file);
    const hr = webImportSheetToHeadersRows(wb.Sheets[wb.SheetNames[0]]);
    return { ok: true, customers: webImportMapRows(hr.headers, hr.rows, WEBIMPORT_CUST_MAP) };
  }

  async function webImportReadStockFile(file) {
    if (!file) return { ok: false, error: 'cancelled' };
    const wb = await webImportReadWorkbook(file);
    const hr = webImportSheetToHeadersRows(wb.Sheets[wb.SheetNames[0]]);
    return { ok: true, stock: webImportMapRows(hr.headers, hr.rows, WEBIMPORT_STOCK_MAP, ['available', 'inTransit']) };
  }

  async function webImportPickAndParseSales() {
    await ensureXlsxLib();
    const files = await webImportPickFiles('.xlsx,.xls,.csv', true);
    if (!files.length) return { ok: false, error: 'cancelled' };
    return webImportReadSalesFiles(files);
  }

  async function webImportPickAndParseMaster() {
    await ensureXlsxLib();
    const files = await webImportPickFiles('.xlsx,.xls,.csv', false);
    if (!files.length) return { ok: false, error: 'cancelled' };
    return webImportReadMasterFile(files[0]);
  }

  async function webImportPickAndParseCustomers() {
    await ensureXlsxLib();
    const files = await webImportPickFiles('.xlsx,.xls,.csv', false);
    if (!files.length) return { ok: false, error: 'cancelled' };
    return webImportReadCustomersFile(files[0]);
  }

  async function webImportPickAndParseStock() {
    await ensureXlsxLib();
    const files = await webImportPickFiles('.xlsx,.xls,.csv', false);
    if (!files.length) return { ok: false, error: 'cancelled' };
    return webImportReadStockFile(files[0]);
  }

  // 备份：下载整库 / 上传整库（后端 /api/backup_download、/api/backup_upload 已实现）
  async function webImportDownloadBackup() {
    const resp=await fetch('/api/backup_download',{headers:{'X-Cost-Token':costVisible()?costToken:''}});
    if(!resp.ok)throw new Error('请先通过 文件 → 解锁成本');
    const url=URL.createObjectURL(await resp.blob()),a=document.createElement('a');
    a.href=url;a.download='bcs_backup.db';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
  }

  function webImportPickBackupFile() {
    return new Promise((resolve) => {
      const inp = document.createElement('input');
      inp.type = 'file';
      inp.accept = '.db';
      inp.onchange = () => resolve((inp.files && inp.files[0]) || null);
      inp.addEventListener('cancel', () => resolve(null));
      inp.click();
    });
  }

  async function webImportUploadBackupFile(file) {
    const buf = await file.arrayBuffer();
    const resp = await fetch('/api/backup_upload', {
      method: 'POST',
      headers: { 'Content-Type': 'application/octet-stream', 'X-Cost-Token':costVisible()?costToken:'' },
      body: buf,
    });
    return resp.json();
  }

  async function importSalesFlow() {
    if(!costVisible()){toast('请先解锁管理权限');return;}
    let res;
    try { res = await importFileWithFallback('import_sales_files', webImportPickAndParseSales); } catch (e) { toast('Import error: ' + e); return; }
    if (!res || !res.ok) { toast((res && res.error) || 'Import cancelled'); return; }
    // 按月聚合：优先取每条记录销售日期的月份，取不到时回退文件名里的 YYYYMM
    const months = {};   // month → {records, file}
    let fileCount = 0;
    const unmappedNotes = [];   // 每个文件没识别的表头（确认框里提示，避免静默丢列）
    for (const f of (res.files || [])) {
      if (!f.ok) { toast((f.fileName || '') + ': ' + (f.error || 'read error')); continue; }
      const recs = parseSalesRows(f.headers || [], f.rows || []);
      if ((recs.unmapped || []).length) {
        unmappedNotes.push('<b>' + esc(f.fileName || '') + '</b>: ' + recs.unmapped.map(esc).join(', '));
      }
      if (!recs.length) continue;
      fileCount++;
      const fm = String(f.fileName || '').match(/(\d{6})/);
      const fileMonth = fm ? fm[1].slice(0, 4) + '-' + fm[1].slice(4, 6) : '';
      for (const r of recs) {
        let m = String(r.soldDate || '').slice(0, 7);
        if (!/^\d{4}-\d{2}$/.test(m)) m = fileMonth;
        if (!/^\d{4}-\d{2}$/.test(m)) continue;
        (months[m] = months[m] || { records: [], file: f.fileName }).records.push(r);
      }
    }
    const monthKeys = Object.keys(months).sort();
    if (!fileCount || !monthKeys.length) { toast(t('noDataShort')); return; }
    const existing = (await api('list_batches')).filter((b) => monthKeys.includes(b.month));
    // 汇总确认框：列出识别到的月份与行数，勾选是否替换已有月份
    const body = document.createElement('div');
    body.className = 'form';
    const listHtml = monthKeys.map((m) =>
      '<tr><td>' + esc(m) + '</td><td class="num">' + months[m].records.length + '</td><td>' + esc(months[m].file) + '</td></tr>').join('');
    body.innerHTML =
      '<div class="confirm-text">' + tf('importFilesFound', { n: monthKeys.length, f: fileCount }) + '</div>' +
      (unmappedNotes.length ? '<div class="form-row full" style="color:#b26a00">⚠ ' + esc(t('unmappedCols')) +
        '<br>' + unmappedNotes.join('<br>') + '</div>' : '') +
      '<div class="table-wrap" style="max-height:220px;margin:6px 0 10px">' +
      '<table class="grid"><thead><tr><th>' + esc(t('colMonth')) + '</th><th>' + esc(t('histRows')) + '</th><th>' + esc(t('histFile')) + '</th></tr></thead><tbody>' + listHtml + '</tbody></table></div>' +
      (existing.length ? '<div class="form-row full"><label class="chk"><input type="checkbox" id="imp-replace" checked> ' +
        esc(tf('monthExistsShort', { n: existing.length })) + '</label></div>' : '');
    showModal(t('dlgImportTitle'), body, [
      { label: t('cancel') },
      {
        label: t('confirm'), primary: true, onClick: async (close) => {
          const replaceExisting = !!body.querySelector('#imp-replace')?.checked;
          close();
          let total = 0;
          const importedMonths = [];
          const results = [];
          for (const m of monthKeys) {
            const has = existing.some((b) => b.month === m);
            if (has && !replaceExisting) continue;
            const r = await api('apply_sales_import', JSON.stringify({ month: m, file: months[m].file, records: months[m].records }));
            results.push(r);
            invalidateCtx();
            if (r && r.ok) { total += r.rows; importedMonths.push(m); }
          }
          if (total) await invalidateStatOutputs();
          S.histSel.clear();
          if(S.nav==='history')await renderHistory();
          else if(S.nav==='sales')await renderSales(true);
          else navigate(S.nav);
          const synced = results.reduce((acc, r) => ({ s1: acc.s1 + ((r && r.repsSynced) ? r.repsSynced.s1 : 0), s2: acc.s2 + ((r && r.repsSynced) ? r.repsSynced.s2 : 0) }), { s1: 0, s2: 0 });
          toast(tf('importDone', { n: total, m: importedMonths.length > 1 ? importedMonths[0] + '~' + importedMonths[importedMonths.length - 1] : (importedMonths[0] || '-') }) +
            (synced.s1 + synced.s2 > 0 ? ' ' + tf('repsSynced', { s1: synced.s1, s2: synced.s2 }) : ''));
          saved();
        },
      },
    ]);
  }

  async function importMasterFlow() {
    let res;
    try { res = await importFileWithFallback('import_master_file', webImportPickAndParseMaster); } catch (e) { toast('Import error: ' + e); return; }
    if (!res || !res.ok) { toast((res && res.error) || 'Import cancelled'); return; }
    const ok = await confirmModal(tf('masterDone', {
      c: (res.customers || []).length,
      s: (res.stock || []).length,
      r: ((res.reps || {}).s1 || []).length + ((res.reps || {}).s2 || []).length,
    }) + ' ?', t('confirm'));
    if (!ok) return;
    await api('apply_master_import', JSON.stringify({ customers: res.customers, stock: res.stock, reps: res.reps }));
    renderTab(S.nav);
    toast(tf('masterDone', { c: (res.customers || []).length, s: (res.stock || []).length, r: ((res.reps || {}).s1 || []).length + ((res.reps || {}).s2 || []).length }));
    saved();
  }

  // 客户/库存单文件导入（表头自动识别）
  async function importCustomersFlow() {
    let res;
    try { res = await importFileWithFallback('import_customers_file', webImportPickAndParseCustomers); } catch (e) { toast('Import error: ' + e); return; }
    if (!res || !res.ok) { toast((res && res.error) || 'Import cancelled'); return; }
    const list = res.customers || [];
    if (!list.length) { toast(t('noDataShort')); return; }
    const ok = await confirmModal(tf('confirmReplaceN', { n: list.length }), t('confirm'));
    if (!ok) return;
    invalidateCtx();
    await api('apply_master_import', JSON.stringify({ customers: list }));
    renderTab('customers');
    saved();
    toast(tf('custImportDone', { n: list.length }));
  }

  async function importStockFlow() {
    let res;
    try { res = await importFileWithFallback('import_stock_file', webImportPickAndParseStock); } catch (e) { toast('Import error: ' + e); return; }
    if (!res || !res.ok) { toast((res && res.error) || 'Import cancelled'); return; }
    const list = res.stock || [];
    if (!list.length) { toast(t('noDataShort')); return; }
    const ok = await confirmModal(tf('confirmReplaceN', { n: list.length }), t('confirm'));
    if (!ok) return;
    invalidateCtx();
    await api('apply_master_import', JSON.stringify({ stock: list }));
    renderTab('inventory');
    saved();
    toast(tf('stockImportDone', { n: list.length }));
  }

  async function loadDemo() {    const demo = makeDemoData();
    await api('clear_all_data');
    for (const b of demo.batches) {
      const recs = demo.sales.filter((r) => r.batch === b.id);
      await api('apply_sales_import', JSON.stringify({ month: b.month, file: b.file, records: recs }));
    }
    await api('apply_master_import', JSON.stringify({ customers: demo.customers, stock: demo.stock, reps: demo.reps }));
    await invalidateStatOutputs();
    navigate('sales');
    toast(t('demoLoaded'));
    saved();
  }

  // ================= backup / restore =================
  async function exportBackup() {
    if(!await checkCostAccess()){toast('请先通过 文件 → 解锁成本');return;}
    const body = document.createElement('div');
    body.className = 'form';
    body.innerHTML =
      '<div class="form-row full"><label class="chk"><input type="radio" name="bkmode" value="full" checked> ' + esc(t('bkFull')) + '</label></div>' +
      '<div class="form-row full"><label class="chk"><input type="radio" name="bkmode" value="sales"> ' + esc(t('bkSales')) + '</label></div>' +
      '<div class="form-row" id="bk-sales-opts" style="display:none">' +
      '  <span>' + esc(t('bkMonthFrom')) + '</span><input type="month" id="bk-m-from">' +
      '  <span>' + esc(t('bkMonthTo')) + '</span><input type="month" id="bk-m-to">' +
      '  <label class="chk"><input type="checkbox" id="bk-withmaster" checked> ' + esc(t('bkWithMaster')) + '</label>' +
      '</div>' +
      '<div class="form-row full"><label class="chk"><input type="radio" name="bkmode" value="master"> ' + esc(t('bkMaster')) + '</label></div>';
    showModal(t('bkExportTitle'), body, [
      { label: t('cancel') },
      {
        label: t('btnGenerate'), primary: true, onClick: async (close) => {
          const mode = body.querySelector('input[name="bkmode"]:checked').value;
          const opts = { mode };
          if (mode === 'sales') {
            opts.monthStart = body.querySelector('#bk-m-from').value || '';
            opts.monthEnd = body.querySelector('#bk-m-to').value || '';
            opts.withMaster = body.querySelector('#bk-withmaster').checked;
          }
          close();
          try {
            const res = await api('export_backup', JSON.stringify(opts));
            if (res && res.ok) { toast(tf('backupDone', { p: res.path })); return; }
            if (res && res.error && !/unknown method/i.test(String(res.error))) {
              toast('Export failed: ' + res.error); return;
            }
            // 远程版后端无 export_backup：直接下载整库文件（/api/backup_download）。
            // 注意：远程版备份统一为整库下载，月份筛选选项暂不生效。
            await webImportDownloadBackup();
            toast(tf('backupDone', { p: 'bcs_backup.db' }));
          } catch (e) {
            // 网络异常等：同样走下载兜底（下载不经过 api()）
            await webImportDownloadBackup();
            toast(tf('backupDone', { p: 'bcs_backup.db' }));
          }
        },
      },
    ]);
    body.querySelectorAll('input[name="bkmode"]').forEach((r) => {
      r.addEventListener('change', () => {
        body.querySelector('#bk-sales-opts').style.display =
          body.querySelector('input[name="bkmode"]:checked').value === 'sales' ? 'flex' : 'none';
      });
    });
  }

  async function restoreBackup() {
    if(!await checkCostAccess()){toast('请先通过 文件 → 解锁成本');return;}
    let info;
    try { info = await api('backup_inspect'); } catch (e) { toast('Restore error: ' + e); return; }
    if (!info || !info.ok) {
      if (info && info.error && info.error !== 'cancelled') toast(tf('restoreBad', { e: info.error }));
      return;
    }
    // 远程版：后端 backup_inspect 返回整库统计（无快照文件信息），走备份文件上传恢复；
    // exe 版返回快照信息（含 months），走下面原流程。
    if (!Array.isArray(info.months)) {
      const file = await webImportPickBackupFile();
      if (!file) return;
      const ok = await confirmModal(t('bkRestoreTitle') + '：' + file.name + ' ？', t('bkExecute'));
      if (!ok) return;
      let up;
      try { up = await webImportUploadBackupFile(file); } catch (e) { toast('Restore failed: ' + e); return; }
      if (!up || !up.ok) { toast('Restore failed: ' + ((up && up.error) || '')); return; }
      // 还原后刷新：与原 replace 分支一致，并清空查询缓存
      LOOKUP = null;
      REPS_CACHE = null;
      const meta = await api('get_meta');
      LANG = meta.lang === 'zh' ? 'zh' : 'en';
      applyI18n();
      await invalidateStatOutputs();
      navigate('sales');
      toast(t('restoreDone'));
      saved();
      return;
    }
    const zh = LANG === 'zh';
    const body = document.createElement('div');
    body.className = 'form';
    const monthList = info.months.map((m) => m.month + ' (' + m.rows + ')').join(', ') || (zh ? '无' : 'none');
    body.innerHTML =
      '<div class="confirm-text">' + esc(t('bkContains')) + ': ' + esc(monthList) + ' — ' + info.salesTotal + ' rows; ' +
      esc(t('custNo')) + ' ' + info.customers + ' · ' + esc(t('invStock')) + ' ' + info.stock + ' · ' + esc(t('repCode')) + ' ' + info.reps + '</div>' +
      '<div class="form-row full"><label class="chk"><input type="radio" name="rsmode" value="merge" checked> ' + esc(t('bkModeMerge')) + '</label></div>' +
      '<div class="form-row" id="rs-merge-opts">' +
      '  <span>' + esc(t('bkMonthConflict')) + '</span><select id="rs-mc"><option value="replace">' + esc(t('cfReplace')) + '</option><option value="skip">' + esc(t('cfSkip')) + '</option></select>' +
      '  <span>' + esc(t('bkMasterConflict')) + '</span><select id="rs-mmc"><option value="overwrite">' + esc(t('cfOverwrite')) + '</option><option value="skip">' + esc(t('cfSkip')) + '</option></select>' +
      '</div>' +
      '<div class="form-row full"><label class="chk"><input type="radio" name="rsmode" value="replace"> ' + esc(t('bkModeReplace')) + '</label></div>';
    showModal(t('bkRestoreTitle'), body, [
      { label: t('cancel') },
      {
        label: t('bkExecute'), primary: true, onClick: async (close) => {
          const mode = body.querySelector('input[name="rsmode"]:checked').value;
          const payload = {
            path: info.path, mode,
            monthConflict: (body.querySelector('#rs-mc') || {}).value || 'skip',
            masterConflict: (body.querySelector('#rs-mmc') || {}).value || 'overwrite',
          };
          close();
          const res = await api('backup_restore', JSON.stringify(payload));
          if (!res || !res.ok) { toast('Restore failed: ' + (res && res.error)); return; }
          if (mode === 'replace') {
            const meta = await api('get_meta');
            LANG = meta.lang === 'zh' ? 'zh' : 'en';
            applyI18n();
            await invalidateStatOutputs();
            navigate('sales');
          } else {
            renderTab(S.nav);
          }
          toast(t('restoreDone'));
          saved();
        },
      },
    ]);
    body.querySelectorAll('input[name="rsmode"]').forEach((r) => {
      r.addEventListener('change', () => {
        body.querySelector('#rs-merge-opts').style.display =
          body.querySelector('input[name="rsmode"]:checked').value === 'merge' ? 'flex' : 'none';
      });
    });
  }

  // ================= i18n apply =================
  function applyI18n() {
    document.querySelectorAll('[data-i18n]').forEach((el) => (el.textContent = t(el.dataset.i18n)));
    document.querySelectorAll('[data-i18n-ph]').forEach((el) => (el.placeholder = t(el.dataset.i18nPh)));
    buildMenus();
    buildSidebar();
    if (S.nav.startsWith('r:')) renderStatCard();
    else renderTab(S.nav);
    document.documentElement.lang = LANG;
    try { window.dispatchEvent(new Event('langchanged')); } catch (e) { /* ignore */ }
  }

  // ================= toolbar bindings =================
  function isoDate(d) {
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }
  // TTM 口径：数据滞后一个月（月底/月初 5 号前导入）——上个月末往前 12 个月
  // 例：今天 2026-09-30 → TTM = 2025-09-01 ~ 2026-08-31
  function ttmRange() {
    const now = new Date();
    const pm = new Date(now.getFullYear(), now.getMonth() - 1, 1);            // 上个月
    const start = new Date(pm.getFullYear(), pm.getMonth() - 11, 1);          // 上上 11 个月的 1 号
    const end = new Date(pm.getFullYear(), pm.getMonth() + 1, 0);             // 上个月最后一天
    return { start: isoDate(start), end: isoDate(end) };
  }
  function collectSalesFilters() {
    const cn = $('#f-custname');
    S.salesFilters = {
      start: $('#f-start').value || '',
      end: $('#f-end').value || '',
      custNo: cn.dataset.custNo || '',
      custName: cn.value || '',
      itemNo: $('#f-itemno').value || '',
      description: $('#f-desc').value || '',
      type: $('#f-type').value || '',
      color: $('#f-color').value || '',
      invoice: $('#f-inv').value || '',
      rep1: $('#f-rep1').value || '',
    };
  }
  const SALES_FILTER_INPUTS = ['f-start', 'f-end', 'f-custname', 'f-itemno', 'f-desc', 'f-type', 'f-color', 'f-inv', 'f-rep1'];
  function bindFuzzyLookups() {
    // 客户名称：候选来自销售史客户清单（对应原 Search_List），点选后回填名称并记住编号做精确联动
    const cn = $('#f-custname');
    attachFuzzy(cn, {
      fetch: fuzzyCustomers,
      onPick: (it, input) => {
        if (it.data && it.data.no) input.dataset.custNo = it.data.no;
        else delete input.dataset.custNo;
      },
    });
    cn.addEventListener('input', () => { delete cn.dataset.custNo; });
    attachFuzzy($('#f-itemno'), { fetch: fuzzyItems });
    attachFuzzy($('#f-desc'), { fetch: fuzzyDescs });
    attachFuzzy($('#f-type'), { fetch: (text) => getLookup().then((lk) => fuzzySimple(lk.types, text)) });
    attachFuzzy($('#f-color'), { fetch: (text) => getLookup().then((lk) => fuzzySimple(lk.colors, text)) });
    attachFuzzy($('#f-inv'), { fetch: fuzzyInvoices });
    attachFuzzy($('#f-rep1'), { fetch: fuzzyRep1 });
  }
  function bindToolbars() {
    $('#sales-query').onclick = () => { collectSalesFilters(); S.salesPage = 0; renderSales(true); };
    $('#sales-reset').onclick = () => {
      SALES_FILTER_INPUTS.forEach((id) => { $('#' + id).value = ''; });
      delete $('#f-custname').dataset.custNo;
      $('#f-year').value = '';
      collectSalesFilters(); S.salesPage = 0; renderSales();
    };
    // 年份快筛：销售记录是分页列表，折叠分组不适用；按年份一键过滤（选项来自已导入月份）
    const yearSel = $('#f-year');
    (async () => {
      const years = new Set([String(new Date().getFullYear())]);
      try { (await api('list_batches')).forEach((b) => years.add(String(b.month || '').slice(0, 4))); } catch (e) { /* ignore */ }
      yearSel.innerHTML = '<option value="">' + esc(t('allYears')) + '</option>' +
        [...years].filter(Boolean).sort().reverse().map((y) => '<option value="' + y + '">' + y + '</option>').join('');
    })();
    yearSel.onchange = () => {
      const y = yearSel.value;
      $('#f-start').value = y ? y + '-01-01' : '';
      $('#f-end').value = y ? y + '-12-31' : '';
      collectSalesFilters(); S.salesPage = 0; renderSales();
    };
    $('#q-lastyear').onclick = () => {
      const y = new Date().getFullYear() - 1;
      $('#f-start').value = y + '-01-01'; $('#f-end').value = y + '-12-31';
    };
    $('#q-thisyear').onclick = () => {
      const y = new Date().getFullYear();
      $('#f-start').value = y + '-01-01'; $('#f-end').value = y + '-12-31';
    };
    $('#q-ttm').onclick = () => {
      const r = ttmRange();
      $('#f-start').value = r.start;
      $('#f-end').value = r.end;
    };
    SALES_FILTER_INPUTS.forEach((id) => {
      $('#' + id).addEventListener('keydown', (e) => { if (e.key === 'Enter') $('#sales-query').click(); });
    });
    bindFuzzyLookups();
    $('#sales-prev').onclick = () => { if (S.salesPage > 0) { S.salesPage--; renderSales(); } };
    $('#sales-next').onclick = () => { S.salesPage++; renderSales(); };
    // 每页数量可选（含"全部"）：选择即生效并记住偏好
    const sizeSel = $('#sales-pagesize');
    const SIZE_OPTS = [100, 200, 500, 1000, 2000, 5000, 0];   // 0=全部（快照建表后万行级可用）
    const fillSizes = () => {
      sizeSel.innerHTML = SIZE_OPTS.map((n) =>
        '<option value="' + n + '"' + (S.pageSize === n ? ' selected' : '') + '>' + (n === 0 ? t('allRows') : n) + '</option>').join('');
    };
    fillSizes();
    sizeSel.onchange = () => {
      S.pageSize = Number(sizeSel.value);
      S.salesPage = 0;
      try { api('set_pref', 'salesPageSize', String(S.pageSize)); } catch (e) { /* ignore */ }
      renderSales();
    };
    // 语言切换时刷新"全部"文案（applyI18n 末尾派发 langchanged）
    window.addEventListener('langchanged', fillSizes);
    let dashTm = null;
    $('#dash-search').addEventListener('input', (e) => {
      S.dash.query = e.target.value;
      clearTimeout(dashTm);
      dashTm = setTimeout(() => {
        if (S.nav === 'dashboard' && S.dash.cache) renderDashList().catch(() => { /* ignore */ });
      }, 300);
    });
    // 客户快速查询带上模糊下拉（与统计条件卡同一组件）：点候选直接选中该客户
    try {
      attachFuzzy($('#dash-search'), {
        fetch: fuzzyCustomersCode,
        onPick: (it, input) => {
          const no = String((it && it.v) || '').trim();
          if (!no) return;
          S.dash.customerNo = no;
          input.value = '';
          S.dash.query = '';
          if (S.nav === 'dashboard' && S.dash.cache) {
            renderDashList().catch(() => { /* ignore */ });
            renderDashAna();
            renderDashChart();
          }
        },
      });
    } catch (e) { /* ignore */ }
    $('#sales-import').onclick = importSalesFlow;
    let custTm = null;
    $('#cust-search').addEventListener('input', (e) => { S.custQuery = e.target.value; clearTimeout(custTm); custTm = setTimeout(renderCustomers, 300); });
    $('#cust-add').onclick = () => editCustomer();
    $('#cust-import').onclick = importCustomersFlow;
    // 客户页同步：只从报价文档 Customer 同步客户，全量替换本地客户表，不动库存/销售
    $('#cust-sync').onclick = async () => {
      const btn = $('#cust-sync');
      if (btn.disabled) return;
      btn.disabled = true;
      const old = btn.textContent;
      btn.textContent = '🔄 ' + t('syncing');
      try {
        const m = await api('sync_master_from_sheets', 'customers');
        if (!m || !m.ok) throw new Error((m && m.error) || 'sync failed');
        invalidateCtx();
        renderCustomers();
        toast(tf('syncCustomersDone', { c: m.customers || 0 }));
      } catch (e) {
        toast(t('syncFail') + ': ' + (e && e.message || e));
      } finally {
        btn.disabled = false;
        btn.textContent = old;
      }
    };
    $('#cust-edit').onclick = () => { const e = sheetSelEntry(); if (e) editCustomer(e); };
    $('#cust-del').onclick = async () => {
      const e = sheetSelEntry();
      if (!e) return;
      if (await confirmModal(tf('delConfirm', { name: e.no + ' ' + e.name }), t('delete'))) {
        invalidateCtx();
        await api('delete_customer', e.no);
        renderCustomers();
        saved();
      }
    };
    let invTm = null;
    $('#inv-search').addEventListener('input', (e) => { S.invQuery = e.target.value; clearTimeout(invTm); invTm = setTimeout(renderInventory, 300); });
    $('#inv-add').onclick = () => editStock();
    $('#inv-import').onclick = importStockFlow;
    // 库存页同步：只从报价文档 Pull 的 SL* 行同步库存，全量替换本地库存表，不动客户/销售
    $('#inv-sync').onclick = async () => {
      const btn = $('#inv-sync');
      if (btn.disabled) return;
      btn.disabled = true;
      const old = btn.textContent;
      btn.textContent = '🔄 ' + t('syncing');
      try {
        const m = await api('sync_master_from_sheets', 'stock');
        if (!m || !m.ok) throw new Error((m && m.error) || 'sync failed');
        invalidateCtx();
        renderInventory();
        toast(tf('syncStockDone', { st: m.stock || 0 }));
      } catch (e) {
        toast(t('syncFail') + ': ' + (e && e.message || e));
      } finally {
        btn.disabled = false;
        btn.textContent = old;
      }
    };
    $('#inv-edit').onclick = () => { const e = sheetSelEntry(); if (e) editStock(e); };
    $('#inv-del').onclick = async () => {
      const e = sheetSelEntry();
      if (!e) return;
      if (await confirmModal(tf('delConfirm', { name: e.itemNo + ' ' + e.description }), t('delete'))) {
        invalidateCtx();
        await api('delete_stock', e.itemNo);
        renderInventory();
        saved();
      }
    };
    $('#reps-add-s1').onclick = () => editRep('s1');
    $('#reps-add-s2').onclick = () => editRep('s2');
    $('#reps-edit').onclick = () => { const e = sheetSelEntry(); if (e && e.kind) editRep(e.kind, e.rec); };
    $('#reps-del').onclick = async () => {
      const e = sheetSelEntry();
      if (!e || !e.kind) return;
      if (await confirmModal(tf('delConfirm', { name: e.rec.code + ' ' + e.rec.name }), t('delete'))) {
        invalidateCtx();
        await api('delete_rep', e.kind, e.rec.code);
        renderReps();
        saved();
      }
    };
    $('#sales-edit').onclick = () => { const e = sheetSelEntry(); if (e) editSalesRow(e); };
    $('#hist-del-sel').onclick = deleteSelectedBatches;
  }

  // ================= boot =================
  let toolbarsBound = false;
  async function applyLocaleAndRender() {
    let meta = {};
    try {
      // 超时兜底：pywebview 桥接早期调用可能排队，3 秒未返回先用默认值渲染 UI（点击不致无响应）
      meta = (await Promise.race([
        api('get_meta'),
        new Promise((res) => setTimeout(() => res({}), 3000)),
      ])) || {};
    } catch (e) { /* ignore */ }
    LANG = meta.lang === 'zh' ? 'zh' : 'en';
    if (meta.prefs && meta.prefs.sidebarCollapsed === '1') S.sidebarCollapsed = true;
    if (meta.prefs && meta.prefs.navCollapsed) {
      try { S.navCollapsed = new Set(JSON.parse(meta.prefs.navCollapsed)); } catch (e) { /* ignore */ }
    }
    if (meta.prefs && meta.prefs.sheetLayouts) {
      try {
        const lay = JSON.parse(meta.prefs.sheetLayouts) || {};
        Object.keys(lay).forEach((k) => { if (Array.isArray(lay[k])) lay[k] = { w: lay[k], h: {} }; });   // 旧格式兼容
        S.sheetLayouts = lay;
      } catch (e) { S.sheetLayouts = {}; }
    }
    if(meta.prefs && meta.prefs.customerBatchLayouts) {try{S.customerBatchLayouts=JSON.parse(meta.prefs.customerBatchLayouts) || {};}catch(e){}}
    if (meta.prefs && meta.prefs.userStyles) {
      try { S.userStyles = JSON.parse(meta.prefs.userStyles) || {}; } catch (e) { S.userStyles = {}; }
    }
    if (meta.prefs && meta.prefs.salesPageSize !== undefined) {
      S.pageSize = Number(meta.prefs.salesPageSize);
      if (!(S.pageSize >= 0)) S.pageSize = 200;
    }
    if (!toolbarsBound) { toolbarsBound = true; bindToolbars(); }
    applyI18n();
    navigate(S.nav || 'sales');
    setStatus(t('appName') + ' · SQLite');
  }

  // 关闭确认（宿主 closing 事件调用）：原生 confirm 同步阻塞，宿主可可靠触发。
  // 确定 → 冲刷全部未保存修改后调用 close_app 真正退出；取消 → 留在软件（closing 已被宿主阻止）。
  window.__armClose = function () {
    // 应用内自绘确认框（原生 confirm 会显示 127.0.0.1 来源横幅，不友好）
    confirmModal(t('closeConfirm'), t('confirm')).then((ok) => {
      if (!ok) return;   // 取消：closing 已被宿主阻止，留在软件
      (async () => {
        try { await window.__flushMemory(); } catch (e) { /* ignore */ }
        try { await api('close_app'); } catch (e) { /* ignore */ }
      })();
    });
    return 'closing';
  };

  // 关闭软件前的统一冲刷：捕获停留页面的最后调整，等所有异步保存落库完成。
  // 宿主（app.py closing 事件）调用；平时保存时机=切换页面（navigate/renderSpec 钩子），不再需要定时器。
  window.__flushMemory = async function () {
    window.__flushDone = false;
    try {
      if (S.nav === 'reps' && S.repGrid) syncRepsFromGrid();
      syncDashMemory();
      if (S.u && S.u.api && S.lastRenderedSheet && !S.isRendering) {
        const wb = S.u.api.getActiveWorkbook();
        const prev = wb.getSheets().find((x) => { try { return x.getSheetName() === S.lastRenderedSheet; } catch (e) { return false; } });
        if (prev) syncSheetMemory(prev);
      }
      await Promise.allSettled(S.pendingSaves.slice());
    } catch (e) { /* ignore */ }
    window.__flushDone = true;
  };

  let bootRenderTimer = null;
  function boot() {
    // 异步异常不再静默：记录到 window.__errs 供排查（渲染/保存链路的 Promise 都不主动抛）
    try {
      window.addEventListener('unhandledrejection', (e) => {
        try { (window.__errs = window.__errs || []).push(String((e.reason && e.reason.message) || e.reason)); } catch (x) { /* ignore */ }
      });
    } catch (e) { /* ignore */ }
    // 页面可能在 pywebview 宿主里但宿主对象注入稍晚：mock 首渲染延迟 600ms 让位——
    // 桥接在窗口期内注入则只做一次真实渲染，避免 mock+真实两次全量渲染并发卡死主线程。
    bootRenderTimer = setTimeout(() => { bootRenderTimer = null; applyLocaleAndRender(); }, 600);
    if (USING_MOCK) {
      const iv = setInterval(() => {
        if (!USING_MOCK) { clearInterval(iv); return; }
        if (window.pywebview && window.pywebview.api && window.pywebview.api.get_meta) {
          clearInterval(iv);
          if (bootRenderTimer) { clearTimeout(bootRenderTimer); bootRenderTimer = null; }   // mock 首渲染未发生：取消，直接真实渲染
          BRIDGE = window.pywebview.api;
          USING_MOCK = false;
          applyLocaleAndRender();
        }
      }, 200);
      setTimeout(() => clearInterval(iv), 60000);
    }
  }

  document.addEventListener('DOMContentLoaded', boot);
  window.__app = { S, navigate, t, parseSalesRows, invalidateStatOutputs, api, setMockSalesFiles: (files) => { S.mockSalesFiles = files; } };
})();
