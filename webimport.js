/* 网页版文件导入/导出 —— SheetJS 本地解析，无需服务器参与。
 * 销售文件沿用 app.js 的 parseSalesRows/HEADER_MAP；主数据/客户/库存的表头映射在此定义。
 */
(function () {
  'use strict';

  function normHeader(h) {
    return String(h == null ? '' : h).toLowerCase().replace(/[\s_\-]+/g, '');
  }

  // ---------- 通用 ----------
  function pickFiles(accept, multiple) {
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

  function readWorkbook(file) {
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

  function sheetToHeadersRows(ws) {
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

  // ---------- 销售文件 ----------
  async function readSalesFiles(files) {
    const out = { ok: true, files: [] };
    for (const f of files) {
      try {
        const wb = await readWorkbook(f);
        // 销售文件取第一张表
        const ws = wb.Sheets[wb.SheetNames[0]];
        const hr = sheetToHeadersRows(ws);
        out.files.push({ fileName: f.name, headers: hr.headers, rows: hr.rows, ok: true });
      } catch (e) {
        out.files.push({ fileName: f.name, ok: false, error: String((e && e.message) || e) });
      }
    }
    return out;
  }

  // ---------- 表头映射:客户/库存/销售员 ----------
  const CUST_MAP = {
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
  const STOCK_MAP = {
    itemno: 'itemNo', item: 'itemNo', 品名: 'itemNo',
    description: 'description', 描述: 'description',
    available: 'available', availableqty: 'available', 库存: 'available',
    intransit: 'inTransit', intransitqty: 'inTransit', 在途: 'inTransit',
  };
  const REP_MAP = {
    code: 'code', repcode: 'code', 销售代码: 'code', 代码: 'code',
    name: 'name', repname: 'name', 姓名: 'name', 名称: 'name',
    group: 'group', 部门: 'group', 分组: 'group',
  };

  function mapRows(headers, rows, hmap, numKeys) {
    const keys = headers.map((h) => hmap[normHeader(h)] || hmap[String(h).trim()] || null);
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

  function findSheet(wb, names) {
    const lower = wb.SheetNames.map((s) => String(s).toLowerCase());
    for (const n of names) {
      const i = lower.findIndex((s) => s.indexOf(n) >= 0);
      if (i >= 0) return wb.Sheets[wb.SheetNames[i]];
    }
    return null;
  }

  async function readMasterFile(file) {
    if (!file) return { ok: false, error: 'cancelled' };
    const wb = await readWorkbook(file);
    const out = { ok: true, customers: [], stock: [], reps: { s1: [], s2: [] } };
    // 客户表
    let ws = findSheet(wb, ['customer', '客户']) || wb.Sheets[wb.SheetNames[0]];
    let hr = sheetToHeadersRows(ws);
    out.customers = mapRows(hr.headers, hr.rows, CUST_MAP);
    // 库存表
    ws = findSheet(wb, ['stock', 'stone', '库存']);
    if (ws) { hr = sheetToHeadersRows(ws); out.stock = mapRows(hr.headers, hr.rows, STOCK_MAP, ['available', 'inTransit']); }
    // 销售员表(s1/s2)
    const s1ws = findSheet(wb, ['sales1', 'rep1', '销售1']);
    const s2ws = findSheet(wb, ['sales2', 'rep2', '销售2']);
    if (s1ws) { hr = sheetToHeadersRows(s1ws); out.reps.s1 = mapRows(hr.headers, hr.rows, REP_MAP).map((r) => ({ kind: 's1', ...r })); }
    if (s2ws) { hr = sheetToHeadersRows(s2ws); out.reps.s2 = mapRows(hr.headers, hr.rows, REP_MAP).map((r) => ({ kind: 's2', ...r })); }
    return out;
  }

  async function readCustomersFile(file) {
    if (!file) return { ok: false, error: 'cancelled' };
    const wb = await readWorkbook(file);
    const hr = sheetToHeadersRows(wb.Sheets[wb.SheetNames[0]]);
    return { ok: true, customers: mapRows(hr.headers, hr.rows, CUST_MAP) };
  }

  async function readStockFile(file) {
    if (!file) return { ok: false, error: 'cancelled' };
    const wb = await readWorkbook(file);
    const hr = sheetToHeadersRows(wb.Sheets[wb.SheetNames[0]]);
    return { ok: true, stock: mapRows(hr.headers, hr.rows, STOCK_MAP, ['available', 'inTransit']) };
  }

  // ---------- 导出 ----------
  function exportXlsx(name, values) {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(values || [[]]);
    // 列宽自适应(按内容长度估算，上限 40)
    const colCount = Math.max.apply(null, [1].concat((values || []).map((r) => r.length)));
    ws['!cols'] = [];
    for (let c = 0; c < colCount; c++) {
      let w = 10;
      (values || []).forEach((r) => {
        const v = r[c];
        const len = String(v === null || v === undefined ? '' : v).length;
        if (len > w) w = Math.min(40, len + 2);
      });
      ws['!cols'].push({ wch: w });
    }
    XLSX.utils.book_append_sheet(wb, ws, String(name || 'Sheet1').slice(0, 31));
    XLSX.writeFile(wb, (name || 'export') + '.xlsx');
  }

  function downloadBackup() {
    const a = document.createElement('a');
    a.href = '/api/backup_download';
    a.download = '';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => a.remove(), 1000);
  }

  function uploadBackup() {
    return new Promise((resolve) => {
      const inp = document.createElement('input');
      inp.type = 'file';
      inp.accept = '.db';
      inp.onchange = async () => {
        const f = inp.files && inp.files[0];
        if (!f) { resolve({ ok: false, error: 'cancelled' }); return; }
        try {
          const buf = await f.arrayBuffer();
          const resp = await fetch('/api/backup_upload', {
            method: 'POST',
            headers: { 'Content-Type': 'application/octet-stream' },
            body: buf,
          });
          resolve(await resp.json());
        } catch (e) { resolve({ ok: false, error: String(e) }); }
      };
      inp.addEventListener('cancel', () => resolve({ ok: false, error: 'cancelled' }));
      inp.click();
    });
  }

  window.WebImport = {
    pickFiles, readSalesFiles, readMasterFile, readCustomersFile, readStockFile,
    exportXlsx, downloadBackup, uploadBackup,
  };
})();
