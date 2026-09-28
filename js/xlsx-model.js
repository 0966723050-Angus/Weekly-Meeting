// xlsx-model.js
// 直接讀寫 .xlsx 內部的 XML(以 fflate 解/壓縮),不經過 SheetJS 重新產生整本活頁簿,
// 因此未修改的工作表、樣式、下拉選單、條件式格式、註解等都原封不動保留。
//
// 列的新增/刪除採「資料區重寫」:只重寫被修改工作表的資料區(標題列之下到最後一筆資料),
// 每一筆資料連同它原本的儲存格樣式與列高一起搬移;資料區以外的列(預先格式化的空白列、
// 統計公式等)完全不動。這樣就不需要去位移合併儲存格、驗證範圍、條件式格式等設定。
(function () {
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
  const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);
  const DAY_MS = 86400000;

  const td = new TextDecoder('utf-8');
  const te = new TextEncoder();

  // ---------- 小工具 ----------
  function colToIdx(letters) {
    let n = 0;
    for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }
  function idxToCol(n) {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function splitRef(ref) {
    const m = /^([A-Z]+)(\d+)$/.exec(ref || '');
    return m ? { col: colToIdx(m[1]), row: Number(m[2]) } : null;
  }
  function kids(el, name) {
    return Array.from(el.childNodes).filter((n) => n.nodeType === 1 && n.localName === name);
  }
  function kid(el, name) { return kids(el, name)[0] || null; }
  function parseXml(str) {
    const doc = new DOMParser().parseFromString(str, 'application/xml');
    const err = doc.getElementsByTagName('parsererror')[0];
    if (err) throw new Error('XML 解析失敗:' + err.textContent.slice(0, 120));
    return doc;
  }
  function serialToISO(serial) {
    const d = new Date(EXCEL_EPOCH_UTC + Math.round(Number(serial)) * DAY_MS);
    return d.toISOString().slice(0, 10);
  }
  function isoToSerial(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
    if (!m) return null;
    return Math.round((Date.UTC(+m[1], +m[2] - 1, +m[3]) - EXCEL_EPOCH_UTC) / DAY_MS);
  }
  function resolvePath(base, target) {
    if (target.startsWith('/')) return target.slice(1);
    const parts = base.split('/'); parts.pop();
    for (const seg of target.split('/')) {
      if (seg === '..') parts.pop(); else if (seg !== '.') parts.push(seg);
    }
    return parts.join('/');
  }
  // 東亞全形字元以 2 個字寬計算(估算列高用)
  function visualLen(str) {
    let n = 0;
    for (const ch of str) n += ch.charCodeAt(0) > 0x2e80 ? 2 : 1;
    return n;
  }

  // ---------- 活頁簿 ----------
  class Workbook {
    constructor(bytes) {
      this.files = fflate.unzipSync(new Uint8Array(bytes));
      this._docs = {};
      this._loadWorkbook();
      this._loadSharedStrings();
      this._loadStyles();
    }

    text(path) {
      const f = this.files[path];
      return f ? td.decode(f) : null;
    }
    doc(path) {
      if (!this._docs[path]) {
        const t = this.text(path);
        if (t == null) throw new Error('找不到 ' + path);
        this._docs[path] = parseXml(t);
      }
      return this._docs[path];
    }

    _loadWorkbook() {
      const wb = this.doc('xl/workbook.xml');
      const rels = this.doc('xl/_rels/workbook.xml.rels');
      const relMap = {};
      for (const r of rels.getElementsByTagName('Relationship')) {
        relMap[r.getAttribute('Id')] = resolvePath('xl/workbook.xml', r.getAttribute('Target'));
      }
      this.sheetPaths = {};
      this.sheetNames = [];
      for (const s of wb.getElementsByTagNameNS(NS, 'sheet')) {
        const name = s.getAttribute('name');
        this.sheetNames.push(name);
        this.sheetPaths[name] = relMap[s.getAttributeNS(NS_R, 'id')];
      }
      this.definedNames = {};
      for (const d of wb.getElementsByTagNameNS(NS, 'definedName')) {
        this.definedNames[d.getAttribute('name')] = d.textContent;
      }
    }

    _loadSharedStrings() {
      this.sst = [];
      if (!this.files['xl/sharedStrings.xml']) return;
      const doc = this.doc('xl/sharedStrings.xml');
      for (const si of kids(doc.documentElement, 'si')) this.sst.push(richText(si));
    }

    _loadStyles() {
      // 每個儲存格樣式(cellXfs 索引)是否為日期 / 百分比格式
      this.xfKind = [];
      if (!this.files['xl/styles.xml']) return;
      const doc = this.doc('xl/styles.xml');
      const custom = {};
      const numFmts = doc.getElementsByTagNameNS(NS, 'numFmts')[0];
      if (numFmts) for (const f of kids(numFmts, 'numFmt')) custom[f.getAttribute('numFmtId')] = f.getAttribute('formatCode');
      const cellXfs = doc.getElementsByTagNameNS(NS, 'cellXfs')[0];
      if (!cellXfs) return;
      for (const xf of kids(cellXfs, 'xf')) {
        const id = Number(xf.getAttribute('numFmtId') || 0);
        let kind = '';
        if ((id >= 14 && id <= 22) || (id >= 45 && id <= 47)) kind = 'date';
        else if (id === 9 || id === 10) kind = 'percent';
        else if (custom[id]) {
          const code = custom[id].replace(/"[^"]*"|\[[^\]]*\]/g, '');
          if (/%/.test(code)) kind = 'percent';
          else if (/[ymd]/i.test(code)) kind = 'date';
        }
        this.xfKind.push(kind);
      }
    }

    // 解析 "專案主檔!$B$4:$B$102" 這種範圍,回傳其中所有非空值(去重)
    rangeValues(ref) {
      const m = /^'?(.+?)'?!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(ref);
      if (!m || !this.sheetPaths[m[1]]) return [];
      const c1 = colToIdx(m[2]), r1 = +m[3], c2 = m[4] ? colToIdx(m[4]) : c1, r2 = m[5] ? +m[5] : r1;
      const doc = this.doc(this.sheetPaths[m[1]]);
      const out = [];
      for (const row of doc.getElementsByTagNameNS(NS, 'row')) {
        const r = +row.getAttribute('r');
        if (r < r1 || r > r2) continue;
        for (const c of kids(row, 'c')) {
          const p = splitRef(c.getAttribute('r'));
          if (!p || p.col < c1 || p.col > c2) continue;
          const v = this.cellValue(c);
          if (v !== '' && v != null && !out.includes(String(v))) out.push(String(v));
        }
      }
      return out;
    }

    cellValue(c) {
      const t = c.getAttribute('t');
      const v = kid(c, 'v');
      if (t === 'inlineStr') { const is = kid(c, 'is'); return is ? richText(is) : ''; }
      if (!v) return '';
      const raw = v.textContent;
      if (t === 's') return this.sst[Number(raw)] ?? '';
      if (t === 'str' || t === 'e') return raw;
      if (t === 'b') return raw === '1' ? 'TRUE' : 'FALSE';
      const n = Number(raw);
      return Number.isFinite(n) ? n : raw;
    }

    // ---------- 讀取一個工作表成為可編輯的模型 ----------
    readSheet(cfg) {
      const path = this.sheetPaths[cfg.name];
      if (!path) throw new Error(`找不到工作表「${cfg.name}」`);
      const doc = this.doc(path);
      const sheetData = doc.getElementsByTagNameNS(NS, 'sheetData')[0];
      const rowEls = kids(sheetData, 'row');
      const rowByNum = new Map(rowEls.map((r) => [Number(r.getAttribute('r')), r]));
      const dataStart = cfg.headerRow + 1;
      const hardMax = cfg.maxRow || Infinity;

      // 欄位定義(由標題列而來)
      const headerEl = rowByNum.get(cfg.headerRow);
      const columns = [];
      if (headerEl) {
        for (const c of kids(headerEl, 'c')) {
          const p = splitRef(c.getAttribute('r'));
          const title = String(this.cellValue(c) || '').trim();
          if (p && title) columns.push({ idx: p.col, letter: idxToCol(p.col), title });
        }
      }

      // 欄寬(估算列高用)
      const widths = {};
      const cols = doc.getElementsByTagNameNS(NS, 'cols')[0];
      if (cols) for (const col of kids(cols, 'col')) {
        for (let i = +col.getAttribute('min'); i <= +col.getAttribute('max'); i++) widths[i] = Number(col.getAttribute('width')) || 9;
      }

      // 下拉選單(資料驗證)
      const lists = {};
      for (const dv of doc.getElementsByTagNameNS(NS, 'dataValidation')) {
        if (dv.getAttribute('type') !== 'list') continue;
        const f1 = kid(dv, 'formula1');
        if (!f1) continue;
        const formula = f1.textContent.trim();
        let values = [];
        if (/^".*"$/.test(formula)) values = formula.slice(1, -1).split(',').map((s) => s.trim()).filter(Boolean);
        else if (this.definedNames[formula]) values = this.rangeValues(this.definedNames[formula]);
        else values = this.rangeValues(formula.replace(/^=/, ''));
        for (const part of (dv.getAttribute('sqref') || '').split(/\s+/)) {
          const a = splitRef(part.split(':')[0]);
          const b = splitRef(part.split(':')[1] || part.split(':')[0]);
          if (!a || !b) continue;
          for (let ci = a.col; ci <= b.col; ci++) {
            // 同一欄以涵蓋資料起始列的驗證為準
            if (!lists[ci] || (a.row <= dataStart && b.row >= dataStart)) lists[ci] = values;
          }
        }
      }

      // 資料列
      const records = [];
      let lastValueRow = cfg.headerRow;
      for (const rowEl of rowEls) {
        const r = Number(rowEl.getAttribute('r'));
        if (r < dataStart || r > hardMax) continue;
        const cells = {};
        const vals = {};
        let hasValue = false;
        for (const c of kids(rowEl, 'c')) {
          const p = splitRef(c.getAttribute('r'));
          if (!p) continue;
          cells[p.col] = c;
          const v = this.cellValue(c);
          if (v !== '' && v != null) { vals[p.col] = v; hasValue = true; }
        }
        records.push({ id: 'r' + r, origRow: r, rowEl, cells, vals, dirty: new Set(), blank: !hasValue });
        if (hasValue) lastValueRow = r;
      }
      // 資料區內的空白列保留為「空白佔位」(不顯示),以免儲存時把下方資料往上擠;
      // 資料區之後的預先格式化空白列則不屬於資料
      while (records.length && records[records.length - 1].blank) records.pop();
      // 原始檔可能省略完全空白的 <row>,補上佔位使「第 i 筆 = 第 dataStart+i 列」
      for (let i = 0; i < records.length; i++) {
        const want = dataStart + i;
        if (records[i].origRow > want) {
          records.splice(i, 0, { id: 'r' + want, origRow: want, rowEl: null, cells: {}, vals: {}, dirty: new Set(), blank: true });
        }
      }

      // 欄位型別
      const styleOf = (col) => {
        for (const rec of records) if (rec.cells[col.idx]) return Number(rec.cells[col.idx].getAttribute('s') || 0);
        const blank = rowByNum.get(lastValueRow + 1);
        const c = blank && kids(blank, 'c').find((x) => splitRef(x.getAttribute('r')).col === col.idx);
        return c ? Number(c.getAttribute('s') || 0) : 0;
      };
      for (const col of columns) {
        const kind = this.xfKind[styleOf(col)] || '';
        if (/%/.test(col.title) || kind === 'percent') col.type = 'percent';
        else if (kind === 'date' || /(日|日期)$/.test(col.title) && !/週次/.test(col.title)) col.type = 'date';
        else col.type = 'text';
        if (col.type === 'date') {
          for (const rec of records) {
            if (typeof rec.vals[col.idx] === 'number') rec.vals[col.idx] = serialToISO(rec.vals[col.idx]);
          }
        }
        // 同欄中第一個日期/百分比格式的樣式,寫入值到「通用格式」儲存格時套用
        for (const rec of records) {
          const c = rec.cells[col.idx];
          const s = c && c.getAttribute('s');
          if (s != null && this.xfKind[Number(s)] === col.type) { col.fmtStyle = s; break; }
        }
        col.options = lists[col.idx] || null;
        col.width = widths[col.idx] || 9;
      }

      return {
        name: cfg.name, path, doc, cfg, columns, records, dataStart, lastValueRow,
        dirty: false, origDataXml: this.text(path),
      };
    }

    // ---------- 將模型寫回工作表 XML ----------
    writeSheet(model) {
      // 每次都從載入時的原始 XML 重新解析,重複儲存/儲存失敗重試都不會累積錯誤
      const { dataStart, records, columns } = model;
      const doc = parseXml(model.origDataXml);
      const sheetData = doc.getElementsByTagNameNS(NS, 'sheetData')[0];
      const rowEls = kids(sheetData, 'row');
      const rowByNum = new Map(rowEls.map((r) => [Number(r.getAttribute('r')), r]));
      const regionEnd = Math.max(model.lastValueRow, dataStart + records.length - 1);
      const colType = {};
      for (const c of columns) colType[c.idx] = c;

      // 空白列樣板:原本最後一筆資料的下一列(預先格式化好的空白列)
      const blankTpl = rowByNum.get(model.lastValueRow + 1)
        || (records.length ? records[records.length - 1].rowEl : null);
      const blankStyles = {};
      if (blankTpl) for (const c of kids(blankTpl, 'c')) blankStyles[splitRef(c.getAttribute('r')).col] = c.getAttribute('s');

      const built = [];
      for (let r = dataStart; r <= regionEnd; r++) {
        const rec = records[r - dataStart];
        if (!rec) {
          // 被刪除而空出來的列 → 用空白列樣板
          built.push(this._blankRow(doc, blankTpl, rowByNum.get(r), r, blankStyles));
          continue;
        }
        const srcRow = rec.rowEl || blankTpl;
        const row = srcRow ? srcRow.cloneNode(false) : doc.createElementNS(NS, 'row');
        row.setAttribute('r', String(r));
        const colSet = new Set([
          ...Object.keys(rec.cells).map(Number),
          ...Object.keys(rec.vals).map(Number),
          ...(rec.rowEl ? [] : Object.keys(blankStyles).map(Number)),
        ]);
        for (const ci of [...colSet].sort((a, b) => a - b)) {
          const ref = idxToCol(ci) + r;
          const orig = rec.cells[ci];
          if (orig && !rec.dirty.has(ci)) {
            const c = orig.cloneNode(true);
            c.setAttribute('r', ref);
            row.appendChild(c);
            continue;
          }
          const c = doc.createElementNS(NS, 'c');
          c.setAttribute('r', ref);
          let s = orig ? orig.getAttribute('s') : blankStyles[ci];
          const col = colType[ci];
          if (col && col.fmtStyle && rec.vals[ci] !== '' && rec.vals[ci] != null &&
              (col.type === 'date' || col.type === 'percent') && this.xfKind[Number(s)] !== col.type) s = col.fmtStyle;
          if (s != null) c.setAttribute('s', s);
          this._setValue(doc, c, rec.vals[ci], colType[ci]);
          row.appendChild(c);
        }
        if ((rec.dirty.size || !rec.rowEl) && !rec.blank) this._fitHeight(row, rec, columns);
        built.push(row);
      }

      // 以重建的列取代資料區
      for (const el of rowEls) {
        const r = Number(el.getAttribute('r'));
        if (r >= dataStart && r <= regionEnd) sheetData.removeChild(el);
      }
      const after = kids(sheetData, 'row').find((el) => Number(el.getAttribute('r')) > regionEnd) || null;
      for (const row of built) sheetData.insertBefore(row, after);

      // 更新 dimension
      const dim = doc.getElementsByTagNameNS(NS, 'dimension')[0];
      if (dim) {
        const [a, b] = (dim.getAttribute('ref') || 'A1').split(':');
        const end = splitRef(b || a);
        if (end && end.row < regionEnd) dim.setAttribute('ref', `${a}:${idxToCol(end.col)}${regionEnd}`);
      }

      // 部分瀏覽器序列化時會自帶 XML 宣告,部分不會 → 統一去掉後再補上標準宣告
      const xml = new XMLSerializer().serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, '');
      this.files[model.path] = te.encode(XML_DECL + xml);
    }

    _blankRow(doc, tpl, existing, r, blankStyles) {
      const row = tpl ? tpl.cloneNode(false) : doc.createElementNS(NS, 'row');
      row.setAttribute('r', String(r));
      if (!tpl && existing) for (const a of Array.from(existing.attributes)) if (a.name !== 'r') row.setAttributeNS(a.namespaceURI, a.name, a.value);
      for (const ci of Object.keys(blankStyles).map(Number).sort((a, b) => a - b)) {
        const c = doc.createElementNS(NS, 'c');
        c.setAttribute('r', idxToCol(ci) + r);
        if (blankStyles[ci] != null) c.setAttribute('s', blankStyles[ci]);
        row.appendChild(c);
      }
      return row;
    }

    _setValue(doc, c, val, col) {
      if (val === '' || val == null) return;
      if (col && col.type === 'date') {
        const serial = isoToSerial(val);
        if (serial != null) { const v = doc.createElementNS(NS, 'v'); v.textContent = String(serial); c.appendChild(v); return; }
      }
      if (typeof val === 'number' && Number.isFinite(val)) {
        const v = doc.createElementNS(NS, 'v'); v.textContent = String(val); c.appendChild(v); return;
      }
      c.setAttribute('t', 'inlineStr');
      const is = doc.createElementNS(NS, 'is');
      const t = doc.createElementNS(NS, 't');
      const str = String(val);
      if (/^\s|\s$|\n/.test(str)) t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
      t.textContent = str;
      is.appendChild(t);
      c.appendChild(is);
    }

    // 依文字長度與欄寬估算列高(自動換行),避免 Excel 開啟後文字被截掉
    _fitHeight(row, rec, columns) {
      let lines = 1;
      for (const col of columns) {
        const v = rec.vals[col.idx];
        if (v == null || v === '' || col.type !== 'text') continue;
        const per = Math.max(4, Math.floor(col.width * 1.1));
        let n = 0;
        for (const para of String(v).split('\n')) n += Math.max(1, Math.ceil(visualLen(para) / per));
        lines = Math.max(lines, n);
      }
      const ht = Math.min(409, Math.max(17.4, lines * 16.2 + 3));
      row.setAttribute('ht', String(Math.round(ht * 10) / 10));
      row.setAttribute('customHeight', '1');
    }

    // 新增一個工作表(預設隱藏),內容以 inline string 寫入;rows 為二維字串陣列
    addSheet(name, rows, { hidden = true, widths = [] } = {}) {
      if (this.sheetPaths[name]) throw new Error(`工作表「${name}」已存在`);
      let n = 1;
      while (this.files[`xl/worksheets/sheet${n}.xml`]) n++;
      const path = `xl/worksheets/sheet${n}.xml`;
      const x = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      const cols = widths.length
        ? '<cols>' + widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('') + '</cols>' : '';
      const body = rows.map((r, ri) => `<row r="${ri + 1}">` + r.map((v, ci) =>
        `<c r="${idxToCol(ci + 1)}${ri + 1}" t="inlineStr"><is><t xml:space="preserve">${x(v)}</t></is></c>`).join('') + '</row>').join('');
      const maxCol = Math.max(1, ...rows.map((r) => r.length));
      this.files[path] = te.encode(XML_DECL +
        `<worksheet xmlns="${NS}" xmlns:r="${NS_R}"><dimension ref="A1:${idxToCol(maxCol)}${Math.max(1, rows.length)}"/>` +
        `<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="16.2"/>${cols}` +
        `<sheetData>${body}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`);

      const ctPath = '[Content_Types].xml';
      const ct = this.doc(ctPath);
      const ov = ct.createElementNS('http://schemas.openxmlformats.org/package/2006/content-types', 'Override');
      ov.setAttribute('PartName', '/' + path);
      ov.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
      ct.documentElement.appendChild(ov);

      const relsPath = 'xl/_rels/workbook.xml.rels';
      const rels = this.doc(relsPath);
      const ids = Array.from(rels.getElementsByTagName('Relationship')).map((r) => Number((r.getAttribute('Id') || '').replace(/\D/g, '')) || 0);
      const rid = 'rId' + (Math.max(0, ...ids) + 1);
      const rel = rels.createElementNS('http://schemas.openxmlformats.org/package/2006/relationships', 'Relationship');
      rel.setAttribute('Id', rid);
      rel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet');
      rel.setAttribute('Target', `worksheets/sheet${n}.xml`);
      rels.documentElement.appendChild(rel);

      const wbPath = 'xl/workbook.xml';
      const wb = this.doc(wbPath);
      const sheets = wb.getElementsByTagNameNS(NS, 'sheets')[0];
      const maxId = Math.max(0, ...Array.from(sheets.getElementsByTagNameNS(NS, 'sheet')).map((s) => Number(s.getAttribute('sheetId')) || 0));
      const sh = wb.createElementNS(NS, 'sheet');
      sh.setAttribute('name', name);
      sh.setAttribute('sheetId', String(maxId + 1));
      if (hidden) sh.setAttribute('state', 'hidden');
      sh.setAttributeNS(NS_R, 'r:id', rid);
      sheets.appendChild(sh);

      for (const p of [ctPath, relsPath, wbPath]) {
        const xml = new XMLSerializer().serializeToString(this._docs[p]).replace(/^<\?xml[^>]*\?>\s*/, '');
        this.files[p] = te.encode(XML_DECL + xml);
      }
      this.sheetNames.push(name);
      this.sheetPaths[name] = path;
    }

    toBytes() {
      return fflate.zipSync(this.files, { level: 6 });
    }
  }

  function richText(el) {
    // <si>/<is> 內可能是 <t> 或多段 <r><t>;略過注音 <rPh>
    let s = '';
    for (const n of el.childNodes) {
      if (n.nodeType !== 1) continue;
      if (n.localName === 't') s += n.textContent;
      else if (n.localName === 'r') { const t = kid(n, 't'); if (t) s += t.textContent; }
    }
    return s;
  }

  // 本週的週次字串,格式同「使用說明」頁公式:2026-W40 / 2026-09-28(週一)
  function currentWeekLabel(date = new Date()) {
    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dow = d.getUTCDay() || 7;
    const monday = new Date(d); monday.setUTCDate(d.getUTCDate() - dow + 1);
    const thursday = new Date(d); thursday.setUTCDate(d.getUTCDate() + 4 - dow);
    const yearStart = new Date(Date.UTC(thursday.getUTCFullYear(), 0, 1));
    const week = Math.ceil(((thursday - yearStart) / DAY_MS + 1) / 7);
    // 與 Excel 公式 YEAR(T2) 一致:年份取當天的年份
    return `${d.getUTCFullYear()}-W${String(week).padStart(2, '0')} / ${monday.toISOString().slice(0, 10)}`;
  }

  window.XlsxModel = { Workbook, currentWeekLabel, serialToISO, isoToSerial };
})();
