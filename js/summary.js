// summary.js - 各部工作彙整
// 依「C:\ATK\週會\彙整\Excel彙整\各部工作彙整.xlsx」的格式:單一工作表「彙整」、欄位 A~L、
// 部門依序排列、每個部門取最新一週的資料列;Excel 為 A3 直印、寬度縮放為一頁;PDF 為單一長頁。
(function () {
  const { idxToCol } = window.XlsxModel;
  const te = new TextEncoder();

  // 彙整順序與來源工作表
  const ORDER = ['太陽能追蹤', '生管部', '業務部', '電氣設計', '管理部', '機構設計'];
  const HEADERS = ['部門', '週次/日期', '專案名稱', '負責人', '上週工作內容', '本週預計工作內容',
    '預計完成進度%', '實際完成進度%', '目前狀態', '需跨部門協調內容', '需總經理裁決事項', '備註/下週延續重點'];
  // 來源欄位標題與彙整欄位不同名時的對應
  const ALIAS = { '專案/案場名稱': '專案名稱', '主責人': '負責人' };
  const WIDTHS = [16.2, 22.4, 22.2, 15.9, 60.9, 67.2, 20.2, 18.6, 14, 25, 25.1, 24.6];
  const PCT = new Set([6, 7]);          // 百分比欄(0 起算)
  const CENTER = new Set([3, 6, 7, 8, 9, 10]);
  const WRAP = new Set([2, 4, 5, 9, 10]);

  const weekKey = (label) => { const m = /(\d{4})-W(\d{1,2})/.exec(String(label || '')); return m ? `${m[1]}-${m[2].padStart(2, '0')}` : ''; };

  // models: { 工作表名稱: model }(來自 xlsx-model 的 readSheet)
  function collectRows(models) {
    const rows = [];
    for (const sheet of ORDER) {
      const m = models[sheet];
      if (!m) continue;
      const colOf = {};
      for (const c of m.columns) colOf[ALIAS[c.title] || c.title] = c;
      const weekCol = colOf['週次/日期'];
      const recs = m.records.filter((r) => !r.blank);
      let latest = '';
      if (weekCol) for (const r of recs) { const k = weekKey(r.vals[weekCol.idx]); if (k > latest) latest = k; }
      for (const r of recs) {
        if (weekCol && latest && weekKey(r.vals[weekCol.idx]) !== latest) continue;
        rows.push(HEADERS.map((h, i) => {
          if (i === 0) return sheet;
          const c = colOf[h];
          const v = c ? r.vals[c.idx] : '';
          return v == null ? '' : v;
        }));
      }
    }
    return rows;
  }

  // ---------- Excel ----------
  const x = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  function buildXlsx(rows) {
    const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    // 樣式:0 預設、1 標題、2 內文靠左、3 內文置中、4 百分比置中、5 內文靠左換行、6 內文置中換行
    const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${NS}">
<fonts count="2"><font><sz val="11"/><name val="新細明體"/><family val="1"/><charset val="136"/></font><font><b/><sz val="11"/><name val="新細明體"/><family val="1"/><charset val="136"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9EAF7"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="7">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="9" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

    const styleOf = (i) => (PCT.has(i) ? 4 : CENTER.has(i) ? (WRAP.has(i) ? 6 : 3) : WRAP.has(i) ? 5 : 2);
    const cell = (ref, v, s) => {
      if (v === '' || v == null) return `<c r="${ref}" s="${s}"/>`;
      if (typeof v === 'number' && Number.isFinite(v)) return `<c r="${ref}" s="${s}"><v>${v}</v></c>`;
      return `<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${x(v)}</t></is></c>`;
    };
    let data = '<row r="1">' + HEADERS.map((h, i) => cell(idxToCol(i + 1) + 1, h, 1)).join('') + '</row>';
    rows.forEach((row, ri) => {
      const r = ri + 2;
      data += `<row r="${r}" ht="100.05" customHeight="1">` + row.map((v, i) => cell(idxToCol(i + 1) + r, v, styleOf(i))).join('') + '</row>';
    });
    const last = rows.length + 1;
    const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:L${last}"/>
<sheetViews><sheetView workbookViewId="0" zoomScale="70" zoomScaleNormal="70"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
<sheetFormatPr defaultRowHeight="16.2"/>
<cols>${WIDTHS.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols>
<sheetData>${data}</sheetData>
<pageMargins left="0.708661417322835" right="0.708661417322835" top="0.748031496062992" bottom="0.748031496062992" header="0.31496062992126" footer="0.31496062992126"/>
<pageSetup paperSize="8" orientation="portrait" fitToHeight="0"/>
</worksheet>`;
    const files = {
      '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
      '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="彙整" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm.Print_Titles" localSheetId="0">彙整!$1:$1</definedName></definedNames></workbook>`,
      'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
      'xl/styles.xml': styles,
      'xl/worksheets/sheet1.xml': sheet,
    };
    const enc = {};
    for (const [k, v] of Object.entries(files)) enc[k] = te.encode(v);
    return fflate.zipSync(enc, { level: 6 });
  }

  // ---------- PDF(單一頁面:頁寬 A3,頁高隨內容)----------
  function loadScript(src) {
    return new Promise((resolve, reject) => {
      if (document.querySelector(`script[src="${src}"]`)) return resolve();
      const s = document.createElement('script');
      s.src = src;
      s.onload = resolve;
      s.onerror = () => reject(new Error('無法載入 PDF 元件,請確認網路'));
      document.head.appendChild(s);
    });
  }

  const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const pct = (v) => (typeof v === 'number' ? Math.round(v * 100) + '%' : esc(v));

  async function buildPdf(rows, title) {
    await loadScript('https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js');
    await loadScript('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js');
    const { jsPDF } = window.jspdf;

    // 單頁:頁寬 A3(297mm),頁高依內容延長,所有欄位與資料列在同一頁
    const PAGE_W = 1600;
    const MARGIN_MM = 12;
    const contentWmm = 297 - MARGIN_MM * 2;
    const total = WIDTHS.reduce((a, b) => a + b, 0);
    const colgroup = '<colgroup>' + WIDTHS.map((w) => `<col style="width:${(w / total * 100).toFixed(3)}%">`).join('') + '</colgroup>';
    const css = `font-family:"PMingLiU","MingLiU","PingFang TC","Noto Serif TC","Microsoft JhengHei",serif;font-size:13px;color:#000;`;
    const td = (v, i) => {
      const align = CENTER.has(i) || PCT.has(i) ? 'center' : 'left';
      const text = PCT.has(i) ? pct(v) : esc(v).replace(/\n/g, '<br>');
      return `<td style="border:1px solid #000;padding:3px 4px;vertical-align:middle;text-align:${align};word-break:break-word;height:128px;box-sizing:border-box">${text}</td>`;
    };
    const th = HEADERS.map((h) => `<th style="border:1px solid #000;background:#D9EAF7;padding:4px;font-weight:bold;text-align:center">${esc(h)}</th>`).join('');

    const host = document.createElement('div');
    host.style.cssText = `position:fixed;left:-99999px;top:0;width:${PAGE_W}px;background:#fff`;
    document.body.appendChild(host);
    try {
      host.innerHTML = `<table style="width:100%;border-collapse:collapse;table-layout:fixed;${css}">${colgroup}` +
        `<thead><tr>${th}</tr></thead><tbody>` + rows.map((r) => `<tr>${r.map(td).join('')}</tr>`).join('') + '</tbody></table>';
      // 手機瀏覽器的畫布面積上限約 1600 萬像素,資料多時自動降低解析度
      const hpx = host.getBoundingClientRect().height;
      const scale = Math.max(0.6, Math.min(1.6, Math.sqrt(15e6 / (PAGE_W * hpx))));
      const canvas = await html2canvas(host, { scale, backgroundColor: '#ffffff', logging: false });
      const img = canvas.toDataURL('image/jpeg', 0.8);
      const hmm = canvas.height / canvas.width * contentWmm;
      const pageH = Math.max(hmm + MARGIN_MM * 2, 100);
      const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: [297, pageH], compress: true });
      pdf.setProperties({ title });
      pdf.addImage(img, 'JPEG', MARGIN_MM, MARGIN_MM, contentWmm, hmm);
      return new Uint8Array(pdf.output('arraybuffer'));
    } finally {
      host.remove();
    }
  }

  window.Summary = { ORDER, HEADERS, collectRows, buildXlsx, buildPdf };
})();
