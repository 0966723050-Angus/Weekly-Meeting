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
  const NO_WRAP = [0, 1, 3, 6, 7, 8]; // PDF 中不換行的短欄位:部門、週次、負責人、預計/實際進度、狀態

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
        if (m.flagIdx != null && String(r.vals[m.flagIdx] ?? '').trim() !== '') continue; // 不列入週報
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

  // ---------- PDF(真正的文字與格線,其他程式可讀取表格內容)----------
  // 以 pdf-lib 直接繪製表格,嵌入繁體中文字型(只嵌入用到的字),PDF 內含可擷取的文字層。
  // 版面:頁寬 A3,頁高依內容延長(所有欄位與資料列在同一頁);超過 PDF 頁面上限才分頁。
  const PDF_LIB = 'https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js';
  const FONTKIT = 'https://cdn.jsdelivr.net/npm/@pdf-lib/fontkit@1.1.1/dist/fontkit.umd.min.js';
  // Noto Sans TC(TrueType;OpenType/CFF 版的中文字型子集化有已知問題)
  const FONT_URL = 'https://cdn.jsdelivr.net/gh/google/fonts@main/ofl/notosanstc/NotoSansTC%5Bwght%5D.ttf';

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

  // 字型約 12MB:第一次下載後存入瀏覽器快取(Cache Storage),之後直接讀取
  let fontBytesPromise = null;
  function loadFontBytes() {
    if (!fontBytesPromise) {
      fontBytesPromise = (async () => {
        let cache = null;
        try { cache = await caches.open('wm-fonts-v1'); } catch { /* 不支援時直接下載 */ }
        const hit = cache && await cache.match(FONT_URL);
        if (hit) return hit.arrayBuffer();
        const r = await fetch(FONT_URL);
        if (!r.ok) throw new Error('無法下載中文字型');
        if (cache) await cache.put(FONT_URL, r.clone()).catch(() => {});
        return r.arrayBuffer();
      })().catch((e) => { fontBytesPromise = null; throw e; });
    }
    return fontBytesPromise;
  }

  // 背景預先下載字型與 PDF 元件(可彙整的帳號登入後呼叫)
  function prefetchPdfAssets() {
    loadFontBytes().catch(() => {});
    loadScript(PDF_LIB).catch(() => {});
    loadScript(FONTKIT).catch(() => {});
  }

  const cellText = (v, i) => {
    if (v == null || v === '') return '';
    if (PCT.has(i)) return typeof v === 'number' ? Math.round(v * 100) + '%' : String(v);
    return String(v).replace(/\r\n?/g, '\n').replace(/\t/g, ' ');
  };

  // 依欄寬自動換行(中文逐字、英數以字為單位)
  function wrapLines(text, font, size, maxW) {
    const out = [];
    for (const para of String(text).split('\n')) {
      if (!para) { out.push(''); continue; }
      const tokens = para.match(/[A-Za-z0-9_.,:;!?%/+\-()'"#&@]+|\s+|./gu) || [];
      let line = '';
      for (const tk of tokens) {
        const tryLine = line + tk;
        if (font.widthOfTextAtSize(tryLine, size) <= maxW) { line = tryLine; continue; }
        if (line.trim()) out.push(line.replace(/\s+$/, ''));
        line = tk.trimStart();
        // 單一 token 太長:逐字切
        while (line && font.widthOfTextAtSize(line, size) > maxW) {
          let n = line.length;
          while (n > 1 && font.widthOfTextAtSize(line.slice(0, n), size) > maxW) n--;
          out.push(line.slice(0, n));
          line = line.slice(n);
        }
      }
      out.push(line.replace(/\s+$/, ''));
    }
    while (out.length > 1 && out[out.length - 1] === '') out.pop();
    while (out.length > 1 && out[0] === '') out.shift();
    return out;
  }

  async function buildPdf(rows, title) {
    await loadScript(PDF_LIB);
    await loadScript(FONTKIT);
    const { PDFDocument, rgb } = window.PDFLib;
    const fontBytes = await loadFontBytes();

    const doc = await PDFDocument.create();
    doc.registerFontkit(window.fontkit);
    const font = await doc.embedFont(fontBytes, { subset: true });
    doc.setTitle(title);
    doc.setSubject('ATK部門週報 各部工作彙整');
    doc.setCreator('ATK部門週報');
    doc.setLanguage('zh-TW');

    // 單位:pt。A3 寬 297mm = 841.89pt
    const PAGE_W = 841.89;
    const M = 34; // 約 12mm 邊界
    const tableW = PAGE_W - M * 2;
    const totalW = WIDTHS.reduce((a, b) => a + b, 0);
    const colW = WIDTHS.map((w) => w / totalW * tableW);
    const SIZE = 8.5;
    const HEAD_SIZE = 8.5;
    const PADX = 3;
    // 短欄位(部門、週次、負責人、進度、狀態)不換行,寬度不足時向兩個工作內容欄借
    const texts = rows.map((r) => r.map(cellText));
    for (const i of NO_WRAP) {
      // 資料值不換行;較長的標題可折成兩行(取一半寬度),5 字以內的標題不折
      const head = HEADERS[i];
      const headHalf = font.widthOfTextAtSize(head.length <= 5 ? head : head.slice(0, Math.ceil(head.length / 2)), HEAD_SIZE);
      const widest = Math.max(headHalf,
        ...texts.map((r) => Math.max(0, ...r[i].split('\n').map((p) => font.widthOfTextAtSize(p, SIZE)))));
      colW[i] = Math.max(colW[i], widest + PADX * 2 + 1);
    }
    // 總寬超過頁寬時,由可換行的欄位等比例縮小;仍不夠才全部等比例縮放
    const sumW = () => colW.reduce((a, b) => a + b, 0);
    let over = sumW() - tableW;
    if (over > 0) {
      const flex = [2, 4, 5, 9, 10, 11];
      const minW = (i) => (i === 4 || i === 5 ? 110 : 48);
      const room = flex.reduce((a, i) => a + Math.max(0, colW[i] - minW(i)), 0);
      if (room > 0) for (const i of flex) colW[i] -= Math.min(colW[i] - minW(i), over * Math.max(0, colW[i] - minW(i)) / room);
      over = sumW() - tableW;
      if (over > 0.5) { const k = tableW / sumW(); for (let i = 0; i < colW.length; i++) colW[i] *= k; }
    }
    const colX = colW.map((_, i) => M + colW.slice(0, i).reduce((a, b) => a + b, 0));
    const PADY = 4;
    const MIN_ROW = 26;
    const MAX_PAGE_H = 14000; // PDF 頁面高度上限 14400pt

    const layoutRow = (cells, size) => {
      const lines = cells.map((t, i) => wrapLines(t, font, size, colW[i] - PADX * 2));
      const h = Math.max(MIN_ROW, Math.max(...lines.map((l) => l.length)) * size * 1.35 + PADY * 2);
      return { lines, h };
    };
    const head = layoutRow(HEADERS, HEAD_SIZE);
    const body = texts.map((r) => layoutRow(r, SIZE));

    // 依高度分頁(一般只有一頁)
    const pages = [];
    let cur = [];
    let h = head.h;
    for (const row of body) {
      if (cur.length && M * 2 + h + row.h > MAX_PAGE_H) { pages.push(cur); cur = []; h = head.h; }
      cur.push(row);
      h += row.h;
    }
    pages.push(cur);

    const black = rgb(0, 0, 0);
    const headFill = rgb(0xD9 / 255, 0xEA / 255, 0xF7 / 255);
    const drawRow = (page, row, top, size, fill, isHead) => {
      row.lines.forEach((lines, i) => {
        page.drawRectangle({
          x: colX[i], y: top - row.h, width: colW[i], height: row.h,
          borderColor: black, borderWidth: 0.6, color: fill || undefined,
        });
        const blockH = lines.length * size * 1.35;
        let y = top - (row.h - blockH) / 2 - size; // 垂直置中
        const center = isHead || CENTER.has(i) || PCT.has(i);
        for (const ln of lines) {
          if (ln) {
            const w = font.widthOfTextAtSize(ln, size);
            const x = center ? colX[i] + (colW[i] - w) / 2 : colX[i] + PADX;
            // 每段文字只畫一次,確保擷取出來的文字不重複
            page.drawText(ln, { x, y: y + size * 0.15, size, font, color: black });
          }
          y -= size * 1.35;
        }
      });
    };

    // 附加 CSV(與表格相同的資料),供程式直接讀取
    const csv = '﻿' + [HEADERS, ...texts].map((r) => r.map((v) => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\r\n');
    await doc.attach(new TextEncoder().encode(csv), `${title}.csv`, {
      mimeType: 'text/csv', description: '各部工作彙整資料(與 PDF 表格內容相同)',
      creationDate: new Date(), modificationDate: new Date(),
    });

    // 字型預設為 Thin 字重:以「填滿+描邊」繪製文字加粗(只影響外觀,擷取的文字不變)
    const { setTextRenderingMode, TextRenderingMode, setLineWidth, setStrokingRgbColor } = window.PDFLib;
    const textWeight = (page, w) => page.pushOperators(
      setTextRenderingMode(TextRenderingMode.FillAndOutline), setLineWidth(w), setStrokingRgbColor(0, 0, 0));

    for (const pageRows of pages) {
      const pageH = M * 2 + head.h + pageRows.reduce((a, r) => a + r.h, 0);
      const page = doc.addPage([PAGE_W, pageH]);
      let top = pageH - M;
      textWeight(page, 0.75); // 標題較粗
      drawRow(page, head, top, HEAD_SIZE, headFill, true);
      textWeight(page, 0.4);
      top -= head.h;
      for (const row of pageRows) {
        drawRow(page, row, top, SIZE, null, false);
        top -= row.h;
      }
    }
    return await doc.save({ useObjectStreams: false }); // 不壓縮物件串流,相容較舊的 PDF 解析程式
  }

  window.Summary = { ORDER, HEADERS, collectRows, buildXlsx, buildPdf, prefetchPdfAssets };
})();
