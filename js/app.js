// app.js - 主流程與 UI
(function () {
  const CFG = window.APP_CONFIG;
  const { Workbook, currentWeekLabel } = window.XlsxModel;
  const $ = (id) => document.getElementById(id);
  const LS_SHEET = 'wm_last_sheet';

  const state = {
    source: null,        // 'drive' | 'local'
    access: {},          // 檔案代號(部門) -> { id, name }:此帳號可存取的雲端檔案
    files: {},           // 檔案代號 -> { key, fileId, name, meta, wb, origSst, models }
    current: null,       // 目前工作表名稱
    user: null,
    seq: 0,
  };
  const sheetCfg = (name) => CFG.SHEETS.find((c) => c.name === name);
  // 工作表所屬的檔案代號;本機模式只有一個檔案
  const fileKeyOf = (name) => (state.source === 'local' ? 'local' : sheetCfg(name).file);
  const curFile = () => state.files[fileKeyOf(state.current)];

  // ---------- 共用 UI ----------
  let toastTimer;
  function toast(msg, ms = 2600) {
    const t = $('toast');
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove('show'), ms);
  }
  function busy(msg) {
    $('busy').hidden = !msg;
    if (msg) $('busyMsg').textContent = msg;
  }
  function confirmDialog(title, msg, buttons) {
    return new Promise((resolve) => {
      $('confirmTitle').textContent = title;
      $('confirmMsg').textContent = msg;
      const box = $('confirmBtns');
      box.innerHTML = '';
      for (const b of buttons) {
        const el = document.createElement('button');
        el.type = 'button';
        el.className = 'btn ' + (b.cls || 'btn-ghost');
        el.textContent = b.label;
        el.onclick = () => { $('confirmBox').hidden = true; resolve(b.value); };
        box.appendChild(el);
      }
      $('confirmBox').hidden = false;
    });
  }
  function lsGet(k) { try { return localStorage.getItem(k); } catch { return null; } }
  function lsSet(k, v) { try { localStorage.setItem(k, v); } catch {} }
  function esc(s) {
    return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  // ---------- 載入 ----------
  // 解析一個檔案,建立其中各工作表的編輯模型
  function parseFile(key, bytes, { name, meta, fileId }) {
    const wb = new Workbook(bytes);
    const models = {};
    for (const cfg of CFG.SHEETS) {
      if (key !== 'local' && cfg.file !== key) continue;
      if (!wb.sheetPaths[cfg.name]) continue;
      const m = wb.readSheet(cfg);
      m.roles = detectRoles(m.columns);
      models[cfg.name] = m;
    }
    state.files[key] = { key, fileId, name, meta, wb, models, origSst: wb.text('xl/sharedStrings.xml') };
    return state.files[key];
  }

  // 此工作表是否可使用(有權限的檔案裡確實有這張表)
  function sheetAvailable(name) {
    if (state.source === 'local') return !!(state.files.local && state.files.local.models[name]);
    const key = sheetCfg(name).file;
    if (!state.access[key]) return false;
    const f = state.files[key];
    return !f || !!f.models[name]; // 尚未下載的檔案先視為可用
  }

  function showSheets() {
    $('welcome').hidden = true;
    const any = CFG.SHEETS.some((c) => sheetAvailable(c.name));
    $('denied').hidden = any;
    $('sheetView').hidden = !any;
    $('sheetTabs').hidden = !any;
    $('addBtn').hidden = !any;
    if (!any) {
      $('deniedMsg').textContent = state.user
        ? `帳號 ${state.user.emailAddress} 目前沒有可使用的週會檔案。若檔案已共用給你,請按下方「選擇週會檔案」,在 ATK 共用雲端硬碟中勾選你的檔案。`
        : '找不到週會表單的工作表,請確認選對檔案。';
      $('deniedPick').hidden = state.source !== 'drive';
      const d = state.source === 'drive' && Drive.debugInfo();
      $('deniedDebug').textContent = d
        ? `診斷:清單 ${d.listed} 個檔案${d.listError ? `(錯誤 ${d.listError.trim()})` : ''},已選取 ${d.pickedOk}/${d.picked}` +
          (d.names ? `,檔名:${d.names}` : '')
        : '';
      updateHeader();
      return;
    }
    const last = lsGet(LS_SHEET);
    const pick = [state.current, last, CFG.DEFAULT_SHEET, ...CFG.SHEETS.map((c) => c.name)]
      .find((n) => n && sheetCfg(n) && sheetAvailable(n));
    switchSheet(pick);
  }

  // 切換分頁;所屬檔案尚未下載時先下載
  async function switchSheet(name) {
    if (!sheetAvailable(name)) return;
    const key = fileKeyOf(name);
    if (!state.files[key]) {
      try { await loadFromDrive(key); } catch (err) { toast(err.message || String(err), 5000); return; }
      if (!state.files[key].models[name]) {
        renderTabs();
        toast(`「${state.access[key].name}」裡沒有「${name}」工作表`, 4000);
        return;
      }
    }
    state.current = name;
    lsSet(LS_SHEET, name);
    renderTabs();
    renderSheet(true);
    updateHeader();
  }

  function detectRoles(columns) {
    const find = (re, except) => columns.find((c) => re.test(c.title) && !(except && except.includes(c)));
    const roles = {};
    roles.week = find(/週次/);
    roles.title = find(/^事項$/) || find(/專案.*名稱|案場名稱/) || find(/^部門$/) || columns[0];
    roles.code = find(/專案編號/);
    roles.owner = find(/負責人|主責人|報告人/);
    roles.status = find(/目前狀態/);
    roles.plan = find(/預計完成進度/);
    roles.actual = find(/實際完成進度/);
    const used = Object.values(roles).filter(Boolean);
    roles.previews = columns.filter((c) => c.type === 'text' && !used.includes(c) && !c.options).slice(0, 2);
    roles.lastWeek = find(/上週工作內容/);
    roles.thisWeek = find(/本週預計工作內容/);
    return roles;
  }

  async function signInAndLoad({ forcePick = false } = {}) {
    try {
      busy('登入 Google…');
      await Drive.ensureToken();
      state.user = await Drive.whoAmI().catch(() => null);
      busy('尋找週會檔案…');
      let access = await Drive.listWeeklyFiles();
      if (forcePick || !Object.keys(access).length) {
        // 第一次使用(或要加選):用 Picker 選取自己有權限的週會檔案,授權本網站存取
        busy('');
        await Drive.pickFiles();
        busy('尋找週會檔案…');
        access = await Drive.listWeeklyFiles();
      }
      if (state.source !== 'drive') state.files = {};
      state.source = 'drive';
      // 已下載但不再可存取的檔案移除
      for (const k of Object.keys(state.files)) if (!access[k]) delete state.files[k];
      state.access = access;
      busy('');
      showSheets();
    } catch (err) {
      busy('');
      toast(err.message || String(err), 5000);
    }
  }

  async function loadFromDrive(key) {
    const info = state.access[key];
    busy(`下載「${info.name}」…`);
    try {
      const { meta, bytes } = await Drive.download(info.id);
      parseFile(key, bytes, { name: meta.name, meta, fileId: info.id });
    } finally {
      busy('');
    }
  }

  function openLocalFile(file) {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        state.source = 'local';
        state.access = {};
        state.files = {};
        state.current = null;
        parseFile('local', reader.result, { name: file.name, meta: null, fileId: null });
        showSheets();
        toast(`已開啟本機檔案「${file.name}」;按「下載」取得修改後的檔案`, 4000);
      } catch (err) {
        toast('無法讀取檔案:' + err.message, 5000);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // ---------- 頁首 / 分頁 ----------
  const fileDirty = (f) => Object.values(f.models).reduce((n, m) => n + (m.changeCount || 0), 0);
  function dirtyCount() {
    return Object.values(state.files).reduce((n, f) => n + fileDirty(f), 0);
  }
  function updateHeader() {
    const n = dirtyCount();
    const badge = $('dirtyBadge');
    badge.hidden = n === 0;
    badge.textContent = n;
    const f = state.current && state.files[fileKeyOf(state.current)];
    $('saveBtn').disabled = !f || (n === 0 && state.source === 'drive');
    $('saveBtn').firstElementChild.textContent = state.source === 'local' ? '下載' : '儲存';
    let line = f ? f.name : '';
    if (f && f.meta && f.meta.modifiedTime) {
      const d = new Date(f.meta.modifiedTime);
      line += ` · ${d.getMonth() + 1}/${d.getDate()} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')} 更新`;
    }
    if (state.source === 'local') line += ' · 本機檔案';
    $('fileLine').textContent = line;
    $('userLine').textContent = state.user ? `${state.user.displayName}\n${state.user.emailAddress}` : (state.source === 'local' ? '本機檔案模式' : '尚未登入');
    for (const tab of $('sheetTabs').children) {
      const tf = state.files[fileKeyOf(tab.dataset.sheet)];
      const m = tf && tf.models[tab.dataset.sheet];
      tab.classList.toggle('dirty', !!(m && m.changeCount));
    }
  }

  // 所有分頁都列出;沒有權限的反白(灰色、不可點)
  function renderTabs() {
    const nav = $('sheetTabs');
    nav.innerHTML = '';
    for (const cfg of CFG.SHEETS) {
      const ok = sheetAvailable(cfg.name);
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'tab' + (cfg.name === state.current ? ' active' : '') + (ok ? '' : ' locked');
      b.dataset.sheet = cfg.name;
      b.textContent = cfg.label || cfg.name;
      b.disabled = !ok;
      if (!ok) b.title = '沒有此工作表的權限';
      b.onclick = () => { if (cfg.name !== state.current) switchSheet(cfg.name); };
      nav.appendChild(b);
    }
    const active = nav.querySelector('.active');
    if (active) active.scrollIntoView({ inline: 'center', block: 'nearest' });
  }

  // ---------- 列表 ----------
  function model() { return curFile().models[state.current]; }

  function valText(col, v) {
    if (v == null || v === '') return '';
    if (col.type === 'percent') return typeof v === 'number' ? Math.round(v * 100) + '%' : String(v);
    return String(v);
  }

  function renderSheet(resetFilter) {
    const m = model();
    const R = m.roles;
    const sel = $('weekFilter');
    if (resetFilter) {
      $('searchBox').value = '';
      sel.innerHTML = '';
      if (R.week) {
        const weeks = [...new Set(m.records.map((r) => String(r.vals[R.week.idx] || '').trim()).filter(Boolean))];
        weeks.sort((a, b) => b.localeCompare(a));
        sel.append(new Option(`全部週次(${weeks.length})`, ''));
        for (const w of weeks) sel.append(new Option(w, w));
      }
      sel.hidden = !R.week;
      $('carryBtn').hidden = !R.week;
    }
    const wk = sel.value;
    const q = $('searchBox').value.trim().toLowerCase();
    const list = $('cardList');
    list.innerHTML = '';
    let shown = 0;
    const total = m.records.filter((r) => !r.blank).length;
    m.records.forEach((rec, i) => {
      if (rec.blank) return;
      if (wk &&String(rec.vals[R.week.idx] || '').trim() !== wk) return;
      if (q) {
        const hay = m.columns.map((c) => valText(c, rec.vals[c.idx])).join(' ').toLowerCase();
        if (!hay.includes(q)) return;
      }
      shown++;
      list.appendChild(card(m, rec, i));
    });
    $('emptyState').hidden = shown > 0;
    $('countLine').textContent = `${m.name} · 共 ${total} 列${shown !== total ? `,顯示 ${shown} 列` : ''}`;
  }

  function statusClass(s) {
    s = String(s || '');
    if (/已完成/.test(s)) return 'st-done';
    if (/進行中/.test(s)) return 'st-doing';
    if (/延遲/.test(s)) return 'st-late';
    if (/待確認/.test(s)) return 'st-wait';
    if (/取消/.test(s)) return 'st-cancel';
    return 'st-none';
  }

  function card(m, rec, i) {
    const R = m.roles;
    const v = (col) => (col ? valText(col, rec.vals[col.idx]) : '');
    const el = document.createElement('article');
    el.className = 'card' + (rec.dirty.size || !rec.origRow ? ' changed' : '');
    const rowNo = m.dataStart + i;
    let html = '<div class="card-top">';
    html += `<span class="rowno">第 ${rowNo} 列</span>`;
    if (R.week && v(R.week)) html += `<span class="chip">${esc(v(R.week))}</span>`;
    if (R.status && v(R.status)) html += `<span class="status ${statusClass(v(R.status))}">${esc(v(R.status))}</span>`;
    html += '<button type="button" class="more" aria-label="列動作">⋯</button></div>';
    html += `<h3 class="card-title">${esc(v(R.title) || '(未填)')}</h3>`;
    const sub = [v(R.code), v(R.owner)].filter(Boolean).join(' · ');
    if (sub) html += `<div class="card-sub">${esc(sub)}</div>`;
    for (const col of R.previews) {
      const t = v(col).trim();
      if (t) html += `<div class="preview"><span class="plabel">${esc(col.title)}</span><div class="ptext">${esc(t)}</div></div>`;
    }
    if (R.plan || R.actual) {
      const bar = (col, cls) => {
        const n = typeof rec.vals[col.idx] === 'number' ? Math.max(0, Math.min(1, rec.vals[col.idx])) : null;
        return `<div class="prog"><span>${esc(col.title.replace('完成進度%', ''))}</span><div class="track"><i class="${cls}" style="width:${n == null ? 0 : n * 100}%"></i></div><b>${n == null ? '—' : Math.round(n * 100) + '%'}</b></div>`;
      };
      html += '<div class="progs">' + (R.plan ? bar(R.plan, 'plan') : '') + (R.actual ? bar(R.actual, 'actual') : '') + '</div>';
    }
    el.innerHTML = html;
    el.onclick = (e) => {
      if (e.target.closest('.more')) openActions(rec);
      else openEditor(rec);
    };
    return el;
  }

  // ---------- 列動作 ----------
  let actionRec = null;
  function openActions(rec) {
    actionRec = rec;
    const m = model();
    const idx = m.records.indexOf(rec);
    $('actionTitle').textContent = `第 ${m.dataStart + idx} 列 · ${valText(m.roles.title, rec.vals[m.roles.title.idx]) || '(未填)'}`;
    $('actionSheet').hidden = false;
  }
  $('actionSheet').addEventListener('click', async (e) => {
    const btn = e.target.closest('[data-row-act]');
    if (!btn && e.target !== $('actionSheet')) return;
    $('actionSheet').hidden = true;
    const act = btn ? btn.dataset.rowAct : 'cancel';
    const rec = actionRec;
    if (!rec || act === 'cancel') return;
    const m = model();
    const idx = m.records.indexOf(rec);
    if (act === 'edit') openEditor(rec);
    else if (act === 'above') openEditor(newRecord(m), idx);
    else if (act === 'below') openEditor(newRecord(m), idx + 1);
    else if (act === 'carry') openEditor(carryRecord(m, rec), idx + 1);
    else if (act === 'delete') deleteRecord(rec);
  });

  function newRecord(m, src) {
    const rec = { id: 'n' + ++state.seq, origRow: null, rowEl: null, cells: {}, vals: {}, dirty: new Set(), isNew: true };
    if (src) {
      rec.rowEl = src.rowEl;
      rec.cells = { ...src.cells };
      rec.vals = { ...src.vals };
    } else {
      const R = m.roles;
      if (R.week) rec.vals[R.week.idx] = currentWeekLabel();
      // 同工作表最常見的負責人 → 預設值
      if (R.owner) {
        const counts = {};
        for (const r of m.records) { const o = r.vals[R.owner.idx]; if (o) counts[o] = (counts[o] || 0) + 1; }
        const top = Object.entries(counts).sort((a, b) => b[1] - a[1])[0];
        if (top) rec.vals[R.owner.idx] = top[0];
      }
      if (R.status) rec.vals[R.status.idx] = R.status.options && R.status.options.includes('進行中') ? '進行中' : '';
    }
    return rec;
  }

  // 延續上週列:週次改為本週、「本週預計工作內容」移到「上週工作內容」
  function carryRecord(m, src) {
    const R = m.roles;
    const rec = newRecord(m, src);
    if (R.week) rec.vals[R.week.idx] = currentWeekLabel();
    if (R.lastWeek && R.thisWeek) {
      rec.vals[R.lastWeek.idx] = src.vals[R.thisWeek.idx] ?? '';
      rec.vals[R.thisWeek.idx] = '';
    }
    return rec;
  }

  async function deleteRecord(rec) {
    const m = model();
    const idx = m.records.indexOf(rec);
    const title = valText(m.roles.title, rec.vals[m.roles.title.idx]) || '(未填)';
    const ok = await confirmDialog('刪除此列?', `第 ${m.dataStart + idx} 列「${title}」將被刪除,下方的列會往上遞補。按「儲存」後才會寫回雲端。`, [
      { label: '取消', value: false },
      { label: '刪除', value: true, cls: 'btn-danger' },
    ]);
    if (!ok) return;
    m.records.splice(idx, 1);
    m.changeCount = (m.changeCount || 0) + 1;
    closeEditor();
    renderSheet(false);
    updateHeader();
    toast('已刪除(尚未儲存)');
  }

  // ---------- 編輯表單 ----------
  let editing = null; // { rec, insertAt }
  function openEditor(rec, insertAt) {
    const m = model();
    editing = { rec, insertAt };
    const isNew = insertAt != null;
    const idx = isNew ? insertAt : m.records.indexOf(rec);
    $('editorTitle').textContent = `${m.name} · ${isNew ? '新增' : '編輯'}第 ${m.dataStart + idx} 列`;
    $('editorDelete').hidden = isNew;
    const box = $('editorFields');
    box.innerHTML = '';
    for (const col of m.columns) {
      const v = rec.vals[col.idx];
      const id = 'f_' + col.idx;
      const wrap = document.createElement('label');
      wrap.className = 'field';
      wrap.htmlFor = id;
      wrap.innerHTML = `<span class="flabel">${esc(col.title)}</span>`;
      let input;
      if (col.type === 'percent') {
        input = document.createElement('div');
        input.className = 'pct';
        const n = typeof v === 'number' ? Math.round(v * 1000) / 10 : '';
        input.innerHTML = `<input id="${id}" type="number" inputmode="decimal" min="0" max="100" step="5" value="${n}"><span>%</span>` +
          `<input type="range" min="0" max="100" step="5" value="${n === '' ? 0 : n}" tabindex="-1" aria-hidden="true">`;
        const [num, range] = input.querySelectorAll('input');
        range.oninput = () => { num.value = range.value; };
        num.oninput = () => { range.value = num.value || 0; };
      } else if (col.type === 'date') {
        input = document.createElement('input');
        input.type = 'date';
        input.id = id;
        input.value = /^\d{4}-\d{2}-\d{2}$/.test(String(v || '')) ? v : '';
        if (v && !input.value) { input.type = 'text'; input.value = v; }
      } else if (col.options && col.options.length) {
        input = document.createElement('div');
        input.className = 'opt-wrap';
        const listId = 'dl_' + col.idx;
        input.innerHTML = `<input id="${id}" type="text" list="${listId}" value="${esc(v ?? '')}">` +
          `<datalist id="${listId}">${col.options.map((o) => `<option value="${esc(o)}">`).join('')}</datalist>`;
        if (col.options.length <= 8) {
          const chips = document.createElement('div');
          chips.className = 'opt-chips';
          for (const o of col.options) {
            const b = document.createElement('button');
            b.type = 'button';
            b.textContent = o;
            b.className = String(v) === o ? 'on' : '';
            b.onclick = (e) => {
              e.preventDefault();
              input.querySelector('input').value = o;
              chips.querySelectorAll('button').forEach((x) => x.classList.toggle('on', x === b));
            };
            chips.appendChild(b);
          }
          input.appendChild(chips);
        }
      } else {
        const long = col.width >= 20 || /內容|協調|裁決|備註|進度|摘要|事項|風險|結論|重點/.test(col.title);
        input = document.createElement(long ? 'textarea' : 'input');
        input.id = id;
        if (long) input.rows = Math.min(8, Math.max(2, String(v ?? '').split('\n').length + 1));
        input.value = v ?? '';
      }
      wrap.appendChild(input);
      box.appendChild(wrap);
    }
    $('editor').hidden = false;
    document.body.classList.add('noscroll');
    box.scrollTop = 0;
    for (const ta of box.querySelectorAll('textarea')) autoGrow(ta);
  }
  function autoGrow(ta) {
    ta.style.height = 'auto';
    ta.style.height = Math.min(ta.scrollHeight + 2, 320) + 'px';
  }
  $('editorFields').addEventListener('input', (e) => { if (e.target.tagName === 'TEXTAREA') autoGrow(e.target); });

  function closeEditor() {
    $('editor').hidden = true;
    document.body.classList.remove('noscroll');
    editing = null;
  }

  function readField(col) {
    const el = $('f_' + col.idx);
    const raw = el.value;
    if (col.type === 'percent') {
      if (raw === '') return '';
      const n = Number(raw);
      return Number.isFinite(n) ? Math.round(n * 10) / 1000 : '';
    }
    if (col.type === 'date' || el.tagName === 'TEXTAREA') return raw.replace(/\r\n/g, '\n');
    return raw.trim();
  }

  $('editorForm').addEventListener('submit', (e) => {
    e.preventDefault();
    if (!editing) return;
    const m = model();
    const { rec, insertAt } = editing;
    const isNew = insertAt != null;
    let changed = 0;
    for (const col of m.columns) {
      const nv = readField(col);
      const ov = rec.vals[col.idx] ?? '';
      if (isNew || String(nv) !== String(ov)) {
        if (nv === '' ) delete rec.vals[col.idx]; else rec.vals[col.idx] = nv;
        if (String(nv) !== String(ov) || isNew) rec.dirty.add(col.idx);
        if (String(nv) !== String(ov)) changed++;
      }
    }
    if (isNew) {
      m.records.splice(insertAt, 0, rec);
      m.changeCount = (m.changeCount || 0) + 1;
      toast(`已新增第 ${m.dataStart + insertAt} 列(尚未儲存)`);
    } else if (changed) {
      m.changeCount = (m.changeCount || 0) + 1;
      toast('已修改(尚未儲存)');
    }
    closeEditor();
    renderSheet(isNew);
    updateHeader();
  });
  $('editorCancel').onclick = closeEditor;
  $('editorDelete').onclick = () => { if (editing) deleteRecord(editing.rec); };

  $('addBtn').onclick = () => {
    const m = model();
    // 篩選某週時,新增到該週最後一列之後;否則加到最後
    const wk = $('weekFilter').value;
    let at = m.records.length;
    if (wk && m.roles.week) {
      for (let i = m.records.length - 1; i >= 0; i--) {
        if (String(m.records[i].vals[m.roles.week.idx] || '').trim() === wk) { at = i + 1; break; }
      }
    }
    openEditor(newRecord(m), at);
  };

  // ---------- 儲存 ----------
  function buildBytes(f, targetWb) {
    for (const m of Object.values(f.models)) if (m.changeCount) targetWb.writeSheet(m);
    return targetWb.toBytes();
  }

  function downloadBytes(bytes, name) {
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name || 'weekly_meeting.xlsx';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  // 儲存單一檔案;回傳 true 表示已存
  async function saveFile(f) {
    busy(`檢查「${f.name}」版本…`);
    let target = f.wb;
    const latest = await Drive.getMeta(f.fileId);
    if (f.meta && latest.version !== f.meta.version) {
      busy('檔案已被他人更新,檢查是否可合併…');
      const { bytes } = await Drive.download(f.fileId);
      const fresh = new Workbook(bytes);
      const dirtyModels = Object.values(f.models).filter((m) => m.changeCount);
      const safe = fresh.text('xl/sharedStrings.xml') === f.origSst &&
        dirtyModels.every((m) => fresh.text(m.path) === m.origDataXml);
      if (safe) {
        target = fresh; // 對方改的是其他工作表 → 套用我的修改到最新版
      } else {
        busy('');
        const who = (latest.lastModifyingUser && latest.lastModifyingUser.displayName) || '其他人';
        const t = new Date(latest.modifiedTime);
        const choice = await confirmDialog(`「${f.name}」已被他人修改`,
          `${who} 於 ${t.toLocaleString('zh-TW')} 更新了同一個工作表。覆寫會蓋掉對方的修改;建議先記下你的修改,重新載入後再編輯。`,
          [
            { label: '取消', value: 'cancel' },
            { label: '重新載入(放棄我的修改)', value: 'reload' },
            { label: '仍要覆寫', value: 'overwrite', cls: 'btn-danger' },
          ]);
        if (choice === 'reload') { await loadFromDrive(f.key); return false; }
        if (choice !== 'overwrite') return false;
      }
    }
    busy(`儲存「${f.name}」…`);
    const bytes = buildBytes(f, new Workbook(target.toBytes()));
    const meta = await Drive.upload(f.fileId, bytes);
    parseFile(f.key, bytes, { name: meta.name, meta, fileId: f.fileId });
    return true;
  }

  async function save() {
    if (state.source === 'local') {
      const f = state.files.local;
      if (f) downloadBytes(buildBytes(f, new Workbook(f.wb.toBytes())), f.name);
      return;
    }
    const dirty = Object.values(state.files).filter((f) => fileDirty(f));
    if (!dirty.length) return;
    let saved = 0;
    try {
      for (const f of dirty) if (await saveFile(f)) saved++;
      if (saved) toast(saved === dirty.length ? '✅ 已儲存到雲端硬碟' : `已儲存 ${saved} 個檔案`);
    } catch (err) {
      toast('儲存失敗:' + (err.message || err), 6000);
    } finally {
      busy('');
      renderTabs();
      if (curFile()) renderSheet(false);
      updateHeader();
    }
  }
  $('saveBtn').onclick = save;

  // ---------- 選單 ----------
  function toggleDrawer(open) {
    $('drawer').classList.toggle('open', open);
    $('drawer').setAttribute('aria-hidden', String(!open));
    $('drawerBackdrop').hidden = !open;
  }
  $('menuBtn').onclick = () => toggleDrawer(true);
  $('drawerBackdrop').onclick = () => toggleDrawer(false);

  async function guardDirty() {
    if (!dirtyCount()) return true;
    return confirmDialog('尚有未儲存的修改', `目前有 ${dirtyCount()} 項修改尚未儲存,繼續將會放棄這些修改。`, [
      { label: '取消', value: false },
      { label: '放棄修改並繼續', value: true, cls: 'btn-danger' },
    ]);
  }

  $('drawer').addEventListener('click', async (e) => {
    const b = e.target.closest('[data-act]');
    if (!b) return;
    toggleDrawer(false);
    const act = b.dataset.act;
    const f = state.current && state.files[fileKeyOf(state.current)];
    if (act === 'download') {
      if (!f) return toast('尚未載入檔案');
      downloadBytes(buildBytes(f, new Workbook(f.wb.toBytes())), f.name);
      return;
    }
    if (act === 'openDrive') {
      if (f && f.meta && f.meta.webViewLink) window.open(f.meta.webViewLink, '_blank', 'noopener');
      else toast('請先登入並載入雲端檔案');
      return;
    }
    if (!(await guardDirty())) return;
    if (act === 'reload') {
      if (state.source === 'local') return toast('本機檔案模式請重新開啟檔案');
      state.files = {}; // 全部重新下載
      signInAndLoad();
    } else if (act === 'pick') signInAndLoad({ forcePick: true });
    else if (act === 'openLocal') $('fileInput').click();
    else if (act === 'signout') {
      Drive.signOut();
      location.reload();
    }
  });

  $('signInBtn').onclick = () => signInAndLoad();

  $('deniedPick').onclick = () => signInAndLoad({ forcePick: true });
  $('deniedSwitch').onclick = () => { Drive.signOut(); location.reload(); };
  $('localBtn').onclick = () => $('fileInput').click();
  $('fileInput').onchange = (e) => {
    const f = e.target.files && e.target.files[0];
    if (f) openLocalFile(f);
    e.target.value = '';
  };
  $('weekFilter').onchange = () => renderSheet(false);
  let searchTimer;
  $('searchBox').oninput = () => { clearTimeout(searchTimer); searchTimer = setTimeout(() => renderSheet(false), 150); };

  document.addEventListener('keydown', (e) => {
    if (e.key !== 'Escape') return;
    if (!$('editor').hidden) closeEditor();
    else if (!$('actionSheet').hidden) $('actionSheet').hidden = true;
    else toggleDrawer(false);
  });

  window.addEventListener('beforeunload', (e) => {
    if (dirtyCount()) { e.preventDefault(); e.returnValue = ''; }
  });

  if ('serviceWorker' in navigator && location.protocol === 'https:') {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }

  // 測試用掛勾(僅供自動化驗證)
  window.__wm = { state, parseFile, showSheets, buildBytes, Workbook };
})();
