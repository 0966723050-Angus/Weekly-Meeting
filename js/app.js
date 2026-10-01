// app.js - 主流程與 UI
(function () {
  const CFG = window.APP_CONFIG;
  const { Workbook, currentWeekLabel, weekLabel, weekLabelDate, todayISO } = window.XlsxModel;
  const $ = (id) => document.getElementById(id);
  const LS_SHEET = 'wm_last_sheet';

  const state = {
    source: null,        // 'drive'(登入雲端硬碟後)
    access: {},          // 檔案代號(部門) -> { id, name }:此帳號可存取的雲端檔案
    files: {},           // 檔案代號 -> { key, fileId, name, meta, wb, origSst, models }
    current: null,       // 目前工作表名稱
    user: null,
    seq: 0,
  };
  const sheetCfg = (name) => CFG.SHEETS.find((c) => c.name === name);
  // 工作表所屬的檔案代號
  const fileKeyOf = (name) => sheetCfg(name).file;
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
      if (cfg.file !== key) continue;
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
        ? `帳號 ${state.user.emailAddress} 目前沒有可使用的週報檔案,請聯絡管理者確認檔案共用設定。`
        : '找不到週報檔案。';
      const d = state.source === 'drive' && Drive.debugInfo();
      $('deniedDebug').textContent = d
        ? `診斷:找到 ${d.listed} 個檔案${d.listError ? `(錯誤 ${d.listError.trim()})` : ''}` +
          (d.names ? `,檔名:${d.names}` : '')
        : '';
      updateHeader();
      return;
    }
    $('summaryBar').hidden = !canSummarize();
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

  async function signInAndLoad() {
    try {
      busy('登入 Google…');
      await Drive.ensureToken();
      state.user = await Drive.whoAmI().catch(() => null);
      busy('尋找週報檔案…');
      const access = await Drive.listWeeklyFiles();
      if (state.source !== 'drive') state.files = {};
      state.source = 'drive';
      // 已下載但不再可存取的檔案移除
      for (const k of Object.keys(state.files)) if (!access[k]) delete state.files[k];
      state.access = access;
      busy('');
      showSheets();
      refreshSubmitState();
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
    let line = f ? f.name : '';
    if (f && f.meta && f.meta.modifiedTime) {
      const d = new Date(f.meta.modifiedTime);
      line += ` · ${d.getMonth() + 1}/${d.getDate()} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')} 更新`;
    }
    $('fileLine').textContent = line;
    $('userLine').textContent = state.user ? `${state.user.displayName}\n${state.user.emailAddress}` : '尚未登入';
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
  const isWeekCol = (col) => /週次/.test(col.title);
  const isProjectCol = (col) => /^專案名稱$|^專案\/案場名稱$/.test(col.title);

  // 專案名稱選項:生管部檔案「專案主檔」B4 以下(無權限時改用本檔的下拉清單)
  // 有生管部檔案權限 → 讀生管部;否則讀自己檔案內的「專案主檔」(由「校正」從生管部同步過來)
  async function projectOptions(col) {
    const key = state.access[CFG.PROJECT_SOURCE.file] ? CFG.PROJECT_SOURCE.file : fileKeyOf(state.current);
    if (state.source === 'drive' && state.access[key] && !state.files[key]) {
      try { await loadFromDrive(key); } catch { /* 下載失敗就用本檔清單 */ }
    }
    const src = state.files[key];
    let list = [];
    const pm = src && src.models[CFG.PROJECT_SOURCE.sheet];
    const nameCol = pm && pm.columns.find(isProjectCol);
    if (nameCol) {
      // 已載入可編輯的專案主檔 → 直接用(含尚未儲存的新增/修改)
      for (const r of pm.records) {
        const v = String(r.vals[nameCol.idx] ?? '').trim();
        if (!r.blank && v && !list.includes(v)) list.push(v);
      }
    } else if (src && src.wb.sheetPaths[CFG.PROJECT_SOURCE.sheet]) {
      list = src.wb.rangeValues(CFG.PROJECT_SOURCE.range);
    }
    return list.length ? list : (col.options || []);
  }

  async function openEditor(rec, insertAt) {
    const m = model();
    // 「專案主檔」本身是選單來源,其專案名稱要能自由輸入
    const projCol = m.name === CFG.PROJECT_SOURCE.sheet ? null : m.columns.find(isProjectCol);
    const projList = projCol ? await projectOptions(projCol) : [];
    // 有預計/實際進度與目前狀態欄的工作表,狀態依進度自動判斷
    const autoStatus = !!(m.roles.status && m.roles.plan && m.roles.actual);
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
      if (isWeekCol(col)) {
        // 點擊跳出日曆,選好日期自動填上週次
        input = document.createElement('div');
        input.className = 'week-pick';
        input.innerHTML = `<input id="${id}" type="text" readonly value="${esc(v ?? '')}" placeholder="點選日期">` +
          `<span class="cal" aria-hidden="true">📅</span><input type="date" class="week-date" value="${weekLabelDate(v) || todayISO()}" aria-label="${esc(col.title)}">`;
        const [text, date] = [input.querySelector('input[type=text]'), input.querySelector('.week-date')];
        date.addEventListener('change', () => { if (date.value) text.value = weekLabel(date.value); });
        input.addEventListener('click', () => { try { date.showPicker(); } catch { date.focus(); } });
      } else if (projCol === col && projList.length) {
        // 可直接輸入,也可按右側 ▾ 從專案主檔清單選取
        input = document.createElement('div');
        input.className = 'combo';
        const cur = v == null ? '' : String(v);
        input.innerHTML = `<input id="${id}" type="text" value="${esc(cur)}" placeholder="輸入或按右側 ▾ 選擇">` +
          `<span class="combo-btn" aria-hidden="true">▾</span>` +
          `<select class="combo-select" aria-label="從清單選擇${esc(col.title)}">` +
          `<option value="" selected disabled>從清單選擇…</option>` +
          projList.map((o) => `<option value="${esc(o)}">${esc(o)}</option>`).join('') + '</select>';
        const [text, pick] = [input.querySelector('input'), input.querySelector('select')];
        pick.addEventListener('change', () => { text.value = pick.value; pick.selectedIndex = 0; });
      } else if (col.type === 'percent') {
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
      } else if (autoStatus && col === m.roles.status) {
        // 目前狀態:依進度自動判斷;「待確認」「取消」手動一鍵輸入
        input = document.createElement('div');
        input.className = 'opt-wrap';
        input.innerHTML = `<input id="${id}" type="text" value="${esc(v ?? '')}">` +
          '<div class="opt-chips">' +
          CFG.STATUS_MANUAL.map((o) => `<button type="button" data-status="${esc(o)}" class="${String(v) === o ? 'on' : ''}">${esc(o)}</button>`).join('') +
          '<button type="button" data-status="" class="auto">↻ 依進度</button></div>' +
          '<div class="hint-line">依預計/實際完成進度自動判斷;「待確認」「取消」請按上方按鈕</div>';
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
    if (autoStatus) wireAutoStatus(m);
    $('editor').hidden = false;
    document.body.classList.add('noscroll');
    box.scrollTop = 0;
    for (const ta of box.querySelectorAll('textarea')) autoGrow(ta);
  }
  // 依進度判斷狀態:實際=100% → 已完成;實際<預計 → 延遲;預計=0% → 未開始;預計>0% → 進行中
  // 進度兩欄都空白時不判斷(回傳 null)
  function statusFromProgress(planPct, actualPct) {
    if (planPct === '' && actualPct === '') return null;
    const plan = planPct === '' ? 0 : Number(planPct);
    const actual = actualPct === '' ? 0 : Number(actualPct);
    if (actual >= 100) return CFG.STATUS_AUTO.done;
    if (actual < plan) return CFG.STATUS_AUTO.late;
    if (plan <= 0) return CFG.STATUS_AUTO.notStarted;
    return CFG.STATUS_AUTO.doing;
  }

  function wireAutoStatus(m) {
    const statusEl = $('f_' + m.roles.status.idx);
    const planEl = $('f_' + m.roles.plan.idx);
    const actualEl = $('f_' + m.roles.actual.idx);
    const chips = statusEl.parentElement.querySelectorAll('[data-status]');
    const mark = () => chips.forEach((b) => b.classList.toggle('on', !!b.dataset.status && b.dataset.status === statusEl.value));
    const apply = () => {
      const s = statusFromProgress(planEl.value.trim(), actualEl.value.trim());
      if (s) statusEl.value = s;
      mark();
    };
    // 數字欄與滑桿任一變動都重新判斷
    for (const el of [planEl, actualEl]) {
      el.addEventListener('input', apply);
      const range = el.parentElement.querySelector('input[type="range"]');
      if (range) range.addEventListener('input', apply);
    }
    chips.forEach((b) => b.addEventListener('click', (e) => {
      e.preventDefault();
      if (b.dataset.status) { statusEl.value = b.dataset.status; mark(); } else apply();
    }));
    statusEl.addEventListener('input', mark);
    // 新增的列直接依進度判斷;既有的列保留原狀態(例如已設為待確認/取消),改動進度時才重新判斷
    if (editing && editing.insertAt != null) apply();
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
    if (el.tagName === 'SELECT') return raw;
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

  function downloadBytes(bytes, name, type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
    const blob = new Blob([bytes], { type });
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

  // ---------- 提交週報 ----------
  // 透過 Apps Script(管理者身分)在生管部檔案「週報提交」工作表註記;每週二 10:00 由 Apps Script 清空
  async function callSubmit(action) {
    const token = await Drive.ensureToken();
    const resp = await fetch(CFG.SUBMIT_URL, { method: 'POST', body: JSON.stringify({ action, token }) });
    const data = await resp.json().catch(() => ({ ok: false, error: `伺服器回應異常(${resp.status})` }));
    if (!data.ok) throw new Error(data.error || '提交失敗');
    return data;
  }

  function showSubmitState(data) {
    const btn = $('submitBtn');
    const sub = data && data.submission;
    const done = !!(sub && sub.time);
    btn.classList.toggle('done', done);
    btn.textContent = done ? '週報已提交 ✓' : '週報提交';
    btn.title = done ? `${data.dept} 已於 ${sub.time} 提交(${sub.who})` : `提交本週${(data && data.dept) || ''}週報`;
  }

  async function refreshSubmitState() {
    const btn = $('submitBtn');
    btn.hidden = !(CFG.SUBMIT_URL && state.source === 'drive');
    if (btn.hidden) return;
    try { showSubmitState(await callSubmit('status')); } catch { /* 狀態取不到時維持「提交」 */ }
  }

  async function submitWeekly() {
    if (!CFG.SUBMIT_URL) return;
    if (dirtyCount()) {
      await save();
      if (dirtyCount()) return toast('有修改尚未儲存,請先儲存後再提交', 4000);
    }
    const btn = $('submitBtn');
    const was = btn.textContent;
    const ok = await confirmDialog('提交週報', was.includes('已提交')
      ? '本週已提交過,要以目前內容重新提交嗎?' : '確認本週週報已更新完成,要提交嗎?', [
      { label: '取消', value: false },
      { label: '確認提交', value: true, cls: 'btn-primary' },
    ]);
    if (!ok) return;
    btn.disabled = true;
    busy('提交中…');
    try {
      const data = await callSubmit('submit');
      showSubmitState(data);
      toast(`✅ 已提交${data.dept}週報(${data.submission.time})`, 4000);
    } catch (err) {
      toast('提交失敗:' + (err.message || err), 6000);
    } finally {
      busy('');
      btn.disabled = false;
    }
  }
  $('submitBtn').onclick = submitWeekly;

  // ---------- 彙整 / 匯出 PDF ----------
  const deptFiles = () => [...new Set(CFG.SHEETS.map((c) => c.file))];
  // 可存取全部部門檔案的帳號才能彙整
  const canSummarize = () => state.source === 'drive' && deptFiles().every((k) => state.access[k]);

  async function runSummary() {
    if (!canSummarize()) return toast('需要可存取全部部門檔案的帳號才能彙整');
    if (dirtyCount()) {
      const ok = await confirmDialog('尚有未儲存的修改', '彙整會使用雲端上的最新內容。要先儲存目前的修改嗎?', [
        { label: '取消', value: false },
        { label: '儲存後彙整', value: true, cls: 'btn-primary' },
      ]);
      if (!ok) return;
      await save();
      if (dirtyCount()) return;
    }
    try {
      // 重新下載各部門檔案,確保是最新內容
      for (const k of deptFiles()) await loadFromDrive(k);
      const models = {};
      for (const cfg of CFG.SHEETS) { const f = state.files[cfg.file]; if (f && f.models[cfg.name]) models[cfg.name] = f.models[cfg.name]; }
      const rows = Summary.collectRows(models);
      busy('產生 Excel…');
      const xlsx = Summary.buildXlsx(rows);
      busy('產生 PDF…');
      const pdf = await Summary.buildPdf(rows, CFG.SUMMARY_NAME);
      busy('存到雲端硬碟…');
      const parents = await Drive.getParents(state.access[CFG.PROJECT_SOURCE.file].id);
      await Drive.saveByName(CFG.SUMMARY_NAME + '.xlsx', xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', parents);
      await Drive.saveByName(CFG.SUMMARY_NAME + '.pdf', pdf, 'application/pdf', parents);
      busy('');
      if (curFile()) { renderTabs(); renderSheet(false); updateHeader(); }
      const dl = await confirmDialog('彙整完成', `已彙整 ${rows.length} 列,並存到雲端硬碟「${CFG.SUMMARY_NAME}.xlsx」與「${CFG.SUMMARY_NAME}.pdf」。`, [
        { label: '關閉', value: false },
        { label: '匯出 PDF', value: true, cls: 'btn-primary' },
      ]);
      if (dl) await exportPdf(pdf);
    } catch (err) {
      toast('彙整失敗:' + (err.message || err), 6000);
    } finally {
      busy('');
    }
  }

  // ---------- 匯出 PDF 到本機資料夾 ----------
  // 電腦版 Chrome/Edge:第一次選擇資料夾(C:\ATK\週會\彙整),之後記住並直接覆寫「本週會議重點.pdf」
  // 不支援的瀏覽器(例如 iPhone)改為下載同名檔案
  function idb(mode, fn) {
    return new Promise((resolve, reject) => {
      const open = indexedDB.open('weekly-meeting', 1);
      open.onupgradeneeded = () => open.result.createObjectStore('kv');
      open.onerror = () => reject(open.error);
      open.onsuccess = () => {
        try {
          const tx = open.result.transaction('kv', mode);
          const req = fn(tx.objectStore('kv'));
          tx.oncomplete = () => resolve(req && req.result);
          tx.onerror = () => reject(tx.error);
          tx.onabort = () => reject(tx.error);
        } catch (err) {
          reject(err);
        }
      };
    });
  }
  const idbGet = (k) => idb('readonly', (st) => st.get(k)).catch(() => null);
  const idbSet = (k, v) => idb('readwrite', (st) => st.put(v, k)).catch(() => {});
  const canWriteFolder = () => typeof window.showDirectoryPicker === 'function';

  // 取得可寫入的匯出資料夾;必須在使用者點擊後立即呼叫(瀏覽器要求)
  async function getExportDir(forcePick = false) {
    if (!canWriteFolder()) return null;
    let dir = forcePick ? null : await idbGet('exportDir');
    if (dir) {
      let perm = await dir.queryPermission({ mode: 'readwrite' });
      if (perm !== 'granted') perm = await dir.requestPermission({ mode: 'readwrite' });
      if (perm === 'granted') return dir;
    }
    toast(`請選擇匯出資料夾:${CFG.EXPORT_DIR_HINT}`, 6000);
    dir = await window.showDirectoryPicker({ id: 'wm-export', mode: 'readwrite', startIn: 'documents' });
    await idbSet('exportDir', dir);
    return dir;
  }

  // 寫入 PDF;回傳顯示用的位置文字
  async function savePdfLocal(bytes, dir) {
    const name = CFG.EXPORT_NAME + '.pdf';
    if (!dir) {
      downloadBytes(bytes, name, 'application/pdf');
      return `已下載「${name}」`;
    }
    const fh = await dir.getFileHandle(name, { create: true });
    const w = await fh.createWritable();
    await w.write(bytes);
    await w.close();
    return `已匯出到「${dir.name}\\${name}」`;
  }

  // pdfBytes 已有時直接寫;否則從雲端下載最新的彙整 PDF
  async function exportPdf(pdfBytes) {
    if (state.source !== 'drive') return toast('請先登入');
    let dir = null;
    try {
      dir = await getExportDir();
    } catch (err) {
      if (err && err.name === 'AbortError') return toast('已取消匯出');
      dir = null; // 無法使用資料夾寫入 → 改下載
    }
    try {
      let bytes = pdfBytes instanceof Uint8Array ? pdfBytes : null;
      if (!bytes) {
        busy('下載彙整 PDF…');
        const f = await Drive.findByName(CFG.SUMMARY_NAME + '.pdf');
        if (!f) { busy(''); return toast('雲端硬碟還沒有彙整 PDF,請先按「彙整」', 4000); }
        bytes = new Uint8Array(await Drive.downloadRaw(f.id));
      }
      busy('');
      toast('✅ ' + await savePdfLocal(bytes, dir), 5000);
    } catch (err) {
      toast('匯出失敗:' + (err.message || err), 6000);
    } finally {
      busy('');
    }
  }

  async function changeExportDir() {
    if (!canWriteFolder()) return toast('此瀏覽器不支援指定資料夾,匯出時會直接下載');
    try {
      const dir = await getExportDir(true);
      toast(`匯出資料夾已設為「${dir.name}」`, 4000);
    } catch (err) {
      if (!(err && err.name === 'AbortError')) toast('設定失敗:' + (err.message || err), 5000);
    }
  }
  // ---------- 校正 ----------
  // 1) 以生管部工作表第 6 列的週次/日期為準,改正各部門工作表每一列的週次/日期
  // 2) 把生管部檔案的「專案主檔」同步到其他部門檔案,讓各部門的專案名稱下拉選單一致
  const weekColOf = (m) => m.columns.find((c) => /週次/.test(c.title));
  const cloneRec = (r) => ({ id: 'n' + ++state.seq, origRow: null, rowEl: null, cells: {}, vals: { ...r.vals },
    dirty: new Set(Object.keys(r.vals).map(Number)), isNew: true, blank: !!r.blank });

  async function runCalibrate() {
    if (!canSummarize()) return toast('需要可存取全部部門檔案的帳號才能校正');
    if (dirtyCount()) {
      const ok = await confirmDialog('尚有未儲存的修改', '校正會使用雲端上的最新內容。要先儲存目前的修改嗎?', [
        { label: '取消', value: false },
        { label: '儲存後校正', value: true, cls: 'btn-primary' },
      ]);
      if (!ok) return;
      await save();
      if (dirtyCount()) return;
    }
    try {
      for (const k of deptFiles()) await loadFromDrive(k);
      const ref = CFG.CALIBRATE;
      const refModel = state.files[ref.file].models[ref.sheet];
      const refCol = refModel && weekColOf(refModel);
      const refRec = refModel && refModel.records[ref.row - refModel.dataStart];
      const refVal = refRec && refCol ? String(refRec.vals[refCol.idx] ?? '').trim() : '';
      if (!refVal) throw new Error(`${ref.sheet}工作表第 ${ref.row} 列的週次/日期是空的,無法校正`);

      const report = [];
      // 週次/日期
      for (const name of ref.sheets) {
        const f = state.files[sheetCfg(name).file];
        const m = f && f.models[name];
        const col = m && weekColOf(m);
        if (!col) continue;
        let n = 0;
        for (const r of m.records) {
          if (r.blank || String(r.vals[col.idx] ?? '').trim() === refVal) continue;
          r.vals[col.idx] = refVal;
          r.dirty.add(col.idx);
          n++;
        }
        if (n) { m.changeCount = (m.changeCount || 0) + 1; report.push(`${sheetCfg(name).label || name}:${n} 列週次/日期`); }
      }
      // 專案主檔
      const srcPm = state.files[CFG.PROJECT_SOURCE.file].models[CFG.PROJECT_SOURCE.sheet];
      const pmCfg = CFG.SHEETS.find((c) => c.name === CFG.PROJECT_SOURCE.sheet);
      const sig = (m) => JSON.stringify(m.records.filter((r) => !r.blank).map((r) => m.columns.map((c) => r.vals[c.idx] ?? '')));
      if (srcPm) {
        for (const k of deptFiles()) {
          if (k === CFG.PROJECT_SOURCE.file) continue;
          const f = state.files[k];
          if (!f.wb.sheetPaths[CFG.PROJECT_SOURCE.sheet]) continue;
          const pm = f.wb.readSheet(pmCfg);
          if (sig(pm) === sig(srcPm)) continue;
          pm.records = srcPm.records.map(cloneRec);
          pm.changeCount = 1;
          f.models[CFG.PROJECT_SOURCE.sheet] = pm; // 隨該檔一起儲存
          report.push(`${k}檔:同步專案主檔`);
        }
      }
      if (!report.length) {
        toast(`✅ 全部一致(基準:${refVal}),不需校正`, 4000);
        return;
      }
      const ok = await confirmDialog('校正內容', `基準:${ref.sheet}第 ${ref.row} 列「${refVal}」\n\n${report.join('\n')}\n\n確定要存回雲端硬碟?`, [
        { label: '取消', value: false },
        { label: '校正並儲存', value: true, cls: 'btn-primary' },
      ]);
      if (!ok) {
        for (const k of deptFiles()) await loadFromDrive(k); // 放棄校正結果
        showSheets();
        return;
      }
      await save();
      toast(dirtyCount() ? '部分檔案未儲存,請再試一次' : '✅ 校正完成', 4000);
    } catch (err) {
      toast('校正失敗:' + (err.message || err), 6000);
    } finally {
      busy('');
      if (curFile()) { renderTabs(); renderSheet(false); updateHeader(); }
    }
  }

  $('summaryBtn').onclick = runSummary;
  $('calibrateBtn').onclick = runCalibrate;
  $('exportPdfBtn').onclick = () => exportPdf();

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
      state.files = {}; // 全部重新下載
      signInAndLoad();
    } else if (act === 'exportDir') changeExportDir();
    else if (act === 'signout') {
      Drive.signOut();
      location.reload();
    }
  });

  $('signInBtn').onclick = () => signInAndLoad();

  $('deniedSwitch').onclick = () => { Drive.signOut(); location.reload(); };
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

  // ---------- 新版偵測:強制重新載入 ----------
  // 開啟時、切回 App 時、每隔幾分鐘檢查 version.json;版本不同就重新載入。
  // 有未儲存的修改時先要求「儲存並更新」(不可略過),避免修改遺失。
  let updatePending = false;
  async function checkForUpdate() {
    if (updatePending || CFG.APP_VERSION.startsWith('__')) return; // 本機開發版不檢查
    let latest;
    try {
      const resp = await fetch('version.json?t=' + Date.now(), { cache: 'no-store' });
      if (!resp.ok) return;
      latest = (await resp.json()).version;
    } catch { return; }
    if (!latest || latest === CFG.APP_VERSION) return;
    updatePending = true;
    if (!dirtyCount()) return reloadToVersion(latest);
    // 有未儲存的修改:只能「儲存並更新」
    for (;;) {
      await confirmDialog('程式已更新', '「ATK部門週報」有新版本,必須更新後才能繼續使用。\n目前有尚未儲存的修改,將先儲存再更新。', [
        { label: '儲存並更新', value: true, cls: 'btn-primary' },
      ]);
      await save();
      if (!dirtyCount()) return reloadToVersion(latest);
      toast('儲存未完成,請處理後再按一次「儲存並更新」', 5000);
    }
  }

  function reloadToVersion(v) {
    toast('程式已更新,正在重新載入…', 3000);
    if (navigator.serviceWorker && navigator.serviceWorker.controller) {
      navigator.serviceWorker.getRegistration().then((r) => r && r.update()).catch(() => {});
    }
    // 網址加上版本參數,確保瀏覽器不會用舊的快取頁面
    const url = new URL(location.href);
    url.searchParams.set('u', v);
    setTimeout(() => location.replace(url.toString()), 600);
  }

  // 載入新版後把網址上的版本參數拿掉
  if (new URLSearchParams(location.search).has('u')) {
    const url = new URL(location.href);
    url.searchParams.delete('u');
    history.replaceState(null, '', url.pathname + url.search + url.hash);
  }
  checkForUpdate();
  setInterval(checkForUpdate, CFG.VERSION_CHECK_MINUTES * 60000);
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') checkForUpdate(); });
  window.addEventListener('focus', checkForUpdate);

  if ('serviceWorker' in navigator && location.protocol === 'https:') {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }

  // 測試用掛勾(僅供自動化驗證)
  window.__wm = { state, parseFile, showSheets, buildBytes, Workbook };
})();
