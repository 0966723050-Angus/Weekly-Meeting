// app.js - 主流程與 UI
(function () {
  const CFG = window.APP_CONFIG;
  const { Workbook, currentWeekLabel } = window.XlsxModel;
  const $ = (id) => document.getElementById(id);
  const LS_SHEET = 'wm_last_sheet';

  const state = {
    source: null,        // 'drive' | 'local'
    fileId: null,
    fileName: '',
    meta: null,          // Drive 檔案資訊(version 用於偵測他人更新)
    wb: null,
    origSst: null,
    models: {},          // sheetName -> model
    current: null,       // 目前工作表名稱
    user: null,
    seq: 0,
  };

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
  function loadWorkbook(bytes, { source, name, meta }) {
    const wb = new Workbook(bytes);
    const models = {};
    for (const cfg of CFG.SHEETS) {
      if (!wb.sheetPaths[cfg.name]) continue;
      const m = wb.readSheet(cfg);
      m.roles = detectRoles(m.columns);
      models[cfg.name] = m;
    }
    if (!Object.keys(models).length) throw new Error('這個檔案裡找不到週會表單的工作表,請確認選對檔案');
    Object.assign(state, { wb, models, source, fileName: name, meta, origSst: wb.text('xl/sharedStrings.xml') });
    const last = lsGet(LS_SHEET);
    state.current = models[state.current] ? state.current : models[last] ? last : models[CFG.DEFAULT_SHEET] ? CFG.DEFAULT_SHEET : Object.keys(models)[0];
    $('welcome').hidden = true;
    $('sheetView').hidden = false;
    $('sheetTabs').hidden = false;
    $('addBtn').hidden = false;
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
      Drive.whoAmI().then((u) => { state.user = u; updateHeader(); }).catch(() => {});
      let fileId = forcePick ? '' : Drive.getFileId();
      if (!fileId) { busy(''); fileId = await Drive.pickFile(); }
      await loadFromDrive(fileId);
    } catch (err) {
      busy('');
      toast(err.message || String(err), 5000);
    }
  }

  async function loadFromDrive(fileId) {
    busy('下載檔案中…');
    try {
      const { meta, bytes } = await Drive.download(fileId);
      state.fileId = fileId;
      loadWorkbook(bytes, { source: 'drive', name: meta.name, meta });
      toast(`已載入「${meta.name}」`);
    } catch (err) {
      if (err.status === 404 || err.status === 403) {
        // 這個帳號尚未透過 Picker 授權此檔 → 重新選檔
        Drive.forgetFile();
        busy('');
        toast('請重新選取檔案以授權存取', 4000);
        const id = await Drive.pickFile();
        return loadFromDrive(id);
      }
      throw err;
    } finally {
      busy('');
    }
  }

  function openLocalFile(file) {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        state.fileId = null;
        loadWorkbook(reader.result, { source: 'local', name: file.name, meta: null });
        toast(`已開啟本機檔案「${file.name}」;按「下載」取得修改後的檔案`, 4000);
      } catch (err) {
        toast('無法讀取檔案:' + err.message, 5000);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // ---------- 頁首 / 分頁 ----------
  function dirtyCount() {
    return Object.values(state.models).reduce((n, m) => n + (m.changeCount || 0), 0);
  }
  function updateHeader() {
    const n = dirtyCount();
    const badge = $('dirtyBadge');
    badge.hidden = n === 0;
    badge.textContent = n;
    $('saveBtn').disabled = !state.wb || (n === 0 && state.source === 'drive');
    $('saveBtn').firstElementChild.textContent = state.source === 'local' ? '下載' : '儲存';
    let line = state.fileName || '';
    if (state.meta && state.meta.modifiedTime) {
      const d = new Date(state.meta.modifiedTime);
      line += ` · ${d.getMonth() + 1}/${d.getDate()} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')} 更新`;
    }
    if (state.source === 'local') line += ' · 本機檔案';
    $('fileLine').textContent = line;
    $('userLine').textContent = state.user ? `${state.user.displayName}\n${state.user.emailAddress}` : (state.source === 'local' ? '本機檔案模式' : '尚未登入');
    for (const tab of $('sheetTabs').children) {
      const m = state.models[tab.dataset.sheet];
      tab.classList.toggle('dirty', !!(m && m.changeCount));
    }
  }

  function renderTabs() {
    const nav = $('sheetTabs');
    nav.innerHTML = '';
    for (const cfg of CFG.SHEETS) {
      if (!state.models[cfg.name]) continue;
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'tab' + (cfg.name === state.current ? ' active' : '');
      b.dataset.sheet = cfg.name;
      b.textContent = cfg.label || cfg.name;
      b.onclick = () => {
        state.current = cfg.name;
        lsSet(LS_SHEET, cfg.name);
        for (const t of nav.children) t.classList.toggle('active', t === b);
        renderSheet(true);
      };
      nav.appendChild(b);
    }
    const active = nav.querySelector('.active');
    if (active) active.scrollIntoView({ inline: 'center', block: 'nearest' });
  }

  // ---------- 列表 ----------
  function model() { return state.models[state.current]; }

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
  function buildBytes(targetWb) {
    for (const m of Object.values(state.models)) if (m.changeCount) targetWb.writeSheet(m);
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

  async function save() {
    if (!state.wb) return;
    if (state.source === 'local') {
      const bytes = buildBytes(new Workbook(state.wb.toBytes()));
      downloadBytes(bytes, state.fileName);
      return;
    }
    if (!dirtyCount()) return;
    busy('檢查雲端檔案版本…');
    try {
      let target = state.wb;
      const latest = await Drive.getMeta(state.fileId);
      if (state.meta && latest.version !== state.meta.version) {
        busy('檔案已被他人更新,檢查是否可合併…');
        const { bytes, meta } = await Drive.download(state.fileId);
        const fresh = new Workbook(bytes);
        const dirtyModels = Object.values(state.models).filter((m) => m.changeCount);
        const safe = fresh.text('xl/sharedStrings.xml') === state.origSst &&
          dirtyModels.every((m) => fresh.text(m.path) === m.origDataXml);
        if (safe) {
          target = fresh; // 對方改的是其他工作表 → 套用我的修改到最新版
          state.meta = meta;
        } else {
          busy('');
          const who = (latest.lastModifyingUser && latest.lastModifyingUser.displayName) || '其他人';
          const t = new Date(latest.modifiedTime);
          const choice = await confirmDialog('檔案已被他人修改',
            `${who} 於 ${t.toLocaleString('zh-TW')} 更新了同一個工作表。覆寫會蓋掉對方的修改;建議先記下你的修改,重新載入後再編輯。`,
            [
              { label: '取消', value: 'cancel' },
              { label: '重新載入(放棄我的修改)', value: 'reload' },
              { label: '仍要覆寫', value: 'overwrite', cls: 'btn-danger' },
            ]);
          if (choice === 'reload') { await loadFromDrive(state.fileId); return; }
          if (choice !== 'overwrite') return;
        }
      }
      busy('儲存到雲端硬碟…');
      const bytes = buildBytes(new Workbook(target.toBytes()));
      const meta = await Drive.upload(state.fileId, bytes);
      loadWorkbook(bytes, { source: 'drive', name: meta.name, meta });
      toast('✅ 已儲存到雲端硬碟');
    } catch (err) {
      toast('儲存失敗:' + (err.message || err), 6000);
    } finally {
      busy('');
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
    if (act === 'download') {
      if (!state.wb) return toast('尚未載入檔案');
      const bytes = buildBytes(new Workbook(state.wb.toBytes()));
      downloadBytes(bytes, state.fileName);
      return;
    }
    if (act === 'openDrive') {
      if (state.meta && state.meta.webViewLink) window.open(state.meta.webViewLink, '_blank', 'noopener');
      else toast('請先登入並載入雲端檔案');
      return;
    }
    if (!(await guardDirty())) return;
    if (act === 'reload') {
      if (state.source === 'drive' && state.fileId) {
        try { await loadFromDrive(state.fileId); } catch (err) { toast(err.message, 5000); }
      } else signInAndLoad();
    } else if (act === 'pick') signInAndLoad({ forcePick: true });
    else if (act === 'openLocal') $('fileInput').click();
    else if (act === 'signout') {
      Drive.signOut();
      location.reload();
    }
  });

  $('signInBtn').onclick = () => signInAndLoad();
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
  window.__wm = { state, loadWorkbook, buildBytes, Workbook };
})();
