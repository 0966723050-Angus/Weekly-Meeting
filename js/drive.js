// drive.js
// Google 登入(drive.file 權限,只能存取使用者用 Picker 選過的檔案)+ Drive 檔案列出/下載/上傳。
// 支援共用雲端硬碟(ATK 工作區),所有呼叫都帶 supportsAllDrives=true。
(function () {
  const CFG = window.APP_CONFIG;
  const LS_HINT = 'wm_login_hint';
  const LS_PICKED = 'wm_picked_ids';
  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  let tokenClient = null;
  let accessToken = null;
  let tokenExpiry = 0;
  let pickerLoaded = false;
  let lastDebug = null;

  function lsGet(k) { try { return localStorage.getItem(k); } catch { return null; } }
  function lsSet(k, v) { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, v); } catch {} }
  function getPicked() { try { return JSON.parse(lsGet(LS_PICKED) || '[]'); } catch { return []; } }
  function setPicked(ids) { lsSet(LS_PICKED, JSON.stringify([...new Set(ids)])); }

  function requestToken(prompt) {
    return new Promise((resolve, reject) => {
      if (!window.google || !google.accounts || !google.accounts.oauth2) {
        reject(new Error('Google 登入元件尚未載入,請確認網路後重試'));
        return;
      }
      if (!tokenClient) {
        tokenClient = google.accounts.oauth2.initTokenClient({
          client_id: CFG.GOOGLE_CLIENT_ID,
          scope: CFG.DRIVE_SCOPE,
          callback: () => {},
        });
      }
      tokenClient.callback = (resp) => {
        if (resp.error) { reject(new Error(resp.error_description || resp.error)); return; }
        accessToken = resp.access_token;
        tokenExpiry = Date.now() + (Number(resp.expires_in) || 3600) * 1000 - 60000;
        resolve(accessToken);
      };
      tokenClient.error_callback = (err) => reject(new Error(err && err.type === 'popup_closed' ? '已關閉登入視窗' : '登入失敗'));
      const opts = { prompt: prompt ?? '' };
      const hint = lsGet(LS_HINT);
      if (hint) opts.login_hint = hint;
      tokenClient.requestAccessToken(opts);
    });
  }

  async function ensureToken() {
    if (accessToken && Date.now() < tokenExpiry) return accessToken;
    return requestToken(accessToken ? '' : undefined);
  }

  function isSignedIn() { return !!accessToken && Date.now() < tokenExpiry; }

  async function api(url, opts = {}, retry = true) {
    await ensureToken();
    const resp = await fetch(url, { ...opts, headers: { ...(opts.headers || {}), Authorization: 'Bearer ' + accessToken } });
    if (resp.status === 401 && retry) {
      accessToken = null;
      await requestToken('');
      return api(url, opts, false);
    }
    return resp;
  }

  async function apiError(resp, what) {
    let msg = '';
    try { msg = (await resp.json()).error.message; } catch {}
    const e = new Error(`${what}失敗(${resp.status})${msg ? ':' + msg : ''}`);
    e.status = resp.status;
    return e;
  }

  function loadPicker() {
    return new Promise((resolve, reject) => {
      if (pickerLoaded) return resolve();
      if (!window.gapi) return reject(new Error('Google Picker 元件尚未載入'));
      gapi.load('picker', { callback: () => { pickerLoaded = true; resolve(); }, onerror: () => reject(new Error('Google Picker 載入失敗')) });
    });
  }

  // 讓使用者在 Google Picker 選取週會 Excel 檔(可複選),同時授權本網站存取這些檔案
  async function pickFiles() {
    if (!CFG.GOOGLE_API_KEY || CFG.GOOGLE_API_KEY.startsWith('__')) {
      throw new Error('網站尚未設定 Google API 金鑰(GitHub Secret GOOGLE_API_KEY),無法開啟檔案選擇視窗');
    }
    await ensureToken();
    await loadPicker();
    return new Promise((resolve, reject) => {
      const shared = new google.picker.DocsView(google.picker.ViewId.DOCS)
        .setMimeTypes(XLSX_MIME).setEnableDrives(true).setIncludeFolders(true)
        .setQuery('weekly_meeting_template');
      const mine = new google.picker.DocsView(google.picker.ViewId.DOCS)
        .setMimeTypes(XLSX_MIME).setIncludeFolders(true).setQuery('weekly_meeting_template');
      const picker = new google.picker.PickerBuilder()
        .enableFeature(google.picker.Feature.SUPPORT_DRIVES)
        .enableFeature(google.picker.Feature.MULTISELECT_ENABLED)
        .setOAuthToken(accessToken)
        .setDeveloperKey(CFG.GOOGLE_API_KEY)
        .setAppId(CFG.GOOGLE_PROJECT_NUMBER)
        .addView(shared)
        .addView(mine)
        .setTitle('請選取 ATK 工作區中你的週會檔案(可複選)')
        .setCallback((data) => {
          if (data.action === google.picker.Action.PICKED) {
            const ids = (data.docs || []).map((d) => d.id);
            setPicked([...getPicked(), ...ids]);
            resolve(ids);
          } else if (data.action === google.picker.Action.CANCEL) {
            reject(new Error('已取消選取檔案'));
          }
        })
        .build();
      picker.setVisible(true);
    });
  }

  // 列出本網站可存取(使用者曾用 Picker 選取過且目前仍有權限)的週會檔案
  // 回傳 { 部門代號: { id, name } },部門代號取自檔名 weekly_meeting_template_<部門>.xlsx
  // 來源:1) files.list(drive.file 只會回傳本網站被授權的檔案,不加名稱條件,由程式比對檔名)
  //       2) 本機記住的 Picker 選取檔案 ID,逐一查詢(不依賴搜尋,確保共用雲端硬碟的檔案也找得到)
  async function listWeeklyFiles() {
    const found = new Map(); // id -> name
    const debug = { listed: 0, listError: '', picked: 0, pickedOk: 0 };
    const fields = encodeURIComponent('nextPageToken,files(id,name)');
    const q = encodeURIComponent(`mimeType = '${XLSX_MIME}' and trashed = false`);
    for (const corpora of ['allDrives', 'user']) {
      const url = `https://www.googleapis.com/drive/v3/files?supportsAllDrives=true&includeItemsFromAllDrives=true` +
        `&corpora=${corpora}&pageSize=200&fields=${fields}&q=${q}`;
      const resp = await api(url);
      if (!resp.ok) { debug.listError += `${corpora}:${resp.status} `; continue; }
      const { files = [] } = await resp.json();
      for (const f of files) found.set(f.id, f.name);
      debug.listed += files.length;
    }
    const picked = getPicked();
    debug.picked = picked.length;
    const keep = [];
    for (const id of picked) {
      if (found.has(id)) { keep.push(id); debug.pickedOk++; continue; }
      const resp = await api(`https://www.googleapis.com/drive/v3/files/${id}?supportsAllDrives=true&fields=id,name,trashed`);
      if (resp.ok) {
        const f = await resp.json();
        if (!f.trashed) { found.set(f.id, f.name); keep.push(id); debug.pickedOk++; }
      } else if (resp.status !== 404 && resp.status !== 403) {
        keep.push(id); // 暫時性錯誤,先保留
      }
    }
    setPicked(keep);
    const out = {};
    const re = new RegExp('^' + CFG.FILE_PREFIX + '(.+?)\\.xlsx$', 'i');
    for (const [id, name] of found) {
      const m = re.exec(String(name).trim());
      if (m && !out[m[1]]) out[m[1]] = { id, name };
    }
    debug.names = [...found.values()].join('、');
    lastDebug = debug;
    return out;
  }

  const META_FIELDS = 'id,name,version,modifiedTime,lastModifyingUser(displayName,emailAddress),capabilities(canEdit),webViewLink';

  async function getMeta(fileId) {
    const resp = await api(`https://www.googleapis.com/drive/v3/files/${fileId}?supportsAllDrives=true&fields=${encodeURIComponent(META_FIELDS)}`);
    if (!resp.ok) throw await apiError(resp, '讀取檔案資訊');
    return resp.json();
  }

  async function download(fileId) {
    const meta = await getMeta(fileId);
    const resp = await api(`https://www.googleapis.com/drive/v3/files/${fileId}?alt=media&supportsAllDrives=true`);
    if (!resp.ok) throw await apiError(resp, '下載檔案');
    const bytes = await resp.arrayBuffer();
    return { meta, bytes };
  }

  async function upload(fileId, bytes) {
    const resp = await api(
      `https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=media&supportsAllDrives=true&fields=${encodeURIComponent(META_FIELDS)}`,
      { method: 'PATCH', headers: { 'Content-Type': XLSX_MIME }, body: bytes }
    );
    if (!resp.ok) throw await apiError(resp, '上傳檔案');
    return resp.json();
  }

  // 目前登入者(記住 email 作為下次登入的提示,可省去選帳號步驟)
  async function whoAmI() {
    const resp = await api('https://www.googleapis.com/drive/v3/about?fields=user(displayName,emailAddress)');
    if (!resp.ok) return null;
    const { user } = await resp.json();
    if (user && user.emailAddress) lsSet(LS_HINT, user.emailAddress);
    return user;
  }

  function signOut() {
    if (accessToken && window.google && google.accounts) google.accounts.oauth2.revoke(accessToken, () => {});
    accessToken = null;
    tokenExpiry = 0;
    lsSet(LS_HINT, null);
  }

  window.Drive = {
    requestToken, ensureToken, isSignedIn, pickFiles, listWeeklyFiles, getMeta, download, upload,
    whoAmI, signOut, debugInfo: () => lastDebug,
  };
})();
