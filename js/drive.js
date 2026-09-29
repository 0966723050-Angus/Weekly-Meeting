// drive.js
// Google 登入 + 雲端硬碟週會檔案的列出/下載/上傳。
// 支援共用雲端硬碟(ATK 工作區),所有呼叫都帶 supportsAllDrives=true。
(function () {
  const CFG = window.APP_CONFIG;
  const LS_HINT = 'wm_login_hint';
  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  let tokenClient = null;
  let accessToken = null;
  let tokenExpiry = 0;
  let lastDebug = null;

  function lsGet(k) { try { return localStorage.getItem(k); } catch { return null; } }
  function lsSet(k, v) { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, v); } catch {} }

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

  // 列出此帳號有權限的週會檔案(依雲端硬碟共用設定),含共用雲端硬碟 ATK
  // 回傳 { 部門代號: { id, name } },部門代號取自檔名 weekly_meeting_template_<部門>.xlsx
  async function listWeeklyFiles() {
    const found = new Map(); // id -> name
    const debug = { listed: 0, listError: '' };
    const fields = encodeURIComponent('files(id,name)');
    const q = encodeURIComponent("name contains 'weekly_meeting_template' and trashed = false");
    for (const corpora of ['allDrives', 'user']) {
      const url = 'https://www.googleapis.com/drive/v3/files?supportsAllDrives=true&includeItemsFromAllDrives=true' +
        `&corpora=${corpora}&pageSize=200&orderBy=modifiedTime desc&fields=${fields}&q=${q}`;
      const resp = await api(url);
      if (!resp.ok) { debug.listError += `${corpora}:${resp.status} `; continue; }
      const { files = [] } = await resp.json();
      for (const f of files) if (!found.has(f.id)) found.set(f.id, f.name);
      debug.listed += files.length;
      if (corpora === 'allDrives') break; // allDrives 已涵蓋全部
    }
    const out = {};
    const re = new RegExp('^' + CFG.FILE_PREFIX + '(.+?)\\.xlsx$', 'i');
    for (const [id, name] of found) {
      const m = re.exec(String(name).trim());
      if (m && !out[m[1]]) out[m[1]] = { id, name }; // 同名取最近修改的
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
    requestToken, ensureToken, isSignedIn, listWeeklyFiles, getMeta, download, upload,
    whoAmI, signOut, debugInfo: () => lastDebug,
  };
})();
