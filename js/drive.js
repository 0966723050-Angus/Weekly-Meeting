// drive.js
// Google 登入 + 雲端硬碟週會檔案的列出/下載/上傳。
// 支援共用雲端硬碟(ATK 工作區),所有呼叫都帶 supportsAllDrives=true。
(function () {
  const CFG = window.APP_CONFIG;
  const LS_HINT = 'wm_login_hint';
  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  function lsGet(k) { try { return localStorage.getItem(k); } catch { return null; } }
  function lsSet(k, v) { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, v); } catch {} }

  const LS_SESSION = 'wm_session'; // 保持登入代碼(由 Apps Script 換發權杖)
  const LS_TOKEN = 'wm_token';     // 目前的存取權杖與到期時間(1 小時內重開不必再換)

  let tokenClient = null;
  let accessToken = null;
  let tokenExpiry = 0;
  let lastDebug = null;
  const LS_SESS_OK = 'wm_sessions_ok';
  // Apps Script 是否已設定好保持登入(先沿用上次得知的結果,背景再確認)
  let sessionsSupported = lsGet(LS_SESS_OK) === '1';

  function setToken(tok, expiresIn) {
    accessToken = tok;
    tokenExpiry = Date.now() + (Number(expiresIn) || 3600) * 1000 - 60000;
    lsSet(LS_TOKEN, JSON.stringify({ t: tok, e: tokenExpiry }));
    return tok;
  }

  // 開啟時先用上次存下、尚未過期的權杖
  try {
    const saved = JSON.parse(lsGet(LS_TOKEN) || 'null');
    if (saved && saved.t && Date.now() < saved.e) { accessToken = saved.t; tokenExpiry = saved.e; }
  } catch {}

  async function authCall(body) {
    const resp = await fetch(CFG.SUBMIT_URL, { method: 'POST', body: JSON.stringify(body) });
    return resp.json();
  }

  // 背景詢問 Apps Script 是否支援保持登入(登入按鈕按下時才不必等)
  if (CFG.SUBMIT_URL) {
    authCall({ action: 'authInfo' })
      .then((d) => { sessionsSupported = !!(d && d.ok && d.sessions); lsSet(LS_SESS_OK, sessionsSupported ? '1' : null); })
      .catch(() => {});
  }

  function hasSession() { return !!lsGet(LS_SESSION); }

  // 用保持登入代碼換新權杖;失敗回傳 null(代碼失效時一併清除)
  let refreshing = null;
  function refreshFromSession() {
    const sid = lsGet(LS_SESSION);
    if (!sid || !CFG.SUBMIT_URL) return Promise.resolve(null);
    if (!refreshing) {
      refreshing = authCall({ action: 'refresh', sid })
        .then((d) => {
          if (d && d.ok && d.access_token) return setToken(d.access_token, d.expires_in);
          if (d && d.error === 'SESSION_EXPIRED') lsSet(LS_SESSION, null);
          return null;
        })
        .catch(() => null)
        .finally(() => { refreshing = null; });
    }
    return refreshing;
  }

  function gisReady() {
    if (!window.google || !google.accounts || !google.accounts.oauth2) throw new Error('Google 登入元件尚未載入,請確認網路後重試');
  }

  // 互動式登入(需在使用者點擊中呼叫):支援保持登入時用授權碼流程,否則用一次性權杖
  function requestToken(prompt) {
    return sessionsSupported ? codeLogin() : tokenLogin(prompt);
  }

  function codeLogin() {
    return new Promise((resolve, reject) => {
      gisReady();
      const hint = lsGet(LS_HINT);
      const client = google.accounts.oauth2.initCodeClient({
        client_id: CFG.GOOGLE_CLIENT_ID,
        scope: CFG.DRIVE_SCOPE,
        ux_mode: 'popup',
        // consent:確保拿到 refresh token(每台裝置只需第一次登入時同意)
        prompt: hint ? 'consent' : 'select_account consent',
        ...(hint ? { login_hint: hint } : {}),
        callback: async (resp) => {
          if (resp.error) { reject(new Error(resp.error_description || resp.error)); return; }
          try {
            const d = await authCall({ action: 'login', code: resp.code });
            if (!d.ok) throw new Error(d.error || '登入失敗');
            if (d.sid) lsSet(LS_SESSION, d.sid);
            if (d.email) lsSet(LS_HINT, d.email);
            resolve(setToken(d.access_token, d.expires_in));
          } catch (err) { reject(err); }
        },
        error_callback: (err) => reject(new Error(err && err.type === 'popup_closed' ? '已關閉登入視窗' : '登入失敗')),
      });
      client.requestCode();
    });
  }

  function tokenLogin(prompt) {
    return new Promise((resolve, reject) => {
      gisReady();
      if (!tokenClient) {
        tokenClient = google.accounts.oauth2.initTokenClient({
          client_id: CFG.GOOGLE_CLIENT_ID,
          scope: CFG.DRIVE_SCOPE,
          callback: () => {},
        });
      }
      tokenClient.callback = (resp) => {
        if (resp.error) { reject(new Error(resp.error_description || resp.error)); return; }
        resolve(setToken(resp.access_token, resp.expires_in));
      };
      tokenClient.error_callback = (err) => reject(new Error(err && err.type === 'popup_closed' ? '已關閉登入視窗' : '登入失敗'));
      const opts = { prompt: prompt ?? '' };
      const hint = lsGet(LS_HINT);
      if (hint) opts.login_hint = hint;
      tokenClient.requestAccessToken(opts);
    });
  }

  // 取得可用權杖:現有 → 保持登入換發 → 互動式登入
  // 沒有保持登入代碼時同步開登入視窗(避免被瀏覽器當成非點擊彈窗擋掉)
  async function ensureToken() {
    if (accessToken && Date.now() < tokenExpiry) return accessToken;
    if (hasSession()) {
      const t = await refreshFromSession();
      if (t) return t;
    }
    return requestToken(accessToken ? '' : undefined);
  }

  // 不開登入視窗,只嘗試恢復登入(App 開啟時自動使用);成功回傳 true
  async function resume() {
    if (accessToken && Date.now() < tokenExpiry) return true;
    return !!(await refreshFromSession());
  }

  function isSignedIn() { return !!accessToken && Date.now() < tokenExpiry; }

  async function api(url, opts = {}, retry = true) {
    await ensureToken();
    const resp = await fetch(url, { ...opts, headers: { ...(opts.headers || {}), Authorization: 'Bearer ' + accessToken } });
    if (resp.status === 401 && retry) {
      accessToken = null;
      tokenExpiry = 0;
      if (!(await refreshFromSession())) await requestToken('');
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

  async function upload(fileId, bytes, mime = XLSX_MIME) {
    const resp = await api(
      `https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=media&supportsAllDrives=true&fields=${encodeURIComponent(META_FIELDS)}`,
      { method: 'PATCH', headers: { 'Content-Type': mime }, body: bytes }
    );
    if (!resp.ok) throw await apiError(resp, '上傳檔案');
    return resp.json();
  }

  // 依完整檔名找檔案(最近修改的優先),找不到回傳 null
  async function findByName(name) {
    const q = encodeURIComponent(`name = '${name.replace(/'/g, "\\'")}' and trashed = false`);
    const resp = await api('https://www.googleapis.com/drive/v3/files?supportsAllDrives=true&includeItemsFromAllDrives=true' +
      `&corpora=allDrives&orderBy=modifiedTime desc&pageSize=10&fields=${encodeURIComponent('files(id,name,parents,webViewLink)')}&q=${q}`);
    if (!resp.ok) throw await apiError(resp, '搜尋檔案');
    const { files = [] } = await resp.json();
    return files[0] || null;
  }

  async function getParents(fileId) {
    const resp = await api(`https://www.googleapis.com/drive/v3/files/${fileId}?supportsAllDrives=true&fields=parents`);
    if (!resp.ok) throw await apiError(resp, '讀取資料夾');
    return (await resp.json()).parents || [];
  }

  // 在指定資料夾建立新檔案
  async function createFile(name, bytes, mime, parents) {
    const boundary = 'wm' + Math.random().toString(36).slice(2);
    const head = `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n` +
      JSON.stringify({ name, mimeType: mime, parents }) + `\r\n--${boundary}\r\nContent-Type: ${mime}\r\n\r\n`;
    const body = new Blob([head, bytes, `\r\n--${boundary}--`]);
    const resp = await api(
      `https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true&fields=${encodeURIComponent(META_FIELDS)}`,
      { method: 'POST', headers: { 'Content-Type': `multipart/related; boundary=${boundary}` }, body }
    );
    if (!resp.ok) throw await apiError(resp, '建立檔案');
    return resp.json();
  }

  // 有同名檔案就覆寫內容,否則在 parents 資料夾建立
  async function saveByName(name, bytes, mime, parents) {
    const f = await findByName(name);
    return f ? upload(f.id, bytes, mime) : createFile(name, bytes, mime, parents);
  }

  async function downloadRaw(fileId) {
    const resp = await api(`https://www.googleapis.com/drive/v3/files/${fileId}?alt=media&supportsAllDrives=true`);
    if (!resp.ok) throw await apiError(resp, '下載檔案');
    return resp.arrayBuffer();
  }

  // 目前登入者(記住 email 作為下次登入的提示,可省去選帳號步驟)
  async function whoAmI() {
    const resp = await api('https://www.googleapis.com/drive/v3/about?fields=user(displayName,emailAddress)');
    if (!resp.ok) return null;
    const { user } = await resp.json();
    if (user && user.emailAddress) lsSet(LS_HINT, user.emailAddress);
    return user;
  }

  // 登出:刪除此裝置的保持登入紀錄(不撤銷整個授權,其他裝置維持登入)
  function signOut() {
    const sid = lsGet(LS_SESSION);
    if (sid && CFG.SUBMIT_URL) {
      try { navigator.sendBeacon(CFG.SUBMIT_URL, JSON.stringify({ action: 'logout', sid })); } catch {}
    } else if (accessToken && window.google && google.accounts) {
      google.accounts.oauth2.revoke(accessToken, () => {});
    }
    accessToken = null;
    tokenExpiry = 0;
    lsSet(LS_SESSION, null);
    lsSet(LS_TOKEN, null);
    lsSet(LS_HINT, null);
  }

  window.Drive = {
    requestToken, ensureToken, resume, hasSession, isSignedIn, listWeeklyFiles, getMeta, download, upload,
    findByName, getParents, saveByName, downloadRaw,
    whoAmI, signOut, debugInfo: () => lastDebug,
  };
})();
