// drive.js
// Google 登入(drive.file 權限,只能存取使用者用 Picker 選過的檔案)+ Drive 檔案下載/上傳。
// 支援共用雲端硬碟(ATK 工作區),所有呼叫都帶 supportsAllDrives=true。
(function () {
  const CFG = window.APP_CONFIG;
  const LS_FILE = 'wm_drive_file_id';
  const LS_HINT = 'wm_login_hint';
  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  let tokenClient = null;
  let accessToken = null;
  let tokenExpiry = 0;
  let pickerLoaded = false;

  function lsGet(k) { try { return localStorage.getItem(k); } catch { return null; } }
  function lsSet(k, v) { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, v); } catch {} }

  function getFileId() { return lsGet(LS_FILE) || CFG.DRIVE_FILE_ID || ''; }

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

  // 讓使用者在 Google Picker 選取週會 Excel 檔(同時授權本網站存取該檔)
  async function pickFile() {
    if (!CFG.GOOGLE_API_KEY || CFG.GOOGLE_API_KEY.startsWith('__')) {
      throw new Error('網站尚未設定 Google API 金鑰(GitHub Secret GOOGLE_API_KEY),無法開啟檔案選擇視窗');
    }
    await ensureToken();
    await loadPicker();
    return new Promise((resolve, reject) => {
      const shared = new google.picker.DocsView(google.picker.ViewId.DOCS)
        .setMimeTypes(XLSX_MIME).setEnableDrives(true).setIncludeFolders(true)
        .setQuery(CFG.FILE_NAME_HINT);
      const mine = new google.picker.DocsView(google.picker.ViewId.DOCS)
        .setMimeTypes(XLSX_MIME).setIncludeFolders(true).setQuery(CFG.FILE_NAME_HINT);
      const picker = new google.picker.PickerBuilder()
        .enableFeature(google.picker.Feature.SUPPORT_DRIVES)
        .setOAuthToken(accessToken)
        .setDeveloperKey(CFG.GOOGLE_API_KEY)
        .setAppId(CFG.GOOGLE_PROJECT_NUMBER)
        .addView(shared)
        .addView(mine)
        .setTitle('請選擇 ATK 工作區中的 weekly_meeting_template_生管.xlsx')
        .setCallback((data) => {
          if (data.action === google.picker.Action.PICKED) {
            const doc = data.docs && data.docs[0];
            if (!doc) return reject(new Error('未選取檔案'));
            lsSet(LS_FILE, doc.id);
            resolve(doc.id);
          } else if (data.action === google.picker.Action.CANCEL) {
            reject(new Error('已取消選取檔案'));
          }
        })
        .build();
      picker.setVisible(true);
    });
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
    requestToken, ensureToken, isSignedIn, pickFile, getMeta, download, upload,
    getFileId, whoAmI, signOut, forgetFile: () => lsSet(LS_FILE, null),
  };
})();
