// ATK 週會表單 - 全域設定
window.APP_CONFIG = {
  // Google Drive 上 weekly_meeting_template_生管.xlsx 的檔案 ID。
  // 留空時,第一次使用會跳出 Google Picker 讓使用者選取檔案,選取後記在瀏覽器中。
  DRIVE_FILE_ID: '',
  FILE_NAME_HINT: 'weekly_meeting_template_生管',

  // 與「ATK近期工作項目」共用同一個 Google Cloud 專案(worklist050)的憑證,
  // 已授權來源 https://0966723050-angus.github.io
  GOOGLE_CLIENT_ID: '873968217418-q5t3i90e4pf04kbd4l6vbpjib13etb7o.apps.googleusercontent.com',
  GOOGLE_PROJECT_NUMBER: '873968217418',
  // API 金鑰(僅 Google Picker 使用)不放在原始碼中:存於 GitHub Secret「GOOGLE_API_KEY」,
  // 由 .github/workflows/pages.yml 部署時替換下面的佔位字串。
  GOOGLE_API_KEY: '__GOOGLE_API_KEY__',

  // 僅能存取使用者透過 Picker 選取的檔案,不會取得整個雲端硬碟權限
  DRIVE_SCOPE: 'https://www.googleapis.com/auth/drive.file',

  // 可編輯的工作表:headerRow = 標題列;maxRow = 資料區最後可用列(之後的列不動,例如總表的統計公式)
  // title / subtitle / preview / status / week 為卡片顯示用的欄位標題
  SHEETS: [
    { name: '生管部', headerRow: 5 },
    { name: '業務部', headerRow: 5 },
    { name: '機構設計', headerRow: 5 },
    { name: '電氣設計', headerRow: 5 },
    { name: '管理部', headerRow: 5 },
    { name: '太陽能追蹤', label: '太陽能', headerRow: 5 },
    { name: '總表', headerRow: 5, maxRow: 13 },
    { name: '行動追蹤', headerRow: 3 },
    { name: '專案主檔', headerRow: 3 },
    { name: '權限', headerRow: 1, adminOnly: true },
  ],

  // 個人權限表:檔案內的隱藏工作表(email 不寫在公開的程式碼裡)
  // A 欄 E-MAIL,B 欄 可閱讀及編輯的工作表(以「、」或逗號分隔,「所有工作表」= 全部並可管理權限)
  PERMISSION_SHEET: '權限',
  ALL_SHEETS_KEYWORD: '所有工作表',
  DEFAULT_SHEET: '生管部',
};
