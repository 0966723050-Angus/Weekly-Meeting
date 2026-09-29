// ATK 週會表單 - 全域設定
window.APP_CONFIG = {
  // 與「ATK近期工作項目」共用同一個 Google Cloud 專案(worklist050)的憑證,
  // 已授權來源 https://0966723050-angus.github.io
  GOOGLE_CLIENT_ID: '873968217418-q5t3i90e4pf04kbd4l6vbpjib13etb7o.apps.googleusercontent.com',

  // 雲端硬碟權限:由程式直接找出使用者有權限的週會檔案(不需選檔視窗)。
  // 實際能開哪些檔案仍由雲端硬碟的共用設定決定;程式只讀寫 weekly_meeting_template_*.xlsx。
  DRIVE_SCOPE: 'https://www.googleapis.com/auth/drive',

  // 分頁 → 所屬檔案(檔名 weekly_meeting_template_<file>.xlsx)。
  // 誰能開哪個檔案由 Google 雲端硬碟的共用權限決定;沒有權限的檔案,其分頁會反白不可點。
  // headerRow = 標題列;maxRow = 資料區最後可用列(之後的列不動,例如總表的統計公式)
  SHEETS: [
    { name: '生管部', file: '生管部', headerRow: 5 },
    { name: '業務部', file: '業務部', headerRow: 5 },
    { name: '機構設計', file: '機構設計', headerRow: 5 },
    { name: '電氣設計', file: '電氣設計', headerRow: 5 },
    { name: '管理部', file: '管理部', headerRow: 5 },
    { name: '太陽能追蹤', label: '太陽能', file: '管理部', headerRow: 5 },
    { name: '總表', file: '生管部', headerRow: 5, maxRow: 13 },
    { name: '行動追蹤', file: '生管部', headerRow: 3 },
    { name: '專案主檔', file: '生管部', headerRow: 3 },
  ],
  FILE_PREFIX: 'weekly_meeting_template_',

  // 編輯時「專案名稱」下拉選單的來源:生管部檔案「專案主檔」B4 以下
  PROJECT_SOURCE: { file: '生管部', sheet: '專案主檔', range: '專案主檔!$B$4:$B$5000' },

  // 「校正」:以生管部工作表第 6 列的週次/日期為準,改正下列工作表每一列的週次/日期
  CALIBRATE: { file: '生管部', sheet: '生管部', row: 6, sheets: ['生管部', '業務部', '機構設計', '電氣設計', '管理部', '太陽能追蹤'] },

  // 彙整輸出檔名(存在生管部檔案所在的雲端資料夾)
  SUMMARY_NAME: '各部工作彙整',
  DEFAULT_SHEET: '生管部',
};
