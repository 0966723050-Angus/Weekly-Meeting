# ATK 週會表單(Weekly Meeting)

手機友善的 RWD / PWA 網頁工具,直接編輯 Google 雲端硬碟「ATK」工作區中的
`weekly_meeting_template_生管.xlsx`。

## 功能

- 工作表分頁:生管部、業務部、機構設計、電氣設計、管理部、太陽能追蹤、總表、行動追蹤、專案主檔
- 卡片式列表(手機好讀),可依週次篩選、關鍵字搜尋
- 點卡片編輯整列;下拉選單欄位(專案名稱、負責人、狀態…)沿用 Excel 的資料驗證清單;進度% 以滑桿/數字輸入
- **列的新增與刪除**:右下「＋ 新增一列」、每張卡片「⋯」→ 在上方/下方插入、複製為本週的新列、刪除此列
- 「複製為本週的新列」:週次自動改為本週(格式同使用說明頁公式,如 `2026-W40 / 2026-09-28`),
  並把「本週預計工作內容」移到「上週工作內容」,方便每週延續
- 按「儲存」寫回雲端硬碟;未儲存的修改以橘色標示
- 也可開啟本機 Excel 檔,修改後下載

## 個人權限

檔案內的隱藏工作表「權限」決定每個帳號在 App 中可閱讀及編輯的工作表
(A 欄 E-MAIL、B 欄工作表名稱,以「、」分隔;「所有工作表」= 全部並可在 App 內管理「權限」表)。
未列在表中的帳號無法使用。email 不寫在公開的程式碼中。

注意:這是 App 介面層的限制。Google 雲端硬碟的共用權限以整個檔案為單位,
具有檔案編輯權的人直接以 Excel / 雲端硬碟開啟仍可看到所有工作表。

## 寫入方式(為什麼不會弄壞 Excel)

`js/xlsx-model.js` 直接修改 xlsx 內的 XML,只重寫「有修改的工作表」的資料區:

- 每一列連同原本的儲存格樣式、列高一起搬移;新增列套用下方預先格式化的空白列樣式
- 資料區以外(預先格式化的空白列、總表統計公式、下拉選單、條件式格式、註解)完全不動
- 未修改的工作表位元組完全相同
- 修改過的文字以 inline string 寫入,不動 sharedStrings.xml;日期以 Excel 日期序號寫入

## 多人同時編輯

儲存前會比對雲端檔案版本:
- 對方改的是其他工作表 → 自動合併(把我的修改套到最新版)
- 對方也改了同一工作表 → 跳出提示,可選「重新載入」或「仍要覆寫」

注意:若有人在電腦上用 Excel 開著同步資料夾中的同一檔案,Excel 存檔時仍可能覆蓋網頁上的修改。

## Google 設定

與「ATK近期工作項目」共用 Google Cloud 專案 `worklist050` 的 OAuth 用戶端
(已授權來源 `https://0966723050-angus.github.io`,OAuth 同意畫面為「內部」,僅 atk.com.tw 帳號可登入)。

API 金鑰(僅 Google Picker 使用)**不放在原始碼**:存於 GitHub Secret `GOOGLE_API_KEY`,
由 `.github/workflows/pages.yml` 部署時注入(GitHub Pages 來源設為 GitHub Actions)。
建議此金鑰限制為:API 僅「Google Picker API」、網站限制 `https://0966723050-angus.github.io/Weekly-Meeting/*`。

權限範圍 `drive.file`:只能存取使用者在 Google Picker 中選取的那一個檔案。
每位使用者第一次登入時需在 Picker 中選一次檔案,之後會記住。

若要免選檔,可把檔案 ID 填入 `js/config.js` 的 `DRIVE_FILE_ID`(每位使用者仍需透過 Picker 授權一次)。

## 本機測試

```bash
python -m http.server 8793
```

本機(localhost)未列入 OAuth 已授權來源,無法登入 Google;可用「開啟本機 Excel 檔」測試編輯與下載。
