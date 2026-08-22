---
name: new-taipei-house
description: >
  新北市建案看屋檢查清單產生器。針對使用者鎖定的多個新北市預售/新成屋建案，
  依「47 項看屋檢查清單」逐項搜尋公開資訊，輸出格式化 Excel 比較表，並自動上傳至
  Google Drive「看房」資料夾。支援每週排程自動執行（Scheduled Tasks / Routine）。

  ALWAYS trigger this skill for prompts containing: "看屋檢查清單", "新北建案",
  "看房比較", "看屋表", "更新看屋", "建案比較", "預售屋比較", or any project name
  in the active list (民生新埔, 敦南之森, 新濠漾, 新濠岳, 大同新紀元 等).
  Also trigger when the user references the file name pattern "看屋檢查清單_新北建案_*.xlsx",
  or when a scheduled Routine named「New taipei house」fires.
---

# 新北市建案看屋檢查清單：執行指南

## 任務目標

針對 `references/projects.md` 裡列出的新北市建案，逐項填寫 `references/checklist.md`
裡的 47 項檢查清單，產出格式化 .xlsx 報表，並**自動上傳到 Google Drive「看房」資料夾**
（資料夾 ID：`12wETI6GI8F5arzLwg7ZkMXd5K5P4Swi-`）。

**核心設計原則：**
- **新增/刪除建案**只需編輯 `references/projects.md`，不必動 SKILL.md 或腳本
- **格式由腳本固定**（深藍標題列、淺藍分類列、未知欄位紅字、欄寬與列高），確保每週輸出視覺一致
- **資料優先以最近一份成果為起點**（增量更新而非從零開始）
- **自動上傳**：已於 2026-08-22 驗證 Google Drive MCP 對「看房」資料夾具備寫入權限
  （`canAddChildren: true`，owner 為 `wsdfex@gmail.com`），直接呼叫
  `Google Drive:create_file` 上傳即可，**不需要再要求使用者手動拖拉**。
  若未來上傳失敗（例如再次出現 `canAddChildren: false` 或 403），退回舊流程：
  呈現檔案下載連結並提醒使用者手動上傳，同時在回覆中說明失敗原因。

---

## Step 1：載入起始範本

依下列優先順序找出「上一次的 xlsx」當作起點：

1. **使用者本次對話有上傳** `看屋檢查清單_新北建案_*.xlsx` → 用該檔案
2. **Google Drive「看房」資料夾裡有上一份報表** → 用 `Google Drive:search_files`
   （`query: "title contains '看屋檢查清單_新北建案_' and parentId = '12wETI6GI8F5arzLwg7ZkMXd5K5P4Swi-'"`）
   找出最新一份，下載內容當起點
3. **都沒有** → 用空白範本（`scripts/build_xlsx.py` 內建的預設結構）

從起始範本萃取每個建案「上次已填寫的資料」當作 baseline，本週只需更新「未知待查詢」項目
或建案近期新公告（新一輪價格、新公設、新交屋進度）。

---

## Step 2：讀取建案清單與檢查清單

讀取三份參考檔：
- `references/projects.md` — **目前要比較的建案列表**（地址、捷運站、生活圈）
- `references/checklist.md` — 47 項檢查清單（6 大分類）
- `references/sources.md` — 搜尋來源優先順序

**如果使用者要新增/刪除/修改建案**，直接編輯 `references/projects.md` 即可。

---

## Step 3：逐建案搜尋資料

對 `projects.md` 裡每一個建案，搜尋並填寫 47 項清單。**請平行搜尋以節省時間。**

### 搜尋策略
- 每個建案先以「建案名稱 + 地區」做一輪 `web_search`，找出 591、樂居、住展、5168 的頁面
- 接著對重要頁面用 `web_fetch` 取得詳細資料（坪數區間、實價登錄、公設、樓層）
- 再對「實價登錄、公設比、總價區間、車位、嫌惡設施」做專項補搜尋
- 找不到的項目**標記「未知待查詢」**（腳本會自動上紅字），**不要編造**

### 填寫規則
- 數字要具體：坪數、米數、距離、價格區間
- 一個欄位塞入多項資訊時用換行（`\n`），腳本會啟用自動換行
- 大於 2 房（含）的資訊**用「藍字粗體」標記** ← 因為使用者主要鎖定套房/1+1 房
- 「未知待查詢」是合法答案，不要硬猜

---

## Step 4：產生 .xlsx 檔案

將每個建案的資料整理成 JSON 後，呼叫 `scripts/build_xlsx.py`：

```bash
cd <repo根目錄，例如 /home/user/New-taipei-house 或 workdir 內的 clone 路徑>
python3 scripts/build_xlsx.py \
  --data /tmp/projects_data.json \
  --output ./output/看屋檢查清單_新北建案_YYYY-MM-DD.xlsx
```

`projects_data.json` 結構（**詳見 `scripts/build_xlsx.py` 開頭註解**）：
```json
{
  "search_date": "2026-08-24",
  "sources": "591新建案、591實價登錄、樂居、住展、5168比價王...",
  "projects": [
    {
      "name": "民生新埔",
      "address": "新北板橋新埔 民生路三段",
      "metro": "新埔站/民生新埔站(步行5分鐘)",
      "data": {
        "建案名稱": "民生新埔",
        "各房型權狀坪數區間": "套房12坪、1+1房15坪\n**2房18~28坪、3房30~32坪**",
        "...": "..."
      }
    }
  ]
}
```

- 用 `**xxx**` 包住 → 腳本自動轉成 **藍字粗體**（大於 2 房用）
- 字串等於 `未知待查詢` 或開頭包含此字串 → 腳本自動轉成 **紅字**
- 其他文字 → 一般黑字

腳本完成後會印出檔案路徑。

---

## Step 5：上傳到 Google Drive「看房」資料夾 + 呈現檔案

1. 讀取產出的 xlsx 檔內容，呼叫 `Google Drive:create_file` 上傳：
   - `title`: `看屋檢查清單_新北建案_YYYY-MM-DD.xlsx`
   - `parentId`: `12wETI6GI8F5arzLwg7ZkMXd5K5P4Swi-`
   - `base64Content`: 檔案的 base64 內容
   - `contentMimeType`: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
   - `disableConversionToGoogleType`: `true`（保留 .xlsx 格式與腳本的自訂樣式，不要被轉成 Google 試算表，否則顏色/格式會跑掉）
2. **若是使用者互動的對話（非排程自動執行）**，額外呼叫 `present_files` 讓使用者也能直接下載。
3. 回覆使用者：本次已更新哪些建案、哪些欄位還是「未知待查詢」、Google Drive 檔案連結
   （`create_file` 回傳的 `viewUrl`）。

### 自動排程執行（Scheduled Task / Routine 觸發，無人在場確認）

- 不要等待任何使用者確認，直接完整跑完 Step 1~5。
- 上傳成功後**不需要**呼叫 `present_files`（沒有互動使用者可以下載）。
- 若上傳失敗，仍要把產出的 xlsx 留在 `./output/` 底下，並在任務摘要中清楚寫明失敗原因，
  以便下次排程或使用者回來查看時能得知狀況。
- 若這次執行是在 git repo 環境下（例如透過 Claude Code Remote 的排程），完成後可視情況
  把 `references/projects.md`（若有更新 baseline 備註）與 `output/` 內的 xlsx 一併 commit，
  但**不要**強制 push 到與使用者衝突的分支；一般排程任務只需完成 Drive 上傳即可，repo commit
  為選配。

**絕對不要做的事：**
- ❌ 不要把 xlsx 轉成 Google Sheets 格式（會遺失自訂顏色/格式），務必
  `disableConversionToGoogleType: true`
- ❌ 找不到的欄位不要編造數字或事實
- ❌ 排程自動執行時不要卡住等待使用者輸入

---

## 輸出格式規格（腳本已固定，僅供參考）

- **第 1 列**：標題「新北市建案 看屋檢查清單」深藍底白字、16pt 粗體、置中、合併儲存格
- **第 2 列**：資料搜尋日期 + 來源清單
- **第 3 列**：表頭（深藍 #1F3864 底白字），每個建案一欄含「建案名稱\n地址\n捷運站」
- **分類列**：淺藍 #BDD7EE 底加粗（一、基本資料 / 二、建商與代銷背景 / 三、產品規劃與坪數 /
  四、生活機能與交通 / 五、嫌惡設施與環境風險 / 六、財務評估）
- **資料列**：行高 45、自動換行、上對齊
- **欄寬**：A=5（分類）、B=18（項目）、C~G=38（各建案）
- **檔名**：`看屋檢查清單_新北建案_YYYY-MM-DD.xlsx`（YYYY-MM-DD = 當天日期）

---

## 常見錯誤排除

| 症狀 | 原因 | 解法 |
|---|---|---|
| 找不到上週的 xlsx | 使用者沒上傳、Drive 裡也沒有 | 從空白範本起跑，全部當未知待查詢慢慢填 |
| 建案名打錯找不到資料 | 591/樂居用全名搜尋失敗 | 改用部分名 + 區域，例如「新濠漾 三重」而非「新濠漾4-英倫公園」 |
| 寫入欄位卻沒上色 | 沒用 `**xxx**` 包大房資訊、或未知待查詢拼字不對 | 確認標記符號正確 |
| Drive 上傳失敗 / canAddChildren 又變 false | 資料夾權限被改回去、OAuth scope 不足 | 退回舊流程：present_files 呈現下載連結 + 提醒手動拖拉，並在回覆說明失敗原因 |
| xlsx 上傳後顏色/格式不見 | 被自動轉成 Google Sheets | 上傳時務必加 `disableConversionToGoogleType: true` |
