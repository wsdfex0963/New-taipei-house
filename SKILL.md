---
name: new-taipei-house
description: >
  新北市建案看屋檢查清單產生器。針對使用者鎖定的預售/新成屋建案（目前 7 案，含一個台北內湖案），
  依「47 項看屋檢查清單」逐項搜尋公開資訊與實價登錄，輸出五份格式化 Excel 比較表，
  並自動上傳至 Google Drive「看房」資料夾。每週五 09:00（台北）排程自動執行。

  ALWAYS trigger this skill for prompts containing: "看屋檢查清單", "新北建案",
  "看房比較", "看屋表", "更新看屋", "建案比較", "預售屋比較", or any project name
  in the active list (民生新埔, 敦南之森, 新濠漾, 新濠岳, 汐止星野之森, 文德好境, 將捷MRT／景平站).
  Also trigger when the user references the file name pattern
  "看屋檢查清單_新北建案_*.xlsx", or when the weekly Routine「New taipei house 每週看屋清單更新」fires.
---

# 新北市建案看屋檢查清單：執行指南

## 任務目標

針對 `references/projects.md` 的建案，逐項填寫 `references/checklist.md` 的 47 項檢查清單，
產出**五份**格式化 .xlsx，並自動上傳至 Google Drive「看房」資料夾
（資料夾 ID：`12wETI6GI8F5arzLwg7ZkMXd5K5P4Swi-`）。

**核心設計原則：**
- **新增/刪除建案、修改看屋條件**只需編輯 `references/projects.md`，不必動本檔或腳本
- **格式由腳本固定**（深藍標題列、淺藍分類列、未知欄位紅字、超出看屋範圍藍字粗體）
- **以上一份 JSON 為起點做增量更新**，不要每次從零開始
- **必須拆成五份 + 加 `--slim` 上傳**（原因見 Step 5，這不是選配）

---

## Step 1：載入起始資料

依下列優先順序取得 baseline：

1. **repo 內 `data/projects_data_*.json`** 取日期最新的一份 → 直接當起點（最省時，格式已對齊）
2. **Google Drive「看房」資料夾**有上一輪的 xlsx → 用 `Google Drive:search_files`
   （`query: "title contains '看屋檢查清單_新北建案_' and parentId = '12wETI6GI8F5arzLwg7ZkMXd5K5P4Swi-'"`）
   讀回內容當 baseline
3. **都沒有** → 從空白起跑，全部標「未知待查詢」再逐步補

本輪只需更新「未知待查詢」欄位、近期新公告（新價格、新公設、交屋進度）與實價登錄最新數字。

---

## Step 2：讀取參考檔

- `references/projects.md` — 建案清單（地址、捷運、**看屋房型條件 `room_filter`**）
- `references/checklist.md` — 47 項檢查清單（6 大分類）
- `references/sources.md` — 搜尋來源優先順序

`room_filter` 是**使用者的看屋篩選條件**，不是建案本身的產品限制。填表時聚焦符合條件的房型；
超出範圍的房型用 `**...**` 標記為「不在看屋範圍」，不要整段刪除（比價時仍有參考價值）。

目前使用者的看屋條件（詳見 projects.md）：
- **文德好境**（台北內湖）只看 1 房
- **汐止星野之森** A~G 全區都看，以 1 房為主、最多看到 2 房
- **將捷MRT**（中和景平站捷運聯開共構宅）注意僅部分樓層可售，其餘為捷運局分回

---

## Step 3：搜尋資料（重點在實價登錄）

對每個建案**平行搜尋**。實價登錄是本表最有價值的部分，務必多平台交叉比對：

### 搜尋順序
1. 「建案名 + 地區」一輪 `WebSearch` → 找出 591／樂居／樂屋網／5168／住展的頁面
2. 「建案名 + 實價登錄 + 成交 + 坪數 + 總價 + 車位 + 公設比」專項搜尋
   → **多平台數字常有出入，全部列出並註明來源**（例：591 均價76萬、樂屋網73.7萬）
3. 「建案名 + 學區 + 生活機能 + 嫌惡設施 + 建材 + 評價」補搜尋
   → Mobile01 與看屋部落格常有負評與缺點，一併收錄
4. 分期／分區建案（如星野之森 A~G）**逐區查**，各區價格與格局差異大
5. 捷運聯開共構案（如將捷MRT）**額外查**：可售戶數 vs 捷運局分回戶數、共構噪音/低頻震動、
   管理費分攤結構——這三項是共構宅特有的風險，一定要單獨列出

### 填寫規則
- 數字要具體：坪數、公尺數、距離、價格區間、實登筆數
- 多項資訊用 `\n` 換行（腳本已開自動換行）
- 超出 `room_filter` 的房型資訊用 `**xxx**` 包起來 → 腳本轉成藍字粗體
- 找不到就寫「未知待查詢」→ 腳本轉紅字。**絕不編造數字或事實**
- 兩份來源矛盾時**兩個都寫並標註**（例：完工年份 2019／2022 兩說，需調謄本）

### 可以合理推算、不算編造的欄位
以下標明「【試算】」後可填，不必留白：
- **貸款成數與利率試算**：依當時利率行情（先搜尋確認）× 各案總價計算月付金
- **稅費估算**：契稅=房屋評定現值×6%、印花稅0.1%、代書費1.5~2萬、規費0.1%
- **保固條款**：依法結構15年、固定建材設備1年（自交屋日起算），新成屋同樣適用
- **履約保證**：預售屋法定五擇一（價金返還保證／價金信託／不動產開發信託／同業連帶擔保／公會連帶保證）
- **斷層帶**：台北盆地列冊活動斷層僅山腳斷層（第二類，樹林—北投—金山一線）

共通的法規與試算前提**寫在 `sources` 欄位**（表格第2列），不要在 6 個欄位重複，
各建案儲存格只留該案專屬的數字。

---

## Step 4：產生五份 .xlsx（一律加 `--slim`）

整理成 JSON 後，用 `--categories` 分五次產出：

```bash
cd <repo 根目錄>
SP=data/projects_data_YYYY-MM-DD.json   # 你整理好的 JSON
D=YYYY-MM-DD

python3 scripts/build_xlsx.py --data $SP --slim --categories 1,2 \
  --output "output/看屋檢查清單_新北建案_${D}_一_基本與建商.xlsx"
python3 scripts/build_xlsx.py --data $SP --slim --categories 3 \
  --output "output/看屋檢查清單_新北建案_${D}_二_產品與坪數.xlsx"
python3 scripts/build_xlsx.py --data $SP --slim --categories 4 \
  --output "output/看屋檢查清單_新北建案_${D}_三_生活機能與交通.xlsx"
python3 scripts/build_xlsx.py --data $SP --slim --categories 5 \
  --output "output/看屋檢查清單_新北建案_${D}_四_環境風險.xlsx"
python3 scripts/build_xlsx.py --data $SP --slim --categories 6 \
  --output "output/看屋檢查清單_新北建案_${D}_五_財務評估.xlsx"
```

五份都保留**全部建案欄位**，只是各切一段項目，橫向比較不受影響。

`--slim` 會移除 xlsx 內的 theme 部件（並把 styles.xml 唯一的 theme 參照換成實色，
避免懸空參照讓 Excel 判定損毀），檔案縮小約 18%，顯示效果完全一樣。

產完先量尺寸，確認每份的 base64 長度都在上限內：

```bash
for f in output/*_${D}_*.xlsx; do
  echo "$(basename "$f") size=$(stat -c%s "$f") b64=$(base64 -w0 "$f" | wc -c)"
done
```

JSON 結構詳見 `scripts/build_xlsx.py` 檔頭註解；`data/projects_data_2026-08-24.json`
是可直接參考的完整範例（7 案 × 47 項）。

---

## Step 5：上傳 Google Drive（必須分五份）

依序上傳五個檔案，**一次回應只傳一份**，每份都用 `Google Drive:create_file`：
- `parentId`: `12wETI6GI8F5arzLwg7ZkMXd5K5P4Swi-`
- `contentMimeType`: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
- `disableConversionToGoogleType`: `true`（**必加**，否則被轉成 Google 試算表，顏色與格式全失）
- `base64Content`: 該檔的 base64（用 `base64 -w0 <檔案>` 取得）

**每次上傳後核對回傳的 `fileSize` 與本機 `stat -c%s` 是否完全相同。**
不一致代表傳輸掉字元 → 用 `Google Drive:trash_file` 刪掉再重傳。

### ⚠️ fileSize 相符「不等於」檔案正確（2026-08-26 實例）
base64 若被替換掉某幾個字元（而非增減），位元組數不變、fileSize 一樣過關，
但解出來的 zip 內容已毀，Excel 打不開。曾有一份檔案 fileSize 完全相符卻無法開啟。

要真正確認內容，用 `Google Drive:read_file_content` 讀回來看有沒有正常表格文字。
**但務必注意：Drive 的文字擷取是非同步索引的，剛上傳的檔案讀回來一律是空字串
（`{"fileContent":""}`），這代表「還沒索引完」，不代表檔案壞掉。**
至少等 15~20 分鐘後再讀，才有判讀價值。

（2026-08-26 曾因為把「剛上傳讀到空字串」誤判成「檔案損毀」，反覆改寫腳本、
上傳十幾份測試檔在追一個不存在的 bug。切勿重蹈覆轍。）

若要立刻確認，改用 `download_file_content` 把檔案抓回來、寫回磁碟後與本機
`md5sum` 比對——這是唯一即時且可靠的驗證方式。

五份都上傳成功後，把上一輪日期的舊檔用 `trash_file` 清掉，避免資料夾混淆。

### ⚠️ base64 傳輸上限（2026-08-23／08-24 兩次實測）
Drive 連接器只接受「內嵌的 base64 檔案內容」，沒有本機路徑上傳，
base64 必須由模型逐字元輸出，太長會被輸出上限截斷或掉字元。

| 實測 | 結果 |
|---|---|
| 24,000 字元（單一 18KB 檔） | 連續四次失敗（掉 1 byte / 整段遺失 / invalid base64） |
| 17,204 字元 | 兩次被擋下「not a valid base64 string」 |
| 14,128 字元 | 輸出到約 11,700 字元被截斷 |
| **10,000~12,000 字元** | **穩定成功** |

**目標：每份 ≤ 12,000 字元 base64（約 9KB 檔案）。**
對話越長可用輸出額度越少，寧可拆更細也不要冒險。
若建案再增加使檔案變大，就繼續往下拆（例如把分類 1 與 2 也分開）。

上傳完成後：
- **互動對話**：另外呼叫 `SendUserFile` 讓使用者也能直接下載，並回報各案更新重點
- **排程執行**：不需 `SendUserFile`；把五個 `viewUrl` 與本輪更新摘要寫進任務結論

---

## Step 6：留存到 repo

把本輪的 JSON 存成 `data/projects_data_YYYY-MM-DD.json`、五份 xlsx 放 `output/`，
一起 commit 並 push 回 `claude/new-taipei-house-scraper-vx1gm1` 分支，
下一輪即可直接沿用（Step 1 的第 1 順位）。

---

## 輸出格式規格（腳本已固定）

- **第 1 列**：標題「新北市建案 看屋檢查清單」深藍 #1F3864 底白字、16pt 粗體、置中、合併
- **第 2 列**：搜尋日期 + 來源清單 + 共通法規/試算說明
- **第 3 列**：表頭（深藍底白字），每欄含「建案名稱\n地址\n捷運站」
- **分類列**：淺藍 #BDD7EE 底加粗
- **資料列**：行高 45、自動換行、上對齊；紅字=未知待查詢、藍字粗體=超出看屋範圍
- **欄寬**：A=5、B=18、C 起每案 38；凍結窗格 C4
- **檔名**：`看屋檢查清單_新北建案_YYYY-MM-DD_{一_基本與建商|二_產品與坪數|三_生活機能與交通|四_環境風險|五_財務評估}.xlsx`

---

## 常見錯誤排除

| 症狀 | 原因 | 解法 |
|---|---|---|
| 上傳後 fileSize 與本機不符 | base64 傳輸掉字元 | 刪掉壞檔重傳；若反覆失敗就把該份再拆小 |
| fileSize 相符但檔案打不開 | base64 被替換字元，位元組數不變 | download_file_content 抓回比對 md5，不符就重傳 |
| read_file_content 回傳空字串 | **Drive 尚未完成索引**，非檔案損毀 | 等 15~20 分鐘再讀；要即時驗證就比對 md5 |
| 上傳回「not a valid base64 string」 | 轉錄有誤（好事，系統擋下了沒產生壞檔） | 直接重傳該份 |
| base64 輸出到一半停住 | 單次回應輸出額度不足 | 該份再拆小；一次回應只傳一份 |
| xlsx 開啟後顏色/格式不見 | 被轉成 Google Sheets | 確認有加 `disableConversionToGoogleType: true` |
| Excel 提示檔案需要修復 | `--slim` 移除 theme 但留下懸空參照 | 用腳本內建的 `slim_xlsx()`，勿手動刪 theme1.xml |
| 建案名找不到資料 | 用全名搜尋失敗 | 改「部分名 + 區域」，例如「新濠漾 三重」而非「新濠漾4-英倫公園」 |
| 分區建案價格差很多 | A~G 各區獨立定價 | 逐區查詢並在同一格分行列出 |
| 未知待查詢比例過高 | 只搜了建案官網 | 補搜實價登錄平台 + Mobile01 + 看屋部落格；並填上可試算的欄位 |
| 找不到上一輪資料 | repo 沒 clone 或分支不對 | 確認在 `claude/new-taipei-house-scraper-vx1gm1` 分支且已 pull |

## 目前仍難以取得的欄位（非失誤，屬正常）

建照/使照字號、建商財務評等、糾紛訴訟紀錄、是否有夾層、公車班次、治安統計、
管理費——這些多半要到接待中心索取或調謄本才有。
標「未知待查詢」即可，並在回報時告訴使用者這些需現場詢問。

目前 7 案 × 47 項 = 329 欄，未知率約 25%。低於 30% 屬正常水準；
若某一輪突然衝高，通常代表實價登錄平台沒搜到，回 Step 3 補搜。
