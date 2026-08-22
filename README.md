# New Taipei House — 新北建案看屋檢查清單自動化

每週自動搜尋新北市預售/新成屋建案（見 `references/projects.md`）的公開資訊，
依 47 項檢查清單（`references/checklist.md`）整理成格式化 Excel，
並自動上傳到 Google Drive「看房」資料夾。

## 檔案結構

- `SKILL.md` — 完整執行指南（Claude 依此逐步搜尋資料、產生報表、上傳 Drive）
- `references/projects.md` — 目前追蹤的建案清單（要新增/刪除建案改這裡）
- `references/checklist.md` — 47 項檢查清單內容（6 大分類）
- `references/sources.md` — 搜尋來源優先順序（591、樂居、住展、5168…）
- `scripts/build_xlsx.py` — 把整理好的 JSON 資料轉成格式化 `.xlsx`
- `output/` — 每週產出的 Excel 報表（本機暫存，正式檔案在 Google Drive「看房」資料夾）

## 每週排程

透過 Claude Code 的 Scheduled Tasks（Routine）每週一早上 9:00（台北時間）自動觸發一個
全新 session，內容為：
1. 讀取本 repo 的 `references/*` 與 `SKILL.md`
2. 針對每個建案搜尋最新資訊，更新 47 項檢查清單
3. 呼叫 `scripts/build_xlsx.py` 產生本週 `.xlsx`
4. 透過 Google Drive MCP 自動上傳到「看房」資料夾（資料夾 ID：
   `12wETI6GI8F5arzLwg7ZkMXd5K5P4Swi-`）

不需要人工介入；若上傳失敗會在該次 session 的任務摘要中說明原因。

## 手動觸發 / 測試

在對話中直接說「更新看屋清單」或提及建案名稱，即可依 `SKILL.md` 的步驟手動跑一次。

## 產生 Excel（本機測試腳本）

```bash
pip install -r requirements.txt
python3 scripts/build_xlsx.py --data projects_data.json --output output/看屋檢查清單_新北建案_YYYY-MM-DD.xlsx
```

`projects_data.json` 的格式請見 `scripts/build_xlsx.py` 檔頭註解。
