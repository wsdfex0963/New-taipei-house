# New Taipei House — 建案看屋檢查清單自動化

每週自動搜尋鎖定建案（見 `references/projects.md`）的公開資訊與實價登錄，
依 47 項檢查清單（`references/checklist.md`）整理成三份格式化 Excel，
並自動上傳到 Google Drive「看房」資料夾。

## 檔案結構

- `SKILL.md` — 完整執行指南（搜尋 → 填表 → 產生三份 Excel → 上傳 Drive → 回存 repo）
- `references/projects.md` — 追蹤的建案清單與**看屋房型條件**（要增減建案改這裡）
- `references/checklist.md` — 47 項檢查清單內容（6 大分類）
- `references/sources.md` — 搜尋來源優先順序（591、樂居、樂屋網、5168、住展…）
- `scripts/build_xlsx.py` — 把整理好的 JSON 轉成格式化 `.xlsx`，`--categories` 可只輸出指定分類
- `data/projects_data_*.json` — 每輪的完整資料，下一輪的增量更新起點
- `output/` — 每輪產出的 Excel（本機留存，正式檔案在 Google Drive「看房」資料夾）

## 每週排程

Claude Code Scheduled Task（Routine）每週五 09:00（台北時間）觸發，流程為：

1. `git pull` 本分支，以 `data/` 最新 JSON 當 baseline
2. 針對每個建案搜尋最新實價登錄與各項資訊，更新 47 項清單
3. 用 `--categories 1,2,3` / `4,5` / `6` 產出三份 `.xlsx`
4. 三份都上傳 Google Drive「看房」資料夾（ID `12wETI6GI8F5arzLwg7ZkMXd5K5P4Swi-`），
   逐份核對檔案大小確認未損壞
5. JSON 與 xlsx commit 回本分支

不需人工介入；若上傳失敗會在該次任務摘要中說明原因與檔案留存位置。

## 為什麼是三份而不是一份

Google Drive 連接器只接受內嵌的 base64 檔案內容，沒有本機路徑上傳。
實測單一 18KB 檔案（約 24,000 字元 base64）無法穩定完整傳輸，
可靠上限約 16,000 字元（約 12KB）。三份各自保留全部建案欄位，只切分項目，
橫向比較不受影響。建案增加導致檔案再變大時，就再往下拆。

## 手動觸發

對話中說「更新看屋清單」或提到任一建案名稱即可依 `SKILL.md` 手動跑一次。

## 本機產生 Excel

```bash
pip install -r requirements.txt
python3 scripts/build_xlsx.py --data data/projects_data_2026-08-23.json \
  --categories 1,2,3 --output "output/看屋檢查清單_新北建案_YYYY-MM-DD_上_基本與產品.xlsx"
```

`--categories` 省略則輸出全部六大分類（單一檔案，適合本機使用，不適合上傳 Drive）。
JSON 格式見 `scripts/build_xlsx.py` 檔頭註解。
