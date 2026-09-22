# 台股波動率輪換策略儀表板

這是一個純前端的靜態網頁（`index.html`），會在載入時自動讀取同一個資料夾內的
`data/latest.xlsx`，在瀏覽器裡直接解析並畫出所有圖表與表格 —— **不需要任何後端伺服器**，
很適合直接放在 GitHub Pages 上讓其他人瀏覽。

## 檔案結構

```
.
├── index.html        # 儀表板本體（已內嵌 Chart.js / SheetJS，不需要額外下載套件）
├── data/
│   └── latest.xlsx   # 每天更新的資料來源（「VIX期貨數據」工作表）
└── .nojekyll          # 避免 GitHub Pages 用 Jekyll 處理這個 repo
```

## 部署到 GitHub Pages（第一次設定）

1. 在 GitHub 上新增一個 repository（Public 或 Private 皆可，Private 需要 GitHub Pro 才能開 Pages）。
2. 把這個資料夾裡的 `index.html`、`data/latest.xlsx`、`.nojekyll` 三個檔案上傳到 repo
   （可以直接在 GitHub 網頁上用「Add file → Upload files」拖曳上傳，或用 `git push`）。
3. 到 repo 的 **Settings → Pages**：
   - Source 選擇 `Deploy from a branch`
   - Branch 選擇你剛剛推上去的分支（通常是 `main`），資料夾選 `/ (root)`
   - 存檔後 GitHub 會給一個網址，通常是
     `https://<你的帳號>.github.io/<repo名稱>/`
4. 等 1～2 分鐘讓 GitHub Pages 建置完成，打開那個網址就能看到儀表板。

## 之後每天怎麼更新資料

**只要把新的 Excel 檔案取代掉 repo 裡的 `data/latest.xlsx`（檔名要完全一樣）：**

- 最簡單的方式：到 GitHub 上該檔案的頁面，點右上角鉛筆／「...」→
  「Upload files」，把當天的「台股波動率輪換策略」Excel 檔案拖進去覆蓋掉 `data/latest.xlsx`，
  然後 Commit。
- 或是本機用 `git` 的話：把新檔案複製過去覆蓋 `data/latest.xlsx` 後
  `git add data/latest.xlsx && git commit -m "更新每日資料" && git push`。

推上去之後，**所有訪客下次打開網頁都會自動看到最新資料**，不需要重新產生或上傳 `index.html`。
GitHub Pages 對靜態檔案有短暫的 CDN 快取（通常幾分鐘內會更新），所以更新後若沒有馬上生效，
稍等一下或強制重新整理（Ctrl/Cmd+Shift+R）即可。

## 「本機預覽 Excel」按鈕是做什麼的？

（此功能已移除，改成單純自動讀取 `data/latest.xlsx`。要更新畫面只需要覆蓋該檔案並推送到
GitHub，見上一節。）

## 效能

- `index.html` 本身只有約 50KB：Chart.js、日期套件與 SheetJS 都改成從 jsDelivr（npm 鏡像）
  以 `<script src>` 外部載入，而不是直接內嵌進頁面。瀏覽器會把這些函式庫快取起來，之後同一
  裝置上不管開這個網站幾次，甚至開到其他也用同版本 Chart.js 的網站，都不用重新下載，第一次
  進站之後會明顯變快。**唯一的取捨是網頁本身需要能連上網路（含 jsDelivr）才能正常運作**，
  不能再單獨存成 .html 檔離線開啟——這點跟前面提到的獨立版（ver1）不同，ver1 仍然是完全
  離線可用的單一檔案。
- 讀取 `data/latest.xlsx` 時改用 `cache: 'no-cache'`（會確認檔案是否有更新，但沒更新的話會
  直接用瀏覽器快取，不用整個 5～6MB 重新下載），並且只解析「VIX期貨數據」這張工作表，跳過
  活頁簿裡其餘僅供 Excel 內部圖表使用的工作表，減少不必要的解析時間。
- 圖表與下方歷史資料表格預設只顯示「近1年」，其餘 20 多年的資料仍然完整保留、按一下「全部」
  就能看到，只是預設不會一次把幾千個資料點都畫出來，讓第一次載入更快。
- 如果想再進一步加速，另一個可以做的調整是：`data/latest.xlsx` 目前是完整的 Excel 活頁簿
  （含你自己在 Excel 裡看的原生圖表工作表），這些圖表網頁完全用不到，卻會讓檔案變大、
  下載/解析變慢。如果你在存檔前另存一份「只保留『VIX期貨數據』工作表」的版本上傳成
  `data/latest.xlsx`，檔案可以縮小不少，但這需要你每天多一個手動步驟，所以目前沒有預設這樣做。

## 資料表格式需求

`data/latest.xlsx` 必須包含一個名為「VIX期貨數據」的工作表，欄位配置需與原始「台股波動率輪換
策略」範本一致（日期、VIX 期貨 M1–M7、VIX9D/VIX/VIX3M/VIX6M/VIX1Y、Roll Yield、0050 相關欄位、
策略持倉欄位等）。只要每天用同一份範本另存新的一天資料，就可以直接覆蓋使用。
