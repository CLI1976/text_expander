# Text Expander

一個基於 AutoHotkey v2 開發的文字擴展工具，用於快速輸入預設的文字片段。

本專案在 Anthropic 的 Claude 3.5 Sonnet (2024.04) 的協助下完成。Claude 提供了完整的程式設計指導、代碼優化建議以及問題排除方案，對專案的開發有重大貢獻。


## 功能特點

- 支持按組管理文字片段（含巢狀子群組）
- 提供兩種輸出模式：直接輸入和剪貼板貼上
- 支持中英文輸入法自動切換，可設定僅在英文輸入法時生效
- 動態標記：片段內容可插入 `{TODAY}`、`{TIME}` 等日期時間標記
- 內建搜尋功能，可搜尋關鍵字或片段內容
- 支持匯入／匯出片段（疊加或取代模式，匯入前自動備份）
- 可以儲存純文字備忘錄（不觸發替換的筆記）
- GUI 介面支持自由縮放，提供列表／預覽兩種顯示模式
- 使用空格、句點或逗號觸發替換，輸出後綴可選「空白」或「觸發字元」
- 可以通過按鈕或快捷鍵（Ctrl+Alt+S）快速開關功能
- 自動保存使用者偏好設定

## 使用方法

1. 管理文字片段：
   - 使用左側樹狀圖管理分組
   - 使用右側列表查看和編輯文字片段
   - 支持新增、編輯、刪除操作

2. 觸發方式：
   - 輸入關鍵字後按空格、句點或逗號即可觸發替換
   - 例如：輸入 "key" 後按空格，會自動替換為預設的文字

3. 輸出模式：
   - Send Text：直接模擬鍵盤輸入
   - Ctrl+V：使用剪貼板貼上（推薦，速度更快；貼上後會自動還原原本的剪貼簿內容）

4. 動態標記（可寫在片段內容中，觸發時自動代入）：

   | 標記 | 說明 |
   |------|------|
   | `{TODAY}` | 今天日期 (yyyy/MM/dd) |
   | `{TODAY_SHORT}` | 今天日期 (yy/MM/dd) |
   | `{YESTERDAY}` / `{TOMORROW}` | 昨天／明天日期 |
   | `{TIME}` / `{TIME_FULL}` | 現在時間 (HH:mm / HH:mm:ss) |
   | `{WEEKDAY}` / `{MONTH}` / `{YEAR}` | 星期／月份／年份 |

## 系統需求

- 作業系統：Windows
- 執行環境：AutoHotkey v2.0 或更高版本
- 函式庫：`LV_Colors.ahk`（已隨附於 `Lib` 資料夾，改寫自 just me 的 LV_Colors 類別）

## 檔案說明

- `text expander.ahk`：主程式檔案
- `Lib\LV_Colors.ahk`：ListView 顏色函式庫
- `snippets.ini`：文字片段儲存檔案
- `textEx_settings.ini`：使用者設定檔案
- `snippets_*.old`：匯入前自動產生的備份檔

## 注意事項

- 建議使用 Ctrl+V 模式，效能較佳
- 可隨時使用 Ctrl+Alt+S 或界面按鈕開關功能
- 視窗大小可自由調整，介面會自動適應
- 關鍵字不可重複、不可包含「=」、不可以「[」開頭；`memo_` 為保留前綴（供備忘項目使用）


![screenshot](Screenshot.png)

If you wish to support me in this and other projects:
[!["Buy Me A Coffee"](https://www.buymeacoffee.com/assets/img/custom_images/orange_img.png)](https://www.buymeacoffee.com/hw98188d)

