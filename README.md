# text_expander
keyword ==> phrase 

主要功能架構：

1. GUI 相關
    * TreeView 分組顯示
    * ListView 項目顯示
        * 左側為tree_view 右側為list_view
    * 基本操作按鈕


2. 資料操作
    * 熱字串（hotstring）管理
    * Memo 功能（純記錄，不觸發）
    * 檔案讀寫（snippets.ini）


3. 事件處理
    * TVSelect 分組選擇
    * 熱字串註冊/解除註冊
    * 輸入法控制


4. 功能控制
    * 按鈕開關
    * 輸入法自動切換
        * 輸出時自動切換為英文輸出避免錯誤，結束時自動切回原輸入法
    * 錯誤處理機制
