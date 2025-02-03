#Requires AutoHotkey v2.0
#SingleInstance Force

; 存儲文字片段的全局變量
global textSnippets := Map()
global hotstringEnabled := true
; 定義設定檔路徑
global settingsFile := A_ScriptDir "\textEx_settings.ini"


; 讀取用戶設定（在 GUI 初始化時）
LoadUserSettings()
{
    global settingsFile, sendRadio, pasteRadio
    
    ; 如果設定檔不存在，創建預設設定
    if !FileExist(settingsFile) {
        try {
            IniWrite("Ctrl+V", settingsFile, "OutputMethod", "DefaultMethod")
        }
    }
    
    ; 讀取設定
    try {
        method := IniRead(settingsFile, "OutputMethod", "DefaultMethod", "Ctrl+V")
        if (method = "Send") {
            sendRadio.Value := 1
        } else {
            pasteRadio.Value := 1
        }
    }
}

; 保存用戶設定（當選擇改變時）
SaveUserSettings(*)
{
    global settingsFile, sendRadio
    
    try {
        method := sendRadio.Value ? "Send" : "Ctrl+V"
        IniWrite(method, settingsFile, "OutputMethod", "DefaultMethod")
    }
}

; 創建主視窗
mainGui := Gui("+Resize +MinSize800x600", "Text Expander")
mainGui.SetFont("s10", "Segoe UI")

GuiResize(thisGui, MinMax, Width, Height) {
    if MinMax = -1  ; 視窗最小化
        return
    
    ; TreeView 固定 150 寬度
    treeWidth := 180
    listWidth := Width - treeWidth - 20  ; 總寬度減去 TreeView 寬度和邊距
    controlHeight := Height - 100  ; 減去按鈕區域的高度
    
    ; 更新控件大小和位置
    TV.Move(5, 5, treeWidth - 10, controlHeight)
    LV.Move(treeWidth + 5, 5, listWidth - 15, controlHeight)

    ; 動態調整欄寬
    LV.ModifyCol(1, 100)
    LV.ModifyCol(2, listWidth - 120)  ; 總寬度減去 Keyword 寬度和一些邊距
    
    ; 更新按鈕位置
    buttonsY := controlHeight + 10
    addButton.Move(5, buttonsY)
    editButton.Move(85, buttonsY)
    deleteButton.Move(135, buttonsY)
    toggleButton.Move(225, buttonsY)
    
    ; 更新 radio 按鈕位置
    radioY := buttonsY + 35
    mainGui.GetPos(,, &guiWidth)
    outputMethodText.Move(15, radioY)
    sendRadio.Move(85, radioY)
    pasteRadio.Move(185, radioY)
    
    ; 更新分組按鈕位置
    groupButtonsY := radioY + 20
    addGroupButton.Move(5, groupButtonsY)
    deleteGroupButton.Move(105, groupButtonsY)

    WinRedraw(mainGui)
}

; 註冊視窗大小改變事件
mainGui.OnEvent("Size", GuiResize)


; 左側樹狀視圖
TV := mainGui.Add("TreeView", "x5 y5 w180 h400")

; 右側列表視圖
LV := mainGui.Add("ListView", "x195 y5 w545 h400", ["Keyword", "Phrase"])
LV.ModifyCol(1, 100)  ; Keyword 欄寬 150
LV.ModifyCol(2, 380)  ; Phrase 欄寬 380

; 添加按鈕
addButton := mainGui.Add("Button", "x5 y410", "Add New")
editButton := mainGui.Add("Button", "x85 y410", "Edit")
deleteButton := mainGui.Add("Button", "x135 y410", "Delete")
toggleButton := mainGui.Add("Button", "x225 y410", "ON")

; 在 GUI 初始化時添加 Radio 按鈕（放在 toggleButton 之後）
; 在 GUI 初始化時添加 Radio 按鈕並設置事件處理
; 輸出方式選擇 (初始位置)
outputMethodText := mainGui.Add("Text", "x15 y450", "輸出方式：")
sendRadio := mainGui.Add("Radio", "x85 y450", "Send Text")
pasteRadio := mainGui.Add("Radio", "x185 y450", "Ctrl+V")

; 分組按鈕 (初始位置)
addGroupButton := mainGui.Add("Button", "x5 y490", "New Group")
deleteGroupButton := mainGui.Add("Button", "x105 y490", "Delete Group")

addGroupButton.OnEvent("Click", AddNewGroup)
deleteGroupButton.OnEvent("Click", DeleteGroup)

; 添加事件處理
sendRadio.OnEvent("Click", SaveUserSettings)
pasteRadio.OnEvent("Click", SaveUserSettings)

; 載入設定
LoadUserSettings()

; 初始化一些示例資料夾和項目
rootFolder := TV.Add("Email Templates")
thanksFolder := TV.Add("Thanks", rootFolder)
TV.Add("Messages", rootFolder)



; 開關按鈕事件處理
ToggleHotstrings(*)
{
    Suspend  ; 切換暫停狀態
    
    ; 更新按鈕文字和顏色
    toggleButton.Text := A_IsSuspended ? "OFF" : "ON"
    toggleButton.Opt(A_IsSuspended ? "+cRed" : "+cGreen")
    
    TrayTip("Text Expander", A_IsSuspended ? "自動替換已禁用" : "自動替換已啟用")
}

; 設置開關按鈕事件
toggleButton.OnEvent("Click", ToggleHotstrings)

; 修改 RegisterHotstrings 函數，加入完整的錯誤處理
RegisterHotstrings(key)
{
    if !textSnippets.Has(key)
        return
        
    text := textSnippets[key]
    
    ; 先嘗試解除註冊（如果存在的話）
    try {
        try Hotstring(":*:" key " ",, "Off")
        try Hotstring(":*:" key ".",, "Off")
    }
    
    Sleep(50)  ; 給系統一點時間處理
    
    ; 註冊新的熱字串
    try {
        Hotstring(":*:" key " ", (*) => SendWithIMEControl(text))
        Hotstring(":*:" key ".", (*) => SendWithIMEControl(text))
    } catch as e {
        TrayTip("Text Expander", "警告：註冊熱字串時發生錯誤")
    }
}

; 修改 UnregisterHotstrings 函數，加入完整的錯誤處理
UnregisterHotstrings(key)
{
    if (key = "")
        return
        
    try {
        try Hotstring(":*:" key " ",, "Off")
    }
    try {
        try Hotstring(":*:" key ".",, "Off")
    }
    Sleep(150)
}

; 添加新片段時，需要選擇分組
AddNewSnippet(*)
{
    ; 檢查是否有選擇分組
    selectedNode := TV.GetSelection()
    if !selectedNode {
        MsgBox("請先選擇一個分組", "提示")
        return
    }
    
    ; 如果選擇的是子節點，獲取其父節點
    if TV.GetParent(selectedNode)
        selectedNode := TV.GetParent(selectedNode)
    
    groupName := TV.GetText(selectedNode)
    
    addGui := Gui()
    addGui.Owner := mainGui
    addGui.SetFont("s10", "Segoe UI")
    
    addGui.Add("Text",, "Group:")
    addGui.Add("Text",, groupName)
    
    addGui.Add("Text", "xm y+10", "Keyword:")
    keywordEdit := addGui.Add("Edit", "w200")
    
    addGui.Add("Text", "xm y+10", "Phrase:")
    phraseEdit := addGui.Add("Edit", "w300 h100")
    
    saveButton := addGui.Add("Button", "xm y+10", "Save")
    
    SaveNewHandler(*)
    {
        keyword := keywordEdit.Value
        phrase := phraseEdit.Value
        
        if (phrase != "") {  ; 只檢查 phrase 是否有內容
            if (keyword == "") {
                ; 為 memo 生成唯一識別碼
                keyword := "memo_" A_Now "_" Random(1000, 9999)
                ; 直接使用識別碼作為顯示文字
                displayText := keyword
            } else {
                displayText := keyword
            }
            
            textSnippets[keyword] := phrase
            LV.Add(, displayText, phrase)
            TV.Add(displayText, selectedNode)
            
            ; 只有當不是 memo_ 開頭的才註冊熱字串
            if !InStr(keyword, "memo_")
                RegisterHotstrings(keyword)
                
            SaveSnippets()
            addGui.Destroy()
        }
    }
    
    saveButton.OnEvent("Click", SaveNewHandler)
    addGui.Show()
}



; 添加新增分組功能
AddNewGroup(*)
{
    groupName := InputBox("請輸入分組名稱:", "新增分組")
    
    if !groupName.Result  ; 用戶按了取消
        return
        
    if groupName.Value != ""  ; 確保輸入不是空白
        TV.Add(groupName.Value)
}

; 添加刪除分組功能
DeleteGroup(*)
{
    ; 獲取選中的節點
    selectedNode := TV.GetSelection()
    if !selectedNode {
        MsgBox("請先選擇要刪除的分組", "提示")
        return
    }
    
    ; 確保選中的是分組（根節點）而不是子項目
    if TV.GetParent(selectedNode) {
        MsgBox("請選擇分組進行刪除（不能選擇分組內的項目）", "提示")
        return
    }
    
    ; 獲取分組名稱
    groupName := TV.GetText(selectedNode)
    
    ; 確認刪除
    if (MsgBox("確定要刪除分組 " groupName " 及其所有項目嗎？", "確認刪除", "YesNo") = "No")
        return
    
    ; 刪除分組內所有項目的熱字串
    childNode := TV.GetChild(selectedNode)
    while childNode {
        key := TV.GetText(childNode)
        if textSnippets.Has(key) {
            UnregisterHotstrings(key)
            textSnippets.Delete(key)
            ; 從 ListView 中也刪除
            Loop LV.GetCount() {
                if (LV.GetText(A_Index, 1) = key) {
                    LV.Delete(A_Index)
                    break
                }
            }
        }
        childNode := TV.GetNext(childNode)
    }
    
    ; 刪除 TreeView 中的分組節點
    TV.Delete(selectedNode)
    
    ; 保存更改
    SaveSnippets()
}

; 編輯片段視窗
EditSnippet(*)
{
    if (LV.GetCount("Selected") = 0) {
        MsgBox("請先選擇要編輯的項目", "提示")
        return
    }

    row := LV.GetNext()
    oldKeyword := LV.GetText(row, 1)
    oldPhrase := LV.GetText(row, 2)
    
    ; 驗證當前熱字串狀態
    if !textSnippets.Has(oldKeyword) {
        MsgBox("警告：找不到原有的熱字串，建議重新啟動程式", "錯誤")
        return
    }
    
    editGui := Gui()
    editGui.Owner := mainGui
    editGui.SetFont("s10", "Segoe UI")
    
    editGui.Add("Text",, "Keyword:")
    keywordEdit := editGui.Add("Edit", "w200", oldKeyword)
    
    editGui.Add("Text", "xm y+10", "Phrase:")
    phraseEdit := editGui.Add("Edit", "w300 h100", oldPhrase)
    
    saveButton := editGui.Add("Button", "xm y+10", "Save")

    ; 修改 SaveEditHandler
    SaveEditHandler(*)
    {
        keyword := keywordEdit.Value
        phrase := phraseEdit.Value

        if (keyword != "" && phrase != "") {
            ; 取消舊的熱字串
            if (oldKeyword != keyword) {
                MsgBox("正在移除舊的熱字串：" oldKeyword)
                try {
                    Hotstring(":*:" oldKeyword " ",, "Off")
                    Hotstring(":*:" oldKeyword ".",, "Off")
                }
                textSnippets.Delete(oldKeyword)
            }

            ; 更新界面
            LV.Modify(row,, keyword, phrase)

            ; 更新 TreeView
            currentNode := TV.GetNext(0)
            while currentNode {
                childNode := TV.GetChild(currentNode)
                while childNode {
                    if (TV.GetText(childNode) = oldKeyword) {
                        TV.Modify(childNode,, keyword)
                        break
                    }
                    childNode := TV.GetNext(childNode)
                }
                currentNode := TV.GetNext(currentNode)
            }

            ; 更新數據並註冊新的熱字串
            textSnippets[keyword] := phrase
            CreateHotstring(keyword, phrase)

            ; 保存更改
            SaveSnippets()

            ; 顯示最終狀態
            debugInfo := "最終檢查：`n"
            for k, v in textSnippets {
                debugInfo .= k " => " SubStr(v, 1, 30) "`n"
            }
            MsgBox(debugInfo)
        
           editGui.Destroy()
        }
    }

    saveButton.OnEvent("Click", SaveEditHandler)
    editGui.Show()
}

; 刪除片段
DeleteSnippet(*)
{
    if (LV.GetCount("Selected") = 0) {
        MsgBox("請先選擇要刪除的項目", "提示")
        return
    }

    row := LV.GetNext()
    keyword := LV.GetText(row, 1)
    
    if (MsgBox("確定要刪除這個項目嗎？", "確認", "YesNo") = "Yes") {
        ; 從 TreeView 中找到並刪除對應的節點
        currentNode := TV.GetNext(0)  ; 從根節點開始搜索
        while currentNode {
            childNode := TV.GetChild(currentNode)
            while childNode {
                if (TV.GetText(childNode) = keyword) {
                    TV.Delete(childNode)
                    break
                }
                childNode := TV.GetNext(childNode)
            }
            currentNode := TV.GetNext(currentNode)
        }

        UnregisterHotstrings(keyword)
        textSnippets.Delete(keyword)
        LV.Delete(row)
        SaveSnippets()
    }
}

; 保存片段到文件
; 改進的保存函數
SaveSnippets()
{
    snippetFile := A_ScriptDir "\snippets.ini"
    try {
        if FileExist(snippetFile)
            FileDelete(snippetFile)
            
        f := FileOpen(snippetFile, "w", "UTF-8")
        processedKeys := Map()  ; 追蹤已處理的鍵值
        
        currentNode := TV.GetNext()
        while currentNode {
            nodeName := TV.GetText(currentNode)
            f.Write("[" nodeName "]`n")
            
            childNode := TV.GetChild(currentNode)
            while childNode {
                displayText := TV.GetText(childNode)
                actualKey := ""
                
                ; 找到對應的實際 key
                for key, value in textSnippets {
                    if (displayText = "(memo)" && InStr(key, "memo_") && !processedKeys.Has(key)) {
                        actualKey := key
                        processedKeys[key] := true  ; 標記為已處理
                        break
                    } else if (key = displayText && !processedKeys.Has(key)) {
                        actualKey := key
                        processedKeys[key] := true  ; 標記為已處理
                        break
                    }
                }
                
                if (actualKey != "" && textSnippets.Has(actualKey))
                    f.Write(actualKey "=" StrReplace(textSnippets[actualKey], "`n", "<<NEWLINE>>") "`n")
                    
                childNode := TV.GetNext(childNode)
            }
            f.Write("`n")
            currentNode := TV.GetNext(currentNode)
        }
        f.Close()
    }
}


; 載入片段從文件
LoadSnippets()
{
    try {
        snippetFile := A_ScriptDir "\snippets.ini"
        if !FileExist(snippetFile)
            return
            
        TV.Delete()
        
        currentSection := ""
        currentNode := 0
        
        Loop read snippetFile, "UTF-8"
        {
            line := Trim(A_LoopReadLine)
            if (line = "")
                continue
                
            if RegExMatch(line, "^\[(.*)\]$", &match) {
                currentSection := match[1]
                currentNode := TV.Add(currentSection)
                continue
            }
            
            parts := StrSplit(line, "=",, 2)
            if (parts.Length >= 2) {
                key := parts[1]
                decodedValue := StrReplace(parts[2], "<<NEWLINE>>", "`n")
                
                textSnippets[key] := decodedValue
                
                ; 保持原始識別碼顯示
                if currentNode
                    TV.Add(key, currentNode)
                    
                LV.Add(, key, decodedValue)
            }
        }
    }
}

; 完全重新註冊所有熱字串的函數
ReloadAllHotstrings()
{
    ; 先記錄所有現有的 key-value 對
    savedPairs := Map()
    for key, value in textSnippets {
        savedPairs[key] := value
    }

    ; 清除所有現有熱字串
    for key, value in textSnippets {
        try {
            Hotstring(":*:" key " ",, "Off")
            Hotstring(":*:" key ".",, "Off")
        }
    }
    Sleep(100)

    ; 清空目前的 Map
    textSnippets.Clear()

    ; 重新填充 Map 並註冊熱字串
    MsgBox("開始重新註冊熱字串...")
    for key, value in savedPairs {
        ; 先將值放入 Map
        textSnippets[key] := value
        
        ; 然後註冊這個特定的熱字串
        MsgBox("正在註冊：" key " => " SubStr(value, 1, 30))
        
        try {
            ; 為這個特定的 key-value 對創建封裝的函數
            CreateHotstring(key, value)
        }
    }
}

; 輔助函數：為特定的 key-value 對創建熱字串
CreateHotstring(key, value)
{
    ; 創建專用的回調函數
    callback(*) {
        SendWithIMEControl(value)
    }
    
    ; 註冊熱字串
    Hotstring(":*:" key " ", callback)
    Hotstring(":*:" key ".", callback)
}


; 修改 SendWithIMEControl 函數以支援兩種模式
SendWithIMEControl(text)
{
    if (pasteRadio.Value) {
        ; 使用剪貼板方式
        backupClipboard := ClipboardAll()
        A_Clipboard := text
        Sleep(50)
        
        Send("^v")
        
        delay := 50 + (StrLen(text) * 6.5)
        Sleep(delay)
        
        A_Clipboard := backupClipboard
    } else {
        ; 使用原本的 Send 方式
        oldDelay := A_KeyDelay
        oldMode := A_SendMode
        
        SetKeyDelay(0)
        SendMode("Input")
        
        prevIME := DllCall("GetKeyboardLayout", "UInt", DllCall("GetWindowThreadProcessId", "UInt", WinExist("A"), "UInt", 0))
        
        SendMessage(0x50, 0, 0x4090409,, "A")
        Sleep(50)
        
        Send(text)
        
        delay := 50 + (StrLen(text) * 5)
        Sleep(delay)
        
        SetKeyDelay(oldDelay)
        SendMode(oldMode)
        
        PostMessage(0x50, 0, prevIME,, "A")
    }
}

; 在創建 TreeView 後添加選擇事件處理
TV.OnEvent("ItemSelect", TVSelect)

; 添加 TreeView 選擇事件處理函數
TVSelect(*)
{
    ; 清空當前 ListView
    LV.Delete()
    
    ; 獲取選中的節點
    selectedNode := TV.GetSelection()
    if !selectedNode
        return
        
    ; 如果選中的是分組節點
    if !TV.GetParent(selectedNode) {
        ; 獲取該分組下的所有項目
        childNode := TV.GetChild(selectedNode)
        while childNode {
            displayText := TV.GetText(childNode)
            
            ; 尋找對應的實際內容
            for key, value in textSnippets {
                if (displayText = "(memo)" && InStr(key, "memo_")) {
                    LV.Add(, displayText, value)
                    break
                } else if (key = displayText) {
                    LV.Add(, key, value)
                    break
                }
            }
            
            childNode := TV.GetNext(childNode)
        }
    }
    ; 如果選中的是項目節點
    else {
        displayText := TV.GetText(selectedNode)
        
        ; 尋找對應的實際內容
        for key, value in textSnippets {
            if (displayText = "(memo)" && InStr(key, "memo_")) {
                LV.Add(, displayText, value)
                break
            } else if (key = displayText) {
                LV.Add(, key, value)
                break
            }
        }
    }
}

; Ctrl+Alt+S 切換熱字串功能
#HotIf
^!s:: {
    global hotstringEnabled
    hotstringEnabled := !hotstringEnabled
    
    for key, value in textSnippets {
        if hotstringEnabled {
            RegisterHotstrings(key)
        } else {
            UnregisterHotstrings(key)
        }
    }
    
    if hotstringEnabled
        TrayTip("Text Expander", "自動替換已啟用")
    else
        TrayTip("Text Expander", "自動替換已禁用")
}

; 設置按鈕事件
addButton.OnEvent("Click", AddNewSnippet)
editButton.OnEvent("Click", EditSnippet)
deleteButton.OnEvent("Click", DeleteSnippet)

; 載入已保存的片段
LoadSnippets()
mainGui.Show()

; 清除所有可能存在的熱字串，然後重新註冊
for key, value in textSnippets {
    try {
        Hotstring(":*:" key " ",, "Off")
        Hotstring(":*:" key ".",, "Off")
    }
}
Sleep(100)

; 重新註冊所有熱字串
for key, value in textSnippets {
    handler := (originalValue := value) => (*) => SendWithIMEControl(originalValue)
    try {
        Hotstring(":*:" key " ", handler())
        Hotstring(":*:" key ".", handler())
    }
}
