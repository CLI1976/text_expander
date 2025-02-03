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
    ; Group Path 約佔 30%，Keyword 約佔 20%，Phrase 佔剩餘空間
    pathWidth := 100 ; Round(listWidth * 0.3)  ; 30% 的寬度給路徑
    keywordWidth := 75  ; Round(listWidth * 0.2)  ; 20% 的寬度給關鍵字
    phraseWidth := listWidth - pathWidth - keywordWidth - 20  ; 剩餘空間給內容

    LV.ModifyCol(1, pathWidth)      ; Group Path
    LV.ModifyCol(2, keywordWidth)   ; Keyword
    LV.ModifyCol(3, phraseWidth)    ; Phrase
    
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

    importButton.Move(285, buttonsY)
    exportButton.Move(345, buttonsY)
    searchButton.Move(405, buttonsY)

    WinRedraw(mainGui)
}

; 註冊視窗大小改變事件
mainGui.OnEvent("Size", GuiResize)


; 左側樹狀視圖
TV := mainGui.Add("TreeView", "x5 y5 w180 h400")

; 右側列表視圖
LV := mainGui.Add("ListView", "x195 y5 w545 h400", ["Group Path", "Keyword", "Phrase"])
LV.ModifyCol(1, 150)  ; Group Path 欄寬
LV.ModifyCol(2, 100)  ; Keyword 欄寬
LV.ModifyCol(3, 280)  ; Phrase 欄寬
LV.OnEvent("DoubleClick", (*) => EditSnippet())  ; 添加這行
; 在主視窗建立時，為 ListView 添加點選事件處理
LV.OnEvent("ItemSelect", LVSelect)

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

; 在按鈕區域添加 Import/Export 按鈕
importButton := mainGui.Add("Button", "x285 y410", "Import")
exportButton := mainGui.Add("Button", "x345 y410", "Export")

; 在主視窗添加搜尋按鈕
searchButton := mainGui.Add("Button", "x405 y410", "Search")

; 設置按鈕事件
addButton.OnEvent("Click", AddNewSnippet)
editButton.OnEvent("Click", EditSnippet)
deleteButton.OnEvent("Click", DeleteSnippet)
importButton.OnEvent("Click", ImportSnippets)  
exportButton.OnEvent("Click", ExportSnippets)  
searchButton.OnEvent("Click", ShowSearchWindow)

; 添加事件處理
sendRadio.OnEvent("Click", SaveUserSettings)
pasteRadio.OnEvent("Click", SaveUserSettings)

; 添加事件處理
addGroupButton.OnEvent("Click", AddNewGroup)
deleteGroupButton.OnEvent("Click", DeleteGroup)



; 載入設定
LoadUserSettings()

; 初始化一些示例資料夾和項目
rootFolder := TV.Add("Email Templates")
thanksFolder := TV.Add("Thanks", rootFolder)
TV.Add("Messages", rootFolder)


; ListView 點選事件處理函數
LVSelect(*) {
    ; 取消 TreeView 的選取狀態
    if TV.GetSelection()
        TV.Modify(TV.GetSelection(), "-Select")
}

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
AddNewSnippet(*) {
    ; 檢查是否有選擇分組
    selectedNode := TV.GetSelection()
    if !selectedNode {
        MsgBox("請先選擇一個群組", "提示")
        return
    }
    

    ; 單一的邏輯來判斷目標群組
    targetNode := selectedNode
    
    ; 如果選中的是一個項目（存在於 textSnippets 中），使用其父群組
    if textSnippets.Has(TV.GetText(selectedNode)) {
        targetNode := TV.GetParent(selectedNode)
    }
    
    targetName := TV.GetText(targetNode)
    groupPath := GetFullGroupPath(targetNode)



    ; 獲取主視窗尺寸
    mainGui.GetPos(&mainX, &mainY, &mainWidth, &mainHeight)
    
    ; 計算新視窗尺寸 (主視窗的 0.8 倍)
    addWidth := Round(mainWidth * 0.8)
    addHeight := Round(mainHeight * 0.8)
    
    addGui := Gui("+Owner" . mainGui.Hwnd . " +Resize", "Add New Snippet")
    addGui.SetFont("s10", "Segoe UI")
    
    ; 計算控件尺寸
    contentWidth := addWidth - 40  ; 減去邊距
    keywordWidth := Min(contentWidth - 20, 300)  ; keyword 輸入框限制最大寬度
    phraseWidth := contentWidth - 20
    phraseHeight := addHeight - 170  ; 保留空間給其他控件
    
    ; 調整控件位置和大小
    ; addGui.Add("Text",, "Group:  " groupName)
    addGui.Add("Text",, "Group:  " targetName)
    ; addGui.Add("Text", "xm y+10", "Keyword:")
    ; keywordEdit := addGui.Add("Edit", "w" keywordWidth)
    addGui.Add("Text", "xm", "Keyword:")
    keywordEdit := addGui.Add("Edit", "w" keywordWidth " yp-3")  ; yp-3 微調對齊
    helpButton := addGui.Add("Button", "x+5 yp h23", "?")  ; 緊接在 Edit 後面
    ; 幫助按鈕事件
    helpButton.OnEvent("Click", ShowDynamicTags)

    addGui.Add("Text", "xm y+10", "Phrase:")
    phraseEdit := addGui.Add("Edit", "w" phraseWidth " h" phraseHeight)
    
    saveButton := addGui.Add("Button", "xm y+10", "Save")
    
    ; 添加視窗大小改變事件處理
    AddGuiSize(thisGui, MinMax, Width, Height) {
        if MinMax = -1  ; 視窗最小化
            return
            
        ; 調整 Phrase 編輯框的大小
        newWidth := Width - 40
        newHeight := Height - 170
        phraseEdit.Move(,, newWidth, newHeight)
        
        ; 獲取 phraseEdit 的位置
        phraseEdit.GetPos(&phraseX, &phraseY)
        
        ; 調整 Save 按鈕的位置
        saveButton.Move(, phraseY + newHeight + 10)
        
        ; 強制重繪
        WinRedraw(thisGui)
    }
    
    addGui.OnEvent("Size", AddGuiSize)
    
    SaveNewHandler(*) {
        keyword := keywordEdit.Value
        phrase := phraseEdit.Value
        
        if (phrase != "") {
            if (keyword == "") {
                keyword := "memo_" A_Now "_" Random(1000, 9999)
                displayText := keyword
            } else {
                displayText := keyword
            }
            
            ; 修改：在保存前確認正確的群組路徑
            targetGroupPath := GetFullGroupPath(targetNode)
            
            textSnippets[keyword] := phrase
            ; 修改：在 ListView 中顯示正確的群組路徑
            LV.Add(, targetGroupPath, displayText, phrase)
            TV.Add(displayText, targetNode)
            
            if !InStr(keyword, "memo_")
                RegisterHotstrings(keyword)
                
            SaveSnippets()
            addGui.Destroy()
        }
    }
    
    saveButton.OnEvent("Click", SaveNewHandler)
    
    ; 設定視窗大小並顯示
    addGui.Move(, , addWidth + 40, addHeight)
    addGui.Show()
}

; 輔助函數：獲取完整群組路徑
GetFullGroupPath(node) {
    path := TV.GetText(node)
    parentNode := TV.GetParent(node)
    while parentNode {
        path := TV.GetText(parentNode) "\" path
        parentNode := TV.GetParent(parentNode)
    }
    return path
}

ShowDynamicTags(*) {
    helpGui := Gui("+Owner" . mainGui.Hwnd, "Dynamic Tags Help")
    helpGui.SetFont("s10", "Segoe UI")

    ; 創建一個固定寬度字體的文字區域
    helpText := "
    (
    標記             說明                   範例
    ============================================================
    {TODAY}          今天日期 (yyyy/MM/dd)  Report dated {TODAY}
    {TODAY_SHORT}    今天日期 (yy/MM/dd)    Doc_{TODAY_SHORT}
    {YESTERDAY}      昨天日期               Previous study on {YESTERDAY}
    {TOMORROW}       明天日期               Follow up on {TOMORROW}
    {TIME}           現在時間 (HH:mm)       Reviewed at {TIME}
    {TIME_FULL}      現在時間 (HH:mm:ss)    Timestamp: {TIME_FULL}
    {WEEKDAY}        今天星期名稱           Meeting on {WEEKDAY}
    {MONTH}          今天月份名稱           Report for {MONTH}
    {YEAR}           今天年份               Annual review {YEAR}
    )"

    ; 使用等寬字體的多行編輯框
    helpEdit := helpGui.Add("Edit", "ReadOnly w550 h200", helpText)
    helpEdit.SetFont("s10", "Consolas")  ; 使用等寬字體確保對齊

    ; 添加確定按鈕
    helpGui.Add("Button", "Default w80", "OK").OnEvent("Click", (*) => helpGui.Destroy())

    helpGui.Show()
}

AddNewGroup(*) {
    ; 創建新群組的視窗
    addGroupGui := Gui("+Owner" . mainGui.Hwnd, "新增分組")
    addGroupGui.SetFont("s10", "Segoe UI")
    
    ; 取得所有頂層群組作為選項
    groupChoices := ["頂層"]  ; 使用陣列來存儲選項
    currentNode := TV.GetNext()
    while currentNode {
        if !TV.GetParent(currentNode)
            groupChoices.Push(TV.GetText(currentNode))
        currentNode := TV.GetNext(currentNode)
    }
    
    ; 新增控件
    addGroupGui.Add("Text",, "新群組名稱:")
    nameEdit := addGroupGui.Add("Edit", "w200")
    
    addGroupGui.Add("Text", "xm y+10", "建立在:")
    parentChoice := addGroupGui.Add("DropDownList", "w200", groupChoices)
    
    ; 設定預設選項
    selectedNode := TV.GetSelection()
    if (selectedNode) {
        if (!TV.GetParent(selectedNode)) {
            ; 尋找對應的索引
            for index, text in groupChoices {
                if (text = TV.GetText(selectedNode)) {
                    parentChoice.Value := index
                    break
                }
            }
        } else {
            ; 如果是子節點，找其父節點
            parentText := TV.GetText(TV.GetParent(selectedNode))
            for index, text in groupChoices {
                if (text = parentText) {
                    parentChoice.Value := index
                    break
                }
            }
        }
    } else {
        parentChoice.Value := 1  ; 預設選擇頂層
    }
    
    addGroupGui.Add("Button", "xm y+20 w80", "確定").OnEvent("Click", SaveNewGroup)
    addGroupGui.Add("Button", "x+10 yp w80", "取消").OnEvent("Click", (*) => addGroupGui.Destroy())
    
    SaveNewGroup(*) {
        groupName := nameEdit.Value
        if (groupName = "") {
            MsgBox("請輸入群組名稱", "提示")
            return
        }
        
        ; 根據選擇決定新增位置
        if (parentChoice.Text = "頂層") {
            newNode := TV.Add(groupName)
        } else {
            ; 尋找選擇的父群組節點
            parentNode := TV.GetNext()
            while parentNode {
                if (!TV.GetParent(parentNode) && TV.GetText(parentNode) = parentChoice.Text) {
                    newNode := TV.Add(groupName, parentNode)
                    break
                }
                parentNode := TV.GetNext(parentNode)
            }
        }
        
        SaveSnippets()  ; 立即保存變更
        addGroupGui.Destroy()
    }
    
    addGroupGui.Show()
}

DeleteGroup(*) {
    ; 檢查是否在 ListView 中選擇了項目
    if LV.GetCount("Selected") > 0 {
        ; 確保取消 TreeView 的選擇
        if TV.GetSelection()
            TV.Modify(TV.GetSelection(), "-Select")
        
        MsgBox("欲刪除群組 請在左側樹狀結構點選", "提示")
        return
    }

    ; 獲取選中的節點
    selectedNode := TV.GetSelection()
    if !selectedNode {
        MsgBox("欲刪除群組 請在左側樹狀結構點選群組", "提示")
        return
    }
    
    ; 如果選中的是一般項目（不是群組），提示使用者
    if !TV.GetChild(selectedNode) && !IsEmptyGroup(selectedNode) {
        MsgBox("欲刪除群組 請在左側樹狀結構點選群組", "提示")
        return
    }
    
    ; 獲取完整路徑
    groupPath := GetFullGroupPath(selectedNode)
    
    ; 確認刪除
    if (MsgBox("確定要刪除分組 " groupPath " 及其所有項目嗎？", "確認刪除", "YesNo") = "No")
        return
    
    ; 遞迴刪除所有子項目
    DeleteGroupItems(selectedNode)
    
    ; 刪除群組節點本身
    TV.Delete(selectedNode)  ; 只刪除選中的節點
    
    ; 保存更改
    SaveSnippets()
}

; 輔助函數：遞迴刪除群組內的所有項目
DeleteGroupItems(node) {
    childNode := TV.GetChild(node)
    while childNode {
        nextChild := TV.GetNext(childNode)  ; 保存下一個節點的引用
        
        ; 如果是子群組，先遞迴刪除其內容
        if TV.GetChild(childNode) {
            DeleteGroupItems(childNode)
        }
        ; 如果是項目，刪除對應的熱字串
        else if !IsEmptyGroup(childNode) {
            key := TV.GetText(childNode)
            if textSnippets.Has(key) {
                UnregisterHotstrings(key)
                textSnippets.Delete(key)
            }
        }
        
        childNode := nextChild
    }
}

; 編輯片段視窗
EditSnippet(*) {
    if (LV.GetCount("Selected") = 0) {
        MsgBox("請先選擇要編輯的項目", "提示")
        return
    }

    ; 獲取主視窗尺寸
    mainGui.GetPos(&mainX, &mainY, &mainWidth, &mainHeight)
    
    ; 計算編輯視窗尺寸 (主視窗的 0.8 倍)
    editWidth := Round(mainWidth * 0.8)
    editHeight := Round(mainHeight * 0.8)
    
    row := LV.GetNext()
    oldKeyword := LV.GetText(row, 2)
    oldPhrase := LV.GetText(row, 3)
    
    if !textSnippets.Has(oldKeyword) {
        MsgBox("警告：找不到原有的熱字串，建議重新啟動程式", "錯誤")
        return
    }
    
    ; 創建編輯視窗並設定大小
    editGui := Gui("+Owner" . mainGui.Hwnd . " +Resize", "Edit Snippet")
    editGui.SetFont("s10", "Segoe UI")
    
    ; 計算控件尺寸
    editWidth -= 40  ; 減去邊距
    keywordWidth := Min(editWidth - 20, 300)  ; keyword 輸入框限制最大寬度
    phraseWidth := editWidth - 20
    phraseHeight := editHeight - 150  ; 保留空間給其他控件
    
    ; 調整控件位置和大小
    editGui.Add("Text",, "Keyword:")
    keywordEdit := editGui.Add("Edit", "w" keywordWidth, oldKeyword)
    
    editGui.Add("Text", "xm y+10", "Phrase:")
    phraseEdit := editGui.Add("Edit", "w" phraseWidth " h" phraseHeight, oldPhrase)
    
    saveButton := editGui.Add("Button", "xm y+10", "Save")

    ; 添加視窗大小改變事件處理
    EditGuiSize(thisGui, MinMax, Width, Height) {
        if MinMax = -1  ; 視窗最小化
            return
            
        ; 只調整 Phrase 編輯框的大小
        newWidth := Width - 40
        newHeight := Height - 150
        phraseEdit.Move(,, newWidth, newHeight)

        ; 獲取 phraseEdit 的位置
        phraseEdit.GetPos(&phraseX, &phraseY)
    
        ; 調整 Save 按鈕的位置 (Y值跟著 phrase 編輯框的底部)
        saveButton.Move(, phraseY + newHeight + 10)
            
         ; 強制重繪
        WinRedraw(thisGui)
    }
    
    editGui.OnEvent("Size", EditGuiSize)
    
    SaveEditHandler(*) {
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
    
    ; 設定視窗大小並顯示
    editGui.Move(, , editWidth + 40, editHeight)
    editGui.Show()
}

; 刪除片段
DeleteSnippet(*) {
    if (LV.GetCount("Selected") = 0) {
        MsgBox("請先選擇要刪除的項目", "提示")
        return
    }

    row := LV.GetNext()
    keyword := LV.GetText(row, 2)  ; 取得 Keyword 欄位
    
    ; 檢查是否為空群組
    if (keyword = "(Empty Group)") {
        MsgBox("空群組無法使用此功能刪除，請使用 Delete Group 按鈕", "操作提示")
        return
    }
    
    if (MsgBox("確定要刪除這個項目嗎？", "確認刪除", "YesNo") = "Yes") {
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
SaveSnippets() {
    snippetFile := A_ScriptDir "\snippets.ini"
    try {
        if FileExist(snippetFile)
            FileDelete(snippetFile)
            
        f := FileOpen(snippetFile, "w", "UTF-8")
        
        ; 首先保存頂層群組的直接項目
        currentNode := TV.GetNext()
        while currentNode {
            if !TV.GetParent(currentNode) {  ; 確保是頂層群組
                nodeName := TV.GetText(currentNode)
                f.Write("[" nodeName "]`n")
                
                ; 先寫入該群組直接的項目（不包括子群組的項目）
                childNode := TV.GetChild(currentNode)
                while childNode {
                    itemName := TV.GetText(childNode)
                    ; 只有當項目直接屬於該群組時才寫入
                    if !TV.GetChild(childNode) && textSnippets.Has(itemName) {
                        f.Write(itemName "=" StrReplace(textSnippets[itemName], "`n", "<<NEWLINE>>") "`n")
                    }
                    childNode := TV.GetNext(childNode)
                }
                
                ; 然後處理子群組
                childNode := TV.GetChild(currentNode)
                while childNode {
                    if TV.GetChild(childNode) {  ; 如果是子群組
                        f.Write("`n")  ; 在子群組前添加空行
                        SaveGroupContent(f, childNode, nodeName)
                    }
                    childNode := TV.GetNext(childNode)
                }
                f.Write("`n")  ; 在每個頂層群組之後添加空行
            }
            currentNode := TV.GetNext(currentNode)
        }
        f.Close()
    }
}

SaveGroupContent(f, parentNode, parentPath) {
    nodeName := TV.GetText(parentNode)
    f.Write("[" parentPath "\" nodeName "]`n")
    
    ; 保存該群組的直接項目
    childNode := TV.GetChild(parentNode)
    while childNode {
        itemName := TV.GetText(childNode)
        if !TV.GetChild(childNode) && textSnippets.Has(itemName) {
            f.Write(itemName "=" StrReplace(textSnippets[itemName], "`n", "<<NEWLINE>>") "`n")
        }
        childNode := TV.GetNext(childNode)
    }
}

; 載入片段從文件
LoadSnippets() {
    try {
        snippetFile := A_ScriptDir "\snippets.ini"
        if !FileExist(snippetFile)
            return
            
        TV.Delete()
        
        currentSection := ""
        currentNode := 0
        currentSubNode := 0
        
        Loop read snippetFile, "UTF-8" {
            line := Trim(A_LoopReadLine)
            if (line = "")
                continue
                
            if RegExMatch(line, "^\[(.*)\]$", &match) {
                groupPath := match[1]
                if InStr(groupPath, "\") {
                    ; 處理子群組
                    groups := StrSplit(groupPath, "\")
                    parentNode := FindOrCreateGroup(groups[1])
                    currentNode := TV.Add(groups[2], parentNode)
                } else {
                    ; 處理頂層群組
                    currentNode := TV.Add(groupPath)
                }
                continue
            }
            
            parts := StrSplit(line, "=",, 2)
            if (parts.Length >= 2) {
                key := parts[1]
                decodedValue := StrReplace(parts[2], "<<NEWLINE>>", "`n")
                
                textSnippets[key] := decodedValue
                if currentNode
                    TV.Add(key, currentNode)
                    
                LV.Add(, key, decodedValue)
            }
        }
    }
}

; 輔助函數：尋找或創建群組
FindOrCreateGroup(groupName) {
    currentNode := TV.GetNext()
    while currentNode {
        if (TV.GetText(currentNode) = groupName)
            return currentNode
        currentNode := TV.GetNext(currentNode)
    }
    return TV.Add(groupName)
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
    ; 支援多種動態內容
    replacements := Map(
        "{TODAY}", FormatTime(, "yyyy/MM/dd"),
        "{TODAY_SHORT}", FormatTime(, "yy/MM/dd"),
        "{TIME}", FormatTime(, "HH:mm"),
        "{TIME_FULL}", FormatTime(, "HH:mm:ss"),
        "{YESTERDAY}", FormatTime(DateAdd(A_Now, -1, "days"), "yyyy/MM/dd"),
        "{TOMORROW}", FormatTime(DateAdd(A_Now, 1, "days"), "yyyy/MM/dd"),
        "{WEEKDAY}", FormatTime(, "dddd"),  ; 完整星期名稱
        "{MONTH}", FormatTime(, "MMMM"),    ; 完整月份名稱
        "{YEAR}", FormatTime(, "yyyy")      ; 完整年份
    )
        
    ; 處理所有動態標記
    for tag, value in replacements
        text := StrReplace(text, tag, value)

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

ExportSnippets(*) {
    ; 讓用戶選擇儲存位置和格式
    savePath := FileSelect("S", A_ScriptDir "\snippets_export", "Export Snippets", "Text Files (*.txt; *.ini)")
    if !savePath
        return
        
    try {
        f := FileOpen(savePath, "w", "UTF-8")
        
        ; 處理所有頂層群組
        currentNode := TV.GetNext()
        while currentNode {
            if !TV.GetParent(currentNode) {  ; 確保是頂層群組
                nodeName := TV.GetText(currentNode)
                f.Write("[" nodeName "]`n")
                
                ; 先寫入該群組直接的項目
                childNode := TV.GetChild(currentNode)
                while childNode {
                    itemName := TV.GetText(childNode)
                    if !TV.GetChild(childNode) && textSnippets.Has(itemName) {
                        f.Write(itemName "=" StrReplace(textSnippets[itemName], "`n", "<<NEWLINE>>") "`n")
                    }
                    childNode := TV.GetNext(childNode)
                }
                
                ; 處理子群組
                childNode := TV.GetChild(currentNode)
                while childNode {
                    if TV.GetChild(childNode) {
                        f.Write("`n")  ; 在子群組前添加空行
                        ExportGroupContent(f, childNode, nodeName)
                    }
                    childNode := TV.GetNext(childNode)
                }
                f.Write("`n")  ; 在每個頂層群組後添加空行
            }
            currentNode := TV.GetNext(currentNode)
        }
        f.Close()
        MsgBox("匯出成功！", "成功")
    } catch as err {
        MsgBox("匯出失敗：" err.Message, "錯誤", "Icon!")
    }
}

; 輔助函數：匯出群組內容
ExportGroupContent(f, parentNode, parentPath) {
    nodeName := TV.GetText(parentNode)
    f.Write("[" parentPath "\" nodeName "]`n")
    
    ; 保存該群組的直接項目
    childNode := TV.GetChild(parentNode)
    while childNode {
        itemName := TV.GetText(childNode)
        if !TV.GetChild(childNode) && textSnippets.Has(itemName) {
            f.Write(itemName "=" StrReplace(textSnippets[itemName], "`n", "<<NEWLINE>>") "`n")
        }
        childNode := TV.GetNext(childNode)
    }
}

ImportSnippets(*) {
    ; 讓用戶選擇檔案
    filePath := FileSelect(1,, "Import Snippets", "Snippet Files (*.txt; *.ini)")
    if !filePath
        return
        

    ; 詢問匯入模式
    ; 在 ImportSnippets 中使用
    result := ImportModeSelect()
    if (result = "取消")
        return

    switch result {  ; 使用 switch 來處理不同的按鈕返回值
        case 6:  ; "是" 按鈕
            result := "疊加"
        case 7:  ; "否" 按鈕
            result := "取代"
        case 2:  ; "取消" 按鈕
            result := "取消"
        return
    }
                 
    if (result = "取消")
        return
        
    try {
        ; 檢查檔案格式
        isValid := ValidateImportFile(filePath)
        if !isValid {
            MsgBox("檔案格式不正確！`n請確保檔案包含正確的分組和項目格式。", "錯誤", "Icon!")
            return
        }

        ; 備份現有的 snippets.ini
        snippetFile := A_ScriptDir "\snippets.ini"
        if FileExist(snippetFile) {
            backupFile := A_ScriptDir "\snippets_" FormatTime(, "yyyyMMddHHmmss") ".old"
            try {
                FileCopy(snippetFile, backupFile)
            } catch as err {
                MsgBox("備份失敗：" err.Message, "警告", "Icon!")
            }
        }
        
        ; 如果是覆蓋模式，清除現有內容
        if (result = "取代") {
            TV.Delete()
            LV.Delete()
            textSnippets.Clear()
        }
        
        ; 讀取並匯入內容
        currentSection := ""
        currentNode := 0
        
        Loop read filePath, "UTF-8" {
            line := Trim(A_LoopReadLine)
            if (line = "")
                continue
                
            if RegExMatch(line, "^\[(.*)\]$", &match) {
                groupPath := match[1]
                if InStr(groupPath, "\") {
                    ; 處理子群組
                    groups := StrSplit(groupPath, "\")
                    parentNode := FindOrCreateGroup(groups[1])
                    currentNode := TV.Add(groups[2], parentNode)
                } else {
                    ; 處理頂層群組
                    currentNode := FindOrCreateGroup(groupPath)
                }
                continue
            }
            
            parts := StrSplit(line, "=",, 2)
            if (parts.Length >= 2) {
                key := parts[1]
                decodedValue := StrReplace(parts[2], "<<NEWLINE>>", "`n")
                
                if (result = "Yes" && textSnippets.Has(key))
                    continue
                
                textSnippets[key] := decodedValue
                if currentNode {
                    TV.Add(key, currentNode)
                    ; 更新 ListView
                    groupPath := GetFullGroupPath(currentNode)
                    LV.Add(, groupPath, key, decodedValue)
                }
                
                if !InStr(key, "memo_")
                    RegisterHotstrings(key)
            }
        }

        ; 保存新的設定到 snippets.ini
        SaveSnippets()
        
        MsgBox("匯入成功！`n原有設定已備份為：" backupFile, "成功")
        
    } catch as err {
        MsgBox("匯入失敗：" err.Message, "錯誤", "Icon!")
    }
}

ImportModeSelect(*) {
    result := ""
    importGui := Gui("+Owner" . mainGui.Hwnd, "匯入模式")
    importGui.SetFont("s10", "Segoe UI")
    
    importGui.Add("Text",, "選擇匯入模式：`n`n疊加 = 保留現有項目`n取代 = 清除現有項目")
    
    importGui.Add("Button", "x10 y80 w80 Default", "疊加").OnEvent("Click", (*) => (result := "疊加", importGui.Destroy()))
    importGui.Add("Button", "x100 y80 w80", "取代").OnEvent("Click", (*) => (result := "取代", importGui.Destroy()))
    importGui.Add("Button", "x190 y80 w80", "取消").OnEvent("Click", (*) => (result := "取消", importGui.Destroy()))
    
    importGui.Show()
    WinWaitClose(importGui)
    return result
}

; 輔助函數：驗證匯入檔案格式
ValidateImportFile(filePath) {
    try {
        hasGroup := false
        hasItem := false
        
        Loop read filePath, "UTF-8"
        {
            line := Trim(A_LoopReadLine)
            if (line = "")
                continue
                
            if RegExMatch(line, "^\[.*\]$")
                hasGroup := true
                
            if InStr(line, "=")
                hasItem := true
        }
        
        return hasGroup && hasItem
    } catch {
        return false
    }
}

; 輔助函數：尋找已存在的分組
FindGroup(groupName) {
    currentNode := TV.GetNext()
    while currentNode {
        if (TV.GetText(currentNode) = groupName)
            return currentNode
        currentNode := TV.GetNext(currentNode)
    }
    return 0
}
; 在創建 TreeView 後添加選擇事件處理
TV.OnEvent("ItemSelect", TVSelect)

; 添加 TreeView 選擇事件處理函數
; TreeView 選擇事件處理函數
; TreeView 選擇事件處理函數
TVSelect(*) {
    ; 清空當前 ListView
    LV.Delete()
    
    ; 獲取選中的節點
    selectedNode := TV.GetSelection()
    if !selectedNode
        return
    
    ; 顯示群組路徑
    if IsEmptyGroup(selectedNode) {
        ; 如果是空群組，顯示群組路徑
        groupPath := GetFullGroupPath(selectedNode)
        LV.Add(, groupPath, "(Empty Group)", "")
    } 
    ; 如果是群組節點（有子節點）
    else if TV.GetChild(selectedNode) {
        ShowGroupContents(selectedNode)
    } else {
        ShowSingleItem(selectedNode)
    }
}

; 輔助函數：判斷是否為空群組
IsEmptyGroup(node) {
    ; 如果節點沒有子節點，且不在 textSnippets 中，則視為空群組
    return !TV.GetChild(node) && !textSnippets.Has(TV.GetText(node))
}

; 輔助函數：顯示群組內容（包括子群組）
ShowGroupContents(groupNode) {
    ; 遞迴處理群組內的所有項目
    ProcessGroupItems(node) {
        childNode := TV.GetChild(node)
        while childNode {
            ; 如果是子群組
            if TV.GetChild(childNode) {
                ; 遞迴處理子群組內容
                ProcessGroupItems(childNode)
            } 
            ; 如果是空群組
            else if IsEmptyGroup(childNode) {
                ; 顯示空群組
                groupPath := GetFullGroupPath(childNode)
                LV.Add(, groupPath, "(Empty Group)", "")
            }
            ; 如果是一般項目
            else {
                ShowSingleItem(childNode)
            }
            childNode := TV.GetNext(childNode)
        }
    }
    
    ; 開始處理群組內容
    ProcessGroupItems(groupNode)
}

; 輔助函數：顯示單個項目
ShowSingleItem(itemNode) {
    displayText := TV.GetText(itemNode)
    
    ; 確保節點不是空群組
    if IsEmptyGroup(itemNode)
        return
        
    parentNode := TV.GetParent(itemNode)
    if !parentNode
        return
        
    groupPath := GetFullGroupPath(parentNode)
    
    ; 尋找對應的實際內容
    for key, value in textSnippets {
        if (displayText = "(memo)" && InStr(key, "memo_")) {
            LV.Add(, groupPath, displayText, value)
            break
        } else if (key = displayText) {
            LV.Add(, groupPath, key, value)
            break
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

; 搜尋視窗函數
ShowSearchWindow(*) {
    searchGui := Gui("+Owner" . mainGui.Hwnd, "Search Snippets")
    searchGui.SetFont("s10", "Segoe UI")
    
    ; 搜尋選項
    searchGui.Add("Text", "x10 y10", "Search in:")
    searchGui.Add("Radio", "x10 y30 vSearchKey Checked", "Keyword")
    searchGui.Add("Radio", "x90 y30 vSearchPhrase", "Phrase")
    searchGui.Add("Radio", "x155 y30 vSearchBoth", "Both")
    
    ; 搜尋輸入
    searchGui.Add("Text", "x10 y60", "Search text:")
    searchEdit := searchGui.Add("Edit", "x10 y80 w300")
    
    ; 結果列表
    resultList := searchGui.Add("ListView", "x10 y120 w500 h300", ["Group", "Keyword", "Phrase"])
    
    ; 即時搜尋 - 修改這裡，傳入 searchGui
    searchEdit.OnEvent("Change", (*) => UpdateSearchResults(searchGui, searchEdit, resultList))
    
    ; 雙擊跳轉
    ; resultList.OnEvent("DoubleClick", (*) => JumpToItem(resultList))
    ; 定義本地函數來處理雙擊事件
    HandleDoubleClick(*) {
        JumpToItem(resultList)
        searchGui.Destroy()
    }
    
    resultList.OnEvent("DoubleClick", HandleDoubleClick)

    searchGui.Show()
}

; 更新搜尋結果
UpdateSearchResults(searchGui, searchEdit, resultList) {
    searchText := searchEdit.Value
    if (searchText = "")
        return
        
    resultList.Delete()

    ; 獲取搜尋選項
    searchInKey := searchGui["SearchKey"].Value
    searchInPhrase := searchGui["SearchPhrase"].Value
    searchInBoth := searchGui["SearchBoth"].Value
    
    currentNode := TV.GetNext()
    while currentNode {
        groupName := TV.GetText(currentNode)
        childNode := TV.GetChild(currentNode)
        
        while childNode {
            key := TV.GetText(childNode)
            if textSnippets.Has(key) {
                value := textSnippets[key]
                
                ; 根據搜尋選項決定搜尋範圍
                matched := false
                if (searchInKey && InStr(key, searchText))
                    matched := true
                else if (searchInPhrase && InStr(value, searchText))
                    matched := true
                else if (searchInBoth && (InStr(key, searchText) || InStr(value, searchText)))
                    matched := true
                
                if matched
                    resultList.Add(, groupName, key, value)
            }
            childNode := TV.GetNext(childNode)
        }
        currentNode := TV.GetNext(currentNode)
    }
}



; 跳轉到項目
JumpToItem(resultList) {
    row := resultList.GetNext()
    if !row
        return
        
    groupName := resultList.GetText(row, 1)
    key := resultList.GetText(row, 2)
    
    ; 處理 TreeView
    currentNode := TV.GetNext()
    while currentNode {
        if (TV.GetText(currentNode) = groupName) {
            TV.Modify(currentNode, "Expand")
            childNode := TV.GetChild(currentNode)
            while childNode {
                if (TV.GetText(childNode) = key) {
                    TV.Modify(childNode, "Select")
                    break
                }
                childNode := TV.GetNext(childNode)
            }
            break
        }
        currentNode := TV.GetNext(currentNode)
    }

    ; 處理 ListView
    LV.Delete()  ; 清空當前 ListView
    ; 更新 ListView 顯示所選群組的內容
    childNode := TV.GetChild(currentNode)
    while childNode {
        itemKey := TV.GetText(childNode)
        if textSnippets.Has(itemKey)
            LV.Add(, itemKey, textSnippets[itemKey])
        childNode := TV.GetNext(childNode)
    }

    ; 在 ListView 中選中目標項目
    Loop LV.GetCount() {
        if (LV.GetText(A_Index, 1) = key) {
            LV.Modify(A_Index, "Select Focus")  ; 選中並設置焦點
            break
        }
    }

}
