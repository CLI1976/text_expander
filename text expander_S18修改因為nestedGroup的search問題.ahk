; === 1. 全局變量和基本設定 ===
; Global variables
; Settings functions (LoadUserSettings, SaveUserSettings)

; === 2. GUI 相關函數及初始化 ===
; 主視窗建立
; GUI 事件處理函數 (GuiResize, LVSelect, TVSelect)
; GUI 控件事件處理函數

; === 3. 熱字串相關函數 ===
; 熱字串註冊和管理
; 文本輸出處理

; === 4. 文件操作相關函數 ===
; 檔案讀寫
; Import/Export 功能

; === 5. 輔助函數 ===
; 路徑處理
; 群組管理
; 顯示處理

; === 6. 程式初始化 ===
; 載入設定
; 註冊熱字串

;==========================================
; === 1. 全局變量和基本設定 ===
;==========================================
#Requires AutoHotkey v2.0
#SingleInstance Force

; Global variables
global textSnippets := Map()
global hotstringEnabled := true
global settingsFile := A_ScriptDir "\textEx_settings.ini"

; Settings functions
LoadUserSettings() {
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

SaveUserSettings(*) {
    global settingsFile, sendRadio
    
    try {
        method := sendRadio.Value ? "Send" : "Ctrl+V"
        IniWrite(method, settingsFile, "OutputMethod", "DefaultMethod")
    }
}

;==========================================
; === 2. GUI 相關函數及初始化 ===
;==========================================

; 創建主視窗
mainGui := Gui("+Resize +MinSize800x600", "Text Expander")
mainGui.SetFont("s10", "Segoe UI")

; 左側樹狀視圖
TV := mainGui.Add("TreeView", "x5 y5 w180 h400")

; 右側列表視圖
LV := mainGui.Add("ListView", "x195 y5 w545 h400", ["Group Path", "Keyword", "Phrase"])
LV.ModifyCol(1, 150)  ; Group Path 欄寬
LV.ModifyCol(2, 100)  ; Keyword 欄寬
LV.ModifyCol(3, 280)  ; Phrase 欄寬

; 添加按鈕
addButton := mainGui.Add("Button", "x5 y410", "Add New")
editButton := mainGui.Add("Button", "x85 y410", "Edit")
deleteButton := mainGui.Add("Button", "x135 y410", "Delete")
toggleButton := mainGui.Add("Button", "x225 y410", "ON")

; 輸出方式選擇
outputMethodText := mainGui.Add("Text", "x15 y450", "輸出方式：")
sendRadio := mainGui.Add("Radio", "x85 y450", "Send Text")
pasteRadio := mainGui.Add("Radio", "x185 y450", "Ctrl+V")

; 分組按鈕
addGroupButton := mainGui.Add("Button", "x5 y490", "New Group")
deleteGroupButton := mainGui.Add("Button", "x105 y490", "Delete Group")

; Import/Export 和搜尋按鈕
importButton := mainGui.Add("Button", "x285 y410", "Import")
exportButton := mainGui.Add("Button", "x345 y410", "Export")
searchButton := mainGui.Add("Button", "x405 y410", "Search")
syncButton := mainGui.Add("Button", "x465 y490", "瀏覽器同步")

; 設置事件處理
mainGui.OnEvent("Size", GuiResize)
LV.OnEvent("DoubleClick", EditSnippet)
LV.OnEvent("ItemSelect", LVSelect)
TV.OnEvent("ItemSelect", TVSelect)

addButton.OnEvent("Click", AddNewSnippet)
editButton.OnEvent("Click", EditSnippet)
deleteButton.OnEvent("Click", DeleteSnippet)
toggleButton.OnEvent("Click", ToggleHotstrings)
importButton.OnEvent("Click", ImportSnippets)
exportButton.OnEvent("Click", ExportSnippets)
searchButton.OnEvent("Click", ShowSearchWindow)
sendRadio.OnEvent("Click", SaveUserSettings)
pasteRadio.OnEvent("Click", SaveUserSettings)
addGroupButton.OnEvent("Click", AddNewGroup)
deleteGroupButton.OnEvent("Click", DeleteGroup)

syncButton.OnEvent("Click", ExportToLocalStorage)

; GUI 事件處理函數
GuiResize(thisGui, MinMax, Width, Height) {
    if MinMax = -1  ; 視窗最小化
        return
    
    ; TreeView 固定寬度
    treeWidth := 180
    listWidth := Width - treeWidth - 20
    controlHeight := Height - 100
    
    ; 更新控件大小和位置
    TV.Move(5, 5, treeWidth - 10, controlHeight)
    LV.Move(treeWidth + 5, 5, listWidth - 15, controlHeight)

    ; 動態調整欄寬
    pathWidth := 100
    keywordWidth := 75
    phraseWidth := listWidth - pathWidth - keywordWidth - 20

    LV.ModifyCol(1, pathWidth)
    LV.ModifyCol(2, keywordWidth)
    LV.ModifyCol(3, phraseWidth)
    
    ; 更新按鈕位置
    buttonsY := controlHeight + 10
    addButton.Move(5, buttonsY)
    editButton.Move(85, buttonsY)
    deleteButton.Move(135, buttonsY)
    toggleButton.Move(225, buttonsY)
    
    ; 更新 radio 按鈕位置
    radioY := buttonsY + 35
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

; TreeView 選擇事件處理
TVSelect(*) {
    LV.Delete()
    
    selectedNode := TV.GetSelection()
    if !selectedNode
        return
    
    if IsEmptyGroup(selectedNode) {
        groupPath := GetFullGroupPath(selectedNode)
        LV.Add(, groupPath, "(Empty Group)", "")
    } 
    else if TV.GetChild(selectedNode) {
        ShowGroupContents(selectedNode)
    } else {
        ShowSingleItem(selectedNode)
    }
}

; ListView 點選事件處理
LVSelect(*) {
    if TV.GetSelection()
        TV.Modify(TV.GetSelection(), "-Select")
}

; === GUI 操作函數 ===
; 添加新片段
AddNewSnippet(*) {
    selectedNode := TV.GetSelection()
    if !selectedNode {
        MsgBox("請先選擇一個群組", "提示")
        return
    }

    targetNode := selectedNode
    if textSnippets.Has(TV.GetText(selectedNode)) {
        targetNode := TV.GetParent(selectedNode)
    }
    
    targetName := TV.GetText(targetNode)
    groupPath := GetFullGroupPath(targetNode)

    mainGui.GetPos(&mainX, &mainY, &mainWidth, &mainHeight)
    addWidth := Round(mainWidth * 0.8)
    addHeight := Round(mainHeight * 0.8)
    
    addGui := Gui("+Owner" . mainGui.Hwnd . " +Resize", "Add New Snippet")
    addGui.SetFont("s10", "Segoe UI")
    
    contentWidth := addWidth - 40
    keywordWidth := Min(contentWidth - 20, 300)
    phraseWidth := contentWidth - 20
    phraseHeight := addHeight - 170
    
    addGui.Add("Text",, "Group:  " targetName)
    addGui.Add("Text", "xm", "Keyword:")
    keywordEdit := addGui.Add("Edit", "w" keywordWidth " yp-3")
    helpButton := addGui.Add("Button", "x+5 yp h23", "?")
    helpButton.OnEvent("Click", ShowDynamicTags)

    addGui.Add("Text", "xm y+10", "Phrase:")
    phraseEdit := addGui.Add("Edit", "w" phraseWidth " h" phraseHeight)
    
    saveButton := addGui.Add("Button", "xm y+10", "Save")
    
    AddGuiSize(thisGui, MinMax, Width, Height) {
        if MinMax = -1
            return
            
        newWidth := Width - 40
        newHeight := Height - 170
        phraseEdit.Move(,, newWidth, newHeight)
        
        phraseEdit.GetPos(&phraseX, &phraseY)
        saveButton.Move(, phraseY + newHeight + 10)
        
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
            
            targetGroupPath := GetFullGroupPath(targetNode)
            
            textSnippets[keyword] := phrase
            LV.Add(, targetGroupPath, displayText, phrase)
            TV.Add(displayText, targetNode)
            
            if !InStr(keyword, "memo_")
                RegisterHotstrings(keyword)
                
            SaveSnippets()
            addGui.Destroy()
        }
    }
    
    saveButton.OnEvent("Click", SaveNewHandler)
    
    addGui.Move(, , addWidth + 40, addHeight)
    addGui.Show()
}

; 編輯片段
EditSnippet(*) {
    if (LV.GetCount("Selected") = 0) {
        MsgBox("請先選擇要編輯的項目", "提示")
        return
    }

    mainGui.GetPos(&mainX, &mainY, &mainWidth, &mainHeight)
    editWidth := Round(mainWidth * 0.8)
    editHeight := Round(mainHeight * 0.8)
    
    row := LV.GetNext()
    oldGroupPath := LV.GetText(row, 1)  ; 儲存群組路徑
    oldKeyword := LV.GetText(row, 2)
    oldPhrase := LV.GetText(row, 3)
    
    ; Debug 記錄
    FileAppend("=== 開始編輯 ===`n", A_ScriptDir "\edit_log.txt")
    FileAppend("原始值：`n", A_ScriptDir "\edit_log.txt")
    FileAppend("群組路徑: " oldGroupPath "`n", A_ScriptDir "\edit_log.txt")
    FileAppend("關鍵字: " oldKeyword "`n", A_ScriptDir "\edit_log.txt")
    FileAppend("短語: " oldPhrase "`n", A_ScriptDir "\edit_log.txt")
    
    editGui := Gui("+Owner" . mainGui.Hwnd . " +Resize", "Edit Snippet")
    editGui.SetFont("s10", "Segoe UI")
    
    editWidth -= 40
    keywordWidth := Min(editWidth - 20, 300)
    phraseWidth := editWidth - 20
    phraseHeight := editHeight - 150
    
    editGui.Add("Text",, "Group: " oldGroupPath)
    editGui.Add("Text",, "Keyword:")
    keywordEdit := editGui.Add("Edit", "w" keywordWidth, oldKeyword)
    
    editGui.Add("Text", "xm y+10", "Phrase:")
    phraseEdit := editGui.Add("Edit", "w" phraseWidth " h" phraseHeight, oldPhrase)
    
    saveButton := editGui.Add("Button", "xm y+10", "Save")

    EditGuiSize(thisGui, MinMax, Width, Height) {
        if MinMax = -1
            return
            
        newWidth := Width - 40
        newHeight := Height - 150
        phraseEdit.Move(,, newWidth, newHeight)
        phraseEdit.GetPos(&phraseX, &phraseY)
        saveButton.Move(, phraseY + newHeight + 10)
        WinRedraw(thisGui)
    }
    
    editGui.OnEvent("Size", EditGuiSize)
    
    if FileExist(A_ScriptDir "\edit_log.txt")
        FileDelete(A_ScriptDir "\edit_log.txt")

    SaveEditHandler(*) {
        keyword := keywordEdit.Value
        phrase := phraseEdit.Value
        
        FileAppend("=== 儲存編輯 ===`n", A_ScriptDir "\edit_log.txt")
        FileAppend("新值：`n", A_ScriptDir "\edit_log.txt")
        FileAppend("關鍵字: " keyword "`n", A_ScriptDir "\edit_log.txt")
        FileAppend("短語: " phrase "`n", A_ScriptDir "\edit_log.txt")
    
        if (keyword != "" && phrase != "") {
            ; 1. 在 textSnippets 中進行更新
            textSnippets[keyword] := phrase
            FileAppend("新關鍵字/短語已加入 Map`n", A_ScriptDir "\edit_log.txt")
            
            if (oldKeyword != keyword) {
                textSnippets.Delete(oldKeyword)
                FileAppend("舊關鍵字已從 Map 中刪除`n", A_ScriptDir "\edit_log.txt")
            }
            
            ; 2. 更新 TreeView
            ; 首先找到對應的節點
            pathParts := StrSplit(oldGroupPath, "\")
            currentNode := TV.GetNext()
            found := false
            
            while currentNode && !found {
                if (TV.GetText(currentNode) = oldKeyword) {
                    TV.Modify(currentNode,, keyword)
                    found := true
                    FileAppend("在頂層找到並更新節點`n", A_ScriptDir "\edit_log.txt")
                } else {
                    found := UpdateTreeNodeRecursive(currentNode, oldKeyword, keyword)
                }
                currentNode := found ? 0 : TV.GetNext(currentNode)
            }
            
            ; 3. 更新 ListView
            LV.Modify(row,, oldGroupPath, keyword, phrase)
            FileAppend("ListView 已更新`n", A_ScriptDir "\edit_log.txt")
            
            ; 4. 如果不是 memo_ 項目，更新熱字串
            if !InStr(keyword, "memo_") {
                CreateHotstring(keyword, phrase)
                FileAppend("熱字串已更新`n", A_ScriptDir "\edit_log.txt")
            }
            
            ; 5. 保存所有更改
            SaveSnippets()
            FileAppend("所有更改已保存`n", A_ScriptDir "\edit_log.txt")
            
            editGui.Destroy()
        }
    }
    
    saveButton.OnEvent("Click", SaveEditHandler)
    
    editGui.Move(, , editWidth + 40, editHeight)
    editGui.Show()
}

; 遞迴更新 TreeView 節點的輔助函數
UpdateTreeNodeRecursive(parentNode, oldText, newText) {
    if !parentNode
        return false
        
    childNode := TV.GetChild(parentNode)
    while childNode {
        if (TV.GetText(childNode) = oldText) {
            TV.Modify(childNode,, newText)
            FileAppend("在子層找到並更新節點: " oldText " -> " newText "`n", A_ScriptDir "\edit_log.txt")
            return true
        }
        
        if (UpdateTreeNodeRecursive(childNode, oldText, newText))
            return true
            
        childNode := TV.GetNext(childNode)
    }
    return false
}

; 刪除片段
DeleteSnippet(*) {
    if (LV.GetCount("Selected") = 0) {
        MsgBox("請先選擇要刪除的項目", "提示")
        return
    }

    row := LV.GetNext()
    keyword := LV.GetText(row, 2)
    
    if (keyword = "(Empty Group)") {
        MsgBox("空群組無法使用此功能刪除，請使用 Delete Group 按鈕", "操作提示")
        return
    }
    
    if (MsgBox("確定要刪除這個項目嗎？", "確認刪除", "YesNo") = "Yes") {
        currentNode := TV.GetNext(0)
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

; === 搜尋相關函數 ===
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
    
    ; 設置事件處理
    searchEdit.OnEvent("Change", (*) => UpdateSearchResults(searchGui, searchEdit, resultList))
    resultList.OnEvent("DoubleClick", (*) => (JumpToItem(resultList), searchGui.Destroy()))

    searchGui.Show()
}

UpdateSearchResults(searchGui, searchEdit, resultList) {
    searchText := searchEdit.Value
    if (searchText = "")
        return
        
    resultList.Delete()

    searchInKey := searchGui["SearchKey"].Value
    searchInPhrase := searchGui["SearchPhrase"].Value
    searchInBoth := searchGui["SearchBoth"].Value
    
    ; 遞迴搜索函數
    SearchNodeRecursive(node, groupPath := "") {
        if !node
            return
            
        currentGroupPath := groupPath
        if (groupPath = "")
            currentGroupPath := TV.GetText(node)
        else
            currentGroupPath := groupPath "\" TV.GetText(node)
            
        ; 處理當前節點的子項目
        childNode := TV.GetChild(node)
        while childNode {
            if TV.GetChild(childNode) {
                ; 如果是群組，遞迴搜索
                SearchNodeRecursive(childNode, currentGroupPath)
            } else {
                ; 如果是項目，檢查是否符合搜索條件
                key := TV.GetText(childNode)
                if textSnippets.Has(key) {
                    value := textSnippets[key]
                    
                    matched := false
                    if (searchInKey && InStr(key, searchText))
                        matched := true
                    else if (searchInPhrase && InStr(value, searchText))
                        matched := true
                    else if (searchInBoth && (InStr(key, searchText) || InStr(value, searchText)))
                        matched := true
                    
                    if matched
                        resultList.Add(, currentGroupPath, key, value)
                }
            }
            childNode := TV.GetNext(childNode)
        }
    }
    
    ; 從頂層群組開始搜索
    currentNode := TV.GetNext()
    while currentNode {
        SearchNodeRecursive(currentNode)
        currentNode := TV.GetNext(currentNode)
    }
}

JumpToItem(resultList) {
    row := resultList.GetNext()
    if !row
        return
        
    groupPath := resultList.GetText(row, 1)
    key := resultList.GetText(row, 2)
    
    ; 根據完整路徑找到並展開群組
    pathParts := StrSplit(groupPath, "\")
    currentNode := TV.GetNext()
    
    ; 遍歷路徑找到目標群組
    targetNode := 0
    for i, part in pathParts {
        while currentNode {
            if (TV.GetText(currentNode) = part) {
                TV.Modify(currentNode, "Expand")
                if (i = pathParts.Length) {
                    targetNode := currentNode
                    break
                }
                currentNode := TV.GetChild(currentNode)
                break
            }
            currentNode := TV.GetNext(currentNode)
        }
    }
    
    ; 在目標群組中找到並選中項目
    if (targetNode) {
        childNode := TV.GetChild(targetNode)
        while childNode {
            if (TV.GetText(childNode) = key) {
                TV.Modify(childNode, "Select")
                break
            }
            childNode := TV.GetNext(childNode)
        }
    }
    
    ; 更新 ListView 顯示
    TVSelect()
    
    ; 在 ListView 中選中相應項目
    Loop LV.GetCount() {
        if (LV.GetText(A_Index, 2) = key) {
            LV.Modify(A_Index, "Select Focus")
            break
        }
    }
}



; === 群組管理相關函數 ===
AddNewGroup(*) {
    addGroupGui := Gui("+Owner" . mainGui.Hwnd, "新增分組")
    addGroupGui.SetFont("s10", "Segoe UI")
    
    ; 取得所有頂層群組作為選項
    groupChoices := ["頂層"]
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
            for index, text in groupChoices {
                if (text = TV.GetText(selectedNode)) {
                    parentChoice.Value := index
                    break
                }
            }
        } else {
            parentText := TV.GetText(TV.GetParent(selectedNode))
            for index, text in groupChoices {
                if (text = parentText) {
                    parentChoice.Value := index
                    break
                }
            }
        }
    } else {
        parentChoice.Value := 1
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
    if LV.GetCount("Selected") > 0 {
        if TV.GetSelection()
            TV.Modify(TV.GetSelection(), "-Select")
        
        MsgBox("欲刪除群組 請在左側樹狀結構點選", "提示")
        return
    }

    selectedNode := TV.GetSelection()
    if !selectedNode {
        MsgBox("欲刪除群組 請在左側樹狀結構點選群組", "提示")
        return
    }
    
    if !TV.GetChild(selectedNode) && !IsEmptyGroup(selectedNode) {
        MsgBox("欲刪除群組 請在左側樹狀結構點選群組", "提示")
        return
    }
    
    groupPath := GetFullGroupPath(selectedNode)
    
    if (MsgBox("確定要刪除分組 " groupPath " 及其所有項目嗎？", "確認刪除", "YesNo") = "No")
        return
    
    DeleteGroupItems(selectedNode)
    TV.Delete(selectedNode)
    SaveSnippets()
}

; 輔助函數：遞迴刪除群組內的所有項目
DeleteGroupItems(node) {
    childNode := TV.GetChild(node)
    while childNode {
        nextChild := TV.GetNext(childNode)
        
        if TV.GetChild(childNode) {
            DeleteGroupItems(childNode)
        }
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

;==========================================
; === 3. 熱字串相關函數 ===
;==========================================

; 註冊熱字串
RegisterHotstrings(key) {
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

; 解除註冊熱字串
UnregisterHotstrings(key) {
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

; 創建熱字串
CreateHotstring(key, value) {
    callback(*) => SendWithIMEControl(value)
    
    Hotstring(":*:" key " ", callback)
    Hotstring(":*:" key ".", callback)
}

; 開關熱字串
ToggleHotstrings(*) {
    Suspend  ; 切換暫停狀態
    
    ; 更新按鈕文字和顏色
    toggleButton.Text := A_IsSuspended ? "OFF" : "ON"
    toggleButton.Opt(A_IsSuspended ? "+cRed" : "+cGreen")
    
    TrayTip("Text Expander", A_IsSuspended ? "自動替換已禁用" : "自動替換已啟用")
}

; 文字輸出處理
SendWithIMEControl(text) {
    ; 支援動態內容
    replacements := Map(
        "{TODAY}", FormatTime(, "yyyy/MM/dd"),
        "{TODAY_SHORT}", FormatTime(, "yy/MM/dd"),
        "{TIME}", FormatTime(, "HH:mm"),
        "{TIME_FULL}", FormatTime(, "HH:mm:ss"),
        "{YESTERDAY}", FormatTime(DateAdd(A_Now, -1, "days"), "yyyy/MM/dd"),
        "{TOMORROW}", FormatTime(DateAdd(A_Now, 1, "days"), "yyyy/MM/dd"),
        "{WEEKDAY}", FormatTime(, "dddd"),
        "{MONTH}", FormatTime(, "MMMM"),
        "{YEAR}", FormatTime(, "yyyy")
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
        ; 使用一般Send方式
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

; 重新註冊所有熱字串
ReloadAllHotstrings() {
    ; 記錄現有的 key-value 對
    savedPairs := Map()
    for key, value in textSnippets {
        savedPairs[key] := value
    }

    ; 清除所有熱字串
    for key, value in textSnippets {
        try {
            Hotstring(":*:" key " ",, "Off")
            Hotstring(":*:" key ".",, "Off")
        }
    }
    Sleep(100)

    ; 重新註冊
    textSnippets.Clear()
    for key, value in savedPairs {
        textSnippets[key] := value
        try CreateHotstring(key, value)
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

;==========================================
; === 4. 文件操作相關函數 ===
;==========================================

; 保存片段到文件
; 保存片段到文件
SaveSnippets() {
    snippetFile := A_ScriptDir "\snippets.ini"
    logFile := A_ScriptDir "\save_log.txt"
    
    try {
        ; 確保刪除舊的日誌文件
        if FileExist(logFile)
            FileDelete(logFile)
            
        FileAppend("=== 開始保存過程 " A_Now " ===`n", logFile)
        FileAppend("正在處理檔案: " snippetFile "`n", logFile)
        
        if FileExist(snippetFile)
            FileDelete(snippetFile)
            
        f := FileOpen(snippetFile, "w", "UTF-8")
        if !f {
            FileAppend("無法開啟檔案進行寫入！`n", logFile)
            return
        }
        
        FileAppend("成功開啟檔案準備寫入`n", logFile)
        
        currentNode := TV.GetNext()
        while currentNode {
            if !TV.GetParent(currentNode) {
                nodeName := TV.GetText(currentNode)
                FileAppend("處理頂層群組: " nodeName "`n", logFile)
                
                f.Write("[" nodeName "]`n")
                
                ; 先處理直接項目
                FileAppend("  處理頂層群組的直接項目:`n", logFile)
                SaveDirectItems(f, currentNode, logFile)
                
                ; 處理子群組
                FileAppend("  處理頂層群組的子群組:`n", logFile)
                SaveNestedGroups(f, currentNode, nodeName, logFile)
                
                f.Write("`n")
            }
            currentNode := TV.GetNext(currentNode)
        }
        f.Close()
        FileAppend("=== 保存完成 ===`n", logFile)
        
    } catch as err {
        FileAppend("錯誤發生: " err.Message "`n", logFile)
        MsgBox("保存時發生錯誤：" err.Message)
    }
}

; 保存直接項目
SaveDirectItems(f, node, logFile) {
    childNode := TV.GetChild(node)
    while childNode {
        itemName := TV.GetText(childNode)
        if !TV.GetChild(childNode) {  ; 如果是一般項目
            FileAppend("    檢查項目: " itemName "`n", logFile)
            
            if textSnippets.Has(itemName) {
                value := textSnippets[itemName]
                FileAppend("    寫入項目：" itemName " = " value "`n", logFile)
                f.Write(itemName "=" StrReplace(value, "`n", "<<NEWLINE>>") "`n")
            } else if InStr(itemName, "memo_") {
                FileAppend("    發現memo項目但沒有對應的值: " itemName "`n", logFile)
            }
        }
        childNode := TV.GetNext(childNode)
    }
}

; 遞迴保存巢狀群組
SaveNestedGroups(f, node, parentPath, logFile) {
    childNode := TV.GetChild(node)
    while childNode {
        if TV.GetChild(childNode) {  ; 如果是群組
            nodeName := TV.GetText(childNode)
            currentPath := parentPath "\" nodeName
            
            FileAppend("    處理巢狀群組: " currentPath "`n", logFile)
            f.Write("[" currentPath "]`n")
            
            ; 保存此群組的直接項目
            FileAppend("      處理巢狀群組的直接項目:`n", logFile)
            SaveDirectItems(f, childNode, logFile)
            f.Write("`n")
            
            ; 遞迴處理更深層的群組
            FileAppend("      處理更深層巢狀群組:`n", logFile)
            SaveNestedGroups(f, childNode, currentPath, logFile)
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
        currentNode := 0
        
        Loop read snippetFile, "UTF-8" {
            line := Trim(A_LoopReadLine)
            if (line = "")
                continue
                
            if RegExMatch(line, "^\[(.*)\]$", &match) {
                groupPath := match[1]
                if InStr(groupPath, "\") {
                    groups := StrSplit(groupPath, "\")
                    parentNode := FindOrCreateGroup(groups[1])
                    currentNode := TV.Add(groups[2], parentNode)
                } else {
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

; 匯出片段
ExportSnippets(*) {
    savePath := FileSelect("S", A_ScriptDir "\snippets_export", "Export Snippets", "Text Files (*.txt; *.ini)")
    if !savePath
        return
        
    try {
        f := FileOpen(savePath, "w", "UTF-8")
        exportLog := A_ScriptDir "\export_log.txt"  ; 建立一個導出專用的日誌檔案
        
        currentNode := TV.GetNext()
        while currentNode {
            if !TV.GetParent(currentNode) {
                nodeName := TV.GetText(currentNode)
                f.Write("[" nodeName "]`n")
                
                ; 先處理直接項目
                SaveDirectItems(f, currentNode, exportLog)
                
                ; 處理子群組
                SaveNestedGroups(f, currentNode, nodeName, exportLog)
                
                f.Write("`n")
            }
            currentNode := TV.GetNext(currentNode)
        }
        f.Close()
        MsgBox("匯出成功！", "成功")
    } catch as err {
        MsgBox("匯出失敗：" err.Message, "錯誤", "Icon!")
    }
}

; 匯入片段
ImportSnippets(*) {
    filePath := FileSelect(1,, "Import Snippets", "Snippet Files (*.txt; *.ini)")
    if !filePath
        return
        
    result := ImportModeSelect()
    if (result = "取消")
        return
                 
    try {
        if !ValidateImportFile(filePath) {
            MsgBox("檔案格式不正確！`n請確保檔案包含正確的分組和項目格式。", "錯誤", "Icon!")
            return
        }

        ; 備份現有檔案
        snippetFile := A_ScriptDir "\snippets.ini"
        if FileExist(snippetFile) {
            backupFile := A_ScriptDir "\snippets_" FormatTime(, "yyyyMMddHHmmss") ".old"
            try FileCopy(snippetFile, backupFile)
            catch as err {
                MsgBox("備份失敗：" err.Message, "警告", "Icon!")
            }
        }
        
        ; 如果是覆蓋模式，清除現有內容
        if (result = "取代") {
            TV.Delete()
            LV.Delete()
            textSnippets.Clear()
        }
        
        currentNode := 0
        
        Loop read filePath, "UTF-8" {
            line := Trim(A_LoopReadLine)
            if (line = "")
                continue
                
            if RegExMatch(line, "^\[(.*)\]$", &match) {
                groupPath := match[1]
                if InStr(groupPath, "\") {
                    groups := StrSplit(groupPath, "\")
                    parentNode := FindOrCreateGroup(groups[1])
                    currentNode := TV.Add(groups[2], parentNode)
                } else {
                    currentNode := FindOrCreateGroup(groupPath)
                }
                continue
            }
            
            parts := StrSplit(line, "=",, 2)
            if (parts.Length >= 2) {
                key := parts[1]
                decodedValue := StrReplace(parts[2], "<<NEWLINE>>", "`n")
                
                if (result = "疊加" && textSnippets.Has(key))
                    continue
                    
                textSnippets[key] := decodedValue
                if currentNode {
                    TV.Add(key, currentNode)
                    groupPath := GetFullGroupPath(currentNode)
                    LV.Add(, groupPath, key, decodedValue)
                }
                
                if !InStr(key, "memo_")
                    RegisterHotstrings(key)
            }
        }

        SaveSnippets()
        MsgBox("匯入成功！`n原有設定已備份為：" backupFile, "成功")
        
    } catch as err {
        MsgBox("匯入失敗：" err.Message, "錯誤", "Icon!")
    }
}





;==========================================
; === 5. 輔助函數 ===
;==========================================

; === 路徑和群組處理相關 ===
GetFullGroupPath(node) {
    path := TV.GetText(node)
    parentNode := TV.GetParent(node)
    while parentNode {
        path := TV.GetText(parentNode) "\" path
        parentNode := TV.GetParent(parentNode)
    }
    return path
}

IsEmptyGroup(node) {
    nodeText := TV.GetText(node)
    ; 如果是 memo_ 開頭的項目，就不該被當作空群組
    if InStr(nodeText, "memo_") {
        return false
    }
    return !TV.GetChild(node) && !textSnippets.Has(nodeText)
}

FindOrCreateGroup(groupName) {
    currentNode := TV.GetNext()
    while currentNode {
        if (TV.GetText(currentNode) = groupName)
            return currentNode
        currentNode := TV.GetNext(currentNode)
    }
    return TV.Add(groupName)
}

; === 內容驗證相關 ===
ValidateImportFile(filePath) {
    try {
        hasGroup := false
        hasItem := false
        
        Loop read filePath, "UTF-8" {
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

; === TreeView內容顯示相關 ===
ShowGroupContents(groupNode) {
    ProcessGroupItems(node) {
        childNode := TV.GetChild(node)
        while childNode {
            if TV.GetChild(childNode) {
                ProcessGroupItems(childNode)
            } 
            else if IsEmptyGroup(childNode) {
                groupPath := GetFullGroupPath(childNode)
                LV.Add(, groupPath, "(Empty Group)", "")
            }
            else {
                ShowSingleItem(childNode)
            }
            childNode := TV.GetNext(childNode)
        }
    }
    
    ProcessGroupItems(groupNode)
}

ShowSingleItem(itemNode) {
    displayText := TV.GetText(itemNode)
    
    ; 如果是 memo_ 項目，直接顯示，不要判斷是否為空群組
    if InStr(displayText, "memo_") {
        if textSnippets.Has(displayText) {
            groupPath := GetFullGroupPath(TV.GetParent(itemNode))
            LV.Add(, groupPath, displayText, textSnippets[displayText])
        }
        return
    }
    
    if IsEmptyGroup(itemNode)
        return
        
    parentNode := TV.GetParent(itemNode)
    if !parentNode
        return
        
    groupPath := GetFullGroupPath(parentNode)
    
    if textSnippets.Has(displayText) {
        LV.Add(, groupPath, displayText, textSnippets[displayText])
    }
}

; === UI相關輔助函數 ===
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

ShowDynamicTags(*) {
    helpGui := Gui("+Owner" . mainGui.Hwnd, "Dynamic Tags Help")
    helpGui.SetFont("s10", "Segoe UI")

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

    helpEdit := helpGui.Add("Edit", "ReadOnly w550 h200", helpText)
    helpEdit.SetFont("s10", "Consolas")

    helpGui.Add("Button", "Default w80", "OK").OnEvent("Click", (*) => helpGui.Destroy())
    helpGui.Show()
}

;==========================================
; === 6. 程式初始化 ===
;==========================================

; 載入用戶設定
LoadUserSettings()

; 初始化示例資料夾
rootFolder := TV.Add("Email Templates")
thanksFolder := TV.Add("Thanks", rootFolder)
TV.Add("Messages", rootFolder)

; 載入已保存的片段
LoadSnippets()

; 顯示主視窗
mainGui.Show()

; 清除並重新註冊所有熱字串
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

;==========================================
; === 7. Sync ===
;==========================================
ExportToLocalStorage(*) {
    ; 1. 檢查並讀取設定檔
    settingPath := A_ScriptDir "\snippets.ini"
    
    if !FileExist(settingPath) {
        MsgBox("找不到設定檔：" settingPath, "錯誤")
        return
    }
    
    try {
        fileContent := FileRead(settingPath, "UTF-8")
    } catch as err {
        MsgBox("讀取檔案失敗：" err.Message, "錯誤")
        return
    }

    ; 轉義特殊字符
    fileContent := StrReplace(fileContent, "\", "\\")
    fileContent := StrReplace(fileContent, "`n", "\n")
    fileContent := StrReplace(fileContent, "`r", "")
    fileContent := StrReplace(fileContent, "'", "\\'")

    ; 2. 創建 HTML
    html := "
    (
<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <title>Text Expander 設定同步</title>
    <style>
        body { font-family: Arial, sans-serif; max-width: 600px; margin: 20px auto; padding: 20px; }
        .control-panel { margin: 20px 0; }
        button { 
            padding: 10px 20px; 
            margin: 5px; 
            cursor: pointer;
            background-color: #4CAF50;
            color: white;
            border: none;
            border-radius: 4px;
        }
        button:hover { background-color: #45a049; }
        .info-box { 
            background-color: #f8f9fa; 
            padding: 15px; 
            border-radius: 4px;
            margin: 20px 0;
        }
        #message { padding: 10px; margin: 10px 0; border-radius: 4px; }
        .success { background-color: #dff0d8; color: #3c763d; }
        .error { background-color: #f2dede; color: #a94442; }
        #storage-info {
            border: 1px solid #ddd;
            padding: 15px;
            margin-top: 20px;
            border-radius: 4px;
        }
        .storage-details { margin-top: 10px; }
    </style>
</head>
<body>
    <h2>Text Expander 設定同步</h2>
    
    <div class='info-box'>
        <h3>使用說明</h3>
        <ul>
            <li>點擊「儲存設定」將目前的設定保存到瀏覽器</li>
            <li>在其他電腦上使用相同的瀏覽器</li>
            <li>點擊「讀取設定」下載已保存的設定檔</li>
        </ul>
    </div>

    <div class='control-panel'>
        <button id='saveBtn'>儲存設定</button>
        <button id='loadBtn'>讀取設定</button>
    </div>
    <div id='message'></div>

    <div id='storage-info'>
        <h3>儲存狀態</h3>
        <div id='storage-details'></div>
    </div>

    <script>
    document.addEventListener('DOMContentLoaded', function() {
        var configData = '" fileContent "';

        function updateStorageInfo() {
            var details = document.getElementById('storage-details');
            var saved = localStorage.getItem('textexpander_config');
            var timestamp = localStorage.getItem('textexpander_timestamp');
            
            if (saved && timestamp) {
                details.innerHTML = '上次儲存時間：' + new Date(timestamp).toLocaleString() + 
                                  '<br>設定大小：' + (saved.length / 1024).toFixed(2) + ' KB';
            } else {
                details.innerHTML = '尚未儲存任何設定';
            }
        }

        function showMessage(text, isError) {
            var msgDiv = document.getElementById('message');
            msgDiv.textContent = text;
            msgDiv.className = isError ? 'error' : 'success';
        }

        document.getElementById('saveBtn').addEventListener('click', function() {
            try {
                localStorage.setItem('textexpander_config', configData);
                localStorage.setItem('textexpander_timestamp', new Date().toISOString());
                showMessage('設定已儲存！', false);
                updateStorageInfo();
            } catch(e) {
                showMessage('儲存失敗：' + e.message, true);
            }
        });

        document.getElementById('loadBtn').addEventListener('click', function() {
            try {
                var saved = localStorage.getItem('textexpander_config');
                if(saved) {
                    var blob = new Blob([saved], {type: 'text/plain'});
                    var a = document.createElement('a');
                    a.href = URL.createObjectURL(blob);
                    a.download = 'snippets.ini';
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    showMessage('設定已下載！', false);
                } else {
                    showMessage('找不到已儲存的設定', true);
                }
            } catch(e) {
                showMessage('讀取失敗：' + e.message, true);
            }
        });

        // 初始顯示儲存狀態
        updateStorageInfo();
    });
    </script>
</body>
</html>
    )"

    ; 3. 寫入並開啟檔案
    tmpFile := A_Temp "\sync.html"
    try {
        if FileExist(tmpFile)
            FileDelete(tmpFile)
        FileAppend(html, tmpFile, "UTF-8")
        Run(tmpFile)
    } catch as err {
        MsgBox("錯誤：" err.Message, "錯誤")
    }
}