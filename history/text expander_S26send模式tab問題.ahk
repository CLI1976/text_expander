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
; 日誌管理函數

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

global treeWidth := 180
global listWidth := 0
global currentHeight := 600  
global controlHeight := 0

; Settings functions
LoadUserSettings() {
    global settingsFile, sendRadio, pasteRadio, spaceSuffixRadio, triggerSuffixRadio
    
    ; 如果設定檔不存在，創建預設設定
    if !FileExist(settingsFile) {
        try {
            IniWrite("Ctrl+V", settingsFile, "OutputMethod", "DefaultMethod")
            IniWrite("Space", settingsFile, "OutputSuffix", "DefaultSuffix")
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

        suffix := IniRead(settingsFile, "OutputSuffix", "DefaultSuffix", "Space")
        if (suffix = "Space") {
            spaceSuffixRadio.Value := 1
        } else {
            triggerSuffixRadio.Value := 1
        }
    }
}

SaveUserSettings(*) {
    global settingsFile, sendRadio, spaceSuffixRadio
    
    try {
        method := sendRadio.Value ? "Send" : "Ctrl+V"
        suffix := spaceSuffixRadio.Value ? "Space" : "Trigger"
        IniWrite(method, settingsFile, "OutputMethod", "DefaultMethod")
        IniWrite(suffix, settingsFile, "OutputSuffix", "DefaultSuffix")
    }
}

;==========================================
; === 2. GUI 相關函數及初始化 ===
;==========================================

; ========== 初始化系統托盤圖標和選單 ==========
; 設置托盤圖標
if FileExist(A_ScriptDir "\square-t.ico") {
    TraySetIcon(A_ScriptDir "\square-t.ico")
} else {
    ; 使用默認圖標
}

; 創建主視窗
mainGui := Gui("+Resize +MinSize800x600", "Text Expander")
mainGui.SetFont("s10", "Segoe UI")

; 左側樹狀視圖
TV := mainGui.Add("TreeView", "x5 y5 w180 h400")

; 右側列表視圖
; 右側分割為兩個部分
; 保持原有佈局，調整視覺樣式
LV := mainGui.Add("ListView", "x195 y5 w545 h200", ["Group Path", "Keyword", "Phrase"])
LV.SetFont("s10", "Segoe UI")

; 預覽區增加邊框和字體設定
previewEdit := mainGui.Add("Edit", "x195 y210 w545 h195 ReadOnly -E0x200")  ; -E0x200 移除灰色背景
previewEdit.SetFont("s11", "Segoe UI")

; 可選：美化欄寬比例
LV.ModifyCol(1, 10)  ; Group Path
LV.ModifyCol(2, 160)  ; Keyword
LV.ModifyCol(3, 360)  ; Phrase

; 添加按鈕
addButton := mainGui.Add("Button", "x5 y410", "Add New")
editButton := mainGui.Add("Button", "x85 y410", "Edit")
deleteButton := mainGui.Add("Button", "x135 y410", "Delete")
toggleButton := mainGui.Add("Button", "x225 y410", "ON")

; 輸出方式選擇
outputMethodText := mainGui.Add("Text", "x15 y450", "輸出方式：")
sendRadio := mainGui.Add("Radio", "x85 y450", "Send Text")
pasteRadio := mainGui.Add("Radio", "x185 y450", "Ctrl+V")

; 新增顯示模式選擇
displayModeText := mainGui.Add("Text", "x305 y450", "顯示模式：")
viewRadio := mainGui.Add("Radio", "x385 y450 Checked", "ListView")
previewRadio := mainGui.Add("Radio", "x465 y450", "Preview")

; 分組按鈕
addGroupButton := mainGui.Add("Button", "x5 y490", "New Group")
deleteGroupButton := mainGui.Add("Button", "x105 y490", "Delete Group")

; Import/Export 和搜尋按鈕
importButton := mainGui.Add("Button", "x285 y410", "Import")
exportButton := mainGui.Add("Button", "x345 y410", "Export")
searchButton := mainGui.Add("Button", "x405 y410", "Search")

; 在原有的輸出方式選擇後面添加後綴選擇
outputSuffixText := mainGui.Add("Text", "x315 y480", "輸出後綴：")
spaceSuffixRadio := mainGui.Add("Radio", "x385 y480 Checked", "空白")
triggerSuffixRadio := mainGui.Add("Radio", "x445 y480", "觸發")

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

viewRadio.OnEvent("Click", SwitchDisplayMode)
previewRadio.OnEvent("Click", SwitchDisplayMode)

addGroupButton.OnEvent("Click", AddNewGroup)
deleteGroupButton.OnEvent("Click", DeleteGroup)

; 在主程式初始化部分添加
spaceSuffixRadio.OnEvent("Click", OnSuffixChange)
triggerSuffixRadio.OnEvent("Click", OnSuffixChange)

; GUI 事件處理函數
GuiResize(thisGui, MinMax, Width, Height) {
        if MinMax = -1  ; 視窗最小化
        return
    
    ; TreeView 固定寬度
    currentHeight := Height  ; 保存當前高度
    treeWidth := 180
    listWidth := Width - treeWidth - 20
    controlHeight := Height - 100
    
    ; 更新控件大小和位置
    TV.Move(5, 5, treeWidth - 10, controlHeight)

    if (viewRadio.Value) {
        ; ListView 模式
        LV.Move(treeWidth + 5, 5, listWidth - 15, controlHeight)
        previewEdit.Visible := false
    } else {
        ; Preview 模式
        listHeight := Floor(controlHeight / 3)
        LV.Move(treeWidth + 5, 5, listWidth - 15, listHeight)
        previewEdit.Visible := true
        previewEdit.Move(treeWidth + 5, listHeight + 10, listWidth - 15, controlHeight - listHeight - 10)
    }
    ; ListView 和 Preview 各佔右側一半高度
    ; listHeight := Floor(controlHeight / 2)
    ; LV.Move(treeWidth + 5, 5, listWidth - 15, listHeight)
    ; previewEdit.Move(treeWidth + 5, listHeight + 10, listWidth - 15, listHeight - 10)
    

    ; 動態調整欄寬
    pathWidth := 10
    keywordWidth := 160
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
    pasteRadio.Move(170, radioY)
    displayModeText.Move(305, radioY)
    viewRadio.Move(385, radioY)
    previewRadio.Move(455, radioY)

    ; 新增：更新輸出後綴選項位置
    suffixY := radioY + 28  ; 在原有 radio 按鈕下方
    outputSuffixText.Move(315, suffixY)
    spaceSuffixRadio.Move(385, suffixY)
    triggerSuffixRadio.Move(445, suffixY)
    
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
    
    ; 更新預覽
    row := LV.GetNext()
    if row {
        phrase := LV.GetText(row, 3)  ; 第3欄是 Phrase
        previewEdit.Value := phrase
    }
    else
        previewEdit.Value := ""
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
    
    ; 編輯日誌
    ManageLogFile(A_ScriptDir "\edit_log.txt", "=== 開始編輯 ===")
    ManageLogFile(A_ScriptDir "\edit_log.txt", "原始值：")
    ManageLogFile(A_ScriptDir "\edit_log.txt", "群組路徑: " oldGroupPath)
    ManageLogFile(A_ScriptDir "\edit_log.txt", "關鍵字: " oldKeyword)
    ManageLogFile(A_ScriptDir "\edit_log.txt", "短語: " oldPhrase)

    
    editGui := Gui("+Owner" . mainGui.Hwnd . " +Resize", "Edit Snippet")
    editGui.SetFont("s10", "Segoe UI")
    
    editWidth -= 40
    keywordWidth := Min(editWidth - 20, 300)
    phraseWidth := editWidth - 20
    phraseHeight := editHeight - 170
    
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
        newHeight := Height - 170
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

; === 切換顯示模式的函數 === 
SwitchDisplayMode(*) {
    ; mainGui.GetPos(,, &Width)
    ; GuiResize(mainGui, 1, Width, currentHeight)
    mainGui.GetPos(,, &currentWidth, &currentHeight)  ; 獲取當前視窗的實際尺寸
    currentHeight -= 40
    ; 使用當前實際尺寸進行重繪
    GuiResize(mainGui, 1, currentWidth, currentHeight)

    listWidth := currentWidth - treeWidth - 20
    controlHeight := currentHeight - 100  ; 使用保存的高度

    if (viewRadio.Value) {
        ; ListView 模式
        previewEdit.Visible := false
        LV.Move(treeWidth + 5, 5, listWidth - 15, controlHeight)
    } else {
        ; Preview 模式
        previewEdit.Visible := true
        listHeight := Floor(controlHeight / 3)
        LV.Move(treeWidth + 5, 5, listWidth - 15, listHeight)
        previewEdit.Move(treeWidth + 5, listHeight + 10, listWidth - 15, controlHeight - listHeight - 10)
    }
    WinRedraw(mainGui)  ; 添加這行來強制刷新 GUI
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

; 修改 CreateHotstring 函數
CreateHotstring(key, value) {
    callback(*) => SendWithIMEControl(value, " ")
    callbackDot(*) => SendWithIMEControl(value, ".")
    callbackComma(*) => SendWithIMEControl(value, ",")  ; 新增逗號觸發
    
    Hotstring(":C*:" key " ", callback)
    Hotstring(":C*:" key ".", callbackDot)
    Hotstring(":C*:" key ",", callbackComma)  ; 新增逗號觸發
}

; 修改 RegisterHotstrings 函數
RegisterHotstrings(key) {
    if !textSnippets.Has(key)
        return
    text := textSnippets[key]
    
    try {
        try Hotstring(":C*:" key " ",, "Off")
        try Hotstring(":C*:" key ".",, "Off")
        try Hotstring(":C*:" key ",",, "Off")  ; 新增逗號
    }
    Sleep(50)
    
    try {
        Hotstring(":C*:" key " ", (*) => SendWithIMEControl(text))
        Hotstring(":C*:" key ".", (*) => SendWithIMEControl(text))
        Hotstring(":C*:" key ",", (*) => SendWithIMEControl(text))  ; 新增逗號
    }
}

; 修改 UnregisterHotstrings 函數
UnregisterHotstrings(key) {
    if (key = "")
        return
    try {
        try Hotstring(":C*:" key " ",, "Off")
        try Hotstring(":C*:" key ".",, "Off")
        try Hotstring(":C*:" key ",",, "Off")  ; 新增逗號
    }
    Sleep(150)
}

; 開關熱字串
ToggleHotstrings(*) {
    Suspend  ; 切換暫停狀態
    
    ; 更新按鈕文字和顏色
    toggleButton.Text := A_IsSuspended ? "OFF" : "ON"
    toggleButton.Opt(A_IsSuspended ? "+cRed" : "+cGreen")
    
    TrayTip("Text Expander", A_IsSuspended ? "自動替換已禁用" : "自動替換已啟用")
}


SendWithIMEControl(text, trigger := " ") {
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
    

    ; *** 在這裡加入換行符標準化 ***
    ; text := StrReplace(text, "`n", "`r`n")
    ; text := StrReplace(text, "`r`r`n", "`r`n")  ; 防止重複的 CR
    
    ; 根據設定決定後綴
    suffix := spaceSuffixRadio.Value ? " " : trigger
    text .= suffix

    if (pasteRadio.Value) {
        ; 改進的剪貼板方式 - 增加重試機制
            ; *** 在這裡加入換行符標準化 ***
        text := StrReplace(text, "`n", "`r`n")
        text := StrReplace(text, "`r`r`n", "`r`n")  ; 防止重複的 CR
        SendTextViaClipboard(text)
    } else {
        ; 改進的 Send 方式
        SendTextDirect(text)
    }
}

; 新增：使用剪貼板的可靠版本
SendTextViaClipboard(text) {
    ; 獲取目標窗口
    targetWindow := WinExist("A")
    if !targetWindow {
        SendTextDirect(text)
        return
    }
    
    ; 備份剪貼板內容
    backupText := A_Clipboard
    
    try {
        ; 設置剪貼板並發送
        A_Clipboard := text
        Sleep(50)
        
        ; 確保窗口焦點
        if (WinExist("A") != targetWindow) {
            WinActivate(targetWindow)
            Sleep(50)
        }
        
        Send("^v")
        
        ; 延遲恢復剪貼板
        SetTimer(() => (A_Clipboard := backupText), -200)
        
    } catch as err {
        ; 出錯時回退到直接發送
        SendTextDirect(text)
    }
}

; 新增：直接發送的可靠版本
SendTextDirect(text) {
    ; 如果文本包含換行+tab組合，切換到剪貼板模式
    if InStr(text, "`n`t") {
        SendTextViaClipboard(text)
        return
    }
    
    ; 保存當前設定
    oldDelay := A_KeyDelay
    oldMode := A_SendMode
    
    try {
        ; 優化發送設定
        SetKeyDelay(1, 1)  ; 稍微增加延遲避免字符丟失
        SendMode("Input")
        
        ; 獲取當前 IME 狀態
        targetWindow := WinExist("A")
        if !targetWindow {
            return
        }
        
        ; 暫時切換到英文輸入模式
        prevIME := DllCall("GetKeyboardLayout", "UInt", 
                   DllCall("GetWindowThreadProcessId", "UInt", targetWindow, "UInt", 0))
        
        ; 強制切換到英文鍵盤
        try {
            SendMessage(0x50, 0, 0x4090409,, "A")
            Sleep(50)  ; 等待 IME 切換完成
        }
        
        ; 分段發送長文本，避免緩衝區問題
        if (StrLen(text) > 100) {
            SendTextInChunks(text)
        } else {
            Send("{Text}" . text)
        }
        
        ; 等待發送完成
        Sleep(50)
        
    } catch as err {
        ; 發生錯誤時的回退處理
        try {
            Send(text)
        }
    } finally {
        ; 恢復原始設定
        SetKeyDelay(oldDelay)
        SendMode(oldMode)
        
        ; 延遲恢復 IME 狀態，避免干擾剛發送的文本
        if (prevIME) {
            SetTimer(() => PostMessage(0x50, 0, prevIME,, "A"), -150)
        }
    }
}

SendTextInChunks(text) {
    chunkSize := 50
    textLen := StrLen(text)
    pos := 1
    
    while (pos <= textLen) {
        chunk := SubStr(text, pos, chunkSize)
        Send("{Text}" . chunk)
        pos += chunkSize
        
        ; 每個區塊之間稍微暫停
        if (pos <= textLen) {
            Sleep(10)
        }
    }
}

; 新增：延遲恢復剪貼板的輔助函數
RestoreClipboard(backupData) {
    try {
        A_Clipboard := backupData
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
        
        ManageLogFile(logFile, "=== 開始保存過程 " A_Now " ===")
        ManageLogFile(logFile, "正在處理檔案: " snippetFile)
                        
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
                ManageLogFile(logFile, "處理頂層群組: " nodeName)
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
                
                ; 註冊大小寫敏感的熱字串
                if !InStr(key, "memo_")
                    RegisterHotstrings(key)
            }
        }
    }
 }

LoadSnippetsWithoutHotstring() {
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
        ManageLogFile(A_ScriptDir "\export_log.txt", "=== 開始導出 " A_Now " ===")
        
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
    
    if InStr(displayText, "memo_") {
        if textSnippets.Has(displayText) {
            groupPath := GetFullGroupPath(TV.GetParent(itemNode))
            modifiedDisplay := "<memo> " StrReplace(displayText, "memo_", "")
            LV.Add(, groupPath, modifiedDisplay, textSnippets[displayText])
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
        LV.Add(, groupPath, displayText, textSnippets[displayText])  ; 正常顏色
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

; === 異步載入讓程式在載入檔案時不會凍結GUI ===
LoadSnippetsAsync() {
    ManageLogFile(A_ScriptDir "\load_debug.txt", "=== 開始異步載入 ===")
    progressGui := Gui("+Owner" mainGui.Hwnd)
    progressGui.Add("Text",, "載入中...")
    progressGui.Show("AutoSize")
    
    
    TV.Opt("-Redraw")
    LV.Opt("-Redraw")
    
    ; 只讀取檔案和更新 UI，不註冊熱字串
    LoadSnippetsWithoutHotstring()  
    
    TV.Opt("+Redraw")
    LV.Opt("+Redraw")
    
    ; 一次性註冊所有熱字串
    FileAppend("=== 註冊熱字串 ===`n", A_ScriptDir "\load_debug.txt")
    for key, value in textSnippets {
        if !InStr(key, "memo_") {
            CreateHotstring(key, value)
            ManageLogFile(A_ScriptDir "\load_debug.txt", "註冊: " key)
        }
    }
    ManageLogFile(A_ScriptDir "\load_debug.txt", "=== 載入完成 ===")
    progressGui.Destroy()
}

; === 日誌管理函數 ===
ManageLogFile(logFile, content, maxSize := 20480) {  ; 20KB = 20480 bytes
    try {
        ; 檢查文件是否存在及其大小
        if FileExist(logFile) {
            fileObj := FileOpen(logFile, "r")
            if fileObj {
                fileSize := fileObj.Length
                fileObj.Close()
                
                ; 如果文件大小超過限制，刪除舊文件
                if (fileSize > maxSize) {
                    FileDelete(logFile)
                }
            }
        }
        
        ; 添加新的日誌內容
        FileAppend(content "`n", logFile)
    } catch Error as err {
        ; 如果發生錯誤，靜默處理或者可以記錄到系統事件日誌
        ; OutputDebug("Log error: " err.Message)
    }
}

; === 切換輸出後綴選項時重新註冊所有熱字串，確保立即生效 ===
OnSuffixChange(*) {
    SaveUserSettings()
    ; 重新註冊所有熱字串
    for key, value in textSnippets {
        if !InStr(key, "memo_") {
            CreateHotstring(key, value)
        }
    }
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

; 顯示主視窗
mainGui.Show()

; 異步載入片段
SetTimer LoadSnippetsAsync, -1

;==========================================
; === 7. Sync ===
;==========================================
; 暫時放棄