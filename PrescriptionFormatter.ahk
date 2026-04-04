; ==============================================================================
; File: PrescriptionFormatter_v6.6.5.ahk
; Version: 6.6.5
; Description: ProcessInitialInput 内の return 文の閉じ括弧不足を修正。
; ==============================================================================

#Requires AutoHotkey v2.0

; --- Win + Alt + S: 用法なし整形 ---
#!s:: {
    if !(info := PrepareFormatting("用法なし"))
        return
    
    text := info.Text
    if (RegExMatch(text, "^商品名")) {
        text := RegExReplace(text, "^商品名\s*", "")
        text := FinalizeText(text)
        A_Clipboard := text
        ToolTip(info.Msg "(用法錠数なし)")
    } else {
        text := ApplyBasicFormatting(text)
        if (RegExMatch(text, "処方日"))
            text := ReorganizeByTrigger(text)
        
        lines := StrSplit(text, "`n", "`r")
        res := ""
        for line in lines {
            if (InStr(line, "@@SPACE@@"))
                res .= line "`n"
        }
        text := RegExReplace(res, "[ \t]+", "")
        text := FinalizeText(text)
        text := RegExReplace(text, "m)\s\d+枚\R外\)\s*(1日\s*\d+枚)$", " $1")
        
        A_Clipboard := text
        ToolTip(info.Msg)
    }
    SetTimer(() => ToolTip(), -2000)
}

; --- Win + Alt + D: 用法あり整形 ---
#!d:: {
    if !(info := PrepareFormatting("用法あり"))
        return
    
    text := info.Text
    if (SubStr(text, 1, 2) == "--") {
        text := FilterOutpatientOrder(text)
    } 
    
    text := ApplyBasicFormatting(text)
    
    if (RegExMatch(text, "処方日"))
        text := ReorganizeByTrigger(text)
    
    text := MergeSpecificPatterns(text)
    text := RegExReplace(text, "[ \t]+", "")
    
    lines := StrSplit(text, "`n", "`r"), processedLines := []
    for line in lines {
        if (line == "") continue
        if (RegExMatch(line, "^(分\d|1日\d回|1日\d枚)")) {
            line := RegExReplace(line, "毎(?=.)|食後", "")
            line := RegExReplace(line, "(?:と)?眠前", "寝"), line := RegExReplace(line, "食前", "前")
            line := RegExReplace(line, "\[食間\]", ""), line := StrReplace(line, "分1", "")
            isBlock := InStr(line, "@@BLOCK@@"), line := StrReplace(line, "@@BLOCK@@", "")
            prevLine := (processedLines.Length > 0) ? processedLines[processedLines.Length] : ""
            if (!isBlock && processedLines.Length > 0 && !RegExMatch(prevLine, "時\s*$"))
                processedLines[processedLines.Length] .= line
            else
                processedLines.Push(line)
        } else {
            processedLines.Push(StrReplace(line, "@@BLOCK@@", ""))
        }
    }
    text := ""
    for line in processedLines
        text .= line "`n"
        
    text := FinalizeText(text)
    text := RegExReplace(text, "m)\s\d+枚\R外\)\s*(1日\s*\d+枚)$", " $1")
    
    A_Clipboard := text
    ToolTip(info.Msg)
    SetTimer(() => ToolTip(), -2000)
}

; --- 統合された前処理関数 ---
PrepareFormatting(suffix) {
    resultObj := ProcessInitialInput()
    if (resultObj.Text == "") {
        NotifyError()
        return false 
    }
    sourceLabel := (resultObj.Source == "Selected") ? "選択を整形" : "クリップボードを整形"
    return {Text: resultObj.Text, Msg: sourceLabel "(" suffix ")"}
}

; --- 共通関数: 入力処理 ---
ProcessInitialInput() {
    savedClip := A
