; ==============================================================================
; File: PrescriptionFormatter_v6.7.1.ahk
; Version: 6.7.1
; Description: メソッドの連結（.Push）における不正な改行を修正。
; ==============================================================================

#Requires AutoHotkey v2.0

; --- Win + Alt + S: 用法なし整形 ---
#!s:: {
    info := PrepareFormatting("用法なし")
    if (!info) {
        return
    }
    
    text := info.Text
    if (RegExMatch(text, "^商品名")) {
        text := RegExReplace(text, "^商品名\s*", "")
        text := FinalizeText(text)
        A_Clipboard := text
        ToolTip(info.Msg "(用法錠数なし)")
    } else {
        text := ApplyBasicFormatting(text)
        if (RegExMatch(text, "処方日")) {
            text := ReorganizeByTrigger(text)
        }
        
        lines := StrSplit(text, "`n", "`r")
        resText := ""
        for line in lines {
            if (InStr(line, "@@SPACE@@")) {
                resText .= line "`n"
            }
        }
        text := RegExReplace(resText, "[ \t]+", "")
        text := FinalizeText(text)
        text := RegExReplace(text, "m)\s\d+枚\R外\)\s*(1日\s*\d+枚)$", " $1")
        
        A_Clipboard := text
        ToolTip(info.Msg)
    }
    SetTimer(() => ToolTip(), -2000)
}

; --- Win + Alt + D: 用法あり整形 ---
#!d:: {
    info := PrepareFormatting("用法あり")
    if (!info) {
        return
    }
    
    text := info.Text
    if (SubStr(text, 1, 2) == "--") {
        text := FilterOutpatientOrder(text)
    } 
    
    text := ApplyBasicFormatting(text)
    
    if (RegExMatch(text, "処方日")) {
        text := ReorganizeByTrigger(text)
    }
    
    text := MergeSpecificPatterns(text)
    text := RegExReplace(text, "[ \t]+", "")
    
    lines := StrSplit(text, "`n", "`r"), processedLines := []
    for line in lines {
        if (line == "") {
            continue
        }
        if (RegExMatch(line, "^(分\d|1日\d回|1日\d枚)")) {
            line := RegExReplace(line, "毎(?=.)|食後", "")
            line := RegExReplace(line, "(?:と)?眠前", "寝")
            line := RegExReplace(line, "食前", "前")
            line := RegExReplace(line, "\[食間\]", "")
            line := StrReplace(line, "分1", "")
            isBlock := InStr(line, "@@BLOCK@@")
            line := StrReplace(line, "@@BLOCK@@", "")
            prevLine := (processedLines.Length > 0) ? processedLines[processedLines.Length] : ""
            if (!isBlock && processedLines.Length > 0 && !RegExMatch(prevLine, "時\s*$")) {
                processedLines[processedLines.Length] .= line
            } else {
                processedLines.Push(line)
            }
        } else {
            processedLines.Push(StrReplace(line, "@@BLOCK@@", ""))
        }
    }
    
    resText := ""
    for line in processedLines {
        resText .= line "`n"
    }
        
    text := FinalizeText(resText)
    text := RegExReplace(text, "m)\s\d+枚\R外\)\s*(1日\s*\d+枚)$", " $1")
    
    A_Clipboard := text
    ToolTip(info.Msg)
    SetTimer(() => ToolTip(), -2000)
}

; --- 共通関数群 ---

PrepareFormatting(suffix) {
    resultObj := ProcessInitialInput()
    if (resultObj.Text == "") {
        NotifyError()
        return false 
    }
    sourceLabel := (resultObj.Source == "Selected") ? "選択を整形" : "クリップボードを整形"; ==============================================================================
; File: PrescriptionFormatter_v6.7.0.ahk
; Version: 6.7.0
; Description: return文におけるオブジェクトリテラルの解釈エラーを回避するため
;              一旦変数に代入してから返す記法へ全面的に修正。
; ==============================================================================

#Requires AutoHotkey v2.0

; --- Win + Alt + S: 用法なし整形 ---
#!s:: {
    info := PrepareFormatting("用法なし")
    if (!info) {
        return
    }
    
    text := info.Text
    if (RegExMatch(text, "^商品名")) {
        text := RegExReplace(text, "^商品名\s*", "")
        text := FinalizeText(text)
        A_Clipboard := text
        ToolTip(info.Msg "(用法錠数なし)")
    } else {
        text := ApplyBasicFormatting(text)
        if (RegExMatch(text, "処方日")) {
            text := ReorganizeByTrigger(text)
        }
        
        lines := StrSplit(text, "`n", "`r")
        resText := ""
        for line in lines {
            if (InStr(line, "@@SPACE@@")) {
                resText .= line "`n"
            }
        }
        text := RegExReplace(resText, "[ \t]+", "")
        text := FinalizeText(text)
        text := RegExReplace(text, "m)\s\d+枚\R外\)\s*(1日\s*\d+枚)$", " $1")
        
        A_Clipboard := text
        ToolTip(info.Msg)
    }
    SetTimer(() => ToolTip(), -2000)
}

; --- Win + Alt + D: 用法あり整形 ---
#!d:: {
    info := PrepareFormatting("用法あり")
    if (!info) {
        return
    }
    
    text := info.Text
    if (SubStr(text, 1, 2) == "--") {
        text := FilterOutpatientOrder(text)
    } 
    
    text := ApplyBasicFormatting(text)
    
    if (RegExMatch(text, "処方日")) {
        text := ReorganizeByTrigger(text)
    }
    
    text := MergeSpecificPatterns(text)
    text := RegExReplace(text, "[ \t]+", "")
    
    lines := StrSplit(text, "`n", "`r"), processedLines := []
    for line in lines {
        if (line == "") {
            continue
        }
        if (RegExMatch(line, "^(分\d|1日\d回|1日\d枚)")) {
            line := RegExReplace(line, "毎(?=.)|食後", "")
            line := RegExReplace(line, "(?:と)?眠前", "寝")
            line := RegExReplace(line, "食前", "前")
            line := RegExReplace(line, "\[食間\]", "")
            line := StrReplace(line, "分1", "")
            isBlock := InStr(line, "@@BLOCK@@")
            line := StrReplace(line, "@@BLOCK@@", "")
            prevLine := (processedLines.Length > 0) ? processedLines[processedLines.Length] : ""
            if (!isBlock && processedLines.Length > 0 && !RegExMatch(prevLine, "時\s*$")) {
                processedLines[processedLines.Length] .= line
            } else {
                processedLines.Push(line)
            }
        } else {
            processedLines.Push(StrReplace(line, "@@BLOCK@@", ""))
        }
    }
    
    text := ""
    for line in processedLines {
        text .= line "`n"
    }
        
    text := FinalizeText(text)
    text := RegExReplace(text, "m)\s\d+枚\R外\)\s*(1日\s*\d+枚)$", " $1")
    
    A_Clipboard := text
    ToolTip(info.Msg)
    SetTimer(() => ToolTip(), -2000)
}

; --- 共通関数群 ---

PrepareFormatting(suffix) {
    resultObj := ProcessInitialInput()
    if (resultObj.Text == "") {
        NotifyError()
        return false 
    }
    sourceLabel := (resultObj.Source == "Selected") ? "選択を整形" : "クリップボードを整形"
    ; 変数に代入してから返すことで構文エラーを回避
    infoObj := {Text: resultObj.Text, Msg: sourceLabel "(" suffix ")"}
    return infoObj
}

ProcessInitialInput() {
    savedClip := A_Clipboard
    A_Clipboard := "" 
    Send("^c")
    if ClipWait(0.5) {
        rawText := A_Clipboard
        src := "Selected"
    } else {
        rawText := savedClip
        src := "Clipboard"
    }
    ; 変数に代入してから返すことで構文エラーを回避
    resObj := {Text: ConvertToHalfWidth(rawText), Source: src}
    return resObj
}

NotifyError() {
    ToolTip("整形する文字列を取得できませんでした")
    SetTimer(() => ToolTip(), -2000)
}

ApplyBasicFormatting(text) {
    text := RegExReplace(text, "s)\s*\([^)]+として\)", "")
    text := RegExReplace(text, "m)\d+\S+分$", "")
    text := StrReplace(text, "吸入用", "")
    text := RegExReplace(text, "i)(\d+)\s*([錠p枚ﾄ]|cap|g)", "@@SPACE@@$1$2")
    text := RegExReplace(text, "i)cap", "c")
    return text
}

ReorganizeByTrigger(text) {
    blocks := [], currentBlock := []
    lines := StrSplit(text, "`n", "`r")
    for line in lines {
        if (RegExMatch(line, "^処方日")) {
            if (currentBlock.Length > 0) {
                blocks.Push(currentBlock)
            }
            currentBlock := []
        } else if (line != "") {
            currentBlock.Push(line)
        }
    }
    if (currentBlock.Length > 0) {
        blocks.
