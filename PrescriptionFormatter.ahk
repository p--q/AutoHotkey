; ==============================================================================
; File: PrescriptionFormatter_v6.6.3.ahk
; Version: 6.6.3
; Description: 重複ロジックを PrepareFormatting に集約。
;              取得元メッセージの動的生成とエラーハンドリングを一括管理。
; ==============================================================================

#Requires AutoHotkey v2.0

; --- Win + Alt + S: 用法なし整形 ---
#!s:: {
    ; 引数に表示用メッセージ(用法錠数なし / 用法なし)を指定
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
        return false ; 呼び出し元で if !(...) を使って中断させる
    }
    
    sourceLabel := (resultObj.Source == "Selected") ? "選択を整形" : "クリップボードを整形"
    return {Text: resultObj.Text, Msg: sourceLabel "(" suffix ")"}
}

; --- 共通関数: 入力処理 ---
ProcessInitialInput() {
    savedClip := A_Clipboard
    A_Clipboard := "" 
    Send("^c")
    if ClipWait(0.5) {
        rawText := A_Clipboard
        source := "Selected"
    } else {
        rawText := savedClip
        source := "Clipboard"
    }
    return {Text: ConvertToHalfWidth(rawText), Source: source}
}

NotifyError() {
    ToolTip("整形する文字列を取得できませんでした")
    SetTimer(() => ToolTip(), -2000)
}

; --- 共通関数: 各種整形ロジック (v6.6.1継承) ---
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
            if (currentBlock.Length > 0) blocks.Push(currentBlock)
            currentBlock := []
        } else if (line != "") {
            currentBlock.Push(line)
        }
    }
    if (currentBlock.Length > 0) blocks.Push(currentBlock)
    finalOutput := ""
    for blockLines in blocks {
        triggerCount := 0
        for l in blockLines {
            if (InStr(l, "@@SPACE@@")) triggerCount++
        }
        buffer := ""
        for line in blockLines {
            if (InStr(line, "@@SPACE@@")) {
                finalOutput .= buffer . line . (triggerCount > 1 ? "@@BLOCK@@" : "") . "`n"
                buffer := ""
            } else {
                if (!InStr(line, " ") && !InStr(line, "　")) buffer .= line
                else finalOutput .= line . (triggerCount > 1 ? "@@BLOCK@@" : "") . "`n"
            }
        }
        if (buffer != "") finalOutput .= buffer . (triggerCount > 1 ? "@@BLOCK@@" : "") . "`n"
    }
    return finalOutput
}

MergeSpecificPatterns(text) {
    lines := StrSplit(text, "`n", "`r"), result := []
    for line in lines {
        if (line == "") continue
        if (InStr(line, "@@BLOCK@@")) {
            result.Push(line)
            continue
        }
        if (RegExMatch(line, "^.+時\s*$")) {
            if (result.Length > 0 && !InStr(result[result.Length], "@@BLOCK@@")) result[result.Length] .= line
            else result.Push(line)
        } else if (RegExMatch(line, "^\s*外\)\s*(.*)$", &m)) {
            if (result.Length > 0 && !InStr(result[result.Length], "@@BLOCK@@")) result[result.Length] .= "@@SPACE@@" . m[1]
            else result.Push("@@SPACE@@" . m[1])
        } else result.Push(line)
    }
    finalOutput := ""
    for l in result
        finalOutput .= l "`n"
    return finalOutput
}

FinalizeText(text) {
    text := StrReplace(text, "@@SPACE@@", " ")
    text := StrReplace(text, "@@BLOCK@@", "")
    text := RegExReplace(text, "\(\Sとして\)", "")
    text := StrReplace(text, "(非持参)", "")
    text := RegExReplace(text, " +", " ")
    return Trim(text, "`n`r")
}

FilterOutpatientOrder(text) {
    lines := StrSplit(text, "`n", "`r"), result := ""
    for line in lines {
        if (line == "" || RegExMatch(line, "^(--|<R|処方箋)")) continue
        result .= line "`n"
    }
    return result
}

ConvertToHalfWidth(str) {
    if (str == "") return ""
    size := DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", 0, "Int", 0)
    buf := Buffer(size * 2)
    DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", buf, "Int", size)
    return StrGet(buf, "UTF-16")
}
