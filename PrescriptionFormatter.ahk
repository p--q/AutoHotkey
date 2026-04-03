; ==============================================================================
; File: PrescriptionFormatter_v6.8.ahk
; Version: 6.8
; Description: 処方整形 (AHK v2) - 薬品名内改行修復 ＆ 用法行(日分)のトリガー除外
; ==============================================================================

#Requires AutoHotkey v2.0

; --- Sキー: 用法なし ---
#!s:: {
    text := ProcessInitialInput()
    if (RegExMatch(text, "^商品名")) {
        text := RegExReplace(text, "^商品名\s*", "")
        text := FinalizeText(text)
        A_Clipboard := text
        ToolTip("整形完了(用法錠数なし)")
    } else {
        text := ApplyBasicFormatting(text)
        if (RegExMatch(text, "処方日"))
            text := ReorganizeByTrigger(text)
        
        lines := StrSplit(text, "`n", "`r")
        result := ""
        for line in lines {
            if (InStr(line, "@@SPACE@@")) {
                result .= line "`n"
            }
        }
        text := RegExReplace(result, "[ \t]+", "")
        text := FinalizeText(text)
        A_Clipboard := text
        ToolTip("整形完了(用法なし)")
    }
    SetTimer(() => ToolTip(), -2000)
}

; --- Dキー: 用法あり ---
#!d:: {
    text := ProcessInitialInput()
    if (SubStr(text, 1, 2) == "--") {
        text := FilterOutpatientOrder(text)
    } 
    
    text := ApplyBasicFormatting(text)
    
    if (RegExMatch(text, "処方日"))
        text := ReorganizeByTrigger(text)
    
    text := MergeSpecificPatterns(text)
    text := RegExReplace(text, "[ \t]+", "")
    
    lines := StrSplit(text, "`n", "`r")
    processedLines := []
    
    for line in lines {
        if (line == "")
            continue

        ; 用法判定 (分n, 1日n回, 1日n枚)
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
    for line in processedLines
        text .= line "`n"
        
    A_Clipboard := FinalizeText(text)
    ToolTip("整形完了(用法あり)")
    SetTimer(() => ToolTip(), -2000)
}

; --- 関数群 ---

ApplyBasicFormatting(text) {
    text := StrReplace(text, "吸入用", "")
    
    ; --- 薬品名内の不自然な改行修復 ---
    ; 1. 単位の途中で切れたもの (500m\ng) を結合
    text := RegExReplace(text, "(\d+)m\s*\r?\n\s*g", "$1mg")
    ; 2. 薬品名(カタカナ)の途中で切れたものを結合
    text := RegExReplace(text, "([ァ-ヶ])\s*\r?\n\s*([ァ-ヶ])", "$1$2")
    ; 3. 括弧書き内の改行を修復
    text := RegExReplace(text, "\(([^)]*)\s*\r?\n\s*([^)]*)\)", "($1$2)")
    
    ; 用法末尾の「分/日分」を一時的に保護するための処理
    text := RegExReplace(text, "m)(*ANYCRLF)\d+\S*分$", "")

    ; 数量マーカー付与ロジック
    ; 否定先読み (?!.*(?:外\)|日分)) により、「外)」や「日分」を含む行をトリガーから除外
    unitPattern := "(\d+\S*[錠p枚ﾄg]|ｷｯﾄ)$"
    text := RegExReplace(text, "m)(*ANYCRLF)^(?!.*(?:外\)|日分))(?=.*" . unitPattern . ").*?\K" . unitPattern, "@@SPACE@@$1")
    
    text := RegExReplace(text, "m)(*ANYCRLF)cap$", "c")
    return text
}

MergeSpecificPatterns(text) {
    lines := StrSplit(text, "`n", "`r")
    result := []
    for line in lines {
        if (line == "")
            continue
        if (InStr(line, "@@BLOCK@@")) {
            result.Push(line)
            continue
        }
        
        if (RegExMatch(line, "^.+時\s*$")) {
            if (result.Length > 0 && !InStr(result[result.Length], "@@BLOCK@@"))
                result[result.Length] .= line
            else
                result.Push(line)
        } 
        else if (RegExMatch(line, "^\s*外\)\s*(.*)$", &m)) {
            if (result.Length > 0 && !InStr(result[result.Length], "@@BLOCK@@"))
                result[result.Length] .= "@@SPACE@@" . m[1]
            else
                result.Push("@@SPACE@@" . m[1])
        } 
        else {
            result.Push(line)
        }
    }
    finalText := ""
    for l in result
        finalText .= l "`n"
    return finalText
}

FinalizeText(text) {
    text := StrReplace(text, "@@SPACE@@", " ")
    
    ; 数量と用法の「枚」重複を整理
    text := RegExReplace(text, "\d+枚\s(1日\d+枚)", "$1")
    
    ; 「(レボフロキサシンとして)」等の注釈および括弧内の残骸を削除
    text := RegExReplace(text, "\([^)]+として\)", "")
    text := RegExReplace(text, "\(\Sとして\)", "")
    text := StrReplace(text, "(非持参)", "")
    
    ; 余分なスペースの掃除
    text := RegExReplace(text, " +", " ")
    
    return Trim(text, "`n`r")
}

ReorganizeByTrigger(text) {
    blocks := []
    currentBlock := []
    lines := StrSplit(text, "`n", "`r")
    for line in lines {
        if (RegExMatch(line, "^処方日")) {
            if (currentBlock.Length > 0)
                blocks.Push(currentBlock)
            currentBlock := []
        } else if (line != "") {
            currentBlock.Push(line)
        }
    }
    if (currentBlock.Length > 0)
        blocks.Push(currentBlock)

    finalOutput := ""
    for blockLines in blocks {
        triggerCount := 0
        for l in blockLines {
            if (InStr(l, "@@SPACE@@"))
                triggerCount++
        }
        buffer := ""
        for line in blockLines {
            if (InStr(line, "@@SPACE@@")) {
                outLine := buffer . line
                if (triggerCount > 1)
                    outLine .= "@@BLOCK@@"
                finalOutput .= outLine "`n"
                buffer := ""
            } else {
                ; スペース（区切り）がない行を薬品名の一部として溜める
                if (!InStr(line, " ")) {
                    buffer .= line
                } else {
                    if (triggerCount > 1)
                        line .= "@@BLOCK@@"
                    finalOutput .= line "`n"
                }
            }
        }
        if (buffer != "")
            finalOutput .= buffer . (triggerCount > 1 ? "@@BLOCK@@" : "") . "`n"
    }
    return finalOutput
}

FilterOutpatientOrder(text) {
    lines := StrSplit(text, "`n", "`r")
    result := ""
    for line in lines {
        if (line == "" || RegExMatch(line, "^(--|<R|処方箋)"))
            continue
        result .= line "`n"
    }
    return result
}

ProcessInitialInput() {
    savedClip := A_Clipboard
    A_Clipboard := ""
    Send("^c")
    if !ClipWait(0.5)
        A_Clipboard := savedClip
    return ConvertToHalfWidth(A_Clipboard)
}

ConvertToHalfWidth(str) {
    size := DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", 0, "Int", 0)
    buf := Buffer(size * 2)
    DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x0040000
