; ==============================================================================
; File Name: PrescriptionFormatter.ahk
; Version:   1.1.0
; Description:
;   処方箋・カルテテキスト整形スクリプト (AutoHotkey v2.0用)
;   
;   【更新履歴】
;   1.1.0: 頓服薬の結合ロジックを強化。
;          「～時」「～まで」といった用法を薬品名・数量と同じ行にまとめるよう調整。
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
    
    ; 頓服などの特殊な改行パターンの結合
    text := MergeSpecificPatterns(text)
    text := RegExReplace(text, "[ \t]+", "")
    
    lines := StrSplit(text, "`n", "`r"), processedLines := []
    for line in lines {
        if (line == "") {
            continue
        }
        
        ; 用法の簡略化
        if (RegExMatch(line, "^(分\d|1日\d回|1日\d枚)")) {
            line := RegExReplace(line, "毎(?=.)|食後", "")
            line := RegExReplace(line, "(?:と)?眠前", "寝")
            line := RegExReplace(line, "食前", "前")
            line := RegExReplace(line, "\[食間\]", "")
            line := StrReplace(line, "分1", "")
            
            isBlock := InStr(line, "@@BLOCK@@")
            line := StrReplace(line, "@@BLOCK@@", "")
            
            prevLine := (processedLines.Length > 0) ? processedLines[processedLines.Length] : ""
            
            ; 直前の行が「～時」で終わっていない場合は結合
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
    ; 最終的な微調整（「～時」の後の不要な改行を詰め、指示通りなどの不要語削除）
    text := RegExReplace(text, "m)時\R+", "時")
    text := RegExReplace(text, "医師の指示通り", "")
    
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
    return {Text: resultObj.Text, Msg: sourceLabel "(" suffix ")"}
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
    return {Text: ConvertToHalfWidth(rawText), Source: src}
}

NotifyError() {
    ToolTip("整形する文字列を取得できませんでした")
    SetTimer(() => ToolTip(), -2000)
}

ApplyBasicFormatting(text) {
    text := RegExReplace(text, "s)\s*\([^)]+として\)", "")
    text := RegExReplace(text, "m)\d+\S+分$", "")
    text := StrReplace(text, "吸入用", "")
    ; 薬品名と数量の間にスペースを入れるためのフラグ
    text := RegExReplace(text, "i)(\d+)\s*([錠p枚ﾄ個]|cap|g|mL)", "@@SPACE@@$1$2")
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
        blocks.Push(currentBlock)
    }
    
    finalOutput := ""
    for blockLines in blocks {
        triggerCount := 0
        for l in blockLines {
            if (InStr(l, "@@SPACE@@")) {
                triggerCount++
            }
        }
        
        buffer := ""
        for line in blockLines {
            if (InStr(line, "@@SPACE@@")) {
                finalOutput .= buffer . line . (triggerCount > 1 ? "@@BLOCK@@" : "") . "`n"
                buffer := ""
            } else {
                ; 空白がない行（薬品名の続きなど）をバッファして結合
                if (!InStr(line, " ") && !InStr(line, "　")) {
                    buffer .= line
                } else {
                    finalOutput .= line . (triggerCount > 1 ? "@@BLOCK@@" : "") . "`n"
                }
            }
        }
        if (buffer != "") {
            finalOutput .= buffer . (triggerCount > 1 ? "@@BLOCK@@" : "") . "`n"
        }
    }
    return finalOutput
}

MergeSpecificPatterns(text) {
    lines := StrSplit(text, "`n", "`r"), result := []
    for line in lines {
        if (line == "") {
            continue
        }
        if (InStr(line, "@@BLOCK@@")) {
            result.Push(line)
            continue
        }
        
        ; 「～時」や「～まで」で始まる行、または「頓）」を含む行の結合処理
        if (RegExMatch(line, "^(頓|.*時|.*まで)\s*$")) {
            if (result.Length > 0 && !InStr(result[result.Length], "@@BLOCK@@")) {
                result[result.Length] .= line
            } else {
                result.Push(line)
            }
        } else if (RegExMatch(line, "^\s*外\)\s*(.*)$", &m)) {
            if (result.Length > 0 && !InStr(result[result.Length], "@@BLOCK@@")) {
                result[result.Length] .= "@@SPACE@@" . m[1]
            } else {
                result.Push("@@SPACE@@" . m[1])
            }
        } else {
            result.Push(line)
        }
    }
    resFinal := ""
    for l in result {
        resFinal .= l "`n"
    }
    return resFinal
}

FinalizeText(text) {
    text := StrReplace(text, "@@SPACE@@", " ")
    text := StrReplace(text, "@@BLOCK@@", "")
    text := RegExReplace(text, "\(\Sとして\)", "")
    text := StrReplace(text, "(非持参)", "")
    text := StrReplace(text, "頓)", "")
    text := RegExReplace(text, " +", " ")
    return Trim(text, "`n`r")
}

FilterOutpatientOrder(text) {
    lines := StrSplit(text, "`n", "`r"), resFilter := ""
    for line in lines {
        if (line == "" || RegExMatch(line, "^(--|<R|処方箋)")) {
            continue
        }
        resFilter .= line "`n"
    }
    return resFilter
}

ConvertToHalfWidth(str) {
    if (str == "") {
        return ""
    }
    size := DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", 0, "Int", 0)
    buf := Buffer(size * 2)
    DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", buf, "Int", size)
    return StrGet(buf, "UTF-16")
}
