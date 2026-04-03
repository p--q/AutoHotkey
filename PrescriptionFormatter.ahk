; ==============================================================================
; File: PrescriptionFormatter_v4.9.ahk
; Version: 4.9
; Description: 処方整形 (AHK v2) - 分1削除・複数薬品ブロック保護版
; ==============================================================================

#Requires AutoHotkey v2.0

#!s:: {
    text := ProcessInitialInput()
    if (RegExMatch(text, "^商品名")) {
        text := RegExReplace(text, "^商品名\s*", "")
        text := FinalizeText(text)
        A_Clipboard := text
        ToolTip("整形完了(用法錠数なし)")
    } else {
        if (RegExMatch(text, "^処方日"))
            text := ReorganizeByTrigger(text)
        
        lines := StrSplit(text, "`n", "`r")
        result := ""
        for line in lines {
            if (RegExMatch(line, "(\d+\S*[錠p枚ﾄ]$|\s\d+\S*g$)")) {
                line := RegExReplace(line, "(\d+\S*[錠p枚ﾄ]$|\s\d+\S*g$)", "@@SPACE@@$1")
                line := RegExReplace(line, "cap$", "c")
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

#!d:: {
    text := ProcessInitialInput()
    if (SubStr(text, 1, 2) == "--") {
        text := FilterOutpatientOrder(text)
    } else if (RegExMatch(text, "^処方日")) {
        text := ReorganizeByTrigger(text)
    }
    
    text := ApplyBasicFormatting(text)
    text := MergeSpecificPatterns(text)
    text := RegExReplace(text, "[ \t]+", "")
    
    lines := StrSplit(text, "`n", "`r")
    processedLines := []
    
    for line in lines {
        if (line == "")
            continue

        ; 保護マーカーがある場合は結合せずにそのままPush
        if (InStr(line, "@@BLOCK@@")) {
            processedLines.Push(StrReplace(line, "@@BLOCK@@", ""))
            continue
        }

        if (RegExMatch(line, "^(分\d|1日\d回|\d回分)")) {
            line := RegExReplace(line, "毎(?=.)|食後", "")
            line := RegExReplace(line, "(?:と)?眠前", "寝")
            line := RegExReplace(line, "食前", "前")
            line := RegExReplace(line, "\[食間\]", "")
            
            prevLine := (processedLines.Length > 0) ? processedLines[processedLines.Length] : ""
            if (processedLines.Length > 0 && !RegExMatch(prevLine, "時$")) {
                processedLines[processedLines.Length] .= line
            } else {
                processedLines.Push(line)
            }
        } else {
            processedLines.Push(line)
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
    text := RegExReplace(text, "m)(*ANYCRLF)\d+\S*分$", "")
    ; 追加：独立した「分1」を削除
    text := RegExReplace(text, "m)(*ANYCRLF)^分1$", "")
    text := RegExReplace(text, "m)(*ANYCRLF)(\d+\S*[錠p枚ﾄ]$|\s\d+\S*g$)", "@@SPACE@@$1")
    text := RegExReplace(text, "m)(*ANYCRLF)cap$", "c")
    return text
}

MergeSpecificPatterns(text) {
    lines := StrSplit(text, "`n", "`r")
    result := []
    for line in lines {
        if (line == "")
            continue

        ; 保護マーカーがある場合は結合をスキップ
        if (InStr(line, "@@BLOCK@@")) {
            result.Push(line)
            continue
        }

        if (RegExMatch(line, "^.+時$")) {
            if (result.Length > 0 && !InStr(result[result.Length], "@@BLOCK@@"))
                result[result.Length] .= line
            else
                result.Push(line)
        } else if (RegExMatch(line, "^分\d+\s\d")) {
            line := RegExReplace(line, "^(分\d+)\s(\d)", "$1@@SPACE@@$2")
            result.Push(line)
        } else if (RegExMatch(line, "^外\)\s(.*)$", &m)) {
            if (result.Length > 0 && !InStr(result[result.Length], "@@BLOCK@@"))
                result[result.Length] .= "@@SPACE@@" . m[1]
            else
                result.Push("@@SPACE@@" . m[1])
        } else if (RegExMatch(line, "^吸入用")) {
            result.Push(RegExReplace(line, "^吸入用", ""))
        } else {
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
    text := RegExReplace(text, "\(\Sとして\)", "")
    text := StrReplace(text, "(非持参)", "")
    return Trim(text, "`n`r")
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

ReorganizeByTrigger(text) {
    blocks := []
    currentBlock := []
    
    ; 1. 処方日ごとにテキストを分割
    lines := StrSplit(text, "`n", "`r")
    for line in lines {
        if (RegExMatch(line, "^処方日")) {
            if (currentBlock.Length > 0)
                blocks.Push(currentBlock)
            currentBlock := []
            continue
        }
        if (line != "")
            currentBlock.Push(line)
    }
    if (currentBlock.Length > 0)
        blocks.Push(currentBlock)

    finalOutput := ""
    for blockLines in blocks {
        ; このブロック内にトリガー行（薬品終端）がいくつあるか数える
        triggerCount := 0
        for l in blockLines {
            if (RegExMatch(l, "(\d+\S*[錠p枚ﾄ]$|\s\d+\S*g$)"))
                triggerCount++
        }

        buffer := ""
        for line in blockLines {
            if (RegExMatch(line, "(\d+\S*[錠p枚ﾄ]$|\s\d+\S*g$)")) {
                outLine := buffer . line
                ; トリガーが複数ある場合は保護マーカーを付与
                if (triggerCount > 1)
                    outLine .= "@@BLOCK@@"
                finalOutput .= outLine "`n"
                buffer := ""
            } else {
                if (!InStr(line, " "))
                    buffer .= line
                else {
                    ; 用法行にもマーカー付与
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

ConvertToHalfWidth(str) {
    size := DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", 0, "Int", 0)
    buf := Buffer(size * 2)
    DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", buf, "Int", size)
    return StrGet(buf, "UTF-16")
}
