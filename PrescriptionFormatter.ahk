; ==============================================================================
; File: PrescriptionFormatter_v5.3.ahk
; Version: 5.3
; Description: 処方整形 (AHK v2) - 関数定義漏れ・構文エラー修正完了版
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
    text := RegExReplace(text, "m)(*ANYCRLF)^分1$", "")
    text := RegExReplace(text, "m)(*ANYCRLF)\d+\S*分$", "")
    text := RegExReplace(text, "m)(*ANYCRLF)(\d+\S*[錠p枚ﾄ]$|\s\d+\S*g$)", "@@SPACE@@$1")
    text := RegExReplace(text, "m)(*ANYCRLF)cap$", "c")
    return text
}

; 修正箇所: 関数名 MergeSpecificPatterns(text) を復旧
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
    text := RegExReplace(text, "吸入用", "")
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
    DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", buf, "Int", size)
    return StrGet(buf, "UTF-16")
}
