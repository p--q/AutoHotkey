; ==============================================================================
; File: PrescriptionFormatter_v5.0.ahk
; Version: 5.0
; Description: 処方整形 (AHK v2) - 処理順序適正化・ブロック保護完全版
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
        ; 先に単位のマーキング（関数B相当）を行う
        text := ApplyBasicFormatting(text)
        if (RegExMatch(text, "処方日"))
            text := ReorganizeByTrigger(text)
        
        lines := StrSplit(text, "`n", "`r")
        result := ""
        for line in lines {
            ; @@SPACE@@ が付いている（＝薬品＋数量）行のみ抽出
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
    
    ; 1. まず各行の基本整形（分1削除、単位マーキング）
    text := ApplyBasicFormatting(text)
    
    ; 2. 処方日がある場合、ブロック処理（関数D）
    if (RegExMatch(text, "処方日"))
        text := ReorganizeByTrigger(text)
    
    ; 3. 「時」の結合（関数C）
    text := MergeSpecificPatterns(text)
    
    ; 4. 余計な空白削除（マーカーは維持）
    text := RegExReplace(text, "[ \t]+", "")
    
    lines := StrSplit(text, "`n", "`r")
    processedLines := []
    
    for line in lines {
        if (line == "") continue

        ; 保護マーカー @@BLOCK@@ があれば結合せず独立
        if (InStr(line, "@@BLOCK@@")) {
            processedLines.Push(StrReplace(line, "@@BLOCK@@", ""))
            continue
        }

        ; 用法マージ判定
        if (RegExMatch(line, "^(分\d|1日\d回|\d回分)")) {
            line := RegExReplace(line, "毎(?=.)|食後", "")
            line := RegExReplace(line, "(?:と)?眠前", "寝")
            line := RegExReplace(line, "食前", "前")
            line := RegExReplace(line, "\[食間\]", "")
            
            prevLine := (processedLines.Length > 0) ? processedLines[processedLines.Length] : ""
            ; 直前が「時」で終わる場合は結合しない
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
    ; 分1（単独行）の削除
    text := RegExReplace(text, "m)(*ANYCRLF)^分1$", "")
    ; 用法末尾の「分」削除
    text := RegExReplace(text, "m)(*ANYCRLF)\d+\S*分$", "")
    ; 単位のマーキング（@@SPACE@@ を挿入）
    text := RegExReplace(text, "m)(*ANYCRLF)(\d+\S*[錠p枚ﾄ]$|\s\d+\S*g$)", "@@SPACE@@$1")
    ; cap -> c
    text := RegExReplace(text, "m)(*ANYCRLF)cap$", "c")
    return text
}

MergeSpecificPatterns(text) {
    lines := StrSplit(text, "`n", "`r")
    result := []
    for line in lines {
        if (line == "") continue
        if (InStr(line, "@@BLOCK@@")) {
            result.Push(line)
            continue
        }
        ; 行末が「時」
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
            if (currentBlock.Length > 0) blocks.Push(currentBlock)
            currentBlock := []
            continue
        }
        if (line != "") currentBlock.Push(line)
    }
    if (currentBlock.Length > 0) blocks.Push(currentBlock)

    finalOutput := ""
    for blockLines in blocks {
        triggerCount := 0
        for l in blockLines {
            ; ApplyBasicFormatting で付与された @@SPACE@@ を元に判定
            if (InStr(l, "@@SPACE@@")) triggerCount++
        }

        buffer := ""
        for line in blockLines {
            if (InStr(line, "@@SPACE@@")) {
                outLine := buffer . line
                if (triggerCount > 1) outLine .= "@@BLOCK@@"
                finalOutput .= outLine "`n"
                buffer := ""
            } else {
                if (!InStr(line, " ")) {
                    buffer .= line
                } else {
                    if (triggerCount > 1) line .= "@@BLOCK@@"
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
        if (line == "" || RegExMatch(line, "^(--|<R|処方箋)")) continue
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
