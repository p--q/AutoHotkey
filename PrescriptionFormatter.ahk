/*
 * File: PrescriptionFormatter.ahk
 * Version: 1.6
 * Description: 処方オーダーおよび薬品DI情報のテキストを整形します。
 * 行頭の「分1-3」「外)」「吸入用」、および行末の「分」を含む行を物理的に削除します。
 */

#Requires AutoHotkey v2.0
#SingleInstance Force

; ------------------------------------------------------------------
; Main Hotkey: Win + Alt + S (Output without Usage)
; ------------------------------------------------------------------
#!s:: {
    text := GetTextAndConvertToHalfWidth()
    
    if (RegExMatch(text, "m)^商品名")) {
        text := RegExReplace(text, "m)^商品名[ 　\t]*", "")
    } else {
        ; --- 行削除セクション ---
        ; 「分1,2,3」で始まる行を削除
        text := RegExReplace(text, "m)^分[123].*(\R|$)", "")
        ; 「外)」で始まる行を削除
        text := RegExReplace(text, "m)^外\).*(\R|$)", "")
        ; 「吸入用」で始まる行を削除
        text := RegExReplace(text, "m)^吸入用.*(\R|$)", "")
        ; 行末が「数字+文字+分」で終わる行を削除
        text := RegExReplace(text, "m)^.*\d+\S+分$(\R|$)", "")
        
        text := RemoveLinesByPrescriptionType(text)
        text := InsertSpaceBeforeUnitAndReplace(text)
        
        ; 改行以外の空白（半角・全角・タブ）をすべて削除
        text := RegExReplace(text, "[ 　\t]+", "")
        ; 保存しておいたスペースマーカーを半角スペースに戻す
        text := StrReplace(text, "@@SPACE@@", " ")
    }
    
    text := RegExReplace(text, "\(\Sとして\)", "")
    A_Clipboard := Trim(text, "`r`n")
    ShowNotification("Formatted (No Usage)")
}

; ------------------------------------------------------------------
; Main Hotkey: Win + Alt + D (Output with Usage)
; ------------------------------------------------------------------
#!d:: {
    text := GetTextAndConvertToHalfWidth()
    text := RemoveLinesByPrescriptionType(text)
    text := InsertSpaceBeforeUnitAndReplace(text)
    text := CombineUsageLinesAndRemoveSpecificWords(text)
    
    ; 改行以外の空白（半角・全角・タブ）をすべて削除
    text := RegExReplace(text, "[ 　\t]+", "")
    
    lines := StrSplit(text, "`n", "`r")
    processedText := ""
    for line in lines {
        if (line == "") {
            continue
        }
        
        currentLine := line
        if (RegExMatch(currentLine, "^分[123]\S+")) {
            currentLine := AbbreviateUsagePatternB(currentLine)
        } else if (RegExMatch(currentLine, "^分\d\S+")) {
            currentLine := AbbreviateUsagePatternC(currentLine)
        }
        processedText .= currentLine "`r`n"
    }
    
    text := StrReplace(processedText, "@@SPACE@@", " ")
    text := RegExReplace(text, "\(\Sとして\)", "")
    
    A_Clipboard := Trim(text, "`r`n")
    ShowNotification("Formatted (With Usage)")
}

; ------------------------------------------------------------------
; Sub-functions
; ------------------------------------------------------------------

GetTextAndConvertToHalfWidth() {
    oldClip := A_Clipboard
    A_Clipboard := ""
    Send "^c"
    if !ClipWait(0.5) {
        A_Clipboard := oldClip
    }
    return ConvertFullToHalf(A_Clipboard)
}

ConvertFullToHalf(str) {
    if (str == "") {
        return ""
    }
    size := DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", 0, "Int", 0)
    buf := Buffer(size * 2)
    DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", buf, "Int", size)
    return StrGet(buf)
}

RemoveLinesByPrescriptionType(text) {
    lines := StrSplit(text, "`n", "`r")
    if (lines.Length == 0) {
        return text
    }
    
    if (SubStr(lines[1], 1, 2) == "--") {
        result := ""
        for line in lines {
            if (RegExMatch(line, "^(--|<R|処方箋)") || line == "") {
                continue
            }
            result .= line "`n"
        }
        return result
    } else if (RegExMatch(lines[1], "^処方日")) {
        text := CombineInpatientOrderLines(text)
        return RegExReplace(text, "m)^処方日.*(\R|$)", "")
    }
    return text
}

CombineInpatientOrderLines(text) {
    sections := []
    currentSection := ""
    lines := StrSplit(text, "`n", "`r")
    
    for line in lines {
        if (RegExMatch(line, "^処方日")) {
            if (currentSection != "") {
                sections.Push(currentSection)
            }
            currentSection := line "`n"
        } else {
            currentSection .= line "`n"
        }
    }
    sections.Push(currentSection)
    
    finalResult := ""
    for sec in sections {
        secLines := StrSplit(sec, "`n", "`r")
        combinedSec := ""
        buffer := ""
        for sLine in secLines {
            if (sLine == "") {
                continue
            }
            if (RegExMatch(sLine, "^処方日")) {
                combinedSec .= sLine "`n"
                continue
            }
            if (InStr(sLine, " ")) { 
                combinedSec .= buffer . sLine "`n"
                buffer := ""
            } else { 
                buffer .= sLine
            }
        }
        finalResult .= combinedSec . buffer "`n"
    }
    return finalResult
}

InsertSpaceBeforeUnitAndReplace(text) {
    ; ここでも用法行の削除（行末が「分」）を補強
    text := RegExReplace(text, "m)^.*\d+\S+分$(\R|$)", "")
    text := RegExReplace(text, "m)(\d+\S+[錠pg枚ﾄ]$)", "@@SPACE@@$1")
    text := RegExReplace(text, "icap$", "c")
    return text
}

CombineUsageLinesAndRemoveSpecificWords(text) {
    lines := StrSplit(text, "`n", "`r")
    newLines := []
    for line in lines {
        if (line == "") {
            continue
        }
        if (RegExMatch(line, "^\S+時") && newLines.Length > 0) {
            newLines[newLines.Length] .= line
        } else {
            newLines.Push(line)
        }
    }
    
    res := ""
    for line in newLines {
        line := RegExReplace(line, "^分(\d+)\s(\d)", "分$1@@SPACE@@$2")
        if (RegExMatch(line, "^外\)")) {
            line := StrReplace(line, "外)", "@@SPACE@@")
            res := RegExReplace(res, "\R$", "") . line 
            continue
        }
        line := StrReplace(line, "吸入用", "")
        res .= line "`n"
    }
    return res
}

AbbreviateUsagePatternB(line) {
    line := RegExReplace(line, "毎(?=.)", "")
    line := StrReplace(line, "食後", "")
    line := RegExReplace(line, "(と)?眠前", "寝")
    line := StrReplace(line, "食前", "前")
    line := StrReplace(line, "[食間]", "")
    line := RegExReplace(line, "1日\d回", "")
    return line
}

AbbreviateUsagePatternC(line) {
    line := RegExReplace(line, "(と)?眠前", "寝")
    line := StrReplace(line, "[食間]", "")
    line := RegExReplace(line, "1日\d回", "")
    return line
}

ShowNotification(msg) {
    ToolTip(msg)
    SetTimer(() => ToolTip(), -1500)
}
