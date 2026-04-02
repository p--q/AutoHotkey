/*
 * File: PrescriptionFormatter.ahk
 * Version: 1.2
 * Description: 処方オーダーおよび薬品DI情報のテキストを整形し、
 * 特定の規則に基づいた用法省略記号への置換や不要行の削除を行います。
 */

#Requires AutoHotkey v2.0
#SingleInstance Force

; ------------------------------------------------------------------
; Main Hotkey: Win + Alt + S (Output without Usage)
; ------------------------------------------------------------------
#!s:: {
    text := GetTextAndConvertToHalfWidth()
    
    ; Determine if it's DI information
    if (RegExMatch(text, "m)^商品名")) {
        ; DI Info: Remove "商品名" and following spaces
        text := RegExReplace(text, "m)^商品名\s*", "")
    } else {
        ; Standard Process
        text := RegExReplace(text, "m)^分[123]\S+", "")
        text := RegExReplace(text, "m)^外\)", "")
        text := RegExReplace(text, "m)^吸入用", "")
        text := RegExReplace(text, "m)\d+\S+分$", "")
        
        text := RemoveLinesByPrescriptionType(text)
        text := InsertSpaceBeforeUnitAndReplace(text)
        
        ; Remove all whitespace (including newlines)
        text := RegExReplace(text, "\s+", "")
        ; Replace placeholder with a half-width space
        text := StrReplace(text, "@@SPACE@@", " ")
    }
    
    text := RegExReplace(text, "\(\Sとして\)", "")
    A_Clipboard := text
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
    
    ; Remove all whitespace
    text := RegExReplace(text, "\s+", "")
    
    ; Process line by line for usage abbreviation
    lines := StrSplit(text, "`n", "`r")
    processedText := ""
    for line in lines {
        if (line == "") continue
        
        if (RegExMatch(line, "^分[123]\S+")) {
            processedText .= AbbreviateUsagePatternB(line)
        } else if (RegExMatch(line, "^分\d\S+")) {
            processedText .= AbbreviateUsagePatternC(line)
        } else {
            processedText .= line
        }
    }
    
    text := StrReplace(processedText, "@@SPACE@@", " ")
    text := RegExReplace(text, "\(\Sとして\)", "")
    
    A_Clipboard := text
    ShowNotification("Formatted (With Usage)")
}

; ------------------------------------------------------------------
; Sub-functions
; ------------------------------------------------------------------

; Corresponds to FuncE: Copy and Convert to Half-width
GetTextAndConvertToHalfWidth() {
    oldClip := A_Clipboard
    A_Clipboard := ""
    Send "^c"
    if !ClipWait(0.5) {
        A_Clipboard := oldClip
    }
    return ConvertFullToHalf(A_Clipboard)
}

; Convert Katakana, Numbers, and Alphabets using Windows API
ConvertFullToHalf(str) {
    if (str == "") return ""
    ; LCMAP_HALFWIDTH = 0x00400000, LOCALE_USER_DEFAULT = 0x400
    size := DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", 0, "Int", 0)
    buf := Buffer(size * 2)
    DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", buf, "Int", size)
    return StrGet(buf)
}

; Corresponds to FuncG: Filter lines based on order type
RemoveLinesByPrescriptionType(text) {
    lines := StrSplit(text, "`n", "`r")
    if (lines.Length == 0) return text
    
    if (SubStr(lines[1], 1, 2) == "--") {
        ; Outpatient Order
        result := ""
        for line in lines {
            if (RegExMatch(line, "^(--|<R|処方箋期限)") || line == "")
                continue
            result .= line "`n"
        }
        return result
    } else if (RegExMatch(lines[1], "^処方日")) {
        ; Inpatient Order
        text := CombineInpatientOrderLines(text)
        return RegExReplace(text, "m)^処方日.*(\R|$)", "")
    }
    return text
}

; Corresponds to FuncD: Specific combination logic for Inpatient Orders
CombineInpatientOrderLines(text) {
    sections := []
    currentSection := ""
    lines := StrSplit(text, "`n", "`r")
    
    for line in lines {
        if (RegExMatch(line, "^処方日")) {
            if (currentSection != "") sections.Push(currentSection)
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
            if (sLine == "") continue
            if (RegExMatch(sLine, "^処方日")) {
                combinedSec .= sLine "`n"
                continue
            }
            if (InStr(sLine, " ")) { ; Trigger line (contains space)
                combinedSec .= buffer . sLine "`n"
                buffer := ""
            } else { ; Continuous lines without space
                buffer .= sLine
            }
        }
        finalResult .= combinedSec . buffer "`n"
    }
    return finalResult
}

; Corresponds to FuncA
InsertSpaceBeforeUnitAndReplace(text) {
    text := RegExReplace(text, "m)\d+\S+分$", "")
    ; Insert placeholder before units
    text := RegExReplace(text, "m)(\d+\S+[錠pg枚ﾄ]$)", "@@SPACE@@$1")
    text := RegExReplace(text, "icap$", "c")
    return text
}

; Corresponds to FuncF
CombineUsageLinesAndRemoveSpecificWords(text) {
    lines := StrSplit(text, "`n", "`r")
    newLines := []
    for line in lines {
        if (line == "") continue
        ; Combine lines starting with "time" to the previous line
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
            res := RegExReplace(res, "\R$", "") . line ; Combine to previous
            continue
        }
        line := StrReplace(line, "吸入用", "")
        res .= line "`n"
    }
    return res
}

; Corresponds to FuncB
AbbreviateUsagePatternB(line) {
    line := RegExReplace(line, "毎(?=.)", "")
    line := StrReplace(line, "食後", "")
    line := RegExReplace(line, "(と)?眠前", "寝")
    line := StrReplace(line, "食前", "前")
    line := StrReplace(line, "[食間]", "")
    line := RegExReplace(line, "1日\d回", "")
    return line
}

; Corresponds to FuncC
AbbreviateUsagePatternC(line) {
    line := RegExReplace(line, "(と)?眠前", "寝")
    line := StrReplace(line, "[食間]", "")
    line := RegExReplace(line, "1日\d回", "")
    return line
}

ShowNotification(msg) {
    ToolTip msg
    SetTimer () => ToolTip(), -1500
}
