; ==============================================================================
; File: PrescriptionFormatter_v3.5.ahk
; Version: 3.5
; Description: 処方整形スクリプト (AHK v2) - 頓服用法(発熱時 等)の抽出条件修正
; ==============================================================================

#Requires AutoHotkey v2.0

; ------------------------------------------------------------------------------
; Win + Alt + S : 用法なし出力
; ------------------------------------------------------------------------------
#!s:: {
    text := ProcessInitialInput()
    
    if (RegExMatch(text, "^商品名")) {
        text := RegExReplace(text, "^商品名\s*", "")
        text := FinalizeText(text)
        A_Clipboard := text
        ToolTip("整形完了(用法錠数なし)")
    } else {
        if (RegExMatch(text, "^処方日")) {
            text := ReorganizeByTrigger(text)
        }
        
        lines := StrSplit(text, "`n", "`r")
        result := ""
        for line in lines {
            if (RegExMatch(line, "\d+\S*[錠pg枚ﾄ]$")) {
                line := RegExReplace(line, "(\d+\S*[錠pg枚ﾄ]$)", "@@SPACE@@$1")
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

; ------------------------------------------------------------------------------
; Win + Alt + D : 用法あり出力
; ------------------------------------------------------------------------------
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
    
    loop lines.Length {
        line := lines[A_Index]
        if (RegExMatch(line, "^分[123]\S+")) {
            line := RegExReplace(line, "毎(?=.)|食後", "")
            line := RegExReplace(line, "(?:と)?眠前", "寝")
            line := RegExReplace(line, "食前", "前")
            line := RegExReplace(line, "\[食間\]", "")
            line := RegExReplace(line, "1日\d回", "")
            if (processedLines.Length > 0)
                processedLines[processedLines.Length] .= line
            else
                processedLines.Push(line)
        }
        else if (RegExMatch(line, "^分\d\S+")) {
            line := RegExReplace(line, "(?:と)?眠前", "寝")
            line := RegExReplace(line, "\[食間\]", "")
            line := RegExReplace(line, "1日\d回", "")
            if (processedLines.Length > 0)
                processedLines[processedLines.Length] .= line
            else
                processedLines.Push(line)
        } else {
            if (line != "")
                processedLines.Push(line)
        }
    }
    
    text := ""
    for line in processedLines
        text .= line "`n"
        
    text := FinalizeText(text)
    A_Clipboard := text
    ToolTip("整形完了(用法あり)")
    SetTimer(() => ToolTip(), -2000)
}

; ------------------------------------------------------------------------------
; 各機能関数
; ------------------------------------------------------------------------------

FinalizeText(text) {
    text := StrReplace(text, "@@SPACE@@", " ")
    text := RegExReplace(text, "\(\Sとして\)", "")
    text := StrReplace(text, "(非持参", "")
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
    if !ClipWait(0.5) {
        A_Clipboard := savedClip
    }
    return ConvertToHalfWidth(A_Clipboard)
}

ReorganizeByTrigger(text) {
    lines := StrSplit(text, "`n", "`r")
    newOutput := ""
    buffer := ""
    
    for line in lines {
        if (RegExMatch(line, "^処方日") || line == "")
            continue
        
        if (RegExMatch(line, "\d+\S*[錠pg枚ﾄ]$")) {
            newOutput .= buffer . line . "`n"
            buffer := ""
        } else {
            if (!InStr(line, " "))
                buffer .= line
            else
                newOutput .= line . "`n"
        }
    }
    return newOutput . buffer
}

ApplyBasicFormatting(text) {
    text := RegExReplace(text, "m)(*ANYCRLF)\d+\S*分$", "")
    text := RegExReplace(text, "m)(*ANYCRLF)(\d+\S*[錠pg枚ﾄ]$)", "@@SPACE@@$1")
    text := RegExReplace(text, "m)(*ANYCRLF)cap$", "c")
    return text
}

MergeSpecificPatterns(text) {
    lines := StrSplit(text, "`n", "`r")
    result := []
    
    for line in lines {
        if (line == "")
            continue
        
        ; 修正：^\S+時\s に変更。時＋空白にマッチさせ、残りを分離
        if (RegExMatch(line, "^(\S+時)\s+(.*)$", &m)) {
            ; 1. 「発熱時」の部分を上の行に結合
            if (result.Length > 0) {
                result[result.Length] .= m[1]
            } else {
                result.Push(m[1])
            }
            ; 2. 空白の後の残り（例：「1日3回まで」）を新しい行として追加
            remaining := Trim(m[2])
            if (remaining != "") {
                result.Push(remaining)
            }
            
        } else if (RegExMatch(line, "^分\d+\s\d")) {
            line := RegExReplace(line, "^(分\d+)\s(\d)", "$1@@SPACE@@$2")
            result.Push(line)
            
        } else if (RegExMatch(line, "^外\)\s(.*)$", &m)) {
            if (result.Length > 0) {
                result[result.Length] .= "@@SPACE@@" . m[1]
            } else {
                result.Push("@@SPACE@@" . m[1])
            }
            
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

ConvertToHalfWidth(str) {
    size := DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", 0, "Int", 0)
    buf := Buffer(size * 2)
    DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", buf, "Int", size)
    return StrGet(buf, "UTF-16")
}; ==============================================================================
; File: PrescriptionFormatter_v3.4.ahk
; Version: 3.4
; Description: 処方整形スクリプト (AHK v2) - 頓服用法(発熱時等)の結合対応版
; ==============================================================================

#Requires AutoHotkey v2.0

; ------------------------------------------------------------------------------
; Win + Alt + S : 用法なし出力
; ------------------------------------------------------------------------------
#!s:: {
    text := ProcessInitialInput()
    
    if (RegExMatch(text, "^商品名")) {
        text := RegExReplace(text, "^商品名\s*", "")
        text := FinalizeText(text)
        A_Clipboard := text
        ToolTip("整形完了(用法錠数なし)")
    } else {
        if (RegExMatch(text, "^処方日")) {
            text := ReorganizeByTrigger(text)
        }
        
        lines := StrSplit(text, "`n", "`r")
        result := ""
        for line in lines {
            if (RegExMatch(line, "\d+\S*[錠pg枚ﾄ]$")) {
                line := RegExReplace(line, "(\d+\S*[錠pg枚ﾄ]$)", "@@SPACE@@$1")
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

; ------------------------------------------------------------------------------
; Win + Alt + D : 用法あり出力
; ------------------------------------------------------------------------------
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
    
    loop lines.Length {
        line := lines[A_Index]
        if (RegExMatch(line, "^分[123]\S+")) {
            line := RegExReplace(line, "毎(?=.)|食後", "")
            line := RegExReplace(line, "(?:と)?眠前", "寝")
            line := RegExReplace(line, "食前", "前")
            line := RegExReplace(line, "\[食間\]", "")
            line := RegExReplace(line, "1日\d回", "")
            if (processedLines.Length > 0)
                processedLines[processedLines.Length] .= line
            else
                processedLines.Push(line)
        }
        else if (RegExMatch(line, "^分\d\S+")) {
            line := RegExReplace(line, "(?:と)?眠前", "寝")
            line := RegExReplace(line, "\[食間\]", "")
            line := RegExReplace(line, "1日\d回", "")
            if (processedLines.Length > 0)
                processedLines[processedLines.Length] .= line
            else
                processedLines.Push(line)
        } else {
            if (line != "")
                processedLines.Push(line)
        }
    }
    
    text := ""
    for line in processedLines
        text .= line "`n"
        
    text := FinalizeText(text)
    A_Clipboard := text
    ToolTip("整形完了(用法あり)")
    SetTimer(() => ToolTip(), -2000)
}

; ------------------------------------------------------------------------------
; 各機能関数
; ------------------------------------------------------------------------------

FinalizeText(text) {
    text := StrReplace(text, "@@SPACE@@", " ")
    text := RegExReplace(text, "\(\Sとして\)", "")
    text := StrReplace(text, "(非持参", "")
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
    if !ClipWait(0.5) {
        A_Clipboard := savedClip
    }
    return ConvertToHalfWidth(A_Clipboard)
}

ReorganizeByTrigger(text) {
    lines := StrSplit(text, "`n", "`r")
    newOutput := ""
    buffer := ""
    
    for line in lines {
        if (RegExMatch(line, "^処方日") || line == "")
            continue
        
        if (RegExMatch(line, "\d+\S*[錠pg枚ﾄ]$")) {
            newOutput .= buffer . line . "`n"
            buffer := ""
        } else {
            if (!InStr(line, " "))
                buffer .= line
            else
                newOutput .= line . "`n"
        }
    }
    return newOutput . buffer
}

ApplyBasicFormatting(text) {
    text := RegExReplace(text, "m)(*ANYCRLF)\d+\S*分$", "")
    text := RegExReplace(text, "m)(*ANYCRLF)(\d+\S*[錠pg枚ﾄ]$)", "@@SPACE@@$1")
    text := RegExReplace(text, "m)(*ANYCRLF)cap$", "c")
    return text
}

MergeSpecificPatterns(text) {
    lines := StrSplit(text, "`n", "`r")
    result := []
    
    for line in lines {
        if (line == "")
            continue
        
        ; 「発熱時」などにマッチする部分を抽出
        if (RegExMatch(line, "^(\S+時)(.*)$", &m)) {
            ; 1. 「発熱時」の部分を上の行に結合
            if (result.Length > 0) {
                result[result.Length] .= m[1]
            } else {
                result.Push(m[1])
            }
            ; 2. その行に続き（例：「1日3回まで」）があれば新しい行として追加
            remaining := Trim(m[2])
            if (remaining != "") {
                result.Push(remaining)
            }
            
        } else if (RegExMatch(line, "^分\d+\s\d")) {
            line := RegExReplace(line, "^(分\d+)\s(\d)", "$1@@SPACE@@$2")
            result.Push(line)
            
        } else if (RegExMatch(line, "^外\)\s(.*)$", &m)) {
            ; 外\) を @@SPACE@@ に置換して上の行に結合
            if (result.Length > 0) {
                result[result.Length] .= "@@SPACE@@" . m[1]
            } else {
                result.Push("@@SPACE@@" . m[1])
            }
            
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

ConvertToHalfWidth(str) {
    size := DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", 0, "Int", 0)
    buf := Buffer(size * 2)
    DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", buf, "Int", size)
    return StrGet(buf, "UTF-16")
}
