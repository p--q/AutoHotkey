; ==============================================================================
; File: PrescriptionFormatter_v6.6.0.ahk
; Version: 6.6.0
; Description: 単位(錠p枚ﾄcg)の直前にマーカーを付与し、薬品名の泣き別れを強力に結合。
;              Win+Alt+S / Win+Alt+D の両方で、外用薬の数量除去置換を実行。
; ==============================================================================

#Requires AutoHotkey v2.0

; --- Win + Alt + S: 用法なし整形 ---
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
            ; 数量マーカーが含まれる行のみを抽出
            if (InStr(line, "@@SPACE@@")) {
                result .= line "`n"
            }
        }
        text := RegExReplace(result, "[ \t]+", "")
        text := FinalizeText(text)
        
        ; 【ご依頼の修正】FinalizeTextの後に外用薬の枚数置換を実行
        text := RegExReplace(text, "m)\s\d+枚\R外\)\s*(1日\d+枚)$", " $1")
        
        A_Clipboard := text
        ToolTip("整形完了(用法なし)")
    }
    SetTimer(() => ToolTip(), -2000)
}

; --- Win + Alt + D: 用法あり整形 ---
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
        
    text := FinalizeText(text)
    
    ; 【ご依頼の修正】FinalizeTextの後に外用薬の枚数置換を実行
    text := RegExReplace(text, "m)\s\d+枚\R外\)\s*(1日\d+枚)$", " $1")
    
    A_Clipboard := text
    ToolTip("整形完了(用法あり)")
    SetTimer(() => ToolTip(), -2000)
}

; --- 共通関数: 前処理 (トリガー定義) ---
ApplyBasicFormatting(text) {
    text := RegExReplace(text, "s)\s*\([^)]+として\)", "")
    text := RegExReplace(text, "m)\d+\S+分$", "")
    text := StrReplace(text, "吸入用", "")
    
    ; 数量マーカーの付与: 数字 + 単位(錠p枚ﾄcg) のパターンを見つけたら前に挿入
    ; 行末($)判定に頼らないことで、後ろに余計なスペースがあっても確実に補足する
    text := RegExReplace(text, "i)(\d+)\s*([錠p枚ﾄ]|cap|g)", "@@SPACE@@$1$2")
    
    ; cap を c に統一
    text := RegExReplace(text, "i)cap", "c")
    return text
}

; --- 共通関数: 処方日ブロック整理 ---
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
                ; バッファ(泣き別れた薬品名) + 数量行 を結合
                finalOutput .= buffer . line . (triggerCount > 1 ? "@@BLOCK@@" : "") . "`n"
                buffer := ""
            } else {
                ; スペースを含まない行は無条件でバッファへ貯める(薬品名の断片)
                if (!InStr(line, " ") && !InStr(line, "　")) {
                    buffer .= line
                } else {
                    finalOutput .= line . (triggerCount > 1 ? "@@BLOCK@@" : "") . "`n"
                }
            }
        }
        if (buffer != "")
            finalOutput .= buffer . (triggerCount > 1 ? "@@BLOCK@@" : "") . "`n"
    }
    return finalOutput
}

; --- 共通関数: 特殊用法結合 ---
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
    finalOutput := ""
    for l in result
        finalOutput .= l "`n"
    return finalOutput
}

; --- 共通関数: 仕上げ ---
FinalizeText(text) {
    text := StrReplace(text, "@@SPACE@@", " ")
    text := StrReplace(text, "@@BLOCK@@", "")
    text := RegExReplace(text, "\(\Sとして\)", "")
    text := StrReplace(text, "(非持参)", "")
    text := RegExReplace(text, " +", " ")
    return Trim(text, "`n`r")
}

; --- 共通関数: フィルタ ---
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

; --- 共通関数: 入力・半角化 ---
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
