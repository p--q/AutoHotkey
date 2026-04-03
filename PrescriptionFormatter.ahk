; ==============================================================================
; File: PrescriptionFormatter_v6.5.2_Commented.ahk
; Version: 6.5.2
; Description: 処方整形 (AHK v2) - 安定版のロジックに解説コメントを付与
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
            ; 薬品名と数量の境界マーカー(@@SPACE@@)がある行のみを抽出
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
        if (line == "") {
            continue
        }

        ; 用法行（1日3回など）のパターンマッチと正規化
        if (RegExMatch(line, "^(分\d|1日\d回|1日\d枚)")) {
            line := RegExReplace(line, "毎(?=.)|食後", "")
            line := RegExReplace(line, "(?:と)?眠前", "寝")
            line := RegExReplace(line, "食前", "前")
            line := RegExReplace(line, "\[食間\]", "")
            line := StrReplace(line, "分1", "")
            
            isBlock := InStr(line, "@@BLOCK@@")
            line := StrReplace(line, "@@BLOCK@@", "")
            
            prevLine := (processedLines.Length > 0) ? processedLines[processedLines.Length] : ""
            
            ; 直前の行が「時」で終わっていない（＝薬品名）なら、用法をその行に結合
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

; --- 共通関数: 前処理 ---
ApplyBasicFormatting(text) {
    ; s)オプションで改行を含む「(〜として)」を削除。[^)]+は「)」以外の1文字以上。
    text := RegExReplace(text, "s)\s*\([^)]+として\)", "")
    
    ; m)オプションで各行末の「7日分」などを削除
    text := RegExReplace(text, "m)\d+\S+分$", "")
    
    text := StrReplace(text, "吸入用", "")

    ; 数量マーカー(@@SPACE@@)の付与
    ; 行末が「数字+単位」または「スペース+数字+g」の場合に目印を付ける
    text := RegExReplace(text, "m)(*ANYCRLF)(\d+\S*[錠p枚ﾄ]$|\s\d+\S*g$)", "@@SPACE@@$1")
    
    text := RegExReplace(text, "m)(*ANYCRLF)cap$", "c")
    return text
}
; --- 共通関数: 外用薬指示や特殊な用法行の結合 ---
MergeSpecificPatterns(text) {
    lines := StrSplit(text, "`n", "`r")
    result := []
    for line in lines {
        if (line == "") {
            continue
        }
        if (InStr(line, "@@BLOCK@@")) {
            result.Push(line)
            continue
        }
        ; 行末が「時」で終わる指示（就寝時など）を薬品名と結合
        if (RegExMatch(line, "^.+時\s*$")) {
            if (result.Length > 0 && !InStr(result[result.Length], "@@BLOCK@@"))
                result[result.Length] .= line
            else
                result.Push(line)
        } 
        ; 「外) 〜」形式の行を薬品名と結合
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

; --- 共通関数: 最終仕上げ ---
FinalizeText(text) {
    text := StrReplace(text, "@@SPACE@@", " ") ; マーカーを半角スペースに戻す
    text := RegExReplace(text, "\(\Sとして\)", "") ; (1として)等の微細な注釈削除
    text := StrReplace(text, "(非持参)", "")
    text := RegExReplace(text, " +", " ") ; 連続するスペースを1つに集約
    return Trim(text, "`n`r")
}

; --- 共通関数: 処方日トリガーによるブロック整理 ---
ReorganizeByTrigger(text) {
    blocks := []
    currentBlock := []
    lines := StrSplit(text, "`n", "`r")
    ; 処方日ごとに内容をグループ化
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
                if (triggerCount > 1) ; 複数薬品に1つの用法がかかる場合の目印
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

; --- 共通関数: 外来オーダ形式の不要行フィルタ ---
FilterOutpatientOrder(text) {
    lines := StrSplit(text, "`n", "`r")
    result := ""
    for line in lines {
        if (line == "" || RegExMatch(line, "^(--|<R|処方箋)")) {
            continue
        }
        result .= line "`n"
    }
    return result
}

; --- 共通関数: 入力処理 ---
ProcessInitialInput() {
    savedClip := A_Clipboard
    A_Clipboard := ""
    Send("^c")
    if !ClipWait(0.5)
        A_Clipboard := savedClip
    return ConvertToHalfWidth(A_Clipboard)
}

; --- 共通関数: 全角半角変換 (WinAPI利用) ---
ConvertToHalfWidth(str) {
    size := DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", 0, "Int", 0)
    buf := Buffer(size * 2)
    DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", buf, "Int", size)
    return StrGet(buf, "UTF-16")
}
