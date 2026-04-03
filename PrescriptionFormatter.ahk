; ==============================================================================
; File: PrescriptionFormatter_v6.5.6.ahk
; Version: 6.5.6
; Description: トリガー行の定義を「錠p枚ﾄ$」または「\s\d+g$」に厳密化。
;              薬品名泣き別れを確実に防ぎ、外用薬の最終置換を完遂する。
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
        
        ; 外用薬の最終置換
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
    
    ; 外用薬の最終置換
    text := RegExReplace(text, "m)\s\d+枚\R外\)\s*(1日\d+枚)$", " $1")
    
    A_Clipboard := text
    ToolTip("整形完了(用法あり)")
    SetTimer(() => ToolTip(), -2000)
}
; --- 共通関数: 前処理 (トリガーマーカーの付与) ---
ApplyBasicFormatting(text) {
    text := RegExReplace(text, "s)\s*\([^)]+として\)", "")
    text := RegExReplace(text, "m)\d+\S+分$", "")
    text := StrReplace(text, "吸入用", "")
    ; トリガー定義: \d+\S*[錠p枚ﾄ]$ または \s\d+\S*g$ に @@SPACE@@ を付与
    text := RegExReplace(text, "m)(*ANYCRLF)(\d+\S*[錠p枚ﾄ]$|\s\d+\S*g$)", "@@SPACE@@$1")
    text := RegExReplace(text, "m)(*ANYCRLF)cap$", "c")
    return text
}

; --- 共通関数: 処方日トリガーによるブロック整理 ---
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
        ; このブロック内にトリガー行(@@SPACE@@)がいくつあるかカウント
        triggerCount := 0
        for l in blockLines {
            if (InStr(l, "@@SPACE@@"))
                triggerCount++
        }
        
        buffer := ""
        for line in blockLines {
            ; 判定: @@SPACE@@マーカーがある行が「数量行（トリガー）」
            if (InStr(line, "@@SPACE@@")) {
                outLine := buffer . line
                if (triggerCount > 1)
                    outLine .= "@@BLOCK@@"
                finalOutput .= outLine "`n"
                buffer := "" ; 結合したのでバッファを空にする
            } else {
                ; マーカーがない行は「薬品名の断片」としてバッファに溜める
                ; ただし、すでにスペースを含んでいる用法行などは除外
                if (!InStr(line, " ") && !InStr(line, "　")) {
                    buffer .= line
                } else {
                    if (triggerCount > 1)
                        line .= "@@BLOCK@@"
                    finalOutput .= line "`n"
                }
            }
        }
        ; 残ったバッファがあれば出力
        if (buffer != "")
            finalOutput .= buffer . (triggerCount > 1 ? "@@BLOCK@@" : "") . "`n"
    }
    return finalOutput
}

; 以下、MergeSpecificPatterns, FinalizeText, FilterOutpatientOrder, 
; ProcessInitialInput, ConvertToHalfWidth は v6.5.5 と同様のため省略（ロジックに変更なし）
