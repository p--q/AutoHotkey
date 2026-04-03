; ==============================================================================
; File: PrescriptionFormatter_v6.3.ahk
; Version: 6.3
; Description: 処方整形 (AHK v2) - 「外)」行のトリガー誤認防止ロジック実装
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

        ; 用法判定 (分n, 1日n回, 1日n枚)
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
        
    A_Clipboard := FinalizeText(text)
    ToolTip("整形完了(用法あり)")
    SetTimer(() => ToolTip(), -2000)
}

; --- 関数群 ---

ApplyBasicFormatting(text) {
    ; 吸入用を先に削除
    text := StrReplace(text, "吸入用", "")
    ; 用法末尾の「分/日分」を削除
    text := RegExReplace(text, "m)(*ANYCRLF)\d+\S*分$", "")

    ; 単位のマーキング (否定先読みで「外)」を含む行を除
