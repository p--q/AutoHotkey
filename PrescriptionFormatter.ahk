; ==============================================================================
; File: PrescriptionFormatter_v3.7.ahk
; Version: 3.7
; Description: 処方整形スクリプト (AHK v2) - 非強欲マッチ採用・最終版
; ==============================================================================

#Requires AutoHotkey v2.0

; ------------------------------------------------------------------------------
; Win + Alt + S : 用法なし出力
; ------------------------------------------------------------------------------
#!s:: {
    text := ProcessInitialInput() ; 関数E相当
    
    if (RegExMatch(text, "^商品名")) {
        ; DI情報の処理
        text := RegExReplace(text, "^商品名\s*", "")
        text := FinalizeText(text)
        A_Clipboard := text
        ToolTip("整形完了(用法錠数なし)")
    } else {
        ; 入院処方オーダー判定
        if (RegExMatch(text, "^処方日")) {
            text := ReorganizeByTrigger(text) ; 関数D相当
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
        text := RegExReplace(result, "[ \t]+", "") ; 改行以外の空白削除
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
    text := ProcessInitialInput() ; 関数E相当
    
    if (SubStr(text, 1, 2) == "--") {
        text := FilterOutpatientOrder(text) ; 関数G相当
    } else if (RegExMatch(text, "^処方日")) {
        text := ReorganizeByTrigger(text) ; 関数D相当
    }
    
    ; 基本整形と特殊パターンの結合（空白削除の前に行う）
    text := ApplyBasicFormatting(text) ; 関数A相当
    text := MergeSpecificPatterns(text) ; 関数F相当
    
    ; 改行を除くすべての空白を削除
    text := RegExReplace(text, "[ \t]+", "")
    
    lines := StrSplit(text, "`n", "`r")
    processedLines := []
    
    loop lines.Length {
        line := lines[A_Index]
        if (RegExMatch(line, "^分[123]\S+")) {
            ; 関数B相当
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
            ; 関数C相当
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
; サポート関数群
; ------------------------------------------------------------------------------

; 関数H相当: 特殊マーカーの復元と最終クリーンアップ
FinalizeText(text) {
    text := StrReplace(text, "@@SPACE@@", " ")
    text := RegExReplace(text, "\(\Sとして\)", "")
    text := StrReplace(text, "(非持参", "")
    return Trim(text, "`n`r")
}

; 関数G相当: 外来処方オーダーの不要行削除
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

; 関数E相当: 選択範囲のコピーと全角半角変換
ProcessInitialInput() {
    savedClip := A_Clipboard
    A_Clipboard := ""
    Send("^c")
    if !ClipWait(0.5) {
        A_Clipboard := savedClip
    }
    return ConvertToHalfWidth(A_Clipboard)
}

; 関数D相当: 入院処方の再構成（トリガー行による結合）
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

; 関数A相当: 基本単位整形（マルチラインオプション適用
