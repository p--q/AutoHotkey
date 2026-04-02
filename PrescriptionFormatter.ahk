; ==============================================================================
; File: PrescriptionFormatter_v3.ahk
; Version: 3.0
; Description: 処方整形スクリプト (AHK v2)
; ==============================================================================

#Requires AutoHotkey v2.0

; ------------------------------------------------------------------------------
; Win + Alt + S : 用法なし出力
; ------------------------------------------------------------------------------
#!s:: {
    text := ProcessInitialInput() ; 関数E
    
    if (RegExMatch(text, "^商品名")) {
        ; DI情報の処理
        text := RegExReplace(text, "^商品名\s*", "")
        text := FinalizeText(text) ; 関数H相当
        A_Clipboard := text
        ToolTip("整形完了(用法錠数なし)")
    } else {
        ; 入院処方オーダー判定
        if (RegExMatch(text, "^処方日")) {
            text := ReorganizeByTrigger(text) ; 関数D
        }
        
        ; 特定の単位行以外を削除し、整形
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
        text := FinalizeText(text) ; 関数H相当
        A_Clipboard := text
        ToolTip("整形完了(用法なし)")
    }
    SetTimer(() => ToolTip(), -2000)
}

; ------------------------------------------------------------------------------
; Win + Alt + D : 用法あり出力
; ------------------------------------------------------------------------------
#!d:: {
    text := ProcessInitialInput() ; 関数E
    
    ; 外来・入院の判定と前処理
    if (SubStr(text, 1, 2) == "--") {
        text := FilterOutpatientOrder(text) ; 関数G
    } else if (RegExMatch(text, "^処方日")) {
        text := ReorganizeByTrigger(text) ; 関数D
    }
    
    text := ApplyBasicFormatting(text) ; 関数A
    text := MergeSpecificPatterns(text) ; 関数F
    
    text := RegExReplace(text, "[ \t]+", "") ; 改行以外の空白削除
    
    ; 用法置換処理 (関数B / 関数C)
    lines := StrSplit(text, "`n", "`r")
    processedLines := []
    
    loop lines.Length {
        line := lines[A_Index]
        if (RegExMatch(line, "^分[123]\S+")) {
            ; 関数B
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
            ; 関数C
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
        
    text := FinalizeText(text) ; 関数H相当
    A_Clipboard := text
    ToolTip("整形完了(用法あり)")
    SetTimer(() => ToolTip(), -2000)
}

; ------------------------------------------------------------------------------
; 各機能関数
; ------------------------------------------------------------------------------

; 関数H: 特殊記号の復元とクリーンアップ
FinalizeText(text) {
    text := StrReplace(text, "@@SPACE@@", " ")
    text := RegExReplace(text, "\(\Sとして\)", "")
    text := StrReplace(text, "(非持参", "")
    return Trim(text, "`n`r")
}

; 関数G: 外来処方オーダーのフィルタリング
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

; 関数E: 取得と全角半角変換
ProcessInitialInput() {
    savedClip := A_Clipboard
    A_Clipboard := ""
    Send("^c")
    if !ClipWait(0.5) {
        A_Clipboard := savedClip
    }
    return ConvertToHalfWidth(A_Clipboard)
}

; 関数D: 処方行の再構成（トリガー行による結合）
ReorganizeByTrigger(text) {
    lines := StrSplit(text, "`n", "`r")
    newOutput := ""
    buffer := ""
    
    for line in lines {
        if (RegExMatch(line, "^処方日") || line == "")
            continue
        
        ; 単位で終わる行がトリガー
        if (RegExMatch(line, "\d+\S*[錠pg枚ﾄ]$")) {
            newOutput .= buffer . line . "`n"
            buffer := ""
        } else {
            ; スペースを含まない行は薬品名の断片として溜める
            if (!InStr(line, " "))
                buffer .= line
            else
                newOutput .= line . "`n"
        }
    }
    return newOutput . buffer
}

; 関数A: 基本単位整形
ApplyBasicFormatting(text) {
    text := RegExReplace(text, "\d+\S*分$", "")
    text := RegExReplace(text, "(\d+\S*[錠pg枚ﾄ]$)", "@@SPACE@@$1")
    text := RegExReplace(text, "cap$", "c")
    return text
}

; 関数F: 特定パターンの行結合
MergeSpecificPatterns(text) {
    lines := StrSplit(text, "`n", "`r")
    result := []
    
    for line in lines {
        if (line == "") continue
        
        if (RegExMatch(line, "^\S+時")) {
            if (result.Length > 0)
                result[result.Length] .= line
            else
                result.Push(line)
        } else if (RegExMatch(line, "^分\d+\s\d")) {
            line := RegExReplace(line, "^(分\d+)\s(\d)", "$1@@SPACE@@$2")
            result.Push(line)
        } else if (RegExMatch(line, "^外\)\s")) {
            line := RegExReplace(line, "^外\)\s", "@@SPACE@@")
            if (result.Length > 0)
                result[result.Length] .= line
            else
                result.Push(line)
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

; Windows APIを使用した全角→半角変換（カタカナ・記号・英数）
ConvertToHalfWidth(str) {
    ; LCMAP_HALFWIDTH = 0x00400000
    size := DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", 0, "Int", 0)
    buf := Buffer(size * 2)
    DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", buf, "Int", size)
    return StrGet(buf, "UTF-16")
}
