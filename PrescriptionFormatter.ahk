; ==============================================================================
; File: PrescriptionFormatter.ahk
; Version: 1.1
; Description: 処方オーダーのテキストを整形し、DI情報や用法を整理して
;              クリップボードに再格納するスクリプト。
; ==============================================================================

#Requires AutoHotkey v2.0

; ------------------------------------------------------------------------------
; Win + Alt + S : 出力に用法がない（DI情報または薬品名のみの抽出）
; ------------------------------------------------------------------------------
#!s:: {
    text := ProcessInitialInput() ; 関数E
    
    if (RegExMatch(text, "^商品名")) {
        ; DI情報の処理
        text := RegExReplace(text, "^商品名\s*", "")
    } else {
        ; 薬品名抽出処理
        text := ReorganizePrescriptionLines(text) ; 関数D
        
        lines := StrSplit(text, "`n", "`r")
        result := ""
        for line in lines {
            if (RegExMatch(line, "\d+\S*[錠pg枚ﾄ]$")) {
                line := RegExReplace(line, "(\d+\S*[錠pg枚ﾄ]$)", "@@SPACE@@$1")
                line := RegExReplace(line, "cap$", "c")
                result .= line "`n"
            }
        }
        text := result
        text := RegExReplace(text, "[ \t]+", "") ; 改行以外の空白削除
    }
    
    FinalizeClipboard(text)
}

; ------------------------------------------------------------------------------
; Win + Alt + D : 出力に用法がある
; ------------------------------------------------------------------------------
#!d:: {
    text := ProcessInitialInput() ; 関数E
    text := FilterOrderType(text)  ; 関数G
    text := ApplyPrescriptionFormatting(text) ; 関数A
    text := MergeSpecificPatterns(text) ; 関数F
    
    text := RegExReplace(text, "[ \t]+", "") ; 改行以外の空白削除
    
    ; 用法置換処理 (関数B / 関数C 相当)
    lines := StrSplit(text, "`n", "`r")
    processedLines := []
    
    loop lines.Length {
        line := lines[A_Index]
        if (RegExMatch(line, "^分[123]\S+")) {
            ; 関数Bの処理
            line := RegExReplace(line, "毎(?=.)|食後", "")
            line := RegExReplace(line, "(?:と)?眠前", "寝")
            line := RegExReplace(line, "食前", "前")
            line := RegExReplace(line, "\[食間\]", "")
            line := RegExReplace(line, "1日\d回", "")
            ; 前の行と結合
            if (processedLines.Length > 0)
                processedLines[processedLines.Length] .= line
            else
                processedLines.Push(line)
        }
        else if (RegExMatch(line, "^分\d\S+")) {
            ; 関数Cの処理
            line := RegExReplace(line, "(?:と)?眠前", "寝")
            line := RegExReplace(line, "\[食間\]", "")
            line := RegExReplace(line, "1日\d回", "")
            ; 前の行と結合
            if (processedLines.Length > 0)
                processedLines[processedLines.Length] .= line
            else
                processedLines.Push(line)
        } else {
            processedLines.Push(line)
        }
    }
    
    ; 配列を文字列に戻す
    text := ""
    for line in processedLines
        text .= line "`n"
        
    FinalizeClipboard(text)
}

; ------------------------------------------------------------------------------
; 各機能関数
; ------------------------------------------------------------------------------

; 関数E: 選択範囲の取得と全角半角変換
ProcessInitialInput() {
    A_Clipboard := ""
    Send("^c")
    if !ClipWait(0.5) {
        ; 選択範囲がなければ現在のクリップボードを使用
    }
    input := A_Clipboard
    return StrConvert(input, "H") 
}

; 関数G: オーダー種別の判別とフィルタリング
FilterOrderType(text) {
    if (SubStr(text, 1, 2) == "--") {
        ; 外来処方
        lines := StrSplit(text, "`n", "`r")
        result := ""
        for line in lines {
            if (line == "" || RegExMatch(line, "^(--|<R|処方箋)"))
                continue
            result .= line "`n"
        }
        return result
    } else if (RegExMatch(text, "^処方日")) {
        ; 入院処方
        text := ReorganizePrescriptionLines(text) ; 関数D
        text := RegExReplace(text, "m)^処方日.*`n?", "")
        return text
    }
    return text
}

; 関数D: 処方行の再構成
ReorganizePrescriptionLines(text) {
    lines := StrSplit(text, "`n", "`r")
    newOutput := ""
    buffer := ""
    
    for line in lines {
        if (line == "")
            continue
        
        if (InStr(line, " ")) { ; トリガー行（スペースあり）
            newOutput .= buffer . line . "`n"
            buffer := ""
        } else {
            buffer .= line
        }
    }
    return newOutput . buffer
}

; 関数A: 基本的な薬品単位の整形
ApplyPrescriptionFormatting(text) {
    text := RegExReplace(text, "\d+\S*分$", "")
    text := RegExReplace(text, "(\d+\S*[錠pg枚ﾄ]$)", "@@SPACE@@$1")
    text := RegExReplace(text, "cap$", "c")
    return text
}

; 関数F: 行の結合と特定パターンの置換
MergeSpecificPatterns(text) {
    lines := StrSplit(text, "`n", "`r")
    result := []
    
    for line in lines {
        if (line == "")
            continue
        
        if (RegExMatch(line, "^\S+時")) {
            if (result.Length > 0)
                result[result.Length] .= line
            else
                result.Push(line)
        } else if (RegExMatch(line, "^分\d+\s\d", &m)) {
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

; 共通の最終処理: 特殊文字の置換とクリーンアップ
FinalizeClipboard(text) {
    text := StrReplace(text, "@@SPACE@@", " ")
    text := RegExReplace(text, "\(\Sとして\)", "")
    text := Trim(text, "`n`r ")
    A_Clipboard := text
    ToolTip("整形完了")
    SetTimer(() => ToolTip(), -2000)
}

; 全角・半角変換補助関数
StrConvert(str, mode) {
    static diff := 0xFEE0
    result := ""
    loop parse, str {
        charAlt := Ord(A_LoopField)
        if (mode = "H") { ; 全角から半角
            if (charAlt >= 0xFF01 && charAlt <= 0xFF5E)
                result .= Chr(charAlt - diff)
            else if (charAlt = 0x3000) ; 全角スペース
                result .= Chr(0x0020)
            else
                result .= A_LoopField
        }
    }
    return result
}
