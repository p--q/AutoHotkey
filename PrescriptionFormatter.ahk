; ==============================================================================
; File: PrescriptionFormatter_v2.ahk
; Version: 2.0
; Description: 処方オーダーおよびDI情報の整形スクリプト。
;              Win+Alt+S (用法なし), Win+Alt+D (用法あり)
; ==============================================================================

#Requires AutoHotkey v2.0

; ------------------------------------------------------------------------------
; Win + Alt + S : 出力に用法がない
; ------------------------------------------------------------------------------
#!s:: {
    text := ProcessInitialInput() ; 関数E
    
    if (RegExMatch(text, "^商品名")) {
        ; DI情報の処理
        text := RegExReplace(text, "^商品名\s*", "")
        FinalizeClipboard(text, "整形完了(用法錠数なし)") ; 関数H
    } else {
        ; 薬品名抽出処理
        text := ReorganizeByTrigger(text) ; 関数D
        
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
        FinalizeClipboard(text, "整形完了(用法なし)") ; 関数H
    }
}

; ------------------------------------------------------------------------------
; Win + Alt + D : 出力に用法がある
; ------------------------------------------------------------------------------
#!d:: {
    text := ProcessInitialInput() ; 関数E
    text := FilterOrderType(text)  ; 関数G
    text := ApplyBasicFormatting(text) ; 関数A
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
        
    FinalizeClipboard(text, "整形完了(用法あり)") ; 関数H
}

; ------------------------------------------------------------------------------
; 各機能関数
; ------------------------------------------------------------------------------

; 関数H: 最終整形とクリップボード格納
FinalizeClipboard(text, message) {
    text := StrReplace(text, "@@SPACE@@", " ")
    text := RegExReplace(text, "\(\Sとして\)", "")
    text := StrReplace(text, "(非持参", "")
    text := Trim(text, "`n`r ")
    A_Clipboard := text
    ToolTip(message)
    SetTimer(() => ToolTip(), -2000)
}

; 関数G: オーダー種別判別
FilterOrderType(text) {
    if (SubStr(text, 1, 2) == "--") {
        ; 外来
        lines := StrSplit(text, "`n", "`r")
        result := ""
        for line in lines {
            if (line == "" || RegExMatch(line, "^(--|<R|処方箋)"))
                continue
            result .= line "`n"
        }
        return result
    } else if (RegExMatch(text, "^処方日")) {
        ; 入院
        return ReorganizeByTrigger(text) ; 関数D呼び出し
    }
    return text
}

; 関数E: 取得と全角半角変換（カタカナ・括弧含む）
ProcessInitialInput() {
    savedClipboard := A_Clipboard
    A_Clipboard := ""
    Send("^c")
    if !ClipWait(0.5) {
        A_Clipboard := savedClipboard
    }
    input := A_Clipboard
    ; StrConvertの"HK"モードでカタカナも半角化
    return StrConvert(input, "HK") 
}

; 関数D: トリガー行による結合（処方日は破棄）
ReorganizeByTrigger(text) {
    lines := StrSplit(text, "`n", "`r")
    newOutput := ""
    buffer := ""
    
    for line in lines {
        if (RegExMatch(line, "^処方日") || line == "")
            continue
        
        ; 「数字...単位」にマッチする行がトリガー
        if (RegExMatch(line, "\d+\S*[錠pg枚ﾄ]$")) {
            newOutput .= buffer . line . "`n"
            buffer := ""
        } else {
            ; スペースを含まない行（薬品名の断片など）をバッファ
            if (!InStr(line, " "))
                buffer .= line
            else
                newOutput .= line . "`n" ; スペースがあるがトリガーでない行はそのまま
        }
    }
    return newOutput . buffer
}

; 関数A: 基本整形
ApplyBasicFormatting(text) {
    text := RegExReplace(text, "\d+\S*分$", "")
    text := RegExReplace(text, "(\d+\S*[錠pg枚ﾄ]$)", "@@SPACE@@$1")
    text := RegExReplace(text, "cap$", "c")
    return text
}

; 関数F: 行結合とパターン置換
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

; 全角・半角変換補助（Win32 APIを使用）
StrConvert(str, mode) {
    ; モード H: 半角化(英数), K: カタカナ半角化
    ; LCMAP_HALFWIDTH = 0x00400000
    static LCMAP_HALFWIDTH := 0x00400000
    size := DllCall("MultiByteToWideChar", "UInt", 65001, "UInt", 0, "Str", str, "Int", -1, "Ptr", 0, "Int", 0)
    buf := Buffer(size * 2)
    DllCall("MultiByteToWideChar", "UInt", 65001, "UInt", 0, "Str", str, "Int", -1, "Ptr", buf, "Int", size)
    
    size2 := DllCall("LCMapStringW", "UInt", 0x400, "UInt", LCMAP_HALFWIDTH, "Ptr", buf, "Int", size, "Ptr", 0, "Int", 0)
    buf2 := Buffer(size2 * 2)
    DllCall("LCMapStringW", "UInt", 0x400, "UInt", LCMAP_HALFWIDTH, "Ptr", buf, "Int", size, "Ptr", buf2, "Int", size2)
    
    return StrGet(buf2, "UTF-16")
}
