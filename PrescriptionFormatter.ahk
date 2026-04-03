; ==============================================================================
; File: PrescriptionFormatter_v6.5.2.ahk
; Version: 6.5.2 (Commented)
; Description: 処方整形 (AHK v2) - 内部処理の解説付き
; ==============================================================================

#Requires AutoHotkey v2.0

; ------------------------------------------------------------------------------
; ホットキー設定: Win + Alt + S (用法なし整形)
; ------------------------------------------------------------------------------
#!s:: {
    text := ProcessInitialInput() ; クリップボード取得と半角化
    if (RegExMatch(text, "^商品名")) {
        text := RegExReplace(text, "^商品名\s*", "")
        text := FinalizeText(text)
        A_Clipboard := text
        ToolTip("整形完了(用法錠数なし)")
    } else {
        text := ApplyBasicFormatting(text) ; 基本的なノイズ除去
        if (RegExMatch(text, "処方日"))
            text := ReorganizeByTrigger(text) ; 処方日を基準にブロック整理
        
        lines := StrSplit(text, "`n", "`r")
        result := ""
        for line in lines {
            ; @@SPACE@@ マーカーがある行（薬品名+数量の行）のみ抽出
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

; ------------------------------------------------------------------------------
; ホットキー設定: Win + Alt + D (用法あり整形)
; ------------------------------------------------------------------------------
#!d:: {
    text := ProcessInitialInput()
    if (SubStr(text, 1, 2) == "--") {
        text := FilterOutpatientOrder(text) ; 外来オーダ形式の不要行削除
    } 
    
    text := ApplyBasicFormatting(text)
    
    if (RegExMatch(text, "処方日"))
        text := ReorganizeByTrigger(text)
    
    text := MergeSpecificPatterns(text) ; 外用薬や「時」行の結合
    text := RegExReplace(text, "[ \t]+", "") ; 余計な空白の削除
    
    lines := StrSplit(text, "`n", "`r")
    processedLines := []
    
    for line in lines {
        if (line == "") continue

        ; 用法行（分1、1日3回など）の正規化
        if (RegExMatch(line, "^(分\d|1日\d回|1日\d枚)")) {
            line := RegExReplace(line, "毎(?=.)|食後", "") ; 「毎食後」→「食後」等
            line := RegExReplace(line, "(?:と)?眠前", "寝") ; 「就寝前」→「寝」
            line := RegExReplace(line, "食前", "前")
            line := RegExReplace(line, "\[食間\]", "")
            line := StrReplace(line, "分1", "") ; 分1は省略
            
            isBlock := InStr(line, "@@BLOCK@@") ; 複数薬品への共通用法フラグ
            line := StrReplace(line, "@@BLOCK@@", "")
            
            prevLine := (processedLines.Length > 0) ? processedLines[processedLines.Length] : ""
            
            ; 直前の行が「時」で終わっていない場合、薬品名と用法を結合
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

; ------------------------------------------------------------------------------
; 関数: ApplyBasicFormatting
; 役割: 薬品名周りのノイズ除去と、数量の境界線(@@SPACE@@)の作成
; ------------------------------------------------------------------------------
ApplyBasicFormatting(text) {
    ; s)オプション: 改行を跨いで「(〜として)」を削除。[^)]+は「)」以外の1文字以上の連続。
    text := RegExReplace(text, "s)\s*\([^)]+として\)", "")
    
    ; m)オプション: 各行の末尾にある「7日分」などを削除
    text := RegExReplace(text, "m)\d+\S+分$", "")
    
    text := StrReplace(text, "吸入用", "")

    ; 数量マーカー付与: 行末にある「錠/p/枚/ﾄ」または「スペース+数字+g」を識別
    ; m)各行判定、(*ANYCRLF)改行コード柔軟対応
    text := RegExReplace(text, "m)(*ANYCRLF)(\d+\S*[錠p枚ﾄ]$|\s\d+\S*g$)", "@@SPACE@@$1")
    
    text := RegExReplace(text, "m)(*ANYCRLF)cap$", "c")
    return text
}
; ------------------------------------------------------------------------------
; 関数: MergeSpecificPatterns
; 役割: 外用薬の指示(外)や、泣き別れた用法(時)を薬品名と結合する
; ------------------------------------------------------------------------------
MergeSpecificPatterns(text) {
    lines := StrSplit(text, "`n", "`r")
    result := []
    for line in lines {
        if (line == "") continue
        
        if (InStr(line, "@@BLOCK@@")) {
            result.Push(line)
            continue
        }
        ; 行末が「時」で終わる場合、直前の薬品名行と結合（1行にまとめる）
        if (RegExMatch(line, "^.+時\s*$")) {
            if (result.Length > 0 && !InStr(result[result.Length], "@@BLOCK@@"))
                result[result.Length] .= line
            else
                result.Push(line)
        } 
        ; 「外) 〜」という形式の行を薬品名と結合
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

; ------------------------------------------------------------------------------
; 関数: FinalizeText
; 役割: 一時的なマーカーの除去と最終的な表記の微調整
; ------------------------------------------------------------------------------
FinalizeText(text) {
    text := StrReplace(text, "@@SPACE@@", " ") ; 数量前のマーカーを半角スペースに戻す
    text := RegExReplace(text, "\(\Sとして\)", "") ; (1として) などの微細な注釈を削除
    text := StrReplace(text, "(非持参)", "")
    text := RegExReplace(text, " +", " ") ; 連続する半角スペースを1つに集約
    return Trim(text, "`n`r")
}

; ------------------------------------------------------------------------------
; 関数: ReorganizeByTrigger
; 役割: 「処方日」をキーにして薬品ごとのブロックを解析し、整形順序を整える
; ------------------------------------------------------------------------------
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
                outLine := buffer . line
                if (triggerCount > 1) ; 1つの用法に複数薬品がある場合のフラグ
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

; ------------------------------------------------------------------------------
; 関数: FilterOutpatientOrder
; 役割: 外来オーダ特有のヘッダー行（ハイフンや特定の記号）を除去する
; ------------------------------------------------------------------------------
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

; ------------------------------------------------------------------------------
; 関数: ProcessInitialInput
; 役割: コピー実行、待機、および全角・半角変換を一括で行う
; ------------------------------------------------------------------------------
ProcessInitialInput() {
    savedClip := A_Clipboard
    A_Clipboard := ""
    Send("^c")
    if !ClipWait(0.5)
        A_Clipboard := savedClip
    return ConvertToHalfWidth(A_Clipboard)
}

; ------------------------------------------------------------------------------
; 関数: ConvertToHalfWidth (WinAPI利用)
; 役割: 全角カタカナ、英数字、記号をすべて半角に変換する
; ------------------------------------------------------------------------------
ConvertToHalfWidth(str) {
    ; LCMapStringW: 0x00400000 は LCMAP_HALFWIDTH (全角から半角へ)
    size := DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", 0, "Int", 0)
    buf := Buffer(size * 2)
    DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", buf, "Int", size)
    return StrGet(buf, "UTF-16")
}
