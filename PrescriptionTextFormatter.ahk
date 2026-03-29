/*
================================================================================
Script Name    : PrescriptionTextFormatter.ahk
Version        : 1.2.0
Description    : 処方箋OCRテキストの整形（全角半角変換・泣き別れ行の結合・特定語句の置換）
Author         : Gemini
Created        : 2026-03-29
Update         : 2026-03-29 - ホットキーを Win + Alt + J に変更（競合回避のため）
--------------------------------------------------------------------------------
Usage:
1. 整形したい処方テキストをクリップボードにコピー
2. Windows + Alt + J を押下
3. 自動整形されたテキストがクリップボードに書き戻される
================================================================================
*/

#Requires AutoHotkey v2.0
#SingleInstance Force

; Windowsキー(#) + Altキー(!) + J で実行
#!j::
{
    try {
        originalText := A_Clipboard
    } catch {
        return
    }

    if (originalText == "")
        return

    ; 1. すべてを半角に変換（カタカナ・英数字・記号）
    str := ToHalfWidth(originalText)

    lines := []
    Loop Parse, str, "`n", "`r"
    {
        trimmed := Trim(A_LoopField)
        if (trimmed != "")
            lines.Push(trimmed)
    }

    if (lines.Length == 0)
        return

    ; 2. 外来処方箋かどうかの判定
    isOutpatient := false
    lastLine := lines[lines.Length]
    
    if (SubStr(lines[1], 1, 2) == "--")
        isOutpatient := true
    else if (SubStr(lines[1], 1, 2) == "<R" || (lines.Length >= 2 && SubStr(lines[2], 1, 2) == "<R"))
        isOutpatient := true
    else if (RegExMatch(lastLine, "^処方箋使用期限"))
        isOutpatient := true

    Result := ""
    if (isOutpatient) {
        ; --- 外来処方箋の処理 ---
        TempLines := []
        for line in lines {
            if (RegExMatch(line, "^(--|<R|処方箋使用期限)"))
                continue
            
            ; cap -> c 置換
            if (RegExMatch(line, "i)cap$"))
                line := RegExReplace(line, "i)cap$", "c")
            
            ; 「分」で終わる行の末尾数字削除
            if (RegExMatch(line, "分$"))
                line := RegExReplace(line, "\d+分$", "")
            
            ; 空白削除
            line := RegExReplace(line, "\s+", "")
            if (line != "")
                TempLines.Push(line)
        }
        Result := MergeMedicalLines(TempLines)
    } else {
        ; --- 外来以外（入院等）の処理 ---
        Blocks := []
        CurrentBlock := []
        for line in lines {
            if (RegExMatch(line, "^処方日")) {
                if (CurrentBlock.Length > 0)
                    Blocks.Push(CurrentBlock)
                CurrentBlock := []
                continue
            }
            CurrentBlock.Push(line)
        }
        if (CurrentBlock.Length > 0)
            Blocks.Push(CurrentBlock)

        FinalOutput := ""
        for blockLines in Blocks {
            ProcessedBlock := []
            i := 1
            while (i <= blockLines.Length) {
                line := blockLines[i]
                
                ; 行頭の特定文字列削除
                line := RegExReplace(line, "^(\(非持参\)|外\))", "")
                
                ; 「として」で終わる丸括弧削除
                line := RegExReplace(line, "\([^)]+として\)", "")

                ; 泣き別れ結合ロジック
                while (i < blockLines.Length) {
                    if (RegExMatch(line, "[錠pﾄ枚g]$"))
                        break
                    
                    nextLine := blockLines[i+1]
                    if (RegExMatch(nextLine, "^分") || RegExMatch(nextLine, "時\s+"))
                        break
                    
                    line .= nextLine
                    i++
                }

                ; cap -> c 置換
                if (RegExMatch(line, "i)cap$"))
                    line := RegExReplace(line, "i)cap$", "c")
                
                ; 「分」末尾の数字削除
                if (RegExMatch(line, "分$"))
                    line := RegExReplace(line, "\d+分$", "")

                line := RegExReplace(line, "\s+", "")
                if (line != "")
                    ProcessedBlock.Push(line)
                i++
            }
            FinalOutput .= MergeMedicalLines(ProcessedBlock) . "`n"
        }
        Result := Trim(FinalOutput, "`n")
    }

    A_Clipboard := Result
    ToolTip("処方箋整形完了 (Win+Alt+J)")
    SetTimer(() => ToolTip(), -1000)
}

; --- サブ関数：分・時の結合ルール ---
MergeMedicalLines(lineArray) {
    output := ""
    for line in lineArray {
        if (SubStr(line, 1, 1) == "分") {
            line := StrReplace(line, "毎食後", "")
            line := StrReplace(line, "食後", "")
            line := StrReplace(line, "眠前", "寝")
            output := RegExReplace(output, "\r?\n$", "") . line . "`n"
        } 
        else if (RegExMatch(line, ".+時.+")) {
            output := RegExReplace(output, "\r?\n$", "") . line . "`n"
        } 
        else {
            output .= line . "`n"
        }
    }
    return Trim(output, "`n")
}

; --- サブ関数：全角から半角へ（カタカナ・英数・記号） ---
ToHalfWidth(str) {
    ; Windows API (LCMapStringW) を使用
    size := DllCall("LCMapStringW", "UInt", 0x0411, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", 0, "Int", 0)
    buf := Buffer(size * 2)
    DllCall("LCMapStringW", "UInt", 0x0411, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", buf, "Int", size)
    result := StrGet(buf, "UTF-16")
    
    result := StrReplace(result, "　", " ")
    return result
}
