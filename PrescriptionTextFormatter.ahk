/*
================================================================================
Script Name    : PrescriptionTextFormatter.ahk
Version        : 1.11.0
Description    : 処方箋整形（正規表現エラーの修正・「時」結合ロジック完成版）
Update         : 2026-03-29 - Missing ")" エラーを修正
--------------------------------------------------------------------------------
Hotkeys: Win + Alt + J
================================================================================
*/

#Requires AutoHotkey v2.0
#SingleInstance Force

#!j::
{
    try {
        originalText := A_Clipboard
    } catch {
        return
    }

    if (originalText == "")
        return

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

    isOutpatient := false
    lastLine := lines[lines.Length]
    if (SubStr(lines[1], 1, 2) == "--" || SubStr(lines[1], 1, 2) == "<R" || (lines.Length >= 2 && SubStr(lines[2], 1, 2) == "<R") || RegExMatch(lastLine, "^処方箋使用期限"))
        isOutpatient := true

    Result := ""
    if (isOutpatient) {
        TempLines := []
        for line in lines {
            if (RegExMatch(line, "^(--|<R|処方箋使用期限)"))
                continue
            
            if (RegExMatch(line, "[錠pﾄ枚g]$")) {
                line := RegExReplace(line, "\s+(?=[^\s]+$)", "TEMP_SPACE")
            }

            if (RegExMatch(line, "i)cap$"))
                line := RegExReplace(line, "i)cap$", "c")
            
            if (RegExMatch(line, "分$"))
                line := RegExReplace(line, "\d+[^\d]*分$", "")
            
            line := RegExReplace(line, "\s+", "")
            
            if (line != "")
                TempLines.Push(line)
        }
        Result := MergeMedicalLines(TempLines)
        Result := StrReplace(Result, "TEMP_SPACE", " ")
        
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
                line := RegExReplace(line, "^(\(非持参\)|外\))", "")
                line := RegExReplace(line, "\([^)]+として\)", "")

                while (i < blockLines.Length) {
                    if (RegExMatch(line, "[錠pﾄ枚g]$"))
                        break
                    nextLine := blockLines[i+1]
                    ; 次行の結合判定
                    if (RegExMatch(nextLine, "^(分|.+時(TEMP_SPACE|$)|発熱時(TEMP_SPACE|$)|疼痛時(TEMP_SPACE|$)|不眠時(TEMP_SPACE|$))"))
                        break
                    line .= nextLine
                    i++
                }

                if (RegExMatch(line, "i)cap$"))
                    line := RegExReplace(line, "i)cap$", "c")
                
                if (RegExMatch(line, "分$"))
                    line := RegExReplace(line, "\d+[^\d]*分$", "")

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
    ToolTip("処方整形完了 (Ver 1.11.0)")
    SetTimer(() => ToolTip(), -1000)
}

; --- サブ関数：行結合ルール ---
MergeMedicalLines(lineArray) {
    output := ""
    ; キーワードの定義（カッコを含まない形に修正）
    kw := "発熱時|疼痛時|不眠時|頓用|の時|時"
    
    for line in lineArray {
        ; 1. 「分」から始まる行
        if (SubStr(line, 1, 1) == "分") {
            line := StrReplace(line, "毎食後", "")
            line := StrReplace(line, "食後", "")
            line := StrReplace(line, "眠前", "寝")
            output := RegExReplace(output, "\r?\n$", "") . line . "`n"
        } 
        ; 2. 厳密な条件での結合（キーワードの直後が空白または行末）
        ; 正規表現を結合して構築する際のエラーを回避
        else if (RegExMatch(line, "i)(" . kw . ")(TEMP_SPACE|$)")) {
            ; マ
