/*
================================================================================
Script Name    : PrescriptionTextFormatter.ahk
Version        : 1.10.0
Description    : 処方箋整形（「時」の直後が空白・行末の場合のみ結合）
Update         : 2026-03-29 - 結合条件を「時＋空白」または「時＋行末」に限定
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
                    ; 次行の結合判定（分、または時＋空白/行末、または発熱時など）
                    if (RegExMatch(nextLine, "^(分|.+時(TEMP_SPACE|$)|発熱時(TEMP_SPACE|$)|疼痛時(TEMP_SPACE|$)|不眠時(TEMP_SPACE|$))"))
                        break
                    line .= nextLine
                    i++
                }

                if (RegExMatch(line, "i)cap$"))
                    line := RegExReplace(line, "i)cap$", "c")
                
                if (RegExMatch
