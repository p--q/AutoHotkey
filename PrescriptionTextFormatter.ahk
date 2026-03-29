/*
================================================================================
Script Name    : PrescriptionTextFormatter.ahk
Version        : 1.27.0
Description    : 処方箋整形（構文エラーを完全に解消した最終安定版）
Update         : 2026-03-29 - 波括弧の対応を物理的に全行再検証
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

    if (originalText == "") {
        return
    }

    str := ToHalfWidth(originalText)
    lines := []
    Loop Parse, str, "`n", "`r"
    {
        trimmed := Trim(A_LoopField)
        if (trimmed != "") {
            lines.Push(trimmed)
        }
    }

    if (lines.Length == 0) {
        return
    }

    isOutpatient := false
    lastLine := lines[lines.Length]
    if (SubStr(lines[1], 1, 2) == "--" || SubStr(lines[1], 1, 2) == "<R" || (lines.Length >= 2 && SubStr(lines[2], 1, 2) == "<R") || RegExMatch(lastLine, "^処方箋使用期限")) {
        isOutpatient := true
    }

    Result := ""
    if (isOutpatient) 
    {
        TempLines := []
        for line in lines {
            if (RegExMatch(line, "^(--|<R|処方箋使用期限)")) {
                continue
            }
            if (RegExMatch(line, "i)[錠pﾄ枚gc]$")) {
                line := RegExReplace(line, "\s+(?=[^\s]+$)", "TEMP_SPACE")
            }
            line := CleanLineEndings(line)
            line := RegExReplace(line, "\s+", "")
            if (line != "") {
                TempLines.Push(line)
            }
        }
        Result := MergeMedicalLines(TempLines)
        Result := StrReplace(Result, "TEMP_SPACE", " ")
    } 
    else 
    {
        Blocks := []
        CurrentBlock := []
        for line in lines {
            if (RegExMatch(line, "^処方日")) {
                if (CurrentBlock.Length > 0) {
                    Blocks.Push(CurrentBlock)
                }
                CurrentBlock := []
                continue
            }
            CurrentBlock.Push(line)
        }
        if (CurrentBlock.Length > 0) {
            Blocks.Push(CurrentBlock)
        }

        FinalOutput := ""
        for blockLines in Blocks {
            ProcessedBlock := []
            i := 1
            while (i <= blockLines.Length) {
                line := blockLines[i]
                line := RegExReplace(line, "^(\(非持参\)|外\))", "")
                line := RegExReplace(line, "\([^)]+として\)", "")

                while (i < blockLines.Length) {
                    if (RegExMatch(line, "i)([錠pﾄ枚gc]|cap)$"))
