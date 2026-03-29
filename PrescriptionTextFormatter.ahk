/*
================================================================================
Script Name    : PrescriptionTextFormatter.ahk
Version        : 1.13.0
Description    : 処方箋整形（構文エラーの完全修正・全機能統合版）
Update         : 2026-03-29 - ToolTipの構文エラーを修正
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

    ; 1. 全角を半角に変換
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

    ; 2. 外来/入院 判定
    isOutpatient := false
    lastLine := lines[lines.Length]
    if (SubStr(lines[1], 1, 2) == "--" || SubStr(lines[1], 1, 2) == "<R" || (lines.Length >= 2 && SubStr(lines[2], 1, 2) == "<R") || RegExMatch(lastLine, "^処方箋使用期限"))
        isOutpatient := true

    Result := ""
    if (isOutpatient) {
        ; --- 外来処方箋の処理 ---
        TempLines := []
        for line in lines {
            if (RegExMatch(line, "^(--|<R|処方箋使用期限)"))
                continue
            
            ; 数量前の空白保護
            if (RegExMatch(line, "[錠pﾄ枚g]$")) {
                line := RegExReplace(line, "\s+(?=[^\s]+$)", "TEMP_SPACE")
            }

            if (RegExMatch(line, "i)cap$"))
                line := RegExReplace(line, "i)cap$", "c")
            
            ; 末尾の「14日分」などを削除
            if (RegExMatch(line, "分$"))
                line := RegExReplace(line, "\d+[^\d]*分$", "")
            
            ; 余分な空白を削除
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
