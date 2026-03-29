/*
================================================================================
Script Name    : PrescriptionTextFormatter.ahk
Version        : 1.24.0
Description    : 処方箋整形（092行目の While 閉じ忘れを完全修正）
Update         : 2026-03-29 - 入院処方ブロックの構造を再整理
--------------------------------------------------------------------------------
Hotkeys: Win + Alt + J
================================================================================
*/

#Requires AutoHotkey v2.0
#SingleInstance Force

; ==============================================================================
; メインホットキー: Win + Alt + J
; ==============================================================================
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
    if (SubStr(lines[1], 1, 2) == "--" || SubStr(lines[1], 1, 2) == "<R" || (lines.Length >= 2 && SubStr(lines[2], 1, 2) == "<R") || RegExMatch(lastLine, "^処方箋使用期限")) {
        isOutpatient := true
    }

    Result := ""
    if (isOutpatient) 
    {
        ; --- 外来処方箋の処理 ---
        TempLines := []
        for line in lines {
            if (RegExMatch(line, "^(--|<R|処方箋使用期限)"))
                continue
            
            if (RegExMatch(line, "i)[錠pﾄ枚gc]$"))
                line := RegExReplace(line, "\s+(?=[^\s]+$)", "TEMP_SPACE")

            line := CleanLineEndings(line)
            line := RegExReplace(line, "\s+", "")
            
            if (line != "")
                TempLines.Push(line)
        }
        Result := MergeMedicalLines(TempLines)
        Result := StrReplace(Result, "TEMP_SPACE", " ")
    } 
    else 
    {
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
        ; --- ブロックごとのループ ---
        for blockLines in Blocks {
            ProcessedBlock := []
            i := 1
            ; --- 行ごとのループ (ここが092行目付近) ---
            while (i <= blockLines.Length) {
                line := blockLines[i]
                line := RegExReplace(line, "^(\(非持参\)|外\))", "")
                line := RegExReplace(line, "\([^)]+として\)", "")

                ; 次の行を薬品名として結合するか判定するループ
                while (i < blockLines.Length) {
                    if (RegExMatch(line, "i)([錠pﾄ枚gc]|cap)$"))
                        break
                    
                    nextLine := blockLines[i+1]
                    if (RegExMatch(nextLine, "^(分|.+時(TEMP_SPACE|$| )|発熱時|疼痛時|不眠時)"))
                        break
                    
                    line .= nextLine
                    i++
