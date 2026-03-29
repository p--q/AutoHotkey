/*
================================================================================
Script Name    : PrescriptionTextFormatter.ahk
Version        : 1.7.0
Description    : 処方箋整形（「時＋空白」行の結合ロジック厳密化版）
Update         : 2026-03-29 - 「時 」で始まる・含む行の結合条件を修正
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

    ; 2. 外来判定
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
            
            if (RegExMatch(line, "分$"))
                line := RegExReplace(line, "\d+[^\d]*分$", "")
            
            ; TEMP_SPACE 以外の空白をすべて削除
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

                ; 泣き別れ結合（数量単位で終わらない場合の次行結合）
                while (i < blockLines.Length) {
                    if (RegExMatch(line, "[錠pﾄ枚g]$"))
                        break
                    nextLine := blockLines[i+1]
                    ; 次の行が「分」または「時 」で始まる場合は結合しない
                    if (RegExMatch(nextLine, "^分") || RegExMatch(nextLine, "^.+時\s+"))
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
    ToolTip("処方整形完了 (Ver 1.7.0)")
    SetTimer(() => ToolTip(), -1000)
}

; --- サブ関数：分・時の結合ルール ---
MergeMedicalLines(lineArray) {
    output := ""
    for line in lineArray {
        ; 「分」から始まる行
        if (SubStr(line, 1, 1) == "分") {
            line := StrReplace(line, "毎食後", "")
            line := StrReplace(line, "食後", "")
            line := StrReplace(line, "眠前", "寝")
            output := RegExReplace(output, "\r?\n$", "") . line . "`n"
        } 
        ; 「時」の直後に何らかの空白（名残）がある行の結合
        ; 例：「1日1回 10時 」が改行されて「 10時 」だけ独立した場合など
        else if (RegExMatch(line, "時(TEMP_SPACE|\s)")) {
            output := RegExReplace(output, "\r?\n$", "") . line . "`n"
        } 
        else {
            output .= line . "`n"
        }
    }
    return Trim(output, "`n")
}

; --- サブ関数：全角から半角へ ---
ToHalfWidth(str) {
    size := DllCall("LCMapStringW", "UInt", 0x0411, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", 0, "Int", 0)
    buf := Buffer(size * 2)
    DllCall("LCMapStringW", "UInt", 0x0411, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", buf, "Int", size)
    result := StrGet(buf, "UTF-16")
    return StrReplace(result, "　", " ")
}
