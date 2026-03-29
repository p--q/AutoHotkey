/*
================================================================================
Script Name    : PrescriptionTextFormatter.ahk
Version        : 1.9.0
Description    : 処方箋整形（用法キーワードで結合を止め、以降を別行にする修正）
Update         : 2026-03-29 - 「発熱時」等で結合する際、その直後で改行を入れるよう修正
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
                    if (RegExMatch(nextLine, "^(分|.+時|発熱時|疼痛時|不眠時|の時)"))
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
    ToolTip("処方整形完了 (Ver 1.9.0)")
    SetTimer(() => ToolTip(), -1000)
}

; --- サブ関数：行結合ルール ---
MergeMedicalLines(lineArray) {
    output := ""
    keywords := "発熱時|疼痛時|不眠時|頓用|の時"
    
    for line in lineArray {
        ; 1. 「分」から始まる行の処理
        if (SubStr(line, 1, 1) == "分") {
            line := StrReplace(line, "毎食後", "")
            line := StrReplace(line, "食後", "")
            line := StrReplace(line, "眠前", "寝")
            output := RegExReplace(output, "\r?\n$", "") . line . "`n"
        } 
        ; 2. 頓用キーワードや「時」を含む行の処理
        else if (RegExMatch(line, "i)(" . keywords . "|時)")) {
            ; キーワードにマッチした部分までを抽出し、それ以降は切り捨てる（改行として扱う）
            ; 例：「発熱時1日2回」→「発熱時」を結合し、「1日2回」は新しい行へ
            if (RegExMatch(line, "i)^.*?(?:" . keywords . "|時)", &match)) {
                matchedPart := match[0]
                remainingPart := LTrim(SubStr(line, StrLen(matchedPart) + 1))
                
                ; キーワードまでを上の行に結合
                output := RegExReplace(output, "\r?\n$", "") . matchedPart . "`n"
                
                ; 残りがあれば新しい行として追加
                if (remainingPart != "")
                    output .= remainingPart . "`n"
            } else {
                output .= line . "`n"
            }
        } 
        else {
            output .= line . "`n"
        }
    }
    return Trim(output, "`n")
}

ToHalfWidth(str) {
    size := DllCall("LCMapStringW", "UInt", 0x0411, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", 0, "Int", 0)
    buf := Buffer(size * 2)
    DllCall("LCMapStringW", "UInt", 0x0411, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", buf, "Int", size)
    result := StrGet(buf, "UTF-16")
    return StrReplace(result, "　", " ")
}
