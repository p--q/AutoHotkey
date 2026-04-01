/*
================================================================================
Script Name    : PrescriptionTextFormatter.ahk
Version        : 1.28.0
Description    : 処方箋整形（波括弧ネスト構造の完全平坦化版）
Update         : 2026-03-29 - 入院処方ループの波括弧エラーを物理的に排除
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
    if (SubStr(lines[1], 1, 2) == "--"
     || SubStr(lines[1], 1, 2) == "<R"
     || (lines.Length >= 2 && SubStr(lines[2], 1, 2) == "<R")
     || RegExMatch(lastLine, "^処方箋使用期限")) {
        isOutpatient := true
    }

    Result := ""
    if (isOutpatient) 
    {
        TempLines := []
        for _, line in lines {
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
        for _, line in lines {
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
        for _, blockLines in Blocks {
            ProcessedBlock := []
            i := 1
            while (i <= blockLines.Length) {
                line := blockLines[i]
                line := RegExReplace(line, "^(\(非持参\)|外\))", "")
                line := RegExReplace(line, "\([^)]+として\)", "")

                ; --- 次行結合ループ ---
                while (i < blockLines.Length) {
                    if (RegExMatch(line, "i)([錠pﾄ枚gc]|cap)$")) {
                        break
                    }
                    nextLine := blockLines[i+1]
                    if (RegExMatch(nextLine, "^(分|.+時(TEMP_SPACE|$| )|発熱時|疼痛時|不眠時)")) {
                        break
                    }
                    line .= nextLine
                    i++
                }

                ; --- 数量前空白保護 ---
                if (RegExMatch(line, "i)([錠pﾄ枚gc]|cap)$")) {
                    line := RegExReplace(line, "\s+(?=[^\s]+$)", "TEMP_SPACE")
                }

                line := CleanLineEndings(line)
                line := RegExReplace(line, "\s+", "")
                if (line != "") {
                    ProcessedBlock.Push(line)
                }
                i++
            }
            FinalOutput .= MergeMedicalLines(ProcessedBlock) . "`n"
        }
        Result := StrReplace(Trim(FinalOutput, "`n"), "TEMP_SPACE", " ")
    }

    A_Clipboard := Result
    ToolTip("処方整形完了 (Ver 1.28.0)")
    SetTimer(() => ToolTip(), -1000)
}

; --- 独立関数群 ---

CleanLineEndings(line) {
    if (RegExMatch(line, "i)cap$")) {
        line := RegExReplace(line, "i)cap$", "c")
    }
    line := RegExReplace(line, "i)\s*[0-9]*分[0-9]*.*$", "")
    return line
}

MergeMedicalLines(lineArray) {
    output := ""
    kw := "発熱時|疼痛時|不眠時|頓用|の時|時"
    for _, line in lineArray {
        if (SubStr(line, 1, 1) == "分") {
            line := StrReplace(line, "毎食後", "")
            line := StrReplace(line, "食後", "")
            line := StrReplace(line, "眠前", "寝")
            output := RegExReplace(output, "\r?\n$", "") . line . "`n"
        } else if (RegExMatch(line, "i)(" . kw . ")(TEMP_SPACE|$)")) {
            if (RegExMatch(line, "i)^.*?(" . kw . ")(TEMP_SPACE|$)", &match)) {
                matchedPart := match[0]
                remainingPart := LTrim(SubStr(line, StrLen(matchedPart) + 1))
                output := RegExReplace(output, "\r?\n$", "") . matchedPart . "`n"
                if (remainingPart != "") {
                    output .= remainingPart . "`n"
                }
            } else {
                output .= line . "`n"
            }
        } else {
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
