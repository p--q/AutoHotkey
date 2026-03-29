/*
================================================================================
Script Name    : PrescriptionTextFormatter.ahk
Version        : 1.19.0
Description    : 処方箋整形（入院処方の c 置換および結合条件の修正）
Update         : 2026-03-29 - 入院処方の結合判定に "c" を追加
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
            
            ; 数量前の空白保護（capがcになる前なので p で判定）
            if (RegExMatch(line, "i)[錠pﾄ枚gc]$"))
                line := RegExReplace(line, "\s+(?=[^\s]+$)", "TEMP_SPACE")

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
    } 
    else 
    {
        ; --- 入院処方箋（外来以外）の処理 ---
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

                ; 次の行を結合するかどうかの判定ループ
                while (i < blockLines.Length) {
                    ; すでに単位（錠pﾄ枚gc）で終わっている、または cap で終わっているなら結合停止
                    if (RegExMatch(line, "i)([錠pﾄ枚gc]|cap)$")) {
                        break
                    }
                    nextLine := blockLines[i+1]
                    ; 次の行が用法（分、時、発熱時など）で始まるなら結合停止
                    if (RegExMatch(nextLine, "^(分|.+時(TEMP_SPACE|$| )|発熱時|疼痛時|不眠時)")) {
                        break
                    }
                    line .= nextLine
                    i++
                }

                ; 【重要】空白を消す前に数量前の空白を退避（c, cap 両方に対応）
                if (RegExMatch(line, "i)([錠pﾄ枚gc]|cap)$")) {
                    line := RegExReplace(line, "\s+(?=[^\s]+$)", "TEMP_SPACE")
                }

                ; cap -> c 置換
                if (RegExMatch(line, "i)cap$"))
                    line := RegExReplace(line, "i)cap$", "c")
                
                if (RegExMatch(line, "分$"))
                    line := RegExReplace(line, "\d+[^\d]*分$", "")

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

    ; 3. クリップボードへ出力
    A_Clipboard := Result
    ToolTip("処方整形完了 (Ver 1.19.0)")
    SetTimer(() => ToolTip(), -1000)
}

; --- サブ関数：行結合ルール ---
MergeMedicalLines(lineArray) {
    output := ""
    kw := "発熱時|疼痛時|不眠時|頓用|の時|時"
    
    for line in lineArray {
        if (SubStr(line, 1, 1) == "分") {
            line := StrReplace(line, "毎食後", "")
            line := StrReplace(line, "食後", "")
            line := StrReplace(line, "眠前", "寝")
            output := RegExReplace(output, "\r?\n$", "") . line . "`n"
        } 
        else if (RegExMatch(line, "i)(" . kw . ")(TEMP_SPACE|$)")) {
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
