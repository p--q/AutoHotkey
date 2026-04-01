; ============================================================
;  File: PrescriptionFormatter.ahk
;  Version: 2.0.0
;  Author: -q
;
;  Description:
;    医療現場で使用する処方テキストを、AutoHotkey v2 により
;    外来処方・入院処方・DI情報のルールに従って整形する。
;
;    Win + Alt + S : 用法なし出力（DI情報判定あり）
;    Win + Alt + D : 用法あり出力（分1/分2/分3 の用法処理）
;
; ============================================================


; ============================================================
; Hotkey: Win + Alt + S（用法なし）
; ============================================================
#!s::{
    text := GetPlainTextAndHankaku()

    ; DI情報判定
    if RegExMatch(text, "^商品名", &m) {
        text := RegExReplace(text, "^商品名[ \t]*")
    } else {
        new := []
        for line in StrSplit(text, "`n") {
            if RegExMatch(line, "^分[123]\S+") 
                continue
            if RegExMatch(line, "^外\)")
                continue
            if RegExMatch(line, "^吸入用")
                continue
            if RegExMatch(line, "\d+\S+分$")
                continue
            new.Push(line)
        }
        text := JoinLines(new)
    }

    text := NormalizeOrderType(text)
    text := NormalizeDoseSuffix(text)

    text := RegExReplace(text, "[ \t]")               ; 空白削除
    text := StrReplace(text, "@@SPACE@@", " ")        ; 復元
    text := RegExReplace(text, "\(\Sとして\)")        ; (○として)削除

    A_Clipboard := text
}


; ============================================================
; Hotkey: Win + Alt + D（用法あり）
; ============================================================
#!d::{
    text := GetPlainTextAndHankaku()
    text := NormalizeOrderType(text)
    text := NormalizeDoseSuffix(text)
    text := NormalizeYohouLines(text)

    text := RegExReplace(text, "[ \t]")               ; 空白削除

    if RegExMatch(text, "^分[123]\S+", &m)
        text := ProcessBun1to3Yohou(text)
    else if RegExMatch(text, "^分\d\S+", &m)
        text := ProcessBunAnyYohou(text)

    text := StrReplace(text, "@@SPACE@@", " ")
    text := RegExReplace(text, "\(\Sとして\)")

    A_Clipboard := text
}


; ============================================================
; 関数：GetPlainTextAndHankaku
; ============================================================
GetPlainTextAndHankaku() {
    old := A_Clipboard
    A_Clipboard := ""
    Send("^c")
    ClipWait(0.2)

    text := (A_Clipboard = "" ? old : A_Clipboard)
    return ToHalfWidth(text)
}


; ============================================================
; 関数：NormalizeOrderType（外来/入院処方）
; ============================================================
NormalizeOrderType(text) {
    lines := StrSplit(text, "`n")
    if lines.Length = 0
        return text

    first := lines[1]

    ; 外来処方
    if RegExMatch(first, "^--") {
        new := []
        for line in lines {
            if line = ""
                continue
            if RegExMatch(line, "^(--|<R|処方箋期限)")
                continue
            new.Push(line)
        }
        return JoinLines(new)
    }

    ; 入院処方
    if RegExMatch(first, "^処方日") {
        text2 := MergeNyuuinShohouBlocks(text)
        new := []
        for line in StrSplit(text2, "`n") {
            if RegExMatch(line, "^処方日")
                continue
            new.Push(line)
        }
        return JoinLines(new)
    }

    return text
}


; ============================================================
; 関数：MergeNyuuinShohouBlocks（入院処方ブロック結合）
; ============================================================
MergeNyuuinShohouBlocks(text) {
    lines := StrSplit(text, "`n")
    result := []
    block := []

    for line in lines {
        if RegExMatch(line, "^処方日") {
            if block.Length
                ProcessNyuuinBlock(block, result)
            result.Push(line)
            block := []
        } else {
            block.Push(line)
        }
    }

    if block.Length
        ProcessNyuuinBlock(block, result)

    return JoinLines(result)
}

ProcessNyuuinBlock(&block, &result) {
    buf := ""
    for line in block {
        if RegExMatch(line, " ") {
            if buf != "" {
                buf .= line
                result.Push(buf)
                buf := ""
            } else {
                result.Push(line)
            }
        } else {
            buf .= line
        }
    }
    if buf != ""
        result.Push(buf)
}


; ============================================================
; 関数：NormalizeDoseSuffix（分・錠・cap）
; ============================================================
NormalizeDoseSuffix(text) {
    text := RegExReplace(text, "m)\d+\S+分$")
    text := RegExReplace(text, "m)(\d+\S+[錠pg枚ﾄ])$", "@@SPACE@@$1")
    text := RegExReplace(text, "m)cap$", "c")
    return text
}


; ============================================================
; 関数：NormalizeYohouLines（用法行整形）
; ============================================================
NormalizeYohouLines(text) {
    new := []
    for line in StrSplit(text, "`n") {

        if RegExMatch(line, "^\S+時") {
            if new.Length
                new[new.Length] .= line
            else
                new.Push(line)
            continue
        }

        if RegExMatch(line, "^分\d+[ \t]\d") {
            line := RegExReplace(line, "[ \t]", "@@SPACE@@")
            new.Push(line)
            continue
        }

        if RegExMatch(line, "^外\)[ \t]") {
            line := RegExReplace(line, "^外\)[ \t]*", "@@SPACE@@")
            if new.Length
                new[new.Length] .= line
            else
                new.Push(line)
            continue
        }

        if RegExMatch(line, "^吸入用")
            continue

        new.Push(line)
    }
    return JoinLines(new)
}


; ============================================================
; 関数：ProcessBun1to3Yohou（分1/2/3）
; ============================================================
ProcessBun1to3Yohou(text) {
    new := []
    for line in StrSplit(text, "`n") {
        if RegExMatch(line, "^分[123]\S+") {
            line := RegExReplace(line, "毎(?=.)")
            line := RegExReplace(line, "食後")
            line := RegExReplace(line, "(?:と)?眠前", "寝")
            line := RegExReplace(line, "食前", "前")
            line := RegExReplace(line, "

\[食間\]

")
            line := RegExReplace(line, "1日\d回")

            if new.Length
                new[new.Length] .= line
            else
                new.Push(line)
        } else {
            new.Push(line)
        }
    }
    return JoinLines(new)
}


; ============================================================
; 関数：ProcessBunAnyYohou（分○）
; ============================================================
ProcessBunAnyYohou(text) {
    new := []
    for line in StrSplit(text, "`n") {
        if RegExMatch(line, "^分\d\S+") {
            line := RegExReplace(line, "(?:と)?眠前", "寝")
            line := RegExReplace(line, "

\[食間\]

")
            line := RegExReplace(line, "1日\d回")

            if new.Length
                new[new.Length] .= line
            else
                new.Push(line)
        } else {
            new.Push(line)
        }
    }
    return JoinLines(new)
}


; ============================================================
; Utility
; ============================================================
JoinLines(arr) {
    return arr.Join("`r`n")
}

ToHalfWidth(s) {
    bufSize := StrLen(s) * 2 + 2
    buf := Buffer(bufSize, 0)
    DllCall("LCMapStringW"
        , "UInt", 0x0411
        , "UInt", 0x00800000
        , "WStr", s
        , "Int", -1
        , "Ptr", buf.Ptr
        , "Int", bufSize)
    return StrGet(buf.Ptr, "UTF-16")
}
