; ============================================================
;  File: PrescriptionFormatter.ahk
;  Version: 2.1.1
;  Author: -q
;  AutoHotkey v2.0.22
;
;  Win + Alt + S : 用法なし出力（DI情報判定あり）
;  Win + Alt + D : 用法あり出力（分1/分2/分3 の用法処理）
; ============================================================


; ============================================================
; Hotkey: Win + Alt + S（用法なし）
; ============================================================
#!s:: {
    text := GetPlainTextAndHankaku()

    ; DI情報判定：先頭行が「商品名」
    if RegExMatch(text, "^商品名") {
        text := RegExReplace(text, "m)^商品名[ \t]*")
    } else {
        new := []
        for line in StrSplit(text, "`n") {
            line := RTrim(line, "`r")

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

    ; すべての空白（スペース/タブ）削除
    text := RegExReplace(text, "[ \t]")

    ; @@SPACE@@ を半角スペースに戻す
    text := StrReplace(text, "@@SPACE@@", " ")

    ; 「(○として)」削除
    text := RegExReplace(text, "\(\Sとして\)")

    A_Clipboard := text
}


; ============================================================
; Hotkey: Win + Alt + D（用法あり）
; ============================================================
#!d:: {
    text := GetPlainTextAndHankaku()
    text := NormalizeOrderType(text)
    text := NormalizeDoseSuffix(text)
    text := NormalizeYohouLines(text)

    ; 空白削除（改行は残す）
    text := RegExReplace(text, "[ \t]")

    ; 分1/2/3 → B処理
    if RegExMatch(text, "m)^分[123]\S+")
        text := ProcessBun1to3Yohou(text)
    ; 分○ → C処理
    else if RegExMatch(text, "m)^分\d\S+")
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

    first := RTrim(lines[1], "`r")

    ; 外来処方オーダー：「--」から始まる
    if RegExMatch(first, "^--") {
        new := []
        for line in lines {
            line := RTrim(line, "`r")
            if line = ""
                continue
            if RegExMatch(line, "^(--|<R|処方箋期限)")
                continue
            new.Push(line)
        }
        return JoinLines(new)
    }

    ; 入院処方オーダー：「処方日」から始まる
    if RegExMatch(first, "^処方日") {
        text2 := MergeNyuuinShohouBlocks(text)
        new := []
        for line in StrSplit(text2, "`n") {
            line := RTrim(line, "`r")
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
        line := RTrim(line, "`r")

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
    ; 「\d+\S+分$」にマッチする文字列を削除
    text := RegExReplace(text, "m)\d+\S+分$")

    ; 「\d+\S+[錠pg枚ﾄ]$」にマッチする文字列の前に「@@SPACE@@」を挿入
    text := RegExReplace(text, "m)(\d+\S+[錠pg枚ﾄ])$", "@@SPACE@@$1")

    ; 「cap$」を「c」に置換
    text := RegExReplace(text, "m)cap$", "c")

    return text
}


; ============================================================
; 関数：NormalizeYohouLines（用法行整形）
; ============================================================
NormalizeYohouLines(text) {
    new := []
    for line in StrSplit(text, "`n") {
        line := RTrim(line, "`r")

        ; 「^\S+時」にマッチする行は上の行に結合
        if RegExMatch(line, "^\S+時") {
            if new.Length
                new[new.Length] .= line
            else
                new.Push(line)
            continue
        }

        ; 「^分\d+\s\d」にマッチする行のスペースを @@SPACE@@ に置換
        if RegExMatch(line, "^分\d+[ \t]\d") {
            line := RegExReplace(line, "[ \t]", "@@SPACE@@")
            new.Push(line)
            continue
        }

        ; 「^外)\s」を @@SPACE@@ に置換して上の行に結合
        if RegExMatch(line, "^外\)[ \t]") {
            line := RegExReplace(line, "^外\)[ \t]*", "@@SPACE@@")
            if new.Length
                new[new.Length] .= line
            else
                new.Push(line)
            continue
        }

        ; 「^吸入用」を削除
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
        line := RTrim(line, "`r")

        if RegExMatch(line, "^分[123]\S+") {
            line := RegExReplace(line, "毎(?=.)")
            line := RegExReplace(line, "食後")
            line := RegExReplace(line, "(?:と)?眠前", "寝")
            line := RegExReplace(line, "食前", "前")
            line := RegExReplace(line, "\[食間\]")
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
        line := RTrim(line, "`r")

        if RegExMatch(line, "^分\d\S+") {
            line := RegExReplace(line, "(?:と)?眠前", "寝")
            line := RegExReplace(line, "\[食間\]")
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
    out := ""
    for i, line in arr {
        if i > 1
            out .= "`r`n"
        out .= line
    }
    return out
}

ToHalfWidth(s) {
    bufSize := StrLen(s) * 2 + 2
    buf := Buffer(bufSize, 0)

    DllCall("LCMapStringW"
        , "UInt", 0x0411          ; Japanese
        , "UInt", 0x00800000      ; LCMAP_HALFWIDTH
        , "WStr", s
        , "Int", -1
        , "Ptr", buf.Ptr
        , "Int", bufSize)

    return StrGet(buf.Ptr, "UTF-16")
}
