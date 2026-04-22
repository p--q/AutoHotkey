; ==============================================================================
; File Name: PrescriptionFormatter.ahk
; Version:   1.1.0 (Base)
; Description:
;   処方箋・電子カルテのテキストを整形するスクリプト。
;   全角を半角に変換し、不要な空白や特定のキーワードを除去・簡略化します。
; ==============================================================================

#Requires AutoHotkey v2.0

; --- Win + Alt + S: 用法なし整形 ---
; 薬品名と数量のみを抽出し、飲み方（用法）をカットしてコピーします。
#!s:: {
    info := PrepareFormatting("用法なし")
    if (!info) {
        return
    }
    
    text := info.Text
    if (RegExMatch(text, "^商品名")) {
        text := RegExReplace(text, "^商品名\s*", "")
        text := FinalizeText(text)
        A_Clipboard := text
        ToolTip(info.Msg "(用法錠数なし)")
    } else {
        text := ApplyBasicFormatting(text) ; 基本的な変換（単位の前にスペースを入れる等）
        if (RegExMatch(text, "処方日")) {
            text := ReorganizeByTrigger(text) ; 改行で分断された薬品名を結合
        }
        
        lines := StrSplit(text, "`n", "`r")
        resText := ""
        for line in lines {
            ; 「@@SPACE@@」が含まれる行（薬品名＋数量の行）だけを抽出
            if (InStr(line, "@@SPACE@@")) {
                resText .= line "`n"
            }
        }
        text := RegExReplace(resText, "[ \t]+", "") ; 不要な空白を削除
        text := FinalizeText(text)
        text := RegExReplace(text, "m)\s\d+枚\R外\)\s*(1日\s*\d+枚)$", " $1")
        
        A_Clipboard := text
        ToolTip(info.Msg)
    }
    SetTimer(() => ToolTip(), -2000)
}

; --- Win + Alt + D: 用法あり整形 ---
; 薬品名、数量に加えて、用法（飲み方）を簡略化してコピーします。
#!d:: {
    info := PrepareFormatting("用法あり")
    if (!info) {
        return
    }
    
    text := info.Text
    ; 特定の外来指示などの不要なヘッダー行を削除
    if (SubStr(text, 1, 2) == "--") {
        text := FilterOutpatientOrder(text)
    } 
    
    text := ApplyBasicFormatting(text)
    
    if (RegExMatch(text, "処方日")) {
        text := ReorganizeByTrigger(text)
    }
    
    ; 頓服や「～時」などの特殊な行の結合処理
    text := MergeSpecificPatterns(text)
    text := RegExReplace(text, "[ \t]+", "")
    
    lines := StrSplit(text, "`n", "`r"), processedLines := []
    for line in lines {
        if (line == "") {
            continue
        }
        ; 用法（分○、1日○回など）の表現を短縮
        if (RegExMatch(line, "^(分\d|1日\d回|1日\d枚)")) {
            line := RegExReplace(line, "毎(?=.)|食後", "") ; 「毎食後」→「食後」を消して簡略化
            line := RegExReplace(line, "(?:と)?眠前", "寝")  ; 「眠前」→「寝」
            line := RegExReplace(line, "食前", "前")         ; 「食前」→「前」
            line := RegExReplace(line, "\[食間\]", "")
            line := StrReplace(line, "分1", "")           ; 「分1」は省略
            
            isBlock := InStr(line, "@@BLOCK@@")
            line := StrReplace(line, "@@BLOCK@@", "")
            
            prevLine := (processedLines.Length > 0) ? processedLines[processedLines.Length] : ""
            ; 薬品名の行に用法をくっつける（直前が「時」で終わる特殊な場合を除く）
            if (!isBlock && processedLines.Length > 0 && !RegExMatch(prevLine, "時\s*$")) {
                processedLines[processedLines.Length] .= line
            } else {
                processedLines.Push(line)
            }
        } else {
            processedLines.Push(StrReplace(line, "@@BLOCK@@", ""))
        }
    }
    
    resText := ""
    for line in processedLines {
        resText .= line "`n"
    }
        
    text := FinalizeText(resText)
    text := RegExReplace(text, "m)\s\d+枚\R外\)\s*(1日\s*\d+枚)$", " $1")
    
    A_Clipboard := text
    ToolTip(info.Msg)
    SetTimer(() => ToolTip(), -2000)
}

; --- 共通関数群 ---

; 整形前の準備（入力取得、エラーチェック、ツールチップ準備）
PrepareFormatting(suffix) {
    resultObj := ProcessInitialInput()
    if (resultObj.Text == "") {
        NotifyError()
        return false 
    }
    sourceLabel := (resultObj.Source == "Selected") ? "選択を整形" : "クリップボードを整形"
    return {Text: resultObj.Text, Msg: sourceLabel "(" suffix ")"}
}

; 入力テキストの取得（選択中ならコピー、なければ現在のクリップボードから）
ProcessInitialInput() {
    savedClip := A_Clipboard
    A_Clipboard := "" 
    Send("^c")
    if ClipWait(0.5) {
        rawText := A_Clipboard
        src := "Selected"
    } else {
        rawText := savedClip
        src := "Clipboard"
    }
    ; 取得した文字列を半角に変換して返す
    return {Text: ConvertToHalfWidth(rawText), Source: src}
}

NotifyError() {
    ToolTip("整形する文字列を取得できませんでした")
    SetTimer(() => ToolTip(), -2000)
}

; 基本的な整形ルール（不要な注釈の削除、単位の前に目印を付与）
ApplyBasicFormatting(text) {
    text := RegExReplace(text, "s)\s*\([^)]+として\)", "")
    text := RegExReplace(text, "m)\d+\S+分$", "") ; 行末の「7日分」などを削除
    text := StrReplace(text, "吸入用", "")
    ; 数量（錠、mLなど）の前に、後でスペースに変換するための「@@SPACE@@」を付与
    text := RegExReplace(text, "i)(\d+)([錠p枚個]|cap|g|mL|ｷｯﾄ)", "@@SPACE@@$1$2")  ; 個は最後に削るが分割行を結合するトリガーに@@SPACE@@を使うためにここでは必要。
    text := RegExReplace(text, "i)cap", "c")
    return text
}

; 改行で分断されている薬品名を一つの行にまとめるロジック
ReorganizeByTrigger(text) {
    blocks := [], currentBlock := []
    lines := StrSplit(text, "`n", "`r")
    ; 処方日ごとにテキストをブロック分けする
    for line in lines {
        if (RegExMatch(line, "^処方日")) {
            if (currentBlock.Length > 0) {
                blocks.Push(currentBlock)
            }
            currentBlock := []
        } else if (line != "") {
            currentBlock.Push(line)
        }
    }
    if (currentBlock.Length > 0) {
        blocks.Push(currentBlock)
    }
    
    finalOutput := ""
    for blockLines in blocks {
        triggerCount := 0
        for l in blockLines {
            if (InStr(l, "@@SPACE@@")) {
                triggerCount++
            }
        }
        
        buffer := ""
        for line in blockLines {
            if (InStr(line, "@@SPACE@@")) {
                ; 数量が見つかったら、それまでに溜めていた薬品名(buffer)と結合
                finalOutput .= buffer . line . (triggerCount > 1 ? "@@BLOCK@@" : "") . "`n"
                buffer := ""
            } else {
                ; 空白がない行は薬品名の一部とみなして溜めておく
                if (!InStr(line, " ") && !InStr(line, "　")) {
                    buffer .= line
                } else {
                    finalOutput .= line . (triggerCount > 1 ? "@@BLOCK@@" : "") . "`n"
                }
            }
        }
        if (buffer != "") {
            finalOutput .= buffer . (triggerCount > 1 ? "@@BLOCK@@" : "") . "`n"
        }
    }
    return finalOutput
}

; 特定のパターン（時、外）を前の行と結合する
MergeSpecificPatterns(text) {
    lines := StrSplit(text, "`n", "`r"), result := []
    for line in lines {
        if (line == "") {
            continue
        }
        ; 「頓)医師の指示通り」という文字列が含まれていたらスキップ
        if (InStr(line, "頓)医師の指示通り")) {
            continue
        }
        if (InStr(line, "@@BLOCK@@")) {
            result.Push(line)
            continue
        }
        ; 「外）」で始まる行を前の行に結合
        if (RegExMatch(line, "^\s*外\)\s*(.*)$", &m)) {
            if (result.Length > 0 && !InStr(result[result.Length], "@@BLOCK@@")) {
                result[result.Length] .= "@@SPACE@@" . m[1]
            } else {
                result.Push("@@SPACE@@" . m[1])
            }
        ; 見つかった位置が 1 より大きい（＝2文字目以降にある）場合のみ実行
        } else if (InStr(line, "時") > 1) {
            if (result.Length > 0 && !InStr(result[result.Length], "@@BLOCK@@")) {
                result[result.Length] .= line
            } else {
                result.Push(line)
            }
        }
    }
    resFinal := ""
    for l in result {
        resFinal .= l "`n"
    }
    return resFinal
}

; 最終的な仕上げ（目印の除去、余分なスペースの整理）
FinalizeText(text) {
    text := RegExReplace(text, "m)@@SPACE@@\d+個$", "")  ; 個で終わる単位を消去。
    text := StrReplace(text, "@@SPACE@@", " ")
    text := StrReplace(text, "@@BLOCK@@", "")
    text := RegExReplace(text, "\(\Sとして\)", "")
    text := StrReplace(text, "(非持参)", "")
    text := RegExReplace(text, " +", " ") ; 連続するスペースを一つに
    return Trim(text, "`n`r")
}

; 外来指示など、不要な区切り線をフィルタリング
FilterOutpatientOrder(text) {
    lines := StrSplit(text, "`n", "`r"), resFilter := ""
    for line in lines {
        if (line == "" || RegExMatch(line, "^(--|<R|処方箋)")) {
            continue
        }
        resFilter .= line "`n"
    }
    return resFilter
}

; WinAPIを使用して全角英数字・記号・カタカナを半角に変換
ConvertToHalfWidth(str) {
    if (str == "") {
        return ""
    }
    ; LCMapStringW: 0x00400000 = LCMAP_HALFWIDTH
    size := DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", 0, "Int", 0)
    buf := Buffer(size * 2)
    DllCall("LCMapStringW", "UInt", 0x400, "UInt", 0x00400000, "Str", str, "Int", -1, "Ptr", buf, "Int", size)
    return StrGet(buf, "UTF-16")
}
