; 関数D: 処方行の再構成
ReorganizePrescriptionLines(text) {
    lines := StrSplit(text, "`n", "`r")
    newOutput := ""
    buffer := ""
    
    for line in lines {
        if (line == "") 
            continue ; 改行して記述することで予約語として正しく認識されます
        
        if (InStr(line, " ")) { ; トリガー行
            newOutput .= buffer . line . "`n"
            buffer := ""
        } else {
            buffer .= line
        }
    }
    return newOutput . buffer
}

; 関数G内も同様に修正
FilterOrderType(text) {
    if (SubStr(text, 1, 2) == "--") {
        lines := StrSplit(text, "`n", "`r")
        result := ""
        for line in lines {
            ; 複数条件のif文でも同様に改行してcontinue
            if (line == "" || RegExMatch(line, "^(--|<R|処方箋)"))
                continue
            result .= line "`n"
        }
        return result
    } else if (RegExMatch(text, "^処方日")) {
        text := ReorganizePrescriptionLines(text)
        text := RegExReplace(text, "m)^処方日.*`n?", "")
        return text
    }
    return text
}

; 関数F内も同様に修正
MergeSpecificPatterns(text) {
    lines := StrSplit(text, "`n", "`r")
    result := []
    
    for line in lines {
        if (line == "")
            continue
        
        if (RegExMatch(line, "^\S+時")) {
            if (result.Length > 0)
                result[result.Length] .= line
            else
                result.Push(line)
        } else if (RegExMatch(line, "^分\d+\s\d", &m)) {
            line := RegExReplace(line, "^(分\d+)\s(\d)", "$1@@SPACE@@$2")
            result.Push(line)
        } else if (RegExMatch(line, "^外\)\s")) {
            line := RegExReplace(line, "^外\)\s", "@@SPACE@@")
            if (result.Length > 0)
                result[result.Length] .= line
            else
                result.Push(line)
        } else if (RegExMatch(line, "^吸入用")) {
            result.Push(RegExReplace(line, "^吸入用", ""))
        } else {
            result.Push(line)
        }
    }
    
    finalText := ""
    for l in result
        finalText .= l "`n"
    return finalText
}
