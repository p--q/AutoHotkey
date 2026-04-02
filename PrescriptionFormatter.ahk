MergeSpecificPatterns(text) {
    lines := StrSplit(text, "`n", "`r")
    result := []
    
    for line in lines {
        if (line == "")
            continue
        
        ; 最初に見つかった「時」で分ける
        if (RegExMatch(line, "^(\S+?時)(.*)$", &m)) {
            ; 1. 「発熱時」などの最初のマッチを上の行に結合
            if (result.Length > 0) {
                result[result.Length] .= m[1]
            } else {
                result.Push(m[1])
            }
            
            ; 2. 同じ行の残りの部分（例：「1日3回まで。頭痛・咽頭痛時も...」）
            ; これを「時」判定を通さずに、そのまま独立した行として追加する
            remaining := Trim(m[2])
            if (remaining != "") {
                result.Push(remaining)
            }
            ; ※ ここで continue は不要（次の line に進むため）
            
        } else if (RegExMatch(line, "^分\d+\s\d")) {
            line := RegExReplace(line, "^(分\d+)\s(\d)", "$1@@SPACE@@$2")
            result.Push(line)
        } else if (RegExMatch(line, "^外\)\s(.*)$", &m)) {
            if (result.Length > 0) {
                result[result.Length] .= "@@SPACE@@" . m[1]
            } else {
                result.Push("@@SPACE@@" . m[1])
            }
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
