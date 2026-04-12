/*
 * @title ContextMenuInfoComparer.ahk
 * @version AutoHotkey v2.0
 * @description 右クリック実行前後のウィンドウ・コントロール情報を比較表示します。
 * メニューの出現やフォーカスの変化を検証するためのツールです。
 */

#Requires AutoHotkey v2.0

; ホットキー: Ctrl + Alt + G (実行)
^!g:: {
    ; 1. 座標の記録
    CoordMode("Mouse", "Screen")
    MouseGetPos(&msX, &msY)

    ; 2. 【右クリック前】の情報取得
    infoBefore := GetInfoAt(msX, msY)

    ; 3. アクション（原点移動 → 座標復帰 → 右クリック）
    MouseMove(0, 0, 0)
    Sleep(50)
    Click(msX, msY, "Right")
    Sleep(200) ; メニューが描画され、情報の準備ができるまで待機

    ; 4. 【右クリック後】の情報取得
    infoAfter := GetInfoAt(msX, msY)

    ; 5. 結果の比較表示用テキスト構築
    comparisonText := "Comparison ( [Before] -> [After] )`n"
                    . "Coords: X" msX " Y" msY "`n"
                    . "--------------------------------------------------`n"
                    . "Title:`n  [B] " infoBefore.title "`n  [A] " infoAfter.title "`n`n"
                    . "Class:`n  [B] " infoBefore.class "`n  [A] " infoAfter.class "`n`n"
                    . "Exe:`n    [B] " infoBefore.exe "`n  [A] " infoAfter.exe "`n`n"
                    . "ID:`n     [B] " infoBefore.id "`n  [A] " infoAfter.id "`n`n"
                    . "ClassNN:`n  [B] " infoBefore.classNN "`n  [A] " infoAfter.classNN "`n`n"
                    . "Text:`n   [B] " textMangle(infoBefore.text) "`n  [A] " textMangle(infoAfter.text)

    MsgBox(comparisonText, "Context Menu Comparison")
}

/**
 * 指定した座標にある情報をオブジェクトとして取得する内部関数
 */
GetInfoAt(x, y) {
    ; 指定座標のウィンドウとコントロールを取得
    ; ※MouseGetPosは現在のカーソル位置を優先するため、一時的に移動させて取得
    MouseMove(x, y, 0)
    targetWin := 0
    targetCtrl := 0
    MouseGetPos(,, &targetWin, &targetCtrl, 2)

    res := {title: "（取得不可）", class: "（取得不可）", exe: "（取得不可）", 
            id: "（取得不可）", classNN: "（取得不可）", text: ""}

    if (targetWin != 0) {
        try res.title   := WinGetTitle("ahk_id " targetWin) || "（空）"
        try res.class   := WinGetClass("ahk_id " targetWin)
        try res.exe     := WinGetProcessName("ahk_id " targetWin)
        res.id          := targetWin
    }
    
    if (targetCtrl != 0) {
        try res.classNN := ControlGetClassNN(targetCtrl)
        try res.text    := ControlGetText(targetCtrl)
    }
    
    return res
}

/**
 * テキスト整形関数
 */
textMangle(x) {
    if (x == "")
        return "（取得不可値または空）"
        
    elli := false
    if (pos := InStr(x, "`n"))
        x := SubStr(x, 1, pos - 1), elli := true
    else if (StrLen(x) > 40)
        x := SubStr(x, 1, 40), elli := true
    
    if (elli)
        x .= " (...)"
    return x
}
