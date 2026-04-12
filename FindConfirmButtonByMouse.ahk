/*
 * @title FindConfirmButtonByMouse.ahk
 * @version AutoHotkey v2.0
 * @description マウス下のプロセスの全コントロールを走査し、「確定(&S)」ボタンの可視状態を確認します。
 */

#Requires AutoHotkey v2.0

; ホットキー: Ctrl + Alt + F (Find)
^!f:: {
    ; 1. マウス下のウィンドウからプロセス名（exe）を取得
    CoordMode("Mouse", "Screen")
    MouseGetPos(,, &mouseWin)
    
    if !(mouseWin) {
        MsgBox("ウィンドウが見つかりませんでした。", "エラー")
        return
    }

    try {
        targetExe := WinGetProcessName("ahk_id " mouseWin)
    } catch {
        MsgBox("プロセスの取得に失敗しました。", "エラー")
        return
    }

    ; ターゲット指定用の文字列 (ahk_exe myapp.exe)
    targetCriteria := "ahk_exe " targetExe
    
    ; 2. そのプロセスの全コントロールを走査
    foundHwnd := 0
    confirmText := "確定(&S)"
    
    try {
        ; 指定プロセスのすべてのコントロールハンドルを取得
        for ctrlHwnd in WinGetControlsHwnd(targetCriteria) {
            ; テキストが一致するか
            if (ControlGetText(ctrlHwnd) == confirmText) {
                ; かつ、可視（Visible）状態か
                if (ControlGetVisible(ctrlHwnd)) {
                    foundHwnd := ctrlHwnd
                    break ; 見つかったらループ終了
                }
            }
        }
    }

    ; 3. 結果の表示
    if (foundHwnd) {
        ctrlClass := ControlGetClassNN(foundHwnd)
        MsgBox(
            Format("【ボタン発見】`n`nプロセス: {1}`nテキスト: {2}`nClassNN: {3}`nHwnd: {4}", 
            targetExe, confirmText, ctrlClass, foundHwnd),
            "検索結果"
        )
    } else {
        MsgBox(
            Format("【未発見】`n`nプロセス「{1}」内に、`n可視状態の「{2}」は見つかりませんでした。", 
            targetExe, confirmText),
            "検索結果"
        )
    }
}
