/*
 * @title ControlMethodChecker.ahk
 * @version 1.0.0
 * @author Gemini
 * @description 
 * マウス下のウィンドウに対して、特定のコントロール（「確定(&S)」ボタン等）を
 * 取得するための3つのアプローチ（Exe走査・ID走査・直接取得）を同時検証します。
 * どの方法が最も確実にターゲットを捕捉できるかを判断するための診断ツールです。
 */

#Requires AutoHotkey v2.0

; ホットキー: Ctrl + Alt + F (Find)
^!f:: {
    ; --- 1. 準備：マウス位置の情報取得 ---
    CoordMode("Mouse", "Screen")
    MouseGetPos(,, &msWin)
    if !msWin {
        MsgBox("ウィンドウが見つかりません。対象のアプリの上にマウスを置いてください。")
        return
    }

    ; プロセス名の取得
    try {
        exeName := WinGetProcessName("ahk_id " msWin)
    } catch {
        exeName := "不明"
    }

    ; 検索対象のテキスト（適宜書き換えてください）
    targetText := "確定(&S)"
    
    report := "【検証レポート】`n"
    report .= "対象プロセス: " exeName "`n"
    report .= "対象テキスト: " targetText "`n"
    report .= "----------------------------------`n`n"

    ; --- 2. 検証：方法1 (ahk_exe でプロセス全体の全コントロールを走査) ---
    res1 := "❌ 失敗"
    try {
        for h in WinGetControlsHwnd("ahk_exe " exeName) {
            if (ControlGetText(h) == targetText && ControlGetVisible(h)) {
                res1 := "✅ 成功 (Hwnd: " h ")"
                break
            }
        }
    }
    report .= "1. ahk_exe で全走査:`n   " res1 "`n`n"

    ; --- 3. 検証：方法2 (ahk_id でマウス下のウィンドウ内のみを全走査) ---
    res2 := "❌ 失敗"
    try {
        for h in WinGetControlsHwnd("ahk_id " msWin) {
            if (ControlGetText(h) == targetText && ControlGetVisible(h)) {
                res2 := "✅ 成功 (Hwnd: " h ")"
                break
            }
        }
    }
    report .= "2. ahk_id で全走査:`n   " res2 "`n`n"

    ; --- 4. 検証：方法3 (ControlGetHwnd で一撃の直接取得) ---
    res3 := "❌ 失敗"
    try {
        h := ControlGetHwnd(targetText, "ahk_id " msWin)
        if ControlGetVisible(h)
            res3 := "✅ 成功 (Hwnd: " h ")"
        else
            res3 := "⚠️ 存在公認 (ただし不可視)"
    } catch {
        res3 := "❌ 失敗 (エラー発生)"
    }
    report .= "3. 直接取得(ControlGetHwnd):`n   " res3 "`n`n"

    ; --- 5. 結果の表示 ---
    MsgBox(report, "ControlMethodChecker v1.0.0")
}
