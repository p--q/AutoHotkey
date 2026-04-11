/*
【右クリックメニュー診断ツール v2.2】
変数展開を修正しました。
*/

#Requires AutoHotkey v2.0

+LButton:: {
    CoordMode("Mouse", "Client")
    
    ; 1. 起動時のメインウィンドウを特定
    MouseGetPos(&mX, &mY, &targetWin)
    mainClass := WinGetClass("ahk_id " . targetWin)
    
    ; 2. スクリプトが右クリックを送る
    Click(mX, mY, "Right")
    
    ; 3. メニュー生成を待つ
    Sleep(500)
    
    ; 4. アクティブなウィンドウの情報を取得
    activeHwnd  := WinActive("A")
    activeClass := WinGetClass("ahk_id " . activeHwnd)
    activeTitle := WinGetTitle("ahk_id " . activeHwnd)
    
    ; 5. 判定
    isDifferent := (activeHwnd != targetWin) ? "YES (別ウィンドウを検知)" : "NO (メイン画面のまま)"

    ; --- 結果を組み立て（v2の式展開フォーマット） ---
    res := "【判定結果】: " . isDifferent . "`n`n"
    res .= "＜検出されたウィンドウ情報＞`n"
    res .= "ID (HWND): " . activeHwnd . "`n"
    res .= "クラス名: " . activeClass . "`n"
    res .= "タイトル: " . (activeTitle = "" ? "(なし)" : activeTitle) . "`n`n"
    res .= "＜比較用：メイン画面の情報＞`n"
    res .= "メインID: " . targetWin . "`n"
    res .= "メインクラス: " . mainClass

    MsgBox(res, "SSIメニュー診断結果")
}

Esc::ExitApp
