/*
【右クリックメニュー診断ツール v2.1】
Shift + 左クリックで起動し、スクリプトが右クリックを行って診断します。
*/

#Requires AutoHotkey v2.0

+LButton:: {
    ; --- 準備 ---
    CoordMode("Mouse", "Client")
    MouseGetPos(&mX, &mY, &targetWin)
    
    ; --- 1. スクリプトの手で右クリック ---
    Click(mX, mY, "Right")
    
    ; --- 2. 出現を待つ（あえて少し長めに待機） ---
    Sleep(500)
    
    ; --- 3. 情報を取得 ---
    activeHwnd  := WinActive("A")
    activeClass := WinGetClass("ahk_id " . activeHwnd)
    activeTitle := WinGetTitle("ahk_id " . activeHwnd)
    
    ; メインウィンドウのクラス名も比較用に取得
    mainClass := WinGetClass("ahk_id " . targetWin)

    ; --- 4. 結果表示 ---
    resultText := "
    (
    【判定成功か？】: " (activeHwnd != targetWin ? "YES (別ウィンドウを検知)" : "NO (メイン画面のまま)") "
    
    ＜検出されたウィンドウ情報＞
    ID (HWND): " activeHwnd "
    クラス名: " activeClass "
    タイトル: " (activeTitle = "" ? "(なし)" : activeTitle) "
    
    ＜比較用：メイン画面の情報＞
    メインクラス: " mainClass "
    
    --- アドバイス ---
    もし「判定成功か？」が NO なら、メニューがウィンドウとして認識されていません。
    もしクラス名がメインと同じなら、ID(HWND)の差分で判定する必要があります。
    )"
    
    MsgBox(resultText, "SSIメニュー診断結果")
}

Esc::ExitApp
