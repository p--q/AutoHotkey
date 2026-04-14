/*
 * @title ContextMenuMethodChecker.ahk
 * @version 1.0.0
 * @author Gemini
 * @description 
 * 右クリックで出現したコンテキストメニューに対し、複数の取得アプローチ
 * （アクティブ判定、クラス名指定、座標指定、コントロール直接取得）を試行し、
 * メニュー情報をどのように捕捉できるかを検証します。
 */

#Requires AutoHotkey v2.0

^!f:: {
    ; --- 1. アクション：座標取得と右クリックの実行 ---
    CoordMode("Mouse", "Screen")
    MouseGetPos(&msX, &msY)

    ; 一旦原点へジャンプ（干渉防止）
    MouseMove(0, 0, 0)
    Sleep(50)

    ; 元の場所に戻って右クリック
    MouseMove(msX, msY, 0)
    Click("Right")
    
    ; メニューが描画されるのを少し待機
    Sleep(300)

    report := "【コンテキストメニュー検証レポート】`n"
    report .= "実行座標: X" msX " Y" msY "`n"
    report .= "----------------------------------`n`n"

    ; --- 2. 検証：方法A (現在のアクティブウィンドウを取得) ---
    ; メニューが出現すると、メニュー自体がアクティブになることが多いです
    activeWin := WinExist("A")
    aTitle := WinGetTitle(activeWin)
    aClass := WinGetClass(activeWin)
    report .= "A. アクティブウィンドウ判定:`n"
    report .= "   Title: " (aTitle = "" ? "(無題)" : aTitle) "`n"
    report .= "   Class: " aClass "`n`n"

    ; --- 3. 検証：方法B (標準メニュークラス #32768 を狙い撃ち) ---
    ; Windows標準の右クリックメニューは通常このクラス名を持ちます
    resB := "❌ 見つかりません"
    if WinExist("ahk_class #32768") {
        hMenu := WinExist("ahk_class #32768")
        resB := "✅ 発見 (Hwnd: " hMenu ")"
    }
    report .= "B. 標準メニュークラス(#32768):`n   " resB "`n`n"

    ; --- 4. 検証：方法C (座標直下のウィンドウ/コントロール取得) ---
    ; 右クリックしたその点に「今」何があるかを確認
    mWin := 0
    mCtrl := 0
    MouseGetPos(,, &mWin, &mCtrl, 2) ; 2 = プレビューではなく実ハンドル取得
    
    cClassNN := "❌ 取得不可"
    try cClassNN := ControlGetClassNN(mCtrl)
    
    report .= "C. 座標直下のオブジェクト:`n"
    report .= "   WinID: " mWin "`n"
    report .= "   Control: " cClassNN "`n`n"

    ; --- 5. 検証：方法D (メニュー内のテキスト取得試行) ---
    ; 特定のテキスト（例:「コピー」や「貼り付け」など）がコントロールとして見えるか
    ; ここでは検証として、何らかのテキストが取得できるか試みます
    resD := "❌ テキスト取得不可"
    try {
        if mCtrl {
            txt := ControlGetText(mCtrl)
            if (txt != "")
                resD := "✅ 成功: " (StrLen(txt) > 20 ? SubStr(txt, 1, 20) "..." : txt)
        }
    }
    report .= "D. コントロールテキスト取得:`n   " resD "`n`n"

    ; --- 結果の表示 ---
    ; メニューが消えないように、MsgBoxの前に少し情報を整理
    MsgBox(report, "ContextMenu Checker v1.0.0")
}
