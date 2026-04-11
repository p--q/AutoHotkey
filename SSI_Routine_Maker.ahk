/*
【SSIルーチン作成補助スクリプト】
バージョン：2.2 (マウス位置基準・最適化版)
*/

#Requires AutoHotkey v2.0

; --- 設定セクション ---
TotalDays   := 6
DialogTitle := "確認"
MenuClass   := "ahk_class Windows.Forms10.Window.20808.app."
BtnConfirm  := "確定(&S)"
DateWinTitle := "基準日から何日前後に登録するか選択"

SetTitleMatchMode(1) ; 前方一致

; --------------------------------------------------

+RButton:: {
    CoordMode("Mouse", "Client")

    ; 1. 今マウスがある座標 (mX, mY) と、コントロールの「ハンドル」を取得
    ; targetClassNN（名前）の代わりに targetHwnd（ハンドル）を取得
    MouseGetPos(&mX, &mY, &targetWin, &targetHwnd, 2)
    
    if (targetHwnd = 0) {
        return
    }

    ; 2. コントロールの「高さ(cH)」だけを取得する
    try {
        ControlGetPos(,,, &cH, targetHwnd, targetWin)
    } catch {
        MsgBox("コントロールの情報が取得できませんでした。")
        return
    }

    ; 複製先の座標を計算（今マウスがある位置から、枠の高さ分だけ下へ）
    ; ※少し余裕を持たせるために (cH / 3) を足していますが、不要なら + cH だけでOKです。
    targetX := mX
    targetY := Integer(mY + cH + (cH / 3))

    Loop TotalDays {
        a := A_Index 

        ; --- 1. 元の薬剤（今のマウス位置）を右クリック ---
        Loop 10 {
            ; 座標 mX, mY（さっき Shift+右クリック した場所）をクリック
            Click(mX, mY, "Right") 
            if WinWait(MenuClass, , 0.5)
                break
            if (A_Index = 10) {
                MsgBox("メニューが出現しませんでした。")
                return
            }
        }
        
        Send("c") ; 複製ショートカット

        ; --- 2. 確定(&S)ボタンが出現するまで待機して送信 ---
        Loop 20 {
            try {
                if ControlGetVisible(BtnConfirm, "ahk_id " . targetWin) {
                    Send("!s")
                    break
                }
            }
            Sleep(100)
        }

        ; --- 3. 確認ダイアログが出た場合は 'y' を送る ---
        if WinWait(DialogTitle, , 0.8) {
            Send("y")
            WinWaitClose(DialogTitle, , 1.0)
        }

        ; --- 4. 複製された薬剤（下の行）を右クリック ---
        Loop 10 {
            ; 計算した targetX, targetY（mY + cH）をクリック
            Click(targetX, targetY, "Right")
            if WinWait(MenuClass, , 0.5)
                break
        }

        ; --- 5. メニュー選択（下3回 ＞ Enter） ---
        Send("{Down 3}{Enter}")

        ; --- 6. 日付選択ウィンドウが出現するまで待機 ---
        if WinWait(DateWinTitle, , 2.0) {
            WinActivate(DateWinTitle)
            Sleep(100)
            Send("{Down " . a . "}")
            Send("!s")
            WinWaitClose(DateWinTitle, , 2.0)
        }
        
        Sleep(200)
    }

    MsgBox(TotalDays "日分の複製が完了しました。", "完了")
}

Esc::ExitApp
