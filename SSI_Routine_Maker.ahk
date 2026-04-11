/*
【SSIルーチン作成補助スクリプト】
バージョン：2.3 (クラス名訂正・座標最適化版)
*/

#Requires AutoHotkey v2.0

; --- 設定セクション ---
TotalDays   := 6
DialogTitle := "確認"
; 訂正後のクラス名
MenuClass   := "ahk_class WindowsForms10.Window.20808.app."
BtnConfirm  := "確定(&S)"
DateWinTitle := "基準日から何日前後に登録するか選択"

; 確実にマッチさせるため、一時的に部分一致(2)を使用
SetTitleMatchMode(2)

; --------------------------------------------------

+RButton:: {
    CoordMode("Mouse", "Client")

    ; 1. 現在のマウス位置 (mX, mY) とハンドルを取得
    MouseGetPos(&mX, &mY, &targetWin, &targetHwnd, 2)
    
    if (targetHwnd = 0)
        return

    ; 2. コントロールの高さ(cH)のみ取得
    try {
        ControlGetPos(,,, &cH, targetHwnd, targetWin)
    } catch {
        MsgBox("コントロールの情報が取得できませんでした。")
        return
    }

    ; 3. 複製先の座標計算（シンプルに1行分下へ）
    targetX := mX
    targetY := mY + cH

    Loop TotalDays {
        a := A_Index 

        ; --- 1. 元の薬剤を右クリック ---
        Loop 10 {
            Click(mX, mY, "Right")
            ; 訂正されたクラス名で出現を待機
            if WinWait(MenuClass, , 0.5)
                break
            if (A_Index = 10) {
                MsgBox("メニュー（" MenuClass "）が出現しませんでした。", "エラー")
                return
            }
        }
        
        Send("c") ; 複製

        ; --- 2. 確定(&S)ボタン出現待ち ---
        Loop 20 {
            try {
                if ControlGetVisible(BtnConfirm, "ahk_id " . targetWin) {
                    Send("!s")
                    break
                }
            }
            Sleep(100)
        }

        ; --- 3. 確認ダイアログ処理 ---
        if WinWait(DialogTitle, , 0.8) {
            Send("y")
            WinWaitClose(DialogTitle, , 1.0)
        }

        ; --- 4. 複製された薬剤（1行下）を右クリック ---
        Loop 10 {
            Click(targetX, targetY, "Right")
            if WinWait(MenuClass, , 0.5)
                break
        }

        ; --- 5. メニュー選択（下3回 ＞ Enter） ---
        Send("{Down 3}{Enter}")

        ; --- 6. 日付選択ウィンドウ待ち ---
        if WinWait(DateWinTitle, , 2.0) {
            WinActivate(DateWinTitle)
            Sleep(150) ; ウィンドウ活性化のための微小な待ち
            Send("{Down " . a . "}")
            Send("!s")
            WinWaitClose(DateWinTitle, , 2.0)
        }
        
        Sleep(200)
    }

    MsgBox(TotalDays "日分の複製が完了しました。", "完了")
}

Esc::ExitApp
