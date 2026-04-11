/*
【SSIルーチン作成補助スクリプト】
ファイル名：SSI_Routine_Maker.ahk
バージョン：2.1 (前方一致・安定化版)
*/

#Requires AutoHotkey v2.0

; --- 設定セクション ---
TotalDays   := 6
DialogTitle := "確認"  ; 先頭が「確認」で始まるウィンドウを待機
; コンテキストメニューのクラス名
MenuClass   := "ahk_class WindowsForms10.Window.20808.app."
; 確定ボタンのテキスト（先頭が「確定(&S)」で始まるボタン）
BtnConfirm  := "確定(&S)"
; 日付選択ウィンドウのタイトル
DateWinTitle := "基準日から何日前後に登録するか選択"

; タイトル・コントロールの一致モードを「1：前方一致」に設定
SetTitleMatchMode(1)

; --------------------------------------------------

+RButton:: {
    CoordMode("Mouse", "Client")

    ; 1. コントロール情報取得
    MouseGetPos(,, &targetWin, &targetClassNN, 2)
    if (targetClassNN = "") {
        MsgBox("コントロール未検出", "SSI_Routine_Maker", "Icon!")
        return
    }

    try {
        ControlGetPos(&cX, &cY, &cW, &cH, targetClassNN, targetWin)
    } catch {
        MsgBox("座標取得失敗", "SSI_Routine_Maker", "Icon!")
        return
    }

    ; 複製先の相対座標計算（1つ下の項目を想定）
    targetX := Integer(cX + (cW / 3))
    targetY := Integer(cY + cH + (cH / 3))

    Loop TotalDays {
        a := A_Index 

        ; --- 1. 元の薬剤を右クリック（メニューが出るまで） ---
        Loop 10 {
            Click(cX, cY, "Right")
            ; 前方一致でクラス名を確認
            if WinWait(MenuClass, , 0.5)
                break
            if (A_Index = 10) {
                MsgBox("メニューが出現しませんでした。", "エラー", "Icon!")
                return
            }
        }
        
        Send("c") ; 複製ショートカット

        ; --- 2. 確定(&S)ボタンが出現するまで待機して送信 ---
        ; 前方一致設定により、「確定(&S)」で始まるコントロールを捕捉
        Loop 20 {
            try {
                if ControlGetVisible(BtnConfirm, "ahk_id " . targetWin) {
                    Send("!s")
                    break
                }
            }
            Sleep(100)
        }

        ; --- 3. 確認ダイアログ（前方一致）が出た場合は 'y' を送る ---
        if WinWait(DialogTitle, , 0.8) {
            Send("y")
            WinWaitClose(DialogTitle, , 1.0)
        }

        ; --- 4. 複製された薬剤を右クリック（メニューが出るまで） ---
        Loop 10 {
            Click(targetX, targetY, "Right")
            if WinWait(MenuClass, , 0.5)
                break
            if (A_Index = 10) {
                MsgBox("複製先のメニューが出現しませんでした。", "エラー", "Icon!")
                return
            }
        }

        ; --- 5. メニュー選択（下3回 ＞ Enter） ---
        Send("{Down 3}{Enter}")

        ; --- 6. 日付選択ウィンドウ（前方一致）が出現するまで待機 ---
        if WinWait(DateWinTitle, , 2.0) {
            WinActivate(DateWinTitle)
            Sleep(100)
            Send("{Down " . a . "}")
            Send("!s")
            WinWaitClose(DateWinTitle, , 2.0)
        }
        
        Sleep(200) ; ループ間の安定化ウェイト
    }

    MsgBox(TotalDays "日分の複製が完了しました。", "完了", "Iconi")
}

Esc::ExitApp
