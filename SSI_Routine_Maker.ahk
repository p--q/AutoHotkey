/*
【SSIルーチン作成補助スクリプト】
ファイル名：SSI_Routine_Maker.ahk
バージョン：1.9.0

■変数の解説：
- SleepAfterC  : 最優先。cキー（レコード作成）後のシステム処理待ち。
- SleepMenu    : 右クリック後、メニューが描画されるまでの待ち。
- SleepAction  : 保存(Alt+S)後や確定(Enter)後、画面が静止するまでの待ち。
- SleepMove    : 矢印キー移動など、画面変化が少ない操作の待ち。
*/

; --- 詳細チューニング変数（ここを調整してください） ---
TotalDays    := 6
DelayKey     := 50     ; キーの押し下げ・間隔 (ms)

SleepAfterC  := 500    ; 「c」キー入力後の待機 (クリティカル)
SleepMenu    := 300    ; 右クリックメニューの描画待ち (やや重い)
SleepAction  := 300    ; 保存(!s)や確定(Enter)後の反映待ち (重い)
SleepMove    := 150    ; 矢印キー(Down)などの移動待ち (軽い)

DialogTitle  := "確認" 
; --------------------------------------------------

#Requires AutoHotkey v2.0

SendMode("Event")
SetKeyDelay(DelayKey, DelayKey)

+RButton:: {
    CoordMode("Mouse", "Client")

    ; 1. コントロール情報取得
    MouseGetPos(,, &targetWin, &targetClassNN, 2)
    if (targetClassNN = "") {
        MsgBox("コントロール未検出", "SSI V1.9.0", "Icon!")
        return
    }

    try {
        ControlGetPos(&cX, &cY, &cW, &cH, targetClassNN, targetWin)
    } catch {
        MsgBox("座標取得失敗", "SSI V1.9.0", "Icon!")
        return
    }

    targetX := Integer(cX + (cW / 3))
    targetY := Integer(cY + cH + (cH / 3))

    ; 2. 初期動作
    Click(cX, cY, "Right")
    Sleep(SleepMenu)
    
    Send("c")
    Sleep(SleepAfterC)

    ; 3. メインループ
    Loop TotalDays {
        a := A_Index 

        ; --- 現在の行を保存 ---
        Send("!s")
        Sleep(SleepAction)
        
        if WinWait(DialogTitle,, 0.5) {
            Send("y")
            Sleep(SleepAction)
        }

        ; --- 次の行（target）を右クリック ---
        Click(targetX, targetY, "Right")
        Sleep(SleepMenu)

        ; --- メニュー選択（下3回 ＞ Enter） ---
        Send("{Down 3}{Enter}")
        Sleep(SleepAction) ; Enter後は画面が変わるためAction

        ; --- 日付選択（下a回） ---
        Send("{Down " . a . "}")
        Sleep(SleepMove)   ; 単なる選択移動なのでMove（短めでOK）
        
        ; --- 保存 ---
        Send("!s")
        if WinWait(DialogTitle,, 0.5) {
            Send("y")
            Sleep(SleepAction)
        }
        
        Sleep(SleepMove)

        ; --- 元の場所（PosA）へ戻る ---
        Click(cX, cY, "Right")
        Sleep(SleepMenu)
        
        Send("c")
        Sleep(SleepAfterC)
    }

    MsgBox(TotalDays "回分の処理が完了しました。", "SSI V1.9.0", "Iconi")
}

Esc::ExitApp
