/*
 * @title SSI_Routine_Maker.ahk
 * @version 4.1
 * @description 
 * 1. 1920x1080画面対応。点滴チェックボックス枠で Shift+右クリック すると
 * 6日分の複製と日付ずらしを自動実行します。
 * 2. 中断ホットキーを Exit に変更。
 * 3. 確定プロセスの異常時は ExecuteFullConfirmation 内で Exit 実行。
 */

#Requires AutoHotkey v2.0

TotalDays := 6

+RButton:: {
    CoordMode("Mouse", "Client")
    
    ; 1. 座標情報を取得（ユーザーのクリック位置を維持）
    pos := GetDrugCoords()
    
    Loop TotalDays {
        currentDay := A_Index
        
        ; --- A. 複製元の薬剤を右クリック ---
        WaitContextMenu(pos.srcX, pos.srcY)
        Sleep(200)
        Send("c") ; 複製(C)
        
        ; --- B. 確定ボタン～確認ダイアログの同期処理 ---
        ; 失敗時は関数内の Exit でルーチン全体が止まります
        ExecuteFullConfirmation("確認")
        
        ; --- C. 複製された薬剤（1行下）を操作 ---
        Sleep(1200) 
        WaitContextMenu(pos.destX, pos.destY)
        
        Sleep(300)
        Send("{Down 3}{Enter}")
        
        ; --- D. 日付変更ウィンドウ処理 ---
        ChangeDate(currentDay)
        
        Sleep(1000)
    }
    
    MsgBox(TotalDays "日分の複製が完了しました。", "SSI_Routine_Maker", "Iconi")
}

; 実行中スレッドのみ中断（スクリプトは常駐継続）
Esc::Exit

; --- 関数群 ---

GetDrugCoords() {
    MouseGetPos(&mX, &mY, &srcWin, &srcClassNN, 2)
    if !(srcClassNN) {
        MsgBox("薬剤コントロール検出失敗")
        Exit
    }
    try {
        ControlGetPos(,,, &cH, srcClassNN, srcWin)
    } catch {
        MsgBox("座標取得失敗")
        Exit
    }
    return {srcX: mX, srcY: mY, destX: mX, destY: mY + cH}
}

ExecuteFullConfirmation(DialogTitle) {
    targetBtnText := "確定(&S)"
    btnHwnd := 0
    
    ; 1. 確定ボタン出現待ち
    startTime := A_TickCount
    Loop 50 {
        MouseGetPos(,, &pWin)
        if (pWin) {
            for hCtrl in WinGetControlsHwnd(pWin) {
                try {
                    if (InStr(ControlGetText(hCtrl), targetBtnText) && ControlGetVisible(hCtrl)) {
                        btnHwnd := hCtrl
                        break 2
                    }
                }
            }
        }
        Sleep(100)
    }
    
    if !btnHwnd {
        MsgBox("確定ボタンが見つかりませんでした。")
        Exit ; ここでスレッドを終了
    }

    ; 2. 確定ボタン押下 ＆ 消失待ち
    Sleep(200)
    Send("!s")
    while (ControlGetVisible(btnHwnd) && A_TickCount - startTime < 8000) {
        Sleep(100)
    }

    ; 3. 確認ダイアログ出現 ＆ 消失待ち
    if WinWait(DialogTitle,, 3) {
        Sleep(500)
        Send("y")
        if !WinWaitClose(DialogTitle,, 5) {
            MsgBox("確認ダイアログが閉じませんでした。")
            Exit ; ここでスレッドを終了
        }
        Sleep(500)
        return ; 成功
    }
    
    MsgBox("確認ダイアログが出現しませんでした。")
    Exit ; ここでスレッドを終了
}

WaitContextMenu(clickX, clickY) {
    Loop 30 {
        Click(clickX, clickY, "Right")
        Sleep(500)
        MouseGetPos(,, &mHwnd)
        if (mHwnd && InStr(WinGetClass(mHwnd), "WindowsForms10.Window.20808"))
            return 
        Sleep(200)
    }
    MsgBox("メニュー表示タイムアウト")
    Exit
}

ChangeDate(dayOffset) {
    dateWinTitle := "基準日から何日前後に登録するか選択"
    if WinWait(dateWinTitle,, 5) {
        Sleep(500)
        Send("{Down " . dayOffset . "}{Enter}{Enter}")
        WinWaitClose(dateWinTitle,, 3)
        return
    }
    MsgBox("日付選択ウィンドウタイムアウト")
    Exit
}
