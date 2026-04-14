/*
 * @title SSI_Routine_Maker.ahk
 * @version 3.5
 * @author Gemini
 * @description 
 * 1. アクティブウィンドウ取得("A")を使わず、マウス下のハンドル(parentWin)のみを使用。
 * 2. メニューが出るまで右クリックをリトライ。
 * 3. 確定ボタンが見つからない、またはウィンドウが出ない場合は即座に Exit。
 */

#Requires AutoHotkey v2.0

TotalDays := 6

+RButton:: {
    CoordMode("Mouse", "Client")
    
    ; 1. 座標とハンドルの初期取得
    pos := GetDrugCoords()
    
    Loop TotalDays {
        currentDay := A_Index
        
        ; --- A. 複製元の薬剤を右クリック（出るまでリトライ） ---
        WaitContextMenu(pos.srcX, pos.srcY)
        Sleep(100)
        Send("c") ; 複製(C)
        
        ; --- B. 確定ボタンを走査（マウス下のウィンドウを基準にする） ---
        EnsureConfirmAndClick()
        
        ; --- C. 確認ダイアログ応答 ---
        ConfirmDialogWithY("確認")
        
        ; --- D. 複製された薬剤（1行下）を右クリック（出るまでリトライ） ---
        Sleep(400) 
        WaitContextMenu(pos.destX, pos.destY)
        
        ; メニュー選択（下3回 ＞ Enter）
        Send("{Down 3}{Enter}")
        
        ; --- E. 日付変更処理 ---
        ChangeDate(currentDay)
        
        Sleep(600)
    }
    
    MsgBox(TotalDays "日分の複製が完了しました。", "SSI_Routine_Maker", "Iconi")
}

; 緊急停止
Esc::ExitApp

; --- 関数群 ---

GetDrugCoords() {
    ; マウス下のハンドルを取得（引数2でHWNDを取得）
    MouseGetPos(&mX, &mY, &srcWin, &srcClassNN, 2)
    
    if !(srcClassNN) {
        MsgBox("薬剤コントロールを検出できません。")
        Exit
    }

    try {
        ControlGetPos(,,, &cH, srcClassNN, srcWin)
    } catch {
        MsgBox("座標情報の取得に失敗。")
        Exit
    }

    return {
        srcX: mX, 
        srcY: mY, 
        destX: mX, 
        destY: mY + cH
    }
}

WaitContextMenu(clickX, clickY) {
    Loop 20 {
        Click(clickX, clickY, "Right")
        Sleep(300)
        
        ; マウス下のハンドルからクラスを直接判定
        MouseGetPos(,, &mHwnd)
        if (mHwnd) {
            try {
                if InStr(WinGetClass(mHwnd), "WindowsForms10.Window.20808")
                    return 
            }
        }
        Sleep(200)
    }
    MsgBox("メニューが表示されませんでした。")
    Exit
}

EnsureConfirmAndClick() {
    targetBtnText := "確定(&S)"
    
    Loop 50 { ; 最大5秒
        ; ★SSI対策: アクティブウィンドウ("A")ではなく、マウス下のウィンドウを取得
        MouseGetPos(,, &parentWin)
        
        if (parentWin) {
            ; parentWinの子コントロールを全走査
            for hCtrl in WinGetControlsHwnd(parentWin) {
                try {
                    if (InStr(ControlGetText(hCtrl), targetBtnText) && ControlGetVisible(hCtrl)) {
                        Send("!s")
                        return
                    }
                }
            }
        }
        Sleep(100)
    }
    MsgBox("確定ボタンが見つかりませんでした。")
    Exit
}

ConfirmDialogWithY(DialogTitle) {
    ; ここはシステムメッセージボックスの場合、WinWaitが効く可能性があります。
    ; もし効かない場合はWinExist(DialogTitle)のLoopに書き換えますが、一旦WinWaitで保持します。
    if WinWait(DialogTitle,, 1.5) {
        Sleep(200)
        Send("y")
    }
}

ChangeDate(dayOffset) {
    dateWinTitle := "基準日から何日前後に登録するか選択"
    Loop 30 {
        ; WinExistでもタイトル指定であればハンドルを介さず捕捉できることが多いです
        if WinExist(dateWinTitle) {
            Sleep(300)
            Send("{Down " . dayOffset . "}{Enter}{Enter}")
            WinWaitClose(dateWinTitle,, 2)
            return
        }
        Sleep(100)
    }
    MsgBox("日付選択ウィンドウが出現しませんでした。")
    Exit
}
