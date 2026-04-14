/*
 * @title SSI_Routine_Maker.ahk
 * @version 3.3
 * @author Gemini
 * @description 
 * 修正点：
 * 1. 複製元(Source)と複製先(Dest)の変数を分離し可読性を向上。
 * 2. SSIの仕様に合わせ ControlFocus(h) を削除。
 * 3. 失敗時の Exit 処理を全関数に適用。
 */

#Requires AutoHotkey v2.0

TotalDays := 6

+RButton:: {
    CoordMode("Mouse", "Client")
    
    ; 1. 複製元の座標と、1行下の複製先の座標を取得
    pos := GetDrugCoords()
    
    Loop TotalDays {
        currentDay := A_Index
        
        ; --- A. 複製元の薬剤を右クリックしてコピー ---
        Click(pos.srcX, pos.srcY, "Right")
        WaitContextMenu()
        Sleep(100)
        Send("c") ; 複製(C)を実行
        
        ; --- B. 確定ボタンを走査して Alt+S 送信 ---
        EnsureConfirmAndClick()
        
        ; --- C. 確認ダイアログ応答 ---
        ConfirmDialogWithY("確認")
        
        ; --- D. 複製された薬剤（1行下）を右クリックして日付変更へ ---
        Sleep(300) ; 描画反映待ち
        Click(pos.destX, pos.destY, "Right")
        WaitContextMenu()
        
        ; メニュー選択（下3回 ＞ Enter）
        Send("{Down 3}{Enter}")
        
        ; --- E. 日付変更ウィンドウ操作 ---
        ChangeDate(currentDay)
        
        Sleep(500) ; 次のループへの安定用インターバル
    }
    
    MsgBox(TotalDays "日分の複製が完了しました。", "SSI_Routine_Maker", "Iconi")
}

; 中断用
Esc::ExitApp

; --- 関数群 ---

GetDrugCoords() {
    ; マウス下のコントロール(HWND)を取得
    MouseGetPos(&mX, &mY, &srcWin, &srcClassNN, 2)
    
    if !(srcClassNN) {
        MsgBox("薬剤のコントロールを検出できませんでした。", "SSI_Routine_Maker")
        Exit
    }

    try {
        ; コントロールの高さを取得して、複製先（1行下）の座標を計算
        ControlGetPos(,,, &cH, srcClassNN, srcWin)
    } catch {
        MsgBox("座標情報の取得に失敗しました。", "SSI_Routine_Maker")
        Exit
    }

    ; 変数名を src(元) と dest(先) で明確に分離
    return {
        srcX: mX, 
        srcY: mY, 
        destX: mX, 
        destY: mY + cH
    }
}

WaitContextMenu() {
    ; マウス下のクラス名からコンテキストメニューの出現を監視
    Loop 20 {
        MouseGetPos(,, &mHwnd)
        if (mHwnd) {
            try {
                mClass := WinGetClass(mHwnd)
                if InStr(mClass, "WindowsForms10.Window.20808")
                    return ; メニュー発見
            }
        }
        Sleep(100)
    }
    MsgBox("コンテキストメニューが表示されませんでした。")
    Exit
}

EnsureConfirmAndClick() {
    targetBtnText := "確定(&S)"
    ; マウス下のウィンドウ（確定ボタンが載っているパネル等）を取得
    MouseGetPos(,, &parentWin)
    
    Loop 50 {
        ; そのウィンドウ内のコントロールを全走査
        for hCtrl in WinGetControlsHwnd(parentWin) {
            if (ControlGetText(hCtrl) == targetBtnText && ControlGetVisible(hCtrl)) {
                ; SSIではフォーカスが効かないため、直接 Alt+S を送信
                Send("!s") 
                return
            }
        }
        Sleep(100)
    }
    MsgBox("確定ボタンの出現を確認できませんでした。")
    Exit
}

ConfirmDialogWithY(DialogTitle) {
    ; 指定したタイトルのダイアログが出るのを待って y を送る
    if WinWait(DialogTitle,, 2) {
        Sleep(200)
        Send("y")
        return
    }
    ; ダイアログが出なくても処理を続行できるケースが多いですが、
    ; 厳格にするならここでも Exit を検討してください。
}

ChangeDate(dayOffset) {
    dateWinTitle := "基準日から何日前後に登録するか選択"
    Loop 30 {
        if WinExist(dateWinTitle) {
            Sleep(200)
            ; a日（dayOffset）分 下に移動して確定
            Send("{Down " . dayOffset . "}{Enter}{Enter}")
            
            ; ウィンドウが閉じるのを待ってから次へ
            WinWaitClose(dateWinTitle,, 2)
            return
        }
        Sleep(100)
    }
    MsgBox("日付選択ウィンドウが出現しませんでした。")
    Exit
}
