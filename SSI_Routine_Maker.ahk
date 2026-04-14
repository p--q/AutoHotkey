/*
 * @title SSI_Routine_Maker.ahk
 * @version 3.4
 * @author Gemini
 * @description 
 * 1. メニューが出るまで右クリックをリトライする機能を追加
 * 2. 確定ボタンをアクティブウィンドウ全体から走査するハイブリッド方式
 * 3. 失敗時は直ちに ExitApp または Exit で処理を中断
 */

#Requires AutoHotkey v2.0

TotalDays := 6

+RButton:: {
    CoordMode("Mouse", "Client")
    
    ; 1. 複製元(src)と複製先(dest)の座標を取得
    pos := GetDrugCoords()
    
    Loop TotalDays {
        currentDay := A_Index
        
        ; --- A. 複製元の薬剤を右クリック（出るまでリトライ） ---
        WaitContextMenu(pos.srcX, pos.srcY)
        Sleep(100)
        Send("c") ; 複製(C)
        
        ; --- B. 確定ボタンを探して Alt+S ---
        EnsureConfirmAndClick()
        
        ; --- C. 確認ダイアログ応答 ---
        ConfirmDialogWithY("確認")
        
        ; --- D. 複製された薬剤（1行下）を右クリック（出るまでリトライ） ---
        Sleep(400) ; 描画安定のための待ち
        WaitContextMenu(pos.destX, pos.destY)
        
        ; メニュー選択（下3回 ＞ Enter）
        Send("{Down 3}{Enter}")
        
        ; --- E. 日付変更処理 ---
        ChangeDate(currentDay)
        
        Sleep(600) ; 次のループへのインターバル
    }
    
    MsgBox(TotalDays "日分の複製が完了しました。", "SSI_Routine_Maker", "Iconi")
}

; 緊急停止
Esc::ExitApp

; --- 以下、機能関数 ---

GetDrugCoords() {
    ; マウス下のハンドル取得
    MouseGetPos(&mX, &mY, &srcWin, &srcClassNN, 2)
    
    if !(srcClassNN) {
        MsgBox("薬剤コントロールが検出できませんでした。")
        Exit
    }

    try {
        ; コントロールの高さを取得
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
    Loop 20 { ; 最大20回リトライ
        Click(clickX, clickY, "Right")
        Sleep(300) ; メニュー描画待ち
        
        MouseGetPos(,, &mHwnd)
        if (mHwnd) {
            try {
                mClass := WinGetClass(mHwnd)
                ; WindowsForms系のメニュークラスを検知
                if InStr(mClass, "WindowsForms10.Window.20808")
                    return 
            }
        }
        Sleep(200)
    }
    MsgBox("コンテキストメニューを表示できませんでした。")
    Exit
}

EnsureConfirmAndClick() {
    targetBtnText := "確定(&S)"
    
    Loop 50 { ; 最大5秒
        ; アクティブウィンドウをターゲットにする
        activeWin := WinExist("A")
        
        if (activeWin) {
            for hCtrl in WinGetControlsHwnd(activeWin) {
                try {
                    txt := ControlGetText(hCtrl)
                    ; 部分一致で判定
                    if (InStr(txt, "確定") && ControlGetVisible(hCtrl)) {
                        Send("!s")
                        return
                    }
                }
            }
        }
        
        ; 念のためマウス下のウィンドウも補完的に走査
        MouseGetPos(,, &mWin)
        if (mWin && mWin != activeWin) {
            for hCtrl in WinGetControlsHwnd(mWin) {
                try {
                    if (InStr(ControlGetText(hCtrl), "確定") && ControlGetVisible(hCtrl)) {
                        Send("!s")
                        return
                    }
                }
            }
        }
        Sleep(100)
    }
    MsgBox("確定ボタンを見失いました。")
    Exit
}

ConfirmDialogWithY(DialogTitle) {
    ; ダイアログは出ない場合もあるので Exit はせずタイムアウトで次へ
    if WinWait(DialogTitle,, 1.5) {
        Sleep(200)
        Send("y")
    }
}

ChangeDate(dayOffset) {
    dateWinTitle := "基準日から何日前後に登録するか選択"
    Loop 30 {
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
