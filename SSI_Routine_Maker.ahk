/*
 * @title SSI_Routine_Maker.ahk
 * @version 3.6
 * @author Gemini
 * @description 
 * 【SSI専用ルーチン】
 * 1920x1080の画面で最大化したときにセットメニューの点滴を6日分コピーします。
 * コピーしたい点滴のチェックボックスのある枠上で「Shift+右クリック」すると開始します。
 * * 仕組み:
 * 1. 薬剤の1行下の座標をコントロールの高さから自動計算します。
 * 2. コンテキストメニューが出るまで右クリックをリトライします。
 * 3. 確定ボタンをマウス下のウィンドウハンドル(HWND)から全走査して特定します。
 * 4. 日付選択ウィンドウで指定日数分だけDownキーを送り日付をずらします。
 * * 中止したい場合は「Esc」キーを押してください。
 */

#Requires AutoHotkey v2.0

TotalDays := 6

+RButton:: {
    CoordMode("Mouse", "Client")
    
    ; 1. 複製元(src)と複製先(dest)の座標・ハンドルを取得
    pos := GetDrugCoords()
    
    Loop TotalDays {
        currentDay := A_Index
        
        ; --- A. 複製元の薬剤を右クリック（出るまでリトライ） ---
        WaitContextMenu(pos.srcX, pos.srcY)
        Sleep(100)
        Send("c") ; 複製(C)を実行
        
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
Esc::Exit

; --- 以下、機能関数群 ---

GetDrugCoords() {
    ; マウス下のコントロール情報を取得（引数2によりHWNDで取得）
    MouseGetPos(&mX, &mY, &srcWin, &srcClassNN, 2)
    
    if !(srcClassNN) {
        MsgBox("薬剤のコントロールを検出できませんでした。", "SSI_Routine_Maker")
        Exit
    }

    try {
        ; コントロールの高さを取得して、1行下の座標を算出
        ControlGetPos(,,, &cH, srcClassNN, srcWin)
    } catch {
        MsgBox("座標情報の取得に失敗しました。", "SSI_Routine_Maker")
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
        
        ; マウス下のハンドルからクラス名を直接判定（アクティブウィンドウに頼らない）
        MouseGetPos(,, &mHwnd)
        if (mHwnd) {
            try {
                if InStr(WinGetClass(mHwnd), "WindowsForms10.Window.20808")
                    return ; メニュー出現を確認
            }
        }
        Sleep(200)
    }
    MsgBox("コンテキストメニューが表示されませんでした。", "SSI_Routine_Maker")
    Exit
}

EnsureConfirmAndClick() {
    targetBtnText := "確定(&S)"
    btnHwnd := 0
    
    ; --- ステップ1：確定ボタンが「出現する」のを待つ ---
    ; ここでボタンが出るまで WaitContextMenu 後の処理をせき止めます
    startTime := A_TickCount
    Loop 50 {
        MouseGetPos(,, &parentWin)
        if (parentWin) {
            for hCtrl in WinGetControlsHwnd(parentWin) {
                try {
                    if (InStr(ControlGetText(hCtrl), targetBtnText) && ControlGetVisible(hCtrl)) {
                        btnHwnd := hCtrl ; ボタンのハンドルを確保
                        break 2
                    }
                }
            }
        }
        Sleep(100)
        if (A_TickCount - startTime > 5000) {
            MsgBox("確定ボタンが出現しませんでした（コピー処理の遅延）。")
            Exit
        }
    }

    ; --- ステップ2：ボタンが見つかったらクリック(Alt+S) ---
    Sleep(200) ; 出現直後の安定待ち
    Send("!s")

    ; --- ステップ3：確定ボタンが「消える」のを待つ ---
    ; これにより、SSIがコピー処理を完全に終えるまで次の WaitContextMenu を走らせません
    startTime := A_TickCount
    Loop 50 {
        try {
            ; ボタンが非表示になるか、存在しなくなれば「処理完了」とみなす
            if !ControlGetVisible(btnHwnd)
                break
        } catch {
            ; コントロール自体が破棄された場合も「処理完了」
            break
        }
        Sleep(100)
        if (A_TickCount - startTime > 5000) {
            MsgBox("確定ボタンが押されましたが、画面が更新されません。")
            Exit
        }
    }
    
    ; 確定処理後の余韻（SSIの内部バッファ解放待ち）
    Sleep(500) 
}

ConfirmDialogWithY(DialogTitle) {
    ; 標準的なダイアログであればタイトル指定のWinWaitが有効
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
            ; dayOffset分だけDownして確定
            Send("{Down " . dayOffset . "}{Enter}{Enter}")
            
            ; 次のループへ行く前にウィンドウが閉じるのを待つ
            WinWaitClose(dateWinTitle,, 2)
            return
        }
        Sleep(100)
    }
    MsgBox("日付選択ウィンドウが出現しませんでした。", "SSI_Routine_Maker")
    Exit
}
