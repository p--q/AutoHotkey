/*
【SSIルーチン作成補助スクリプト】
ファイル名：SSI_Routine_Maker.ahk
バージョン：3.1

1920x1080の画面で最大化したときにセットメニューの点滴を6日分コピーする。
コピーしたい点滴のチェックボックスのある枠でShift+右クリックで開始。
中止はESC

■変数の解説：
- SleepAfterC  : 最優先。cキー（レコード作成）後のシステム処理待ち。
- SleepMenu    : 右クリック後、メニューが描画されるまでの待ち。
- SleepAction  : 保存(Alt+S)後や確定(Enter)後、画面が静止するまでの待ち。
- SleepMove    : 矢印キー移動など、画面変化が少ない操作の待ち。
*/

; --- 詳細チューニング変数（ここを調整してください） ---
TotalDays    := 6


;SleepAfterC  := 500    ; 「c」キー入力後の待機 (クリティカル)
SleepMenu    := 300    ; 右クリックメニューの描画待ち (やや重い)
SleepAction  := 300    ; 保存(!s)や確定(Enter)後の反映待ち (重い)
SleepMove    := 150    ; 矢印キー(Down)などの移動待ち (軽い)

;DialogTitle  := "確認" 
; --------------------------------------------------

#Requires AutoHotkey v2.0

+RButton:: {
    CoordMode("Mouse", "Client")
    pos := GetTargetCoords()  ; mX, mY: 複製元薬剤の座標、targetX, targetY: 複製した薬剤の座標
    Loop TotalDays {
        a := A_Index 
        Click(pos.mX, pos.mY, "Right")  ; コピー元の薬剤のコンテクストメニューを表示
        WaitContextMenu()  ; コンテクストメニューの出現を確認。
        Send("c")  ; 複製
        EnsureConfirm()  ; 確定ボタンの出現を待って確定する
        ConfirmDialogWithY("確認")  ; 引数のタイトルのダイアログのYボタンを押す。
        ; 複製した薬剤（target）を右クリック 
        Click(pos.targetX, pos.targetY, "Right")
        WaitContextMenu()
        ; --- メニュー選択（下3回 ＞ Enter） ---
        Send("{Down 3}{Enter}")
        ChangeDate(a)  ; a日ずらす    
        Sleep(SleepMove)
    }
    MsgBox(TotalDays "日分の複製が完了しました。", "SSI_Routine_Maker", "Iconi")
}

Esc::Exit

; --- 右クリック後の出現監視関数 ---
WaitContextMenu() {
    Loop 20 { ; 最大2秒程度待機
        MouseGetPos(,, &mHwnd)
        if (mHwnd) {
            try {
                mClass := WinGetClass(mHwnd)
                if InStr(mClass, "WindowsForms10.Window.20808.app.") {
                    return
                }
            }
        }
        Sleep(100)
    }
    MsgBox("コンテクストメニューの取得できませんでした")
    Exit ; 処理を中断して終了
}


GetTargetCoords() {
    MouseGetPos(&mX, &mY, &targetWin, &targetClassNN, 2) ; マウスカーソルの座標とその下のコントロールの属性を取得
    if (targetClassNN = "") {
        MsgBox("コントロール未検出", "SSI_Routine_Maker", "Icon!")
        Exit
    }
    try {
        ControlGetPos(,,, &cH, targetClassNN, targetWin) ; 属性を渡してマウスカーソル下のコントロールの高さを取得
    } catch {
        MsgBox("座標取得失敗", "SSI_Routine_Maker", "Icon!")
        Exit
    }
    return {mX: mX, mY: mY, targetX: mX, targetY: mY+cH}
}

EnsureConfirm() {  ; 確定ボタンの出現を待って確定する
    ; 最大5秒間（0.1秒 × 50回）「確定」ボタンを監視
    Loop 50 {
        ; SSIの内部名に合わせて "確定" で判定
        if ControlGetVisible("確定", "A") {
            Sleep(50)  ; 描画の安定待ち
            Send("!s") ; 確定する
            Sleep(SleepAction)
            return
        }
        Sleep(100)
    }
    MsgBox("確定ボタンが見つかりませんでした（複製失敗）")
    Exit ; 処理を中断して終了
}

ConfirmDialogWithY(DialogTitle) {  ; 引数のタイトルのダイアログのYボタンを押す。
    if WinWait(DialogTitle,, 0.5) {
        Send("y")
        Sleep(SleepAction)
    }
}

ChangeDate(a) {  ; a日ずらす    
    ; --- 日付選択ウィンドウの監視 ---
    dateWinTitle := "基準日から何日前後に登録するか選択"
    ; 最大3秒間（0.1秒 × 30回）ウィンドウの出現を待つ
    Loop 30 {
        if WinExist(dateWinTitle) {
            Send("{Down " . a . "}")
            Sleep(SleepMove)
            Send("!s")
            return
        }
        Sleep(100)
    }
    ; ウィンドウが出現しなかったら終了
    MsgBox("日付選択ウィンドウが出現しませんでした。処理を終了します。")
    Exit
}
