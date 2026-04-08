/*
【SSIルーチン作成補助スクリプト（ダイアログ対応版）】
ファイル名：SSI_Routine_Maker.ahk
*/
TotalDays := 6  ; 繰り返す日数
; 確認ダイアログのタイトル（Window Spyで確認して正確に記入してください）
DialogTitle := "確認" 
/*
■追加機能：
Alt+S のあとに「保存しますか？」などのダイアログが出た場合、
自動的に「y」キーを押して続行します。
*/

#Requires AutoHotkey v2.0

+RButton:: {
    CoordMode("Mouse", "Client")

    ; 1. コントロール取得
    MouseGetPos(,, &targetWin, &targetClassNN, 2)
    if (targetClassNN = "") {
        MsgBox("コントロールが見つかりません。", "エラー", "Icon!")
        return
    }

    ; 2. 座標取得と計算
    try {
        ControlGetPos(&cX, &cY, &cW, &cH, targetClassNN, targetWin)
    } catch {
        MsgBox("座標取得に失敗しました。", "エラー", "Icon!")
        return
    }

    targetX := Integer(cX + (cW / 3))
    targetY := Integer(cY + cH + (cH / 3))

    ; 3. 開始
    Click("Right")
    Sleep(400)

    ; 4. ループ処理
    Loop TotalDays {
        a := A_Index 

        Send("c")
        Sleep(200)

        ; --- Alt+S と ダイアログ処理 ---
        Send("!s")
        
        ; ダイアログが出るのを最大1.0秒待つ
        if WinWait(DialogTitle,, 1.0) {
            Send("y")    ; ダイアログが出たら y を押す
            Sleep(300)   ; 閉じるのを待つ
        }
        Sleep(400)       ; 次の動作への待機

        ; 計算済みの座標で右クリック
        Click(targetX, targetY, "Right")
        Sleep(400)

        Send("{Down 3}{Enter}")
        Sleep(400)

        Send("{Down " . a . "}")
        Sleep(200)
        Send("!s")
        
        ; ここでもダイアログが出る可能性がある場合は、同様に追記可能です
        if WinWait(DialogTitle,, 1.0) {
            Send("y")
            Sleep(300)
        }
        Sleep(600)

        Click(cX, cY, "Right")
        Sleep(400)
    }

    MsgBox(TotalDays "回分の処理が完了しました。", "完了", "Iconi")
}

Esc::ExitApp
