/*
【SSIルーチン作成補助スクリプト】
ファイル名：SSI_Routine_Maker.ahk

■設定（ここを変更してください）：
*/
TotalDays := 6  ; 繰り返す日数をここで設定できます
/*

■概要：
点滴のルーチンを指定日数分（a=1〜設定値）自動作成します。
対象のコントロール（PosA）を起点に、右クリックメニューの操作と
日付選択（下矢印キーによる移動）を繰り返します。

■使い方：
1. SSI電子カルテの対象となるボタンや入力欄の上にマウスを置く。
2. 「Shiftキー」を押しながら「右クリック」でスクリプト開始。
3. ループが終了するまでマウスやキーボードを操作せずにお待ちください。
4. 緊急停止したい場合は「Escキー」を押してください。
*/

#Requires AutoHotkey v2.0

; Shift + 右クリックで開始
+RButton:: {
    ; 座標モードをClientに設定
    CoordMode("Mouse", "Client")

    ; 1. マウス直下のコントロール(ClassNN)とウィンドウハンドルを取得
    ; 第5引数「2」でClassNNを取得
    MouseGetPos(,, &targetWin, &targetClassNN, 2)

    if (targetClassNN = "") {
        MsgBox("【エラー】`nマウスの下にコントロール（ClassNN）が見つかりません。`n対象のボタンや入力欄の上にマウスを置いて実行してください。", "スクリプトエラー", "Icon!")
        return
    }

    ; 2. PosA（コントロールの座標とサイズ）を取得
    try {
        ControlGetPos(&cX, &cY, &cW, &cH, targetClassNN, targetWin)
    } catch {
        MsgBox("【エラー】`nコントロール '" targetClassNN "' の座標取得に失敗しました。`nウィンドウが最小化されていないか確認してください。", "スクリプトエラー", "Icon!")
        return
    }

    ; 3. 現在のマウス位置（PosA）で右クリック
    Click("Right")
    Sleep(400) ; メニュー表示待機

    ; 4. ループ処理（a=1 から TotalDays まで）
    Loop TotalDays {
        a := A_Index ; A_Indexは1から始まるループカウンタ

        ; キー入力: c
        Send("c")
        Sleep(200)

        ; キー入力: Alt + s
        Send("!s")
        Sleep(400)

        ; 指定座標で右クリック
        ; x: PosA.x + PosA.w/3
        ; y: PosA.y + PosA.h + PosA.h/3
        targetX := Integer(cX + (cW / 3))
        targetY := Integer(cY + cH + (cH / 3))
        
        Click(targetX, targetY, "Right")
        Sleep(400)

        ; 下矢印キーを3回押してEnter
        Send("{Down 3}{Enter}")
        Sleep(400)

        ; 下矢印キーを a 回押して Alt + s
        Send("{Down " . a . "}")
        Sleep(200)
        Send("!s")
        Sleep(600) ; SSI側の登録処理を待機

        ; PosA（コントロールの左上座標）へ移動して右クリック
        Click(cX, cY, "Right")
        Sleep(400)
    }

    MsgBox(TotalDays "回（日分）のルーチン作成が完了しました。", "完了", "Iconi")
}

; 緊急停止ショートカット（Escキー）
Esc::ExitApp
