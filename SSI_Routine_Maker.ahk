/*
 * @title SSI_Routine_Maker.ahk
 * @version 5.5
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
global lapData := []

+RButton:: {
    CoordMode("Mouse", "Screen")
    lapData.Length := 0
    
    ; 1. 座標情報を取得
    pos := GetDrugCoords()
    
    Loop TotalDays {
        currentDay := A_Index
        ; 試行回数(Try)と時間(ms)を記録
        currentLap := { contextTry:0, contextMs:0, btnAppTry:0, btnAppMs:0, diagCloseMs:0, dateWinTry:0, dateWinMs:0 }
        
        ; --- A. 複製元の薬剤を右クリック ---
        t1 := A_TickCount
        resA := WaitContextMenu(pos.srcX, pos.srcY)
        currentLap.contextMs := A_TickCount - t1
        currentLap.contextTry := resA.tries
        
        Sleep(100)
        Send("c") ; 複製(C)
        
        ; --- B & C. 確定ボタン走査 ＆ 確認ダイアログ応答 ---
        resBC := EnsureConfirmAndClick()
        currentLap.btnAppMs := resBC.btnAppMs
        currentLap.btnAppTry := resBC.btnAppTry
        currentLap.diagCloseMs := resBC.diagCloseMs
        
        ; --- D. 複製された薬剤（1行下）を右クリック ---
        Sleep(400) 
        t2 := A_TickCount
        resD := WaitContextMenu(pos.destX, pos.destY)
        currentLap.contextMs += (A_TickCount - t2)
        currentLap.contextTry += resD.tries
        
        Send("{Down 3}{Enter}")
        
        ; --- E. 日付変更処理 ---
        t3 := A_TickCount
        resE := ChangeDate(currentDay)
        currentLap.dateWinMs := A_TickCount - t3
        currentLap.dateWinTry := resE.tries
        
        lapData.Push(currentLap)
        Sleep(600)
    }
    
    ; 統計の表示
    L1 := lapData[1]
    L6 := lapData[TotalDays]
    
    res := "【1回目 vs " TotalDays "回目 負荷比較レポート】`n`n"
    res .= "項目 [試行回数 / 所要時間]`t1回目`t" TotalDays "回目`n"
    res .= "----------------------------------------------------------------------`n"
    res .= "右クリック合計:`t[" L1.contextTry "回 / " L1.contextMs "ms]`t[" L6.contextTry "回 / " L6.contextMs "ms]`n"
    res .= "確定ボタン出現:`t[" L1.btnAppTry "回 / " L1.btnAppMs "ms]`t[" L6.btnAppTry "回 / " L6.btnAppMs "ms]`n"
    res .= "ダイアログ消失:`t[ -- / " L1.diagCloseMs "ms]`t[ -- / " L6.diagCloseMs "ms]`n"
    res .= "日付窓出現:`t`t[" L1.dateWinTry "回 / " L1.dateWinMs "ms]`t[" L6.dateWinTry "回 / " L6.dateWinMs "ms]`n"
    res .= "----------------------------------------------------------------------`n"
    res .= "※試行回数が 1 を超える場合は、SSIの描画待ちが発生しています。"
    
    MsgBox(res, "SSI詳細パフォーマンス統計 (v5.5)", "Iconi")
}

; 緊急停止
Esc::Exit

; --- 以下、機能関数群 ---

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

WaitContextMenu(clickX, clickY
