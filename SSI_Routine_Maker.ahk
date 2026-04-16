/*
 * @title SSI_Routine_Maker.ahk
 * @version 5.7
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
    
    pos := GetDrugCoords()
    
    Loop TotalDays {
        currentDay := A_Index
        ; 記録項目をさらに細分化
        currentLap := { srcTry:0, srcMs:0, dstTry:0, dstMs:0, btnAppTry:0, btnAppMs:0, diagCloseMs:0, dateWinTry:0, dateWinMs:0 }
        
        ; --- A. 複製元の薬剤を右クリック (src) ---
        t1 := A_TickCount
        resA := WaitContextMenu(pos.srcX, pos.srcY)
        currentLap.srcMs := A_TickCount - t1
        currentLap.srcTry := resA.tries
        
        Sleep(150)
        Send("c") 
        
        ; --- B & C. 確定ボタン走査 ＆ 確認ダイアログ応答 ---
        resBC := EnsureConfirmAndClick()
        currentLap.btnAppMs := resBC.btnAppMs
        currentLap.btnAppTry := resBC.btnAppTry
        currentLap.diagCloseMs := resBC.diagCloseMs
        
        ; --- D. 複製された薬剤（1行下）を右クリック (dst) ---
        Sleep(500) ; 確定直後のビジー回避
        t2 := A_TickCount
        resD := WaitContextMenu(pos.destX, pos.destY)
        currentLap.dstMs := A_TickCount - t2
        currentLap.dstTry := resD.tries
        
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
    L1 := lapData[1], L6 := lapData[TotalDays]
    
    res := "【1回目 vs " TotalDays "回目 詳細比較 (v5.7)】`n`n"
    res .= "項目 [試行/時間]`t`t1回目`t`t" TotalDays "回目`n"
    res .= "----------------------------------------------------------------------`n"
    res .= "右クリック(複製元):`t[" L1.srcTry " / " L1.srcMs "ms]`t[" L6.srcTry " / " L6.srcMs "ms]`n"
    res .= "右クリック(複製先):`t[" L1.dstTry " / " L1.dstMs "ms]`t[" L6.dstTry " / " L6.dstMs "ms]`n"
    res .= "確定ボタン出現:`t`t[" L1.btnAppTry " / " L1.btnAppMs "ms]`t[" L6.btnAppTry " / " L6.btnAppMs "ms]`n"
    res .= "ダイアログ消失:`t`t[ -- / " L1.diagCloseMs "ms]`t[ -- / " L6.diagCloseMs "ms]`n"
    res .= "日付窓出現:`t`t`t[" L1.dateWinTry " / " L1.dateWinMs "ms]`t[" L6.dateWinTry " / " L6.dateWinMs "ms]`n"
    res .= "----------------------------------------------------------------------`n"
    res .= "※右クリックが 1回/600ms以下 であれば理想的な速度です。"
    
    MsgBox(res, "SSI詳細パフォーマンス統計", "Iconi")
}

Esc::Exit

; --- 待機時間のバランス調整：500ms（1回で仕留める境界線） ---

WaitContextMenu(clickX, clickY) {
    Loop 20 {
        currentTry := A_Index
        Click(clickX, clickY, "Right")
        
        ; 以前の 300(空振り) と 700(余裕) の間、500msでテスト。
        ; 2回/625ms よりも、1回/500ms台 を目指します。
        Sleep(500) 
        
        MouseGetPos(,, &mHwnd)
        if (mHwnd && InStr(WinGetClass(mHwnd), "WindowsForms10.Window.20808"))
            return {tries: currentTry}
        
        Sleep(200) ; 失敗時のリカバー
    }
    Exit
}

; --- 以下の関数は安定版を維持 ---

GetDrugCoords() {
    MouseGetPos(&mX, &mY, &srcWin, &srcClassNN, 2)
    ControlGetPos(,,, &cH, srcClassNN, srcWin)
    return {srcX: mX, srcY: mY, destX: mX, destY: mY + cH}
}

EnsureConfirmAndClick() {
    targetBtnText := "確定(&S)"
    res := {btnAppMs: 0, btnAppTry: 0, diagCloseMs: 0}
    tStart := A_TickCount
    Loop 50 {
        currTry := A_Index
        MouseGetPos(,, &pWin)
        if (pWin) {
            for hCtrl in WinGetControlsHwnd(pWin) {
                try {
                    if (InStr(ControlGetText(hCtrl), targetBtnText) && ControlGetVisible(hCtrl)) {
                        res.btnAppMs := A_TickCount - tStart
                        res.btnAppTry := currTry
                        Sleep(150), Send("!s"), Sleep(150)
                        ConfirmDialogWithY("確認")
                        tWaitClose := A_TickCount
                        Loop 50 {
                            if (!ControlGetVisible(hCtrl)) {
                                res.diagCloseMs := A_TickCount - tWaitClose
                                return res
                            }
                            Sleep(100)
                        }
                    }
                }
            }
        }
        Sleep(100)
    }
    Exit
}

ConfirmDialogWithY(DialogTitle) {
    if WinWait(DialogTitle,, 2) {
        Sleep(300), Send("y")
    }
}

ChangeDate(dayOffset) {
    dateWinTitle := "基準日から何日前後に登録するか選択"
    tStart := A_TickCount
    Loop 30 {
        currTry := A_Index
        if WinExist(dateWinTitle) {
            ms := A_TickCount - tStart
            Sleep(300)
            Send("{Down " . dayOffset . "}{Enter}{Enter}")
            WinWaitClose(dateWinTitle,, 2)
            return {tries: currTry, ms: ms}
        }
        Sleep(100)
    }
    Exit
}
