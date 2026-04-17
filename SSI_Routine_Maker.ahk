/*
 * @title SSI_Routine_Maker.ahk
 * @version 6.2
 * @author Gemini
 * @description 
 * 【SSI専用ルーチン】
 * 1920x1080の画面で最大化したときにセットメニューの点滴を6日分コピーします。
 * コピーしたい点滴のチェックボックスのある枠上で「Shift+右クリック」すると開始します。
 * * 仕組み:
 * 1. 薬剤の1行下の座標をコントロールの高さから自動計算します。
 * 2. コンテキストメニューが出るまで超高速でリトライ判定を行います。
 * 3. 確定ボタンを全走査して特定し、出現した瞬間にクリックします。
 * 4. 日付選択ウィンドウの出現をミリ秒単位で監視し、即座にキーを送ります。
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
        currentLap := { srcTry:0, srcMs:0, dstTry:0, dstMs:0, btnAppTry:0, btnAppMs:0, diagCloseMs:0, dateWinTry:0, dateWinMs:0 }
        
        ; --- A. 複製元の右クリック (src) ---
        t1 := A_TickCount
        resA := MachineGunClick(pos.srcX, pos.srcY)
        currentLap.srcMs := A_TickCount - t1
        currentLap.srcTry := resA.tries
        
        Send("c") 
        
        ; --- B & C. 確定ボタン走査 ＆ 確認ダイアログ応答 ---
        resBC := EnsureConfirmAndClick()
        currentLap.btnAppMs := resBC.btnAppMs
        currentLap.btnAppTry := resBC.btnAppTry
        currentLap.diagCloseMs := resBC.diagCloseMs
        
        ; --- D. 複製された薬剤（1行下）を右クリック (dst) ---
        ; 確定処理後のビジー時間を考慮しつつ、最短で再開
        Sleep(300) 
        t2 := A_TickCount
        resD := MachineGunClick(pos.destX, pos.destY)
        currentLap.dstMs := A_TickCount - t2
        currentLap.dstTry := resD.tries
        
        Send("{Down 3}{Enter}")
        
        ; --- E. 日付変更処理 ---
        t3 := A_TickCount
        resE := ChangeDate(currentDay)
        currentLap.dateWinMs := A_TickCount - t3
        currentLap.dateWinTry := resE.tries
        
        lapData.Push(currentLap)
        
        ; 次のループへの待機も最小限に
        Sleep(200)
    }
    
    L1 := lapData[1], L6 := lapData[TotalDays]
    
    res := "【マシンガン・ポーリング解析 (v6.2)】`n`n"
    res .= "項目 [試行/時間]`t`t1回目`t`t6回目`n"
    res .= "----------------------------------------------------------------------`n"
    res .= "右クリック(複製元):`t[" L1.srcTry " / " L1.srcMs "ms]`t[" L6.srcTry " / " L6.srcMs "ms]`n"
    res .= "右クリック(複製先):`t[" L1.dstTry " / " L1.dstMs "ms]`t[" L6.dstTry " / " L6.dstMs "ms]`n"
    res .= "確定ボタン出現:`t`t[" L1.btnAppTry " / " L1.btnAppMs "ms]`t[" L6.btnAppTry " / " L6.btnAppMs "ms]`n"
    res .= "日付窓出現:`t`t`t[" L1.dateWinTry " / " L1.dateWinMs "ms]`t[" L6.dateWinTry " / " L6.dateWinMs "ms]`n"
    res .= "----------------------------------------------------------------------`n"
    res .= "※このモードではTry回数が増えるほど「最速で隙間を突いた」ことになります。"
    
    MsgBox(res, "SSI詳細パフォーマンス統計", "Iconi")
}

Esc::ExitApp

; --- 改善の核：高速リトライ・ポーリング ---

MachineGunClick(cX, cY) {
    ; 初回クリック
    Click(cX, cY, "Right")
    Loop 50 {
        currentTry := A_Index
        ; 判定スパンを極短(50ms)に設定。
        ; メニューが出ていなければ、200msごとに追加クリックを叩き込む「攻め」の姿勢
        if (Mod(A_Index, 4) == 0) {
            Click(cX, cY, "Right")
        }
        
        Sleep(50) 
        
        MouseGetPos(,, &mHwnd)
        if (mHwnd && InStr(WinGetClass(mHwnd), "WindowsForms10.Window.20808"))
            return {tries: currentTry}
    }
    Exit
}

EnsureConfirmAndClick() {
    targetBtnText := "確定(&S)"
    res := {btnAppMs: 0, btnAppTry: 0, diagCloseMs: 0}
    tStart := A_TickCount
    Loop 100 {
        currTry := A_Index
        MouseGetPos(,, &pWin)
        if (pWin) {
            for hCtrl in WinGetControlsHwnd(pWin) {
                try {
                    if (InStr(ControlGetText(hCtrl), targetBtnText) && ControlGetVisible(hCtrl)) {
                        res.btnAppMs := A_TickCount - tStart
                        res.btnAppTry := currTry
                        ; ボタンが見えた瞬間に !s を送り、ダイアログを待ち受ける
                        Send("!s")
                        ConfirmDialogWithY("確認")
                        
                        tWaitClose := A_TickCount
                        Loop 50 {
                            if (!ControlGetVisible(hCtrl)) {
                                res.diagCloseMs := A_TickCount - tWaitClose
                                return res
                            }
                            Sleep(30)
                        }
                    }
                }
            }
        }
        Sleep(30)
    }
    Exit
}

ConfirmDialogWithY(DialogTitle) {
    ; ダイアログが出るまで 30ms 刻みで監視
    Loop 50 {
        if WinExist(DialogTitle) {
            Send("y")
            return
        }
        Sleep(30)
    }
}

ChangeDate(dayOffset) {
    dateWinTitle := "基準日から何日前後に登録するか選択"
    tStart := A_TickCount
    Loop 50 {
        currTry := A_Index
        if WinExist(dateWinTitle) {
            ms := A_TickCount - tStart
            ; ウィンドウ検知からキー送信までのラグを抹殺
            Send("{Down " . dayOffset . "}{Enter}{Enter}")
            WinWaitClose(dateWinTitle,, 1)
            return {tries: currTry, ms: ms}
        }
        Sleep(30)
    }
    Exit
}

GetDrugCoords() {
    MouseGetPos(&mX, &mY, &srcWin, &srcClassNN, 2)
    ControlGetPos(,,, &cH, srcClassNN, srcWin)
    return {srcX: mX, srcY: mY, destX: mX, destY: mY + cH}
}
