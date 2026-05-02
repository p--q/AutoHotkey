/*
 * @title SSI_Routine_Maker.ahk
 * @version 7.4
 * @author Gemini
 * @description 
 * 確定ボタン押下後、再試行は一切行わず、右クリックメニューが出るかどうかで
 * 次のステップへ進むべきか、確認窓を待つべきかを判定します。
 */

#Requires AutoHotkey v2.0

TotalDays := 6
global lapData := []

+RButton:: {
    CoordMode("Mouse", "Screen")
    lapData.Length := 0
    startTime := A_TickCount
    
    pos := GetDrugCoords()
    
    Loop TotalDays {
        currentDay := A_Index
        currentLap := { srcMs:0, dstMs:0, btnAppMs:0, dateWinMs:0 }
        
        ; --- A. 複製元の薬剤を右クリック ---
        t1 := A_TickCount
        WaitContextMenu(pos.srcX, pos.srcY)
        currentLap.srcMs := A_TickCount - t1
        
        Sleep(50)
        Send("c") 
        
        ; --- B & C. 確定処理 ＆ 分岐判定 ---
        t2 := A_TickCount
        ; 確定を一度だけ送り、その後の状態をチェック
        if !ConfirmAndCheckState(pos.destX, pos.destY) {
            ; もし右クリックメニューが出ず、かつ確認窓が出た場合の処理
            ; (ConfirmAndCheckState内でWinExist判定を行っています)
        }
        currentLap.btnAppMs := A_TickCount - t2
        
        ; --- D. 複製された薬剤の処理 ---
        ; 既にメニューが出ている状態、あるいは確認窓を閉じた状態からスタート
        Send("{Down 3}{Enter}") 
        
        ; --- E. 日付変更処理 ---
        t4 := A_TickCount
        ChangeDate(currentDay)
        currentLap.dateWinMs := A_TickCount - t4
        
        lapData.Push(currentLap)
        Sleep(300)
    }
    
    totalTime := (A_TickCount - startTime) / 1000
    MsgBox("完了！`n総時間: " Format("{:.2f}", totalTime) " 秒", "SSI統計", "Iconi")
}

Esc::ExitApp

; --- 改善：確定後の状態を右クリックで判定 ---

ConfirmAndCheckState(dstX, dstY) {
    ; 1. 確定を初回送信
    Send("!s")
    Sleep(200) ; SSIが内部処理を進めるための最低限の待ち

    ; 2. ループ監視
    Loop 20 {
        ; A. まず右クリックを試みる（生存・遷移確認）
        Click(dstX, dstY, "Right")
        Sleep(60)
        
        MouseGetPos(,, &mHwnd)
        if (mHwnd && InStr(WinGetClass(mHwnd), "WindowsForms10.Window.20808")) {
            ; 【成功パターン1】右クリックメニューが出た！
            ; ＝既に次の操作ができる状態なので、そのまま次に進む
            return true
        }

        ; B. メニューが出ないなら、確認窓が出ていないかチェック
        if WinExist("確認") {
            ; 【成功パターン2】確認窓が出た！
            Send("y")
            Sleep(200)
            ; 確認窓を閉じた後は、改めてメニューを出す必要があるため
            ; WaitContextMenuを呼び出すか、呼び出し元で対応
            WaitContextMenu(dstX, dstY)
            return true
        }
        
        Sleep(40)
    }
    return false
}

; --- 共通関数 ---

WaitContextMenu(cX, cY) {
    Loop 50 {
        if (Mod(A_Index, 6) == 1)
            Click(cX, cY, "Right")
        Sleep(50) 
        MouseGetPos(,, &mHwnd)
        if (mHwnd && InStr(WinGetClass(mHwnd), "WindowsForms10.Window.20808"))
            return
    }
}

ChangeDate(dayOffset) {
    Loop 50 {
        if WinExist("基準日から何日前後に登録するか選択") {
            Sleep(150)
            Send("{Down " . dayOffset . "}{Enter}{Enter}")
            WinWaitClose("基準日から何日前後に登録するか選択",, 1)
            return
        }
        Sleep(30)
    }
}

GetDrugCoords() {
    MouseGetPos(&mX, &mY, &srcWin, &srcClassNN, 2)
    ControlGetPos(,,, &cH, srcClassNN, srcWin)
    return {srcX: mX, srcY: mY, destX: mX, destY: mY + cH}
}
