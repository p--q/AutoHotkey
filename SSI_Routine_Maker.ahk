/*
 * @title SSI_Routine_Maker.ahk
 * @version 6.7
 * @author Gemini
 * @description 
 * 【SSI専用ルーチン：最速安定版】
 * 1. 処理開始から終了までの「総経過時間」を計測・表示します。
 * 2. SSIをフリーズさせるパーツ走査を完全に排除。
 * 3. 確定(!s)と確認(y)を最短リズムで流し込みます。
 */

#Requires AutoHotkey v2.0

TotalDays := 6
global lapData := []

+RButton:: {
    CoordMode("Mouse", "Screen")
    lapData.Length := 0
    
    ; 全体の開始時間を記録
    GlobalStart := A_TickCount
    
    pos := GetDrugCoords()
    
    Loop TotalDays {
        currentDay := A_Index
        currentLap := { srcTry:0, srcMs:0, dstTry:0, dstMs:0, btnAppMs:0, dateWinMs:0 }
        
        ; --- A. 複製元の右クリック ---
        resA := MachineGunClick(pos.srcX, pos.srcY)
        currentLap.srcMs := resA.ms
        currentLap.srcTry := resA.tries
        
        Send("c") 
        
        ; --- B & C. 確定処理 (走査なし・最短流し込み) ---
        tB := A_TickCount
        Sleep(350) ; 描画待ち
        Loop 5 {
            Send("!s") 
            Sleep(60)
            if WinExist("確認") {
                Send("y")
                break
            }
            Sleep(40)
        }
        currentLap.btnAppMs := A_TickCount - tB
        
        ; --- D. 複製された薬剤（1行下）を右クリック ---
        Sleep(400) 
        resD := MachineGunClick(pos.destX, pos.destY)
        currentLap.dstMs := resD.ms
        currentLap.dstTry := resD.tries
        
        Send("{Down 3}{Enter}")
        
        ; --- E. 日付変更処理 ---
        resE := ChangeDate(currentDay)
        currentLap.dateWinMs := resE.ms
        
        lapData.Push(currentLap)
        Sleep(200)
    }
    
    ; 全体の経過時間を算出
    TotalElapsed := (A_TickCount - GlobalStart) / 1000
    
    L1 := lapData[1], L6 := lapData[TotalDays]
    res := "【SSI詳細パフォーマンス統計 (v6.7)】`n`n"
    res .= "★総経過時間: " . Format("{:.2f}", TotalElapsed) . " 秒`n`n"
    res .= "項目 [時間]`t`t1回目`t`t6回目`n"
    res .= "----------------------------------------------------------------------`n"
    res .= "右クリック(元):`t" L1.srcMs "ms`t`t" L6.srcMs "ms`n"
    res .= "右クリック(先):`t" L1.dstMs "ms`t`t" L6.dstMs "ms`n"
    res .= "確定ボタン処理:`t" L1.btnAppMs "ms`t`t" L6.btnAppMs "ms`n"
    res .= "日付窓出現:`t`t" L1.dateWinMs "ms`t`t" L6.dateWinMs "ms`n"
    res .= "----------------------------------------------------------------------`n"
    res .= "Escキーでいつでも終了(ExitApp)できます。"
    
    MsgBox(res, "SSI詳細パフォーマンス統計", "Iconi")
}

; 緊急停止（完全に終了）
Esc::ExitApp

; --- 最速化関数群 ---

MachineGunClick(cX, cY) {
    tS := A_TickCount
    Click(cX, cY, "Right")
    Loop 20 {
        curTry := A_Index
        if (Mod(A_Index, 4) == 0)
            Click(cX, cY, "Right")
        Sleep(50) 
        MouseGetPos(,, &mHwnd)
        if (mHwnd && InStr(WinGetClass(mHwnd), "WindowsForms10.Window.20808"))
            return {tries: curTry, ms: A_TickCount - tS}
    }
    return {tries: 20, ms: A_TickCount - tS}
}

ChangeDate(dayOffset) {
    dateWinTitle := "基準日から何日前後に登録するか選択"
    tStart := A_TickCount
    Loop 50 {
        if WinExist(dateWinTitle) {
            Send("{Down " . dayOffset . "}{Enter}{Enter}")
            return {tries: A_Index, ms: A_TickCount - tStart}
        }
        Sleep(30)
    }
    return {tries: 50, ms: A_TickCount - tStart}
}

GetDrugCoords() {
    MouseGetPos(&mX, &mY, &srcWin, &srcClassNN, 2)
    ControlGetPos(,,, &cH, srcClassNN, srcWin)
    return {srcX: mX, srcY: mY, destX: mX, destY: mY + cH}
}
