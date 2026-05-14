/*
 * @title SSI_Routine_Maker.ahk
 * @version 7.6
 * @author Gemini
 * @description 
 * 【手動切り分け版】
 * ・Shift + 右クリック ＝ 「確定(!s)」のみ送信。
 * ・Ctrl  + 右クリック ＝ 「確定(!s) ＋ 確認(y)」をループ送信。
 */

#Requires AutoHotkey v2.0

TotalDays := 6
global lapData := []

; --- パターン1：Shift + 右クリック (確定のみ) ---
+RButton:: {
    RunRoutine("Simple")
}

; --- パターン2：Ctrl + 右クリック (確定 + 確認) ---
^RButton:: {
    RunRoutine("Full")
}

; 共通ルーチン本体
RunRoutine(Mode) {
    CoordMode("Mouse", "Screen")
    lapData.Length := 0
    GlobalStart := A_TickCount
    
    pos := GetDrugCoords()
    
    Loop TotalDays {
        currentDay := A_Index
        currentLap := { srcMs:0, dstMs:0, btnAppMs:0, dateWinMs:0 }
        
        ; --- A. 複製元の右クリック ---
        resA := MachineGunClick(pos.srcX, pos.srcY)
        currentLap.srcMs := resA.ms
        
        Send("c") 
        
        ; --- B & C. 確定処理 (手動分岐) ---
        tB := A_TickCount
        Sleep(350) ; 描画待ち
        
        if (Mode == "Simple") {
            ; Shift時の挙動：確定のみ
            Sleep(400) ; うまくいかないときはこれを増やすとうまくいくかも。
            Send("!s")
            Sleep(150)
        } else {
            ; Ctrl時の挙動：従来のリトライあり
            Loop 5 {
                Send("!s") 
                Sleep(60)
                if WinExist("確認") {
                    Send("y")
                    break
                }
                Sleep(40)
            }
        }
        currentLap.btnAppMs := A_TickCount - tB
        
        ; --- D. 複製された薬剤（1行下）を右クリック ---
        Sleep(400) 
        resD := MachineGunClick(pos.destX, pos.destY)
        currentLap.dstMs := resD.ms
        
        Send("{Down 3}{Enter}")
        
        ; --- E. 日付変更処理 ---
        resE := ChangeDate(currentDay)
        currentLap.dateWinMs := resE.ms
        
        lapData.Push(currentLap)
        Sleep(200)
    }
    
    TotalElapsed := (A_TickCount - GlobalStart) / 1000
    L1 := lapData[1], L6 := lapData[TotalDays]
    
    res := "【SSI詳細統計 (v7.5) / モード: " . Mode . "】`n`n"
    res .= "★総経過時間: " . Format("{:.2f}", TotalElapsed) . " 秒`n`n"
    res .= "項目 [時間]`t`t1回目`t`t6回目`n"
    res .= "----------------------------------------------------`n"
    res .= "右クリック(元):`t" L1.srcMs "ms`t`t" L6.srcMs "ms`n"
    res .= "右クリック(先):`t" L1.dstMs "ms`t`t" L6.dstMs "ms`n"
    res .= "確定ボタン処理:`t" L1.btnAppMs "ms`t`t" L6.btnAppMs "ms`n"
    res .= "日付窓出現:`t`t" L1.dateWinMs "ms`t`t" L6.dateWinMs "ms`n"
    res .= "----------------------------------------------------`n"
    res .= "Escキーで終了(ExitApp)できます。"
    
    MsgBox(res, "SSIパフォーマンスレポート", "Iconi")
}

; 緊急停止
Esc::ExitApp

; --- 共通関数群 ---

MachineGunClick(cX, cY) {
    tS := A_TickCount
    Click(cX, cY, "Right")
    Loop 20 {
        if (Mod(A_Index, 4) == 0)
            Click(cX, cY, "Right")
        Sleep(50) 
        MouseGetPos(,, &mHwnd)
        if (mHwnd && InStr(WinGetClass(mHwnd), "WindowsForms10.Window.20808"))
            return {ms: A_TickCount - tS}
    }
    return {ms: A_TickCount - tS}
}

ChangeDate(dayOffset) {
    dateWinTitle := "基準日から何日前後に登録するか選択"
    tStart := A_TickCount
    Loop 50 {
        if WinExist(dateWinTitle) {
            Send("{Down " . dayOffset . "}{Enter}{Enter}")
            return {ms: A_TickCount - tStart}
        }
        Sleep(30)
    }
    return {ms: A_TickCount - tStart}
}

GetDrugCoords() {
    MouseGetPos(&mX, &mY, &srcWin, &srcClassNN, 2)
    ControlGetPos(,,, &cH, srcClassNN, srcWin)
    return {srcX: mX, srcY: mY, destX: mX, destY: mY + cH}
}
