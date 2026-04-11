; ============================================================
; SSI電子カルテ コンテクストメニュー検出テストスクリプト（AHK v2）
; 3方式を同時にチェックして、どれが反応するかを表示する
; ============================================================

#SingleInstance Force

global lastWinList := []

; 起動時に既存の WinForms ウィンドウ一覧を保存
Init() {
    global lastWinList := WinGetList("ahk_class WindowsForms10.Window.*")
}
Init()

; 右クリック後に新規ウィンドウを検出
~RButton::
{
    SetTimer(CheckNewWinFormsWindow, -80)
}

CheckNewWinFormsWindow() {
    global lastWinList

    newList := WinGetList("ahk_class WindowsForms10.Window.*")

    for hwnd in newList {
        if !(hwnd in lastWinList) {
            ShowDetection("方式A：右クリック後に新規 WinForms ウィンドウを検出", hwnd)
            lastWinList := newList
            return
        }
    }
}

; 方式B：タイトルなし & 小さい WinForms ウィンドウを検出
SetTimer(CheckSmallWinFormsWindow, 100)

CheckSmallWinFormsWindow() {
    winList := WinGetList("ahk_class WindowsForms10.Window.*")

    for hwnd in winList {
        title := WinGetTitle(hwnd)
        if (title = "") {
            rect := WinGetPos(hwnd)
            w := rect[3], h := rect[4]

            if (w < 500 && h < 700) {
                ShowDetection("方式B：タイトルなし & 小型 WinForms ウィンドウを検出", hwnd)
                return
            }
        }
    }
}

; 方式C：項目ウィンドウ（20808.app.*）を直接検出
SetTimer(CheckMenuItemWindow, 100)

CheckMenuItemWindow() {
    hwnd := WinExist("ahk_class WindowsForms10.Window.20808.app.*")
    if hwnd {
        ShowDetection("方式C：項目ウィンドウ（20808.app.*）を検出", hwnd)
    }
}

; 検出結果を表示（重複表示防止）
global lastDetected := 0

ShowDetection(method, hwnd) {
    global lastDetected
    if (hwnd = lastDetected)
        return

    lastDetected := hwnd
    MsgBox method "`nHWND: " hwnd
}
