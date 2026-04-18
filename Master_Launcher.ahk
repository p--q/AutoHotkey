/*
 * @title Master_Launcher.ahk
 * @version 1.3
 * @description 
 * 下記3つの自動化スクリプトを一括起動・管理します。
 * 1. SSI_Routine_Maker.ahk (SSI爆速ルーチン)
 * 2. PrescriptionFormatter.ahk (処方フォーマッター)
 * 3. Muhenkan_Y_Repeater.ahk (無変換+Y リピーター)
 */

#Requires AutoHotkey v2.0
#SingleInstance Force

; --- 起動対象のリスト ---
ScriptsToRun := [
    "SSI_Routine_Maker.ahk",
    "PrescriptionFormatter.ahk",
    "Muhenkan_Y_Repeater.ahk"
]

; スクリプトの起動実行
RunScripts(ScriptsToRun)

; --- トレイメニューの構築 ---
A_TrayMenu.Delete() ; 標準メニューを一旦クリア
A_TrayMenu.Add("全てのスクリプトを再起動", (*) => ReloadAll(ScriptsToRun))
A_TrayMenu.Add() ; 区切り線
A_TrayMenu.Add("ランチャーを終了", (*) => ExitApp())

; --- 関数定義 ---

RunScripts(scriptArray) {
    for scriptName in scriptArray {
        if FileExist(scriptName) {
            try {
                Run(scriptName)
            } catch {
                MsgBox("起動に失敗しました:`n" . scriptName, "Launcher Error")
            }
        } else {
            ; ファイルが見つからない場合は警告を表示
            MsgBox("ファイルが見つかりません:`n" . scriptName . "`n`nスクリプトが同じフォルダにあるか確認してください。", "Launcher Error")
        }
    }
}

ReloadAll(scriptArray) {
    ; 子スクリプトを再起動（各スクリプト側の SingleInstance Force により上書きされます）
    RunScripts(scriptArray)
    ; ランチャー自身もリロード
    Reload()
}
