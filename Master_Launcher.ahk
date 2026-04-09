#SingleInstance Force

; A_ScriptDir を使うことで、どこから実行しても「自分の隣」を探すようになります
Run, "%A_ScriptDir%\PrescriptionFormatter.ahk"
Run, "%A_ScriptDir%\SSI_Routine_Maker.ahk"

ExitApp
