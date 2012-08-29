#include "..\..\app\TeraTerm.au3"
#include "..\..\utility\ConsoleUtility.au3"
#include "..\..\utility\FileUtility.au3"

#region グローバル変数の設定
$TeraTerm_CmdPath = "C:\usr\local\teraterm-4.74\ttermpro.exe"
$TeraTerm_MacroCmdPath = "C:\usr\local\teraterm-4.74\ttpmacro.exe"
$TeraTerm_DebugLog = 1
#endregion グローバル変数の設定

TestMain()

#region テストコード
Func TestMain()
	ConsoleUtility_WriteLn(TeraTerm_MacroRun_Test())
	ConsoleUtility_WriteLn(TeraTerm_MacroRunWait_Test())
EndFunc   ;==>TestMain

;
; TeraTerm_MacroRun() テスト.
;
Func TeraTerm_MacroRun_Test()
	Local $result = "NG"
	Local $time = TimerInit()

	Local $ttl = FileUtility_ScriptDirFilePath("test.ttl")
	TeraTerm_MacroRun($ttl)
	TeraTerm_MacroWaitClose()

	If Ceiling(TimerDiff($time) / 1000) > 5 Then
		$result = "OK"
	EndIf
	Return $result
EndFunc   ;==>TeraTerm_MacroRun_Test

;
; TeraTerm_MacroRunWait() テスト.
;
Func TeraTerm_MacroRunWait_Test()
	Local $result = "NG"
	Local $time = TimerInit()

	Local $ttl = FileUtility_ScriptDirFilePath("test.ttl")
	TeraTerm_MacroRunWait($ttl)

	If Ceiling(TimerDiff($time) / 1000) > 5 Then
		$result = "OK"
	EndIf
	Return $result
EndFunc   ;==>TeraTerm_MacroRunWait_Test
#endregion テストコード