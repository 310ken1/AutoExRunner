#include "..\AutoItTest.au3"
#include "..\..\app\TeraTerm.au3"
#include "..\..\utility\FileUtility.au3"

#region Globale_Argument_Setting
$TeraTerm_MacroCmdPath = "C:\usr\local\teraterm-4.74\ttpmacro.exe"
$TeraTerm_DebugLog = 0
#endregion Globale_Argument_Setting

Local $TeraTermTest[2][5] = [ _
		["", "TeraTerm_MacroRun_Test", "AutoItTest_Assert", True, ""], _
		["", "TeraTerm_MacroRunWait_Test", "AutoItTest_Assert", True, ""] _
		]

AutoItTest_Runner($TeraTermTest)

#region TeraTerm_MacroRun_Test
;
; TeraTerm_MacroRun() テスト.
;
Func TeraTerm_MacroRun_Test()
	Local $result = False
	Local $time = TimerInit()

	Local $ttl = FileUtility_ScriptDirFilePath("test.ttl")
	TeraTerm_MacroRun($ttl)
	TeraTerm_MacroWaitClose()

	If 5 < Ceiling(TimerDiff($time) / 1000)  Then
		$result = True
	EndIf
	Return $result
EndFunc   ;==>TeraTerm_MacroRun_Test
#endregion

#region TeraTerm_MacroRunWait_Test
;
; TeraTerm_MacroRunWait() テスト.
;
Func TeraTerm_MacroRunWait_Test()
	Local $result = False
	Local $time = TimerInit()

	Local $ttl = FileUtility_ScriptDirFilePath("test.ttl")
	TeraTerm_MacroRunWait($ttl)

	If 5 < Ceiling(TimerDiff($time) / 1000) Then
		$result = True
	EndIf
	Return $result
EndFunc   ;==>TeraTerm_MacroRunWait_Test
#endregion Test