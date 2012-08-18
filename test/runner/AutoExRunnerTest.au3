#include "..\..\runner\AutoExRunner.au3"

AutoExRunner(@ScriptDir & "\AutoExRunnerTest.xls", "ƒV[ƒg1", "CallBackFunc")

Func CallBackFunc($sheet, $line)
	Local $factor[3] = ["ˆöŽq‚P", "ˆöŽq‚Q", "ˆöŽq‚R"]
	ConsoleWrite("No." & GetNo($sheet, $line) & "   ")
	For $f In $factor
		Local $cell = GetCell($sheet, $line, $f)
		ConsoleWrite($cell.value & " : ")
	Next
	ConsoleWrite(@CRLF)
EndFunc   ;==>CallBackFunc

