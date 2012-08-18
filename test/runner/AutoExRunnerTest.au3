#include "..\..\runner\AutoExRunner.au3"

AutoExRunner(@ScriptDir & "\AutoExRunnerTest.xls", "シート1", "CallBackFunc")

Func CallBackFunc($sheet, $line)
	Local $factor[3] = ["因子１", "因子２", "因子３"]
	ConsoleWrite("No." & GetNo($sheet, $line) & "   ")
	For $f In $factor
		Local $cell = GetCell($sheet, $line, $f)
		ConsoleWrite($cell.value & "：")
	Next
	ConsoleWrite(@CRLF)
EndFunc   ;==>CallBackFunc

