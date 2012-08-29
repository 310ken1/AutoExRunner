#include "..\..\runner\AutoOpenOfficeRunner.au3"

AutoOpenOfficeRunner(@ScriptDir & "\AutoOpenOfficeRunnerTest.ods", "シート１", "CallBackFunc")

Func CallBackFunc($sheet, $line)
	Local $factor[3] = ["因子１", "因子２", "因子３"]
	ConsoleWrite("No." & GetNo($sheet, $line) & "   ")
	For $f In $factor
		Local $cell = GetCell($sheet, $line, $f)
        ConsoleWrite($cell.String & " : ")
	Next
	ConsoleWrite(@CRLF)
EndFunc   ;==>CallBackFunc

