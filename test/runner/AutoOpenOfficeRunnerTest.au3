#include "..\..\runner\AutoOpenOfficeRunner.au3"

AutoOpenOfficeRunner(@ScriptDir & "\AutoOpenOfficeRunnerTest.ods", "ƒV[ƒg‚P", "CallBackFunc")

Func CallBackFunc($sheet, $line)
	Local $factor[3] = ["ˆöŽq‚P", "ˆöŽq‚Q", "ˆöŽq‚R"]
	ConsoleWrite("No." & GetNo($sheet, $line) & "   ")
	For $f In $factor
		Local $cell = GetCell($sheet, $line, $f)
        ConsoleWrite($cell.String & " : ")
	Next
	ConsoleWrite(@CRLF)
EndFunc   ;==>CallBackFunc

