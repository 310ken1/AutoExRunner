#include "..\..\runner\AutoOpenOfficeRunner.au3"

AutoOpenOfficeRunner(@ScriptDir & "\AutoOpenOfficeRunnerTest.ods", "�V�[�g�P", "CallBackFunc")

Func CallBackFunc($sheet, $line)
	Local $factor[3] = ["���q�P", "���q�Q", "���q�R"]
	ConsoleWrite("No." & GetNo($sheet, $line) & "   ")
	For $f In $factor
		Local $cell = GetCell($sheet, $line, $f)
        ConsoleWrite($cell.String & " : ")
	Next
	ConsoleWrite(@CRLF)
EndFunc   ;==>CallBackFunc

