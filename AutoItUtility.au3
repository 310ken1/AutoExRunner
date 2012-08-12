#include-once
;
; 改行コードを付加し, コンソールに出力する.
;
; @param $str 出力文字列.
;
Func ConsoleWriteLn($str)
	ConsoleWrite($str & @CRLF)
EndFunc   ;==>ConsoleWriteLn

