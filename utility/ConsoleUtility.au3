#include-once
#region パブリックメソッド
;
; 改行コードを付加し, コンソールに出力する.
;
; @param $str 出力文字列.
;
Func ConsoleUtility_WriteLn($str)
	ConsoleWrite($str & @CRLF)
EndFunc   ;==>ConsoleUtility_WriteLn

;
; デバッグ用のログをコンソールに出力する.
; $flag が 0 より大きい数値の場合に出力する.
;
; @param $flag 0より大きい数値の場合に出力する.
; @param $str 出力文字列.
;
Func ConsoleUtility_DebugLog($flag, $str)
	If 0 < $flag Then
		ConsoleWrite($str)
	EndIf
EndFunc   ;==>ConsoleUtility_DebugLog

;
; デバッグ用のログを, 改行コードを付加し, コンソールに出力する.
; $flag が 0 より大きい数値の場合に出力する.
;
; @param $flag 0より大きい数値の場合に出力する.
; @param $str 出力文字列.
;
Func ConsoleUtility_DebugLogLn($flag, $str)
	If 0 < $flag Then
		ConsoleUtility_WriteLn($str)
	EndIf
EndFunc   ;==>ConsoleUtility_DebugLogLn
#endregion パブリックメソッド