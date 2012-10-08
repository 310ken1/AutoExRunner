#include-once
#include "..\Utility\ConsoleUtility.au3"

#region Globale_Argument_Define
;
; デバッグログフラグ.
; 1 を指定することで, デバッグログが出力される.
;
Global $AutoItTest_DebugLog = 0
#endregion Globale_Argument_Define

#region Public_Method
;
; テストを実行する.
;
; @param $array 次の要素を持つ２次元配列.
;                         [テスト前関数,
;                          テスト関数,
;                          成否判定関数,
;                          期待する値,
;                          テスト後関数]
;
Func AutoItTest_Runner(ByRef $array)
	Local $ok = 0
	Local $count = 0
	For $count = 0 To UBound($array) - 1
		Local $before = $array[$count][0]
		Local $test = $array[$count][1]
		Local $judge = $array[$count][2]
		Local $expected = $array[$count][3]
		Local $after = $array[$count][4]

		Call($before)
		Local $actual = Call($test)
		Local $ret = Call($judge, $expected, $actual, @error, @extended)
		Call($after)

		Local $result = "NG"
		If True = $ret Then
			$result = "OK"
			$ok +=1
		EndIf
		ConsoleUtility_WriteLn($test & " : " & $result)
	Next
	ConsoleUtility_WriteLn("Total:" & $count & "  OK:" & $ok & "  NG:" & $count - $ok)
EndFunc   ;==>AutoItTest_Runner

;
; 期待する値と一致するか判定する.
;
; @param $expected 期待する値.
; @param $actual 実際の値.
; @param $error @errorの値(未使用).
; @param $extended @extendedの値(未使用).
; @return 成否(True/False)
;
Func AutoItTest_Assert($expected, $actual, $error = 0, $extended = 0)
	Local $result = False
	If $expected = $actual Then
		$result = True
	EndIf
	ConsoleUtility_DebugLogLn($AutoItTest_DebugLog, "AutoItTest_Assert : " & $result)
	Return $result
EndFunc   ;==>AutoItTest_Assert

;
; 期待する値と一致するか判定する.
;
; @param $expected 期待する値.
; @param $actual 実際の値.
; @param $error @errorの値(未使用).
; @param $extended @extendedの値(未使用).
; @return 成否(True/False)
;
Func AutoItTest_AssertArrayEquals(Const $expected, Const $actual, $error = 0, $extended = 0)
	Local $result = False
	If UBound($expected) = UBound($actual) Then
		$result = True
		For $i = 0 To UBound($expected) - 1
			If $expected[$i] <> $actual[$i] Then
				$result = False
				ExitLoop
			EndIf
		Next
	EndIf
	ConsoleUtility_DebugLogLn($AutoItTest_DebugLog, "AutoItTest_AssertArrayEquals : " & $result)
	Return $result
EndFunc   ;==>AutoItTest_AssertArrayEquals

;
; オブジェクトが生成されているか判定する.
;
; @param $expected 期待する値.
; @param $actual 実際の値.
; @param $error @errorの値(未使用).
; @param $extended @extendedの値(未使用).
; @return 成否(True/False)
;
Func AutoItTest_IsObj($expected, $actual, $error = 0, $extended = 0)
	Local $result = False
	If $expected = IsObj($actual) Then
		$result = True
	EndIf
	ConsoleUtility_DebugLogLn($AutoItTest_DebugLog, "AutoItTest_IsObj : " & $result)
	Return $result
EndFunc   ;==>AutoItTest_IsObj

;
; 配列か判定する.
;
; @param $expected 期待する値.
; @param $actual 実際の値.
; @param $error @errorの値(未使用).
; @param $extended @extendedの値(未使用).
; @return 成否(True/False)
;
Func AutoItTest_IsArray($expected, $actual, $error = 0, $extended = 0)
	Local $result = False
	If $expected = IsArray($actual) Then
		$result = True
	EndIf
	ConsoleUtility_DebugLogLn($AutoItTest_DebugLog, "AutoItTest_IsArray : " & $result)
	Return $result
EndFunc   ;==>AutoItTest_IsObj

;
; ファイルが存在していることを判定する.
;
; @param $expected 期待する値(ファイル名)
; @param $actual 実際の値(未使用).
; @param $error @errorの値(未使用).
; @param $extended @extendedの値(未使用).
; @return 成否(True/False)
;
Func AutoItTest_FileExists($expected, $actual, $error = 0, $extended = 0)
	Local $result = False
	If FileExists($expected) Then
		$result = True
	EndIf
	ConsoleUtility_DebugLogLn($AutoItTest_DebugLog, "AutoItTest_FileExists : " & $result)
	Return $result
EndFunc   ;==>AutoItTest_FileExists

;
;  ファイルが存在していないことを判定する..
;
; @param $expected 期待する値(ファイル名)
; @param $actual 実際の値(未使用).
; @param $error @errorの値(未使用).
; @param $extended @extendedの値(未使用).
; @return 成否(True/False)
;
Func AutoItTest_FileNotExists($expected, $actual, $error = 0, $extended = 0)
	Local $result = False
	If Not FileExists($expected) Then
		$result = True
	EndIf
	ConsoleUtility_DebugLogLn($AutoItTest_DebugLog, "AutoItTest_FileNotExists : " & $result)
	Return $result
EndFunc   ;==>AutoItTest_FileNotExists

;
; ウィンドウが存在することを判定する.
;
; @param $expected 期待する値(ウィンドウテキスト).
; @param $actual 実際の値(未使用).
; @param $error @errorの値(未使用).
; @param $extended @extendedの値(未使用).
; @return 成否(True/False)
;
Func AutoItTest_WinExists($expected, $actual, $error = 0, $extended = 0)
	Local $result = False
	If WinExists($expected) Then
		$result = True
	EndIf
	ConsoleUtility_DebugLogLn($AutoItTest_DebugLog, "AutoItTest_WinExists : " & $result)
	Return $result
EndFunc   ;==>AutoItTest_WinExists
;
; ウィンドウが存在しないことを判定する.
;
; @param $expected 期待する値(ウィンドウテキスト).
; @param $actual 実際の値(未使用).
; @param $error @errorの値(未使用).
; @param $extended @extendedの値(未使用).
; @return 成否(True/False)
;
Func AutoItTest_WinNotExists($expected, $actual, $error = 0, $extended = 0)
	Local $result = False
	If Not WinExists($expected) Then
		$result = True
	EndIf
	ConsoleUtility_DebugLogLn($AutoItTest_DebugLog, "AutoItTest_WinNotExists : " & $result)
	Return $result
EndFunc   ;==>AutoItTest_WinNotExists
#endregion Public_Method