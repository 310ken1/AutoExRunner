#include "..\AutoItTest.au3"
#include "..\..\utility\DateUtility.au3"

;
; !! 注意 !!
; このテストは, 実行するために管理者権限が必要.
;

#region Globale_Argument_Define
; デバッグログフラグ.
; 1 を指定することで, デバッグログが出力される.
$AutoItTest_DebugLog = 0
#endregion Globale_Argument_Define

#region Constant_Define
Const $DateUtility_DateStringToArray_Test_Date = "2001/02/03"
Const $DateUtility_DateStringToArray_Test_Date_Answer[6] = [2, 3, 2001, 0, 0, 0]
Const $DateUtility_DateStringToArray_Test_Time = "2001/02/03 10:20"
Const $DateUtility_DateStringToArray_Test_Time_Answer[6] = [2, 3, 2001, 10, 20, 0]
Const $DateUtility_DateStringToArray_Test_Sec = "2001/02/03 10:20:30"
Const $DateUtility_DateStringToArray_Test_Sec_Answer[6] = [2, 3, 2001, 10, 20, 30]
Const $DateUtility_DateStringToArray_Test_OneColumn = "2001/2/3 1:2:3"
Const $DateUtility_DateStringToArray_Test_OneColumn_Answer[6] = [2, 3, 2001, 1, 2, 3]
Const $DateUtility_DateStringToArray_Test_InvalidFormat = "2001/2/310:20"
#endregion Constant_Define

Local $DateUtilityTest[8][5] = [ _
		["", "DateUtility_DateStringToArray_Test_Date", "AutoItTest_AssertArrayEquals", $DateUtility_DateStringToArray_Test_Date_Answer, ""], _
		["", "DateUtility_DateStringToArray_Test_Time", "AutoItTest_AssertArrayEquals", $DateUtility_DateStringToArray_Test_Time_Answer, ""], _
		["", "DateUtility_DateStringToArray_Test_Sec", "AutoItTest_AssertArrayEquals", $DateUtility_DateStringToArray_Test_Sec_Answer, ""], _
		["", "DateUtility_DateStringToArray_Test_OneColumn", "AutoItTest_AssertArrayEquals", $DateUtility_DateStringToArray_Test_OneColumn_Answer, ""], _
		["", "DateUtility_DateStringToArray_Test_InvalidFormat", "AutoItTest_Assert", 1, ""], _
		["", "DateUtility_SetLocalTimeString_Test", "AutoItTest_Assert", True, ""], _
		["", "DateUtility_SetLocalTimeString_Test_Invalid", "AutoItTest_Assert", False, ""], _
		["", "DateUtility_Restore_Test", "AutoItTest_Assert", True, ""] _
		]
AutoItTest_Runner($DateUtilityTest)

#region DateUtility_DateStringToArray_Test
; 日付のみ(2001/02/03)の変換.
; 期待する結果:時刻配列([2, 3, 2001, 0, 0, 0])の取得.
Func DateUtility_DateStringToArray_Test_Date()
	Return DateUtility_DateStringToArray($DateUtility_DateStringToArray_Test_Date)
EndFunc   ;==>DateUtility_DateStringToArray_Test_Date
; 時刻(秒なし)を含む(2001/02/03 10:20)の変換.
; 期待する結果:時刻配列([2, 3, 2001, 10, 20, 0])の取得.
Func DateUtility_DateStringToArray_Test_Time()
	Return DateUtility_DateStringToArray($DateUtility_DateStringToArray_Test_Time)
EndFunc   ;==>DateUtility_DateStringToArray_Test_Time
; 時刻(秒あり)を含む(2001/02/03 10:20:30)の変換.
; 期待する結果:時刻配列([2, 3, 2001, 10, 20, 30])の取得.
Func DateUtility_DateStringToArray_Test_Sec()
	Return DateUtility_DateStringToArray($DateUtility_DateStringToArray_Test_Sec)
EndFunc   ;==>DateUtility_DateStringToArray_Test_Sec
; 一桁の日付(2001/2/3 1:2:3)の変換.
; 期待する結果:時刻配列([2, 3, 2001, 1, 2, 3])の取得.
Func DateUtility_DateStringToArray_Test_OneColumn()
	Return DateUtility_DateStringToArray($DateUtility_DateStringToArray_Test_OneColumn)
EndFunc   ;==>DateUtility_DateStringToArray_Test_OneColumn
; 無効なフォーマット(2001/2/310:20)の変換.
; 期待する結果:@error=1
Func DateUtility_DateStringToArray_Test_InvalidFormat()
	DateUtility_DateStringToArray($DateUtility_DateStringToArray_Test_InvalidFormat)
	Return @error
EndFunc   ;==>DateUtility_DateStringToArray_Test_InvalidFormat
#endregion DateUtility_DateStringToArray_Test

#region DateUtility_SetLocalTimeString_Test
; 正常な日付(2001/02/03 10:20)の設定.
; 期待する結果:日付が設定される.
Func DateUtility_SetLocalTimeString_Test()
	Local $result = False
	Local $handle = DateUtility_Save()

	DateUtility_SetLocalTimeString("2001/02/03 10:20")
	Local $change = _Date_Time_SystemTimeToArray(_Date_Time_GetLocalTime())

	If 2001 = $change[$DateUtilityYear] And _
			2 = $change[$DateUtilityMonth] And _
			3 = $change[$DateUtilityDay] And _
			10 = $change[$DateUtilityHour] And _
			20 = $change[$DateUtilityMinute] Then
		$result = True
	EndIf

	DateUtility_Restore($handle)
	Return $result
EndFunc   ;==>DateUtility_SetLocalTimeString_Test
; 無効な日付()の設定.
; 期待する結果:
Func DateUtility_SetLocalTimeString_Test_Invalid()
	Return DateUtility_SetLocalTimeString("aaaa")
EndFunc   ;==>DateUtility_SetLocalTimeString_Test_Invalid
#endregion DateUtility_SetLocalTimeString_Test

#region DateUtility_Restore_Test
; 現在時刻を保持し, 元に戻す.
; 期待する結果:現在時刻に戻る.
Func DateUtility_Restore_Test()
	Local $result = False
	Local $elapsed = 5000

	Local $handle = DateUtility_Save()
	Local $begin = _Date_Time_SystemTimeToArray(_Date_Time_GetLocalTime())

	Local $new = _Date_Time_EncodeSystemTime(1, 1, 2000, @HOUR + 1, @MIN + 1, @SEC + 5)
	_Date_Time_SetLocalTime($new)
	Local $change = _Date_Time_SystemTimeToArray(_Date_Time_GetLocalTime())

	If $begin[$DateUtilityDay] <> $change[$DateUtilityDay] And _
			$begin[$DateUtilityHour] <> $change[$DateUtilityHour] And _
			$begin[$DateUtilityMinute] <> $change[$DateUtilityMinute] And _
			$begin[$DateUtilitySecond] <> $change[$DateUtilitySecond] Then
		Sleep($elapsed)

		DateUtility_Restore($handle)
		Local $after = _Date_Time_SystemTimeToArray(_Date_Time_GetLocalTime())
		If $begin[$DateUtilityDay] = $after[$DateUtilityDay] And _
				$begin[$DateUtilityHour] = $after[$DateUtilityHour] And _
				$begin[$DateUtilityMinute] = $after[$DateUtilityMinute] And _
				$begin[$DateUtilitySecond] + ($elapsed / 1000) = $after[$DateUtilitySecond] Then
			$result = True
		EndIf
	EndIf
	Return $result
EndFunc   ;==>DateUtility_Restore_Test
#endregion DateUtility_Restore_Test