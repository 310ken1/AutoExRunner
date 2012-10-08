#include "..\AutoItTest.au3"
#include "..\..\runner\AutoOpenOfficeRunner.au3"
#include "..\..\app\OpenOfficeCalc.au3"
#include "..\..\utility\FileUtility.au3"

#region Constant_Define
; テスト用の無効なファイル.
Const $AutoOpenOfficeRunnerTestInvalidFile = FileUtility_ScriptDirFilePath("AutoOpenOfficeRunnerTest00.ods")
; テスト用の有効なファイル.
Const $AutoOpenOfficeRunnerTestFile1 = FileUtility_ScriptDirFilePath("AutoOpenOfficeRunnerTest01.ods")
; テスト用の有効なファイル.
Const $AutoOpenOfficeRunnerTestFile2 = FileUtility_ScriptDirFilePath("AutoOpenOfficeRunnerTest02.ods")
; テスト用の一時ファイル.
Const $AutoOpenOfficeRunnerTestTempFile = FileUtility_ScriptDirFilePath("AutoOpenOfficeRunnerTestTemp.ods")
; テスト用ファイル(AutoOpenOfficeRunnerTest01.ods)の有効なシート名.
Const $AutoOpenOfficeRunnerTestSheetName = "シート1"
; テスト用ファイル(AutoOpenOfficeRunnerTest01.ods)の無効なシート名.
Const $AutoOpenOfficeRunnerTestInvalidSheetName = "シート10"
; テスト用ファイルの最大項目目のキー.
Const $AutoOpenOfficeRunnerTestMaxKey = "項目10"
; テスト用ファイルの最大+1項目目のキー.
Const $AutoOpenOfficeRunnerTestMaxOverKey = "項目11"
; テスト用ファイルの無効なキー.
Const $AutoOpenOfficeRunnerTestInvalidKey = "項目01"
; テスト用の書き込み文字列.
Const $AutoOpenOfficeRunnerTestString = "Test"
; テスト用ファイル(AutoOpenOfficeRunnerTest01.ods)の行数.
Const $AutoOpenOfficeRunnerRow = 45
; コールバック関数実行時に取得した値(KeyMax).
Const $AutoOpenOfficeRunnerMaxValueAnswer[$AutoOpenOfficeRunnerRow] = [ _
		"項目10_1", "項目10_2", "項目10_3", "項目10_4", "項目10_5", "項目10_6", "項目10_7", "項目10_8", "項目10_9", _
		"項目10_11", "項目10_12", "項目10_13", "項目10_14", "項目10_15", "項目10_16", "項目10_17", "項目10_18", "項目10_19", _
		"項目10_21", "項目10_22", "項目10_23", "項目10_24", "項目10_25", "項目10_26", "項目10_27", "項目10_28", "項目10_29", _
		"項目10_31", "項目10_32", "項目10_33", "項目10_34", "項目10_35", "項目10_36", "項目10_37", "項目10_38", "項目10_39", _
		"項目10_41", "項目10_42", "項目10_43", "項目10_44", "項目10_45", "項目10_46", "項目10_47", "項目10_48", "項目10_49" _
		]
; AutoOpenOfficeRunner_GetCell関数で, 最大+1項目目のキーを指定した場合の答え.
Const $AutoOpenOfficeRunnerGetCellError[$AutoOpenOfficeRunnerRow] = [ _
		0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 _
		]
; AutoOpenOfficeRunner_GetString関数で, 最大+1項目目のキーを指定した場合の答え.
Const $AutoOpenOfficeRunnerGetStringError[$AutoOpenOfficeRunnerRow] = [ _
		"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" _
		]
; 配列の初期値.
Const $AutoOpenOfficeRunnerArrayInit[$AutoOpenOfficeRunnerRow] = [ _
		0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 _
		]
#endregion Constant_Define

#region Static_Argument_Define
; コールバック関数の実行回数.
Static Local $AutoOpenOfficeRunnerCount = 0
; コールバック関数実行時に取得した値.
Static Local $AutoOpenOfficeRunnerValue[$AutoOpenOfficeRunnerRow]
#endregion Static_Argument_Define

Local $AutoOpenOfficeRunnerTest[17][5] = [ _
		["AutoOpenOfficeRunnerTest_After", "AutoOpenOfficeRunner_Open_Test_InvalidFile", "AutoItTest_Assert", 0, "AutoOpenOfficeRunnerTest_Before"], _
		["AutoOpenOfficeRunnerTest_After", "AutoOpenOfficeRunner_Open_Test_MultipleOpen_different", "AutoItTest_WinExists", "AutoOpenOfficeRunnerTest02.ods", "AutoOpenOfficeRunnerTest_Before"], _
		["AutoOpenOfficeRunnerTest_After", "AutoOpenOfficeRunner_Open_Test_MultipleOpen_same", "AutoItTest_IsArray", True, "AutoOpenOfficeRunnerTest_Before"], _
		["AutoOpenOfficeRunnerTest_After", "AutoOpenOfficeRunner_Run_Test_Count", "AutoItTest_Assert", $AutoOpenOfficeRunnerRow, ""], _
		["AutoOpenOfficeRunnerTest_After", "AutoOpenOfficeRunner_Run_Test_InvalidHandle", "AutoItTest_Assert", False, ""], _
		["AutoOpenOfficeRunnerTest_After", "AutoOpenOfficeRunner_Run_Test_InvalidSheetName", "AutoItTest_Assert", False, ""], _
		["AutoOpenOfficeRunnerTest_After", "AutoOpenOfficeRunner_Run_Test_InvalidCallback", "AutoItTest_Assert", False, ""], _
		["AutoOpenOfficeRunnerTest_After", "AutoOpenOfficeRunner_Close_Test_Save", "AutoItTest_Assert", $AutoOpenOfficeRunnerTestString, "AutoOpenOfficeRunnerTest_Before"], _
		["AutoOpenOfficeRunnerTest_After", "AutoOpenOfficeRunner_Close_Test_InvalidHandle", "AutoItTest_WinExists", "AutoOpenOfficeRunnerTest01.ods", "AutoOpenOfficeRunnerTest_Before"], _
		["AutoOpenOfficeRunnerTest_After", "AutoOpenOfficeRunner_GetCell_Test_KeyColumnMax", "AutoOpenOfficeRunner_AssertCellArrayEquals", $AutoOpenOfficeRunnerMaxValueAnswer, "AutoOpenOfficeRunnerTest_Before"], _
		["AutoOpenOfficeRunnerTest_After", "AutoOpenOfficeRunner_GetCell_Test_KeyColumnMaxOver", "AutoItTest_AssertArrayEquals", $AutoOpenOfficeRunnerGetCellError, "AutoOpenOfficeRunnerTest_Before"], _
		["AutoOpenOfficeRunnerTest_After", "AutoOpenOfficeRunner_GetCell_Test_InvalidHandle", "AutoItTest_Assert", 0, "AutoOpenOfficeRunnerTest_Before"], _
		["AutoOpenOfficeRunnerTest_After", "AutoOpenOfficeRunner_GetCell_Test_InvalidKey", "AutoItTest_Assert", 0, "AutoOpenOfficeRunnerTest_Before"], _
		["AutoOpenOfficeRunnerTest_After", "AutoOpenOfficeRunner_GetString_Test_KeyColumnMax", "AutoItTest_AssertArrayEquals", $AutoOpenOfficeRunnerMaxValueAnswer, "AutoOpenOfficeRunnerTest_Before"], _
		["AutoOpenOfficeRunnerTest_After", "AutoOpenOfficeRunner_GetString_Test_KeyColumnMaxOver", "AutoItTest_AssertArrayEquals", $AutoOpenOfficeRunnerGetStringError, "AutoOpenOfficeRunnerTest_Before"], _
		["AutoOpenOfficeRunnerTest_After", "AutoOpenOfficeRunner_GetString_Test_InvalidHandle", "AutoItTest_Assert", "", "AutoOpenOfficeRunnerTest_Before"], _
		["AutoOpenOfficeRunnerTest_After", "AutoOpenOfficeRunner_GetString_Test_InvalidKey", "AutoItTest_Assert", "", "AutoOpenOfficeRunnerTest_Before"] _
		]
AutoItTest_Runner($AutoOpenOfficeRunnerTest)

#region Common
; 前処理.
Func AutoOpenOfficeRunnerTest_After()
	$AutoOpenOfficeRunnerCount = 0
	$AutoOpenOfficeRunnerValue = $AutoOpenOfficeRunnerArrayInit
EndFunc   ;==>AutoOpenOfficeRunnerTest_After
; 後処理.
Func AutoOpenOfficeRunnerTest_Before()
	While WinExists($OpenOfficeCalc_ClassName)
		WinClose($OpenOfficeCalc_ClassName)
	WEnd
	FileDelete($AutoOpenOfficeRunnerTestTempFile)
EndFunc   ;==>AutoOpenOfficeRunnerTest_Before
; コールバック関数.
Func CallBackFunc(Const $handle)
	$AutoOpenOfficeRunnerCount += 1
EndFunc   ;==>CallBackFunc
; セル配列が一致するか判定する.
Func AutoOpenOfficeRunner_AssertCellArrayEquals($expected, $actual, $error = 0, $extended = 0)
	Local $result = True
	For $i = 0 To UBound($expected) - 1
		Local $actual_cell = $actual[$i]
		If $expected[$i] <> $actual_cell.getString() Then
			$result = False
			ExitLoop
		EndIf
	Next
	Return $result
EndFunc   ;==>AutoOpenOfficeRunner_AssertCellArrayEquals
#endregion Common

#region AutoOpenOfficeRunner_Open_Test
; 無効な(存在しない)ファイルを指定した場合.
; 期待する結果: 戻り値として 0 を返す.
Func AutoOpenOfficeRunner_Open_Test_InvalidFile()
	Return AutoOpenOfficeRunner_Open($AutoOpenOfficeRunnerTestInvalidFile)
EndFunc   ;==>AutoOpenOfficeRunner_Open_Test_InvalidFile
; 別々のファイルを多重オープンした場合.
; 期待する結果: ファイルが多重にオープンされる.
Func AutoOpenOfficeRunner_Open_Test_MultipleOpen_different()
	Local $hd1 = AutoOpenOfficeRunner_Open($AutoOpenOfficeRunnerTestFile1)
	Local $hd2 = AutoOpenOfficeRunner_Open($AutoOpenOfficeRunnerTestFile2)
	Return True
EndFunc   ;==>AutoOpenOfficeRunner_Open_Test_MultipleOpen_different
; 同じファイルを多重オープンした場合.
; 期待する結果: 戻り値として ハンドル を返す.
Func AutoOpenOfficeRunner_Open_Test_MultipleOpen_same()
	Local $hd1 = AutoOpenOfficeRunner_Open($AutoOpenOfficeRunnerTestFile1)
	Return AutoOpenOfficeRunner_Open($AutoOpenOfficeRunnerTestFile1)
EndFunc   ;==>AutoOpenOfficeRunner_Open_Test_MultipleOpen_same
#endregion AutoOpenOfficeRunner_Open_Test

#region AutoOpenOfficeRunner_Run_Test
; コールバック関数が, 正しい回数(No項目が空白の場合は除外)実行される.
; 期待する結果: コールバック関数が45回実行される.
Func AutoOpenOfficeRunner_Run_Test_Count()
	Local $hd = AutoOpenOfficeRunner_Open($AutoOpenOfficeRunnerTestFile1)
	AutoOpenOfficeRunner_Run($hd, $AutoOpenOfficeRunnerTestSheetName, "CallBackFunc")
	AutoOpenOfficeRunner_Close($hd)
	Return $AutoOpenOfficeRunnerCount
EndFunc   ;==>AutoOpenOfficeRunner_Run_Test_Count
; 無効なハンドルを指定した場合.
; 期待する結果: 戻り値としてFalseを返す.
Func AutoOpenOfficeRunner_Run_Test_InvalidHandle()
	Return AutoOpenOfficeRunner_Run(0, $AutoOpenOfficeRunnerTestSheetName, "CallBackFunc")
EndFunc   ;==>AutoOpenOfficeRunner_Run_Test_InvalidHandle
; 無効なシート名を指定した場合.
; 期待する結果: 戻り値としてFalseを返す.
Func AutoOpenOfficeRunner_Run_Test_InvalidSheetName()
	Local $hd = AutoOpenOfficeRunner_Open($AutoOpenOfficeRunnerTestFile1)
	Local $result = AutoOpenOfficeRunner_Run($hd, $AutoOpenOfficeRunnerTestInvalidSheetName, "CallBackFunc")
	AutoOpenOfficeRunner_Close($hd)
	Return $result
EndFunc   ;==>AutoOpenOfficeRunner_Run_Test_InvalidSheetName
; 無効なコールバック関数を指定した場合.
; 期待する結果: 戻り値としてFalseを返す.
Func AutoOpenOfficeRunner_Run_Test_InvalidCallback()
	Local $hd = AutoOpenOfficeRunner_Open($AutoOpenOfficeRunnerTestFile1)
	Local $result = AutoOpenOfficeRunner_Run($hd, $AutoOpenOfficeRunnerTestInvalidSheetName, "CallBackFunc1")
	AutoOpenOfficeRunner_Close($hd)
	Return $result
EndFunc   ;==>AutoOpenOfficeRunner_Run_Test_InvalidCallback
#endregion AutoOpenOfficeRunner_Run_Test

#region AutoOpenOfficeRunner_Close_Test
; $save=Trueを指定した場合.
; 期待する結果: ファイルが保存される.
Func AutoOpenOfficeRunner_Close_Test_Save()
	FileCopy($AutoOpenOfficeRunnerTestFile2, $AutoOpenOfficeRunnerTestTempFile)
	Local $hd = AutoOpenOfficeRunner_Open($AutoOpenOfficeRunnerTestTempFile)
	AutoOpenOfficeRunner_Run($hd, $AutoOpenOfficeRunnerTestSheetName, _
			"AutoOpenOfficeRunner_Close_Test_Save_CallBackFunc")
	AutoOpenOfficeRunner_Close($hd, True)

	Local $document = OpenOfficeCalc_Open($AutoOpenOfficeRunnerTestTempFile)
	Local $sheet = OpenOfficeCalc_GetSheet($AutoOpenOfficeRunnerTestSheetName)
	Local $cell = OpenOfficeCalc_GetCell(2, 10)
	Return $cell.getString
EndFunc   ;==>AutoOpenOfficeRunner_Close_Test_Save
; コールバック関数.
Func AutoOpenOfficeRunner_Close_Test_Save_CallBackFunc(Const $handle)
	Local $cell = AutoOpenOfficeRunner_GetCell($handle, $AutoOpenOfficeRunnerTestMaxKey)
	$cell.setString($AutoOpenOfficeRunnerTestString)
EndFunc   ;==>AutoOpenOfficeRunner_Close_Test_Save_CallBackFunc
; 無効なハンドルを指定した場合.
; 期待する結果: 何もしない(OpenOfficeをクローズしない).
Func AutoOpenOfficeRunner_Close_Test_InvalidHandle()
	Local $hd = AutoOpenOfficeRunner_Open($AutoOpenOfficeRunnerTestFile1)
	AutoOpenOfficeRunner_Close(0)
	Return True
EndFunc   ;==>AutoOpenOfficeRunner_Close_Test_InvalidHandle
#endregion AutoOpenOfficeRunner_Close_Test

#region AutoOpenOfficeRunner_GetCell_Test
; キーの最大個数目を指定した場合.
; 期待する結果: 戻り値として セルオブジェクト を返す.
Func AutoOpenOfficeRunner_GetCell_Test_KeyColumnMax()
	Local $hd = AutoOpenOfficeRunner_Open($AutoOpenOfficeRunnerTestFile1)
	AutoOpenOfficeRunner_Run($hd, $AutoOpenOfficeRunnerTestSheetName, _
			"AutoOpenOfficeRunner_GetCell_Test_KeyColumnMax_CallBackFunc")
	Return $AutoOpenOfficeRunnerValue
EndFunc   ;==>AutoOpenOfficeRunner_GetCell_Test_KeyColumnMax
; コールバック関数.
Func AutoOpenOfficeRunner_GetCell_Test_KeyColumnMax_CallBackFunc(Const $handle)
	$AutoOpenOfficeRunnerValue[$AutoOpenOfficeRunnerCount] = _
			AutoOpenOfficeRunner_GetCell($handle, $AutoOpenOfficeRunnerTestMaxKey)
	$AutoOpenOfficeRunnerCount += 1
EndFunc   ;==>AutoOpenOfficeRunner_GetCell_Test_KeyColumnMax_CallBackFunc
; キーの最大+1個数目を指定した場合.
; 期待する結果: 戻り値として 0 を返す.
Func AutoOpenOfficeRunner_GetCell_Test_KeyColumnMaxOver()
	Local $hd = AutoOpenOfficeRunner_Open($AutoOpenOfficeRunnerTestFile1)
	AutoOpenOfficeRunner_Run($hd, $AutoOpenOfficeRunnerTestSheetName, _
			"AutoOpenOfficeRunner_GetCell_Test_KeyColumnMaxOver_CallBackFunc")
	Return $AutoOpenOfficeRunnerValue
EndFunc   ;==>AutoOpenOfficeRunner_GetCell_Test_KeyColumnMaxOver
; コールバック関数.
Func AutoOpenOfficeRunner_GetCell_Test_KeyColumnMaxOver_CallBackFunc(Const $handle)
	$AutoOpenOfficeRunnerValue[$AutoOpenOfficeRunnerCount] = _
			AutoOpenOfficeRunner_GetCell($handle, $AutoOpenOfficeRunnerTestMaxOverKey)
	$AutoOpenOfficeRunnerCount += 1
EndFunc   ;==>AutoOpenOfficeRunner_GetCell_Test_KeyColumnMaxOver_CallBackFunc
; 無効なハンドルを指定した場合.
; 期待する結果: 戻り値として 0 を返す.
Func AutoOpenOfficeRunner_GetCell_Test_InvalidHandle()
	Local $hd = AutoOpenOfficeRunner_Open($AutoOpenOfficeRunnerTestFile2)
	AutoOpenOfficeRunner_Run($hd, $AutoOpenOfficeRunnerTestSheetName, _
			"AutoOpenOfficeRunner_GetCell_Test_InvalidHandle_CallBackFunc")
	Return $AutoOpenOfficeRunnerValue[0]
EndFunc   ;==>AutoOpenOfficeRunner_GetCell_Test_InvalidHandle
; コールバック関数.
Func AutoOpenOfficeRunner_GetCell_Test_InvalidHandle_CallBackFunc(Const $handle)
	$AutoOpenOfficeRunnerValue[$AutoOpenOfficeRunnerCount] = _
			AutoOpenOfficeRunner_GetCell(0, $AutoOpenOfficeRunnerTestMaxKey)
	$AutoOpenOfficeRunnerCount += 1
EndFunc   ;==>AutoOpenOfficeRunner_GetCell_Test_InvalidHandle_CallBackFunc
; 無効なキーを指定した場合.
; 期待する結果: 戻り値として 0 を返す.
Func AutoOpenOfficeRunner_GetCell_Test_InvalidKey()
	Local $hd = AutoOpenOfficeRunner_Open($AutoOpenOfficeRunnerTestFile2)
	AutoOpenOfficeRunner_Run($hd, $AutoOpenOfficeRunnerTestSheetName, _
			"AutoOpenOfficeRunner_GetCell_Test_InvalidHandle_CallBackFunc")
	Return $AutoOpenOfficeRunnerValue[0]
EndFunc   ;==>AutoOpenOfficeRunner_GetCell_Test_InvalidKey
; コールバック関数.
Func AutoOpenOfficeRunner_GetCell_Test_InvalidKey_CallBackFunc(Const $handle)
	$AutoOpenOfficeRunnerValue[$AutoOpenOfficeRunnerCount] = _
			AutoOpenOfficeRunner_GetCell($handle, $AutoOpenOfficeRunnerTestInvalidKey)
	$AutoOpenOfficeRunnerCount += 1
EndFunc   ;==>AutoOpenOfficeRunner_GetCell_Test_InvalidKey_CallBackFunc
#endregion AutoOpenOfficeRunner_GetCell_Test

#region AutoOpenOfficeRunner_GetString_Test
; キーの最大個数目を指定した場合.
; 期待する結果: 戻り値として 値() を返す.
Func AutoOpenOfficeRunner_GetString_Test_KeyColumnMax()
	Local $hd = AutoOpenOfficeRunner_Open($AutoOpenOfficeRunnerTestFile1)
	AutoOpenOfficeRunner_Run($hd, $AutoOpenOfficeRunnerTestSheetName, _
			"AutoOpenOfficeRunner_GetString_Test_KeyColumnMax_CallBackFunc")
	Return $AutoOpenOfficeRunnerValue
EndFunc   ;==>AutoOpenOfficeRunner_GetString_Test_KeyColumnMax
; コールバック関数.
Func AutoOpenOfficeRunner_GetString_Test_KeyColumnMax_CallBackFunc(Const $handle)
	$AutoOpenOfficeRunnerValue[$AutoOpenOfficeRunnerCount] = _
			AutoOpenOfficeRunner_GetString($handle, $AutoOpenOfficeRunnerTestMaxKey)
	$AutoOpenOfficeRunnerCount += 1
EndFunc   ;==>AutoOpenOfficeRunner_GetString_Test_KeyColumnMax_CallBackFunc
; キーの最大+1個数目を指定した場合.
; 期待する結果: 戻り値として 0 を返す.
Func AutoOpenOfficeRunner_GetString_Test_KeyColumnMaxOver()
	Local $hd = AutoOpenOfficeRunner_Open($AutoOpenOfficeRunnerTestFile1)
	AutoOpenOfficeRunner_Run($hd, $AutoOpenOfficeRunnerTestSheetName, _
			"AutoOpenOfficeRunner_GetString_Test_KeyColumnMaxOver_CallBackFunc")
	Return $AutoOpenOfficeRunnerValue
EndFunc   ;==>AutoOpenOfficeRunner_GetString_Test_KeyColumnMaxOver
; コールバック関数.
Func AutoOpenOfficeRunner_GetString_Test_KeyColumnMaxOver_CallBackFunc(Const $handle)
	$AutoOpenOfficeRunnerValue[$AutoOpenOfficeRunnerCount] = _
			AutoOpenOfficeRunner_GetString($handle, $AutoOpenOfficeRunnerTestMaxOverKey)
	$AutoOpenOfficeRunnerCount += 1
EndFunc   ;==>AutoOpenOfficeRunner_GetString_Test_KeyColumnMaxOver_CallBackFunc
; 無効なハンドルを指定した場合.
; 期待する結果: 戻り値として 0 を返す.
Func AutoOpenOfficeRunner_GetString_Test_InvalidHandle()
	Local $hd = AutoOpenOfficeRunner_Open($AutoOpenOfficeRunnerTestFile2)
	AutoOpenOfficeRunner_Run($hd, $AutoOpenOfficeRunnerTestSheetName, _
			"AutoOpenOfficeRunner_GetString_Test_InvalidHandle_CallBackFunc")
	Return $AutoOpenOfficeRunnerValue[0]
EndFunc   ;==>AutoOpenOfficeRunner_GetString_Test_InvalidHandle
; コールバック関数.
Func AutoOpenOfficeRunner_GetString_Test_InvalidHandle_CallBackFunc(Const $handle)
	$AutoOpenOfficeRunnerValue[$AutoOpenOfficeRunnerCount] = _
			AutoOpenOfficeRunner_GetString($handle, $AutoOpenOfficeRunnerTestMaxOverKey)
	$AutoOpenOfficeRunnerCount += 1
EndFunc   ;==>AutoOpenOfficeRunner_GetString_Test_InvalidHandle_CallBackFunc
; 無効なキーを指定した場合.
; 期待する結果: 戻り値として 0 を返す.
Func AutoOpenOfficeRunner_GetString_Test_InvalidKey()
	Local $hd = AutoOpenOfficeRunner_Open($AutoOpenOfficeRunnerTestFile2)
	AutoOpenOfficeRunner_Run($hd, $AutoOpenOfficeRunnerTestSheetName, _
			"AutoOpenOfficeRunner_GetString_Test_InvalidKey_CallBackFunc")
	Return $AutoOpenOfficeRunnerValue[0]
EndFunc   ;==>AutoOpenOfficeRunner_GetString_Test_InvalidKey
; コールバック関数.
Func AutoOpenOfficeRunner_GetString_Test_InvalidKey_CallBackFunc(Const $handle)
	$AutoOpenOfficeRunnerValue[$AutoOpenOfficeRunnerCount] = _
			AutoOpenOfficeRunner_GetString($handle, $AutoOpenOfficeRunnerTestMaxOverKey)
	$AutoOpenOfficeRunnerCount += 1
EndFunc   ;==>AutoOpenOfficeRunner_GetString_Test_InvalidKey_CallBackFunc
#endregion AutoOpenOfficeRunner_GetString_Test