#include "..\AutoItTest.au3"
#include "..\..\runner\AutoExcelRunner.au3"
#include "..\..\utility\FileUtility.au3"

#region Constant_Define
; Excel のプロセス名.
Const $Excel_ProcessName = "EXCEL.EXE"
; テスト用の無効なファイル.
Const $AutoExcelRunnerTestInvalidFile = FileUtility_ScriptDirFilePath("AutoExcelRunnerTest00.xls")
; テスト用の有効なファイル.
Const $AutoExcelRunnerTestFile1 = FileUtility_ScriptDirFilePath("AutoExcelRunnerTest01.xls")
; テスト用の有効なファイル.
Const $AutoExcelRunnerTestFile2 = FileUtility_ScriptDirFilePath("AutoExcelRunnerTest02.xls")
; テスト用の一時ファイル.
Const $AutoExcelRunnerTestTempFile = FileUtility_ScriptDirFilePath("AutoExcelRunnerTestTemp.xls")
; テスト用ファイル(AutoExcelRunnerTest01.ods)の有効なシート名.
Const $AutoExcelRunnerTestSheetName = "シート1"
; テスト用ファイル(AutoExcelRunnerTest01.ods)の無効なシート名.
Const $AutoExcelRunnerTestInvalidSheetName = "シート10"
; テスト用ファイルの最大項目目のキー.
Const $AutoExcelRunnerTestMaxKey = "項目10"
; テスト用ファイルの最大+1項目目のキー.
Const $AutoExcelRunnerTestMaxOverKey = "項目11"
; テスト用ファイルの無効なキー.
Const $AutoExcelRunnerTestInvalidKey = "項目01"
; テスト用の書き込み文字列.
Const $AutoExcelRunnerTestString = "Test"
; テスト用ファイル(AutoExcelRunnerTest01.ods)の行数.
Const $AutoExcelRunnerRow = 45
; コールバック関数実行時に取得した値(KeyMax).
Const $AutoExcelRunnerMaxValueAnswer[$AutoExcelRunnerRow] = [ _
		"項目10_1", "項目10_2", "項目10_3", "項目10_4", "項目10_5", "項目10_6", "項目10_7", "項目10_8", "項目10_9", _
		"項目10_11", "項目10_12", "項目10_13", "項目10_14", "項目10_15", "項目10_16", "項目10_17", "項目10_18", "項目10_19", _
		"項目10_21", "項目10_22", "項目10_23", "項目10_24", "項目10_25", "項目10_26", "項目10_27", "項目10_28", "項目10_29", _
		"項目10_31", "項目10_32", "項目10_33", "項目10_34", "項目10_35", "項目10_36", "項目10_37", "項目10_38", "項目10_39", _
		"項目10_41", "項目10_42", "項目10_43", "項目10_44", "項目10_45", "項目10_46", "項目10_47", "項目10_48", "項目10_49" _
		]
; AutoExcelRunner_GetCell関数で, 最大+1項目目のキーを指定した場合の答え.
Const $AutoExcelRunnerGetCellError[$AutoExcelRunnerRow] = [ _
		0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 _
		]
; AutoExcelRunner_GetString関数で, 最大+1項目目のキーを指定した場合の答え.
Const $AutoExcelRunnerGetStringError[$AutoExcelRunnerRow] = [ _
		"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" _
		]
; 配列の初期値.
Const $AutoExcelRunnerArrayInit[$AutoExcelRunnerRow] = [ _
		0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 _
		]
#endregion Constant_Define

#region Static_Argument_Define
; コールバック関数の実行回数.
Static Local $AutoExcelRunnerCount = 0
; コールバック関数実行時に取得した値.
Static Local $AutoExcelRunnerValue[$AutoExcelRunnerRow]
#endregion Static_Argument_Define

Local $AutoExcelRunnerTest[17][5] = [ _
		["AutoExcelRunnerTest_After", "AutoExcelRunner_Open_Test_InvalidFile", "AutoItTest_Assert", 0, "AutoExcelRunnerTest_Before"], _
		["AutoExcelRunnerTest_After", "AutoExcelRunner_Open_Test_MultipleOpen_different", "AutoItTest_WinExists", "Microsoft Excel - AutoExcelRunnerTest02.xls", "AutoExcelRunnerTest_Before"], _
		["AutoExcelRunnerTest_After", "AutoExcelRunner_Open_Test_MultipleOpen_same", "AutoItTest_IsArray", True, "AutoExcelRunnerTest_Before"], _
		["AutoExcelRunnerTest_After", "AutoExcelRunner_Run_Test_Count", "AutoItTest_Assert", $AutoExcelRunnerRow, ""], _
		["AutoExcelRunnerTest_After", "AutoExcelRunner_Run_Test_InvalidHandle", "AutoItTest_Assert", False, ""], _
		["AutoExcelRunnerTest_After", "AutoExcelRunner_Run_Test_InvalidSheetName", "AutoItTest_Assert", False, ""], _
		["AutoExcelRunnerTest_After", "AutoExcelRunner_Run_Test_InvalidCallback", "AutoItTest_Assert", False, ""], _
		["AutoExcelRunnerTest_After", "AutoExcelRunner_Close_Test_Save", "AutoItTest_Assert", $AutoExcelRunnerTestString, "AutoExcelRunnerTest_Before"], _
		["AutoExcelRunnerTest_After", "AutoExcelRunner_Close_Test_InvalidHandle", "AutoItTest_WinExists", "Microsoft Excel - AutoExcelRunnerTest01.xls", "AutoExcelRunnerTest_Before"], _
		["AutoExcelRunnerTest_After", "AutoExcelRunner_GetCell_Test_KeyColumnMax", "AutoExcelRunner_AssertCellArrayEquals", $AutoExcelRunnerMaxValueAnswer, "AutoExcelRunnerTest_Before"], _
		["AutoExcelRunnerTest_After", "AutoExcelRunner_GetCell_Test_KeyColumnMaxOver", "AutoItTest_AssertArrayEquals", $AutoExcelRunnerGetCellError, "AutoExcelRunnerTest_Before"], _
		["AutoExcelRunnerTest_After", "AutoExcelRunner_GetCell_Test_InvalidHandle", "AutoItTest_Assert", 0, "AutoExcelRunnerTest_Before"], _
		["AutoExcelRunnerTest_After", "AutoExcelRunner_GetCell_Test_InvalidKey", "AutoItTest_Assert", 0, "AutoExcelRunnerTest_Before"], _
		["AutoExcelRunnerTest_After", "AutoExcelRunner_GetString_Test_KeyColumnMax", "AutoItTest_AssertArrayEquals", $AutoExcelRunnerMaxValueAnswer, "AutoExcelRunnerTest_Before"], _
		["AutoExcelRunnerTest_After", "AutoExcelRunner_GetString_Test_KeyColumnMaxOver", "AutoItTest_AssertArrayEquals", $AutoExcelRunnerGetStringError, "AutoExcelRunnerTest_Before"], _
		["AutoExcelRunnerTest_After", "AutoExcelRunner_GetString_Test_InvalidHandle", "AutoItTest_Assert", "", "AutoExcelRunnerTest_Before"], _
		["AutoExcelRunnerTest_After", "AutoExcelRunner_GetString_Test_InvalidKey", "AutoItTest_Assert", "", "AutoExcelRunnerTest_Before"] _
		]
AutoItTest_Runner($AutoExcelRunnerTest)

#region Common
; 前処理.
Func AutoExcelRunnerTest_After()
	$AutoExcelRunnerCount = 0
	$AutoExcelRunnerValue = $AutoExcelRunnerArrayInit
EndFunc   ;==>AutoExcelRunnerTest_After
; 後処理.
Func AutoExcelRunnerTest_Before()
	While ProcessExists($Excel_ProcessName)
		ProcessClose($Excel_ProcessName)
	WEnd
	FileDelete($AutoExcelRunnerTestTempFile)
EndFunc   ;==>AutoExcelRunnerTest_Before
; コールバック関数.
Func AutoExcelRunnerTest_CallBackFunc(Const $handle)
	$AutoExcelRunnerCount += 1
EndFunc   ;==>AutoExcelRunnerTest_CallBackFunc
; セル配列が一致するか判定する.
Func AutoExcelRunner_AssertCellArrayEquals($expected, $actual, $error = 0, $extended = 0)
	Local $result = True
	For $i = 0 To UBound($expected) - 1
		Local $actual_cell = $actual[$i]
		If $expected[$i] <> $actual_cell.value Then
			ConsoleWrite($actual_cell.value & @CRLF)
			$result = False
			ExitLoop
		EndIf
	Next
	Return $result
EndFunc   ;==>AutoExcelRunner_AssertCellArrayEquals
#endregion Common

#region AutoExcelRunner_Open_Test
; 無効な(存在しない)ファイルを指定した場合.
; 期待する結果: 戻り値として 0 を返す.
Func AutoExcelRunner_Open_Test_InvalidFile()
	Return AutoExcelRunner_Open($AutoExcelRunnerTestInvalidFile)
EndFunc   ;==>AutoExcelRunner_Open_Test_InvalidFile
; 別々のファイルを多重オープンした場合.
; 期待する結果: ファイルが多重にオープンされる.
Func AutoExcelRunner_Open_Test_MultipleOpen_different()
	Local $hd1 = AutoExcelRunner_Open($AutoExcelRunnerTestFile1)
	Local $hd2 = AutoExcelRunner_Open($AutoExcelRunnerTestFile2)
	Return True
EndFunc   ;==>AutoExcelRunner_Open_Test_MultipleOpen_different
; 同じファイルを多重オープンした場合.
; 期待する結果: 戻り値として ハンドル を返す.
Func AutoExcelRunner_Open_Test_MultipleOpen_same()
	Local $hd1 = AutoExcelRunner_Open($AutoExcelRunnerTestFile1)
	Return AutoExcelRunner_Open($AutoExcelRunnerTestFile1)
EndFunc   ;==>AutoExcelRunner_Open_Test_MultipleOpen_same
#endregion AutoExcelRunner_Open_Test

#region AutoExcelRunner_Run_Test
; コールバック関数が, 正しい回数(No項目が空白の場合は除外)実行される.
; 期待する結果: コールバック関数が45回実行される.
Func AutoExcelRunner_Run_Test_Count()
	Local $hd = AutoExcelRunner_Open($AutoExcelRunnerTestFile1)
	AutoExcelRunner_Run($hd, $AutoExcelRunnerTestSheetName, _
			"AutoExcelRunnerTest_CallBackFunc")
	AutoExcelRunner_Close($hd)
	Return $AutoExcelRunnerCount
EndFunc   ;==>AutoExcelRunner_Run_Test_Count
; 無効なハンドルを指定した場合.
; 期待する結果: 戻り値としてFalseを返す.
Func AutoExcelRunner_Run_Test_InvalidHandle()
	Return AutoExcelRunner_Run(0, $AutoExcelRunnerTestSheetName, _
			"AutoExcelRunnerTest_CallBackFunc")
EndFunc   ;==>AutoExcelRunner_Run_Test_InvalidHandle
; 無効なシート名を指定した場合.
; 期待する結果: 戻り値としてFalseを返す.
Func AutoExcelRunner_Run_Test_InvalidSheetName()
	Local $hd = AutoExcelRunner_Open($AutoExcelRunnerTestFile1)
	Local $result = AutoExcelRunner_Run($hd, $AutoExcelRunnerTestInvalidSheetName, _
			"AutoExcelRunnerTest_CallBackFunc")
	AutoExcelRunner_Close($hd)
	Return $result
EndFunc   ;==>AutoExcelRunner_Run_Test_InvalidSheetName
; 無効なコールバック関数を指定した場合.
; 期待する結果: 戻り値としてFalseを返す.
Func AutoExcelRunner_Run_Test_InvalidCallback()
	Local $hd = AutoExcelRunner_Open($AutoExcelRunnerTestFile1)
	Local $result = AutoExcelRunner_Run($hd, $AutoExcelRunnerTestInvalidSheetName, _
			"AutoExcelRunnerTest_CallBackFunc1")
	AutoExcelRunner_Close($hd)
	Return $result
EndFunc   ;==>AutoExcelRunner_Run_Test_InvalidCallback
#endregion AutoExcelRunner_Run_Test

#region AutoExcelRunner_Close_Test
; $save=Trueを指定した場合.
; 期待する結果: ファイルが保存される.
Func AutoExcelRunner_Close_Test_Save()
	FileCopy($AutoExcelRunnerTestFile2, $AutoExcelRunnerTestTempFile)
	Local $hd = AutoExcelRunner_Open($AutoExcelRunnerTestTempFile)
	AutoExcelRunner_Run($hd, $AutoExcelRunnerTestSheetName, _
			"AutoExcelRunner_Close_Test_Save_CallBackFunc")
	AutoExcelRunner_Close($hd, True)

	Local $document = ObjGet($AutoExcelRunnerTestTempFile, "Excel.Application")
	Local $sheet = $document.Worksheets($AutoExcelRunnerTestSheetName)
	Return $sheet.cells(3, 11).value
EndFunc   ;==>AutoExcelRunner_Close_Test_Save
; コールバック関数.
Func AutoExcelRunner_Close_Test_Save_CallBackFunc(Const $handle)
	Local $cell = AutoExcelRunner_GetCell($handle, $AutoExcelRunnerTestMaxKey)
	$cell.value = $AutoExcelRunnerTestString
EndFunc   ;==>AutoExcelRunner_Close_Test_Save_CallBackFunc
; 無効なハンドルを指定した場合.
; 期待する結果: 何もしない(OpenOfficeをクローズしない).
Func AutoExcelRunner_Close_Test_InvalidHandle()
	Local $hd = AutoExcelRunner_Open($AutoExcelRunnerTestFile1)
	AutoExcelRunner_Close(0)
	Return True
EndFunc   ;==>AutoExcelRunner_Close_Test_InvalidHandle
#endregion AutoExcelRunner_Close_Test

#region AutoExcelRunner_GetCell_Test
; キーの最大個数目を指定した場合.
; 期待する結果: 戻り値として セルオブジェクト を返す.
Func AutoExcelRunner_GetCell_Test_KeyColumnMax()
	Local $hd = AutoExcelRunner_Open($AutoExcelRunnerTestFile1)
	AutoExcelRunner_Run($hd, $AutoExcelRunnerTestSheetName, _
			"AutoExcelRunner_GetCell_Test_KeyColumnMax_CallBackFunc")
	Return $AutoExcelRunnerValue
EndFunc   ;==>AutoExcelRunner_GetCell_Test_KeyColumnMax
; コールバック関数.
Func AutoExcelRunner_GetCell_Test_KeyColumnMax_CallBackFunc(Const $handle)
	$AutoExcelRunnerValue[$AutoExcelRunnerCount] = _
			AutoExcelRunner_GetCell($handle, $AutoExcelRunnerTestMaxKey)
	$AutoExcelRunnerCount += 1
EndFunc   ;==>AutoExcelRunner_GetCell_Test_KeyColumnMax_CallBackFunc
; キーの最大+1個数目を指定した場合.
; 期待する結果: 戻り値として 0 を返す.
Func AutoExcelRunner_GetCell_Test_KeyColumnMaxOver()
	Local $hd = AutoExcelRunner_Open($AutoExcelRunnerTestFile1)
	AutoExcelRunner_Run($hd, $AutoExcelRunnerTestSheetName, _
			"AutoExcelRunner_GetCell_Test_KeyColumnMaxOver_CallBackFunc")
	Return $AutoExcelRunnerValue
EndFunc   ;==>AutoExcelRunner_GetCell_Test_KeyColumnMaxOver
; コールバック関数.
Func AutoExcelRunner_GetCell_Test_KeyColumnMaxOver_CallBackFunc(Const $handle)
	$AutoExcelRunnerValue[$AutoExcelRunnerCount] = _
			AutoExcelRunner_GetCell($handle, $AutoExcelRunnerTestMaxOverKey)
	$AutoExcelRunnerCount += 1
EndFunc   ;==>AutoExcelRunner_GetCell_Test_KeyColumnMaxOver_CallBackFunc
; 無効なハンドルを指定した場合.
; 期待する結果: 戻り値として 0 を返す.
Func AutoExcelRunner_GetCell_Test_InvalidHandle()
	Local $hd = AutoExcelRunner_Open($AutoExcelRunnerTestFile2)
	AutoExcelRunner_Run($hd, $AutoExcelRunnerTestSheetName, _
			"AutoExcelRunner_GetCell_Test_InvalidHandle_CallBackFunc")
	Return $AutoExcelRunnerValue[0]
EndFunc   ;==>AutoExcelRunner_GetCell_Test_InvalidHandle
; コールバック関数.
Func AutoExcelRunner_GetCell_Test_InvalidHandle_CallBackFunc(Const $handle)
	$AutoExcelRunnerValue[$AutoExcelRunnerCount] = _
			AutoExcelRunner_GetCell(0, $AutoExcelRunnerTestMaxKey)
	$AutoExcelRunnerCount += 1
EndFunc   ;==>AutoExcelRunner_GetCell_Test_InvalidHandle_CallBackFunc
; 無効なキーを指定した場合.
; 期待する結果: 戻り値として 0 を返す.
Func AutoExcelRunner_GetCell_Test_InvalidKey()
	Local $hd = AutoExcelRunner_Open($AutoExcelRunnerTestFile2)
	AutoExcelRunner_Run($hd, $AutoExcelRunnerTestSheetName, _
			"AutoExcelRunner_GetCell_Test_InvalidHandle_CallBackFunc")
	Return $AutoExcelRunnerValue[0]
EndFunc   ;==>AutoExcelRunner_GetCell_Test_InvalidKey
; コールバック関数.
Func AutoExcelRunner_GetCell_Test_InvalidKey_CallBackFunc(Const $handle)
	$AutoExcelRunnerValue[$AutoExcelRunnerCount] = _
			AutoExcelRunner_GetCell($handle, $AutoExcelRunnerTestInvalidKey)
	$AutoExcelRunnerCount += 1
EndFunc   ;==>AutoExcelRunner_GetCell_Test_InvalidKey_CallBackFunc
#endregion AutoExcelRunner_GetCell_Test

#region AutoExcelRunner_GetString_Test
; キーの最大個数目を指定した場合.
; 期待する結果: 戻り値として 値() を返す.
Func AutoExcelRunner_GetString_Test_KeyColumnMax()
	Local $hd = AutoExcelRunner_Open($AutoExcelRunnerTestFile1)
	AutoExcelRunner_Run($hd, $AutoExcelRunnerTestSheetName, _
			"AutoExcelRunner_GetString_Test_KeyColumnMax_CallBackFunc")
	Return $AutoExcelRunnerValue
EndFunc   ;==>AutoExcelRunner_GetString_Test_KeyColumnMax
; コールバック関数.
Func AutoExcelRunner_GetString_Test_KeyColumnMax_CallBackFunc(Const $handle)
	$AutoExcelRunnerValue[$AutoExcelRunnerCount] = _
			AutoExcelRunner_GetString($handle, $AutoExcelRunnerTestMaxKey)
	$AutoExcelRunnerCount += 1
EndFunc   ;==>AutoExcelRunner_GetString_Test_KeyColumnMax_CallBackFunc
; キーの最大+1個数目を指定した場合.
; 期待する結果: 戻り値として 0 を返す.
Func AutoExcelRunner_GetString_Test_KeyColumnMaxOver()
	Local $hd = AutoExcelRunner_Open($AutoExcelRunnerTestFile1)
	AutoExcelRunner_Run($hd, $AutoExcelRunnerTestSheetName, _
			"AutoExcelRunner_GetString_Test_KeyColumnMaxOver_CallBackFunc")
	Return $AutoExcelRunnerValue
EndFunc   ;==>AutoExcelRunner_GetString_Test_KeyColumnMaxOver
; コールバック関数.
Func AutoExcelRunner_GetString_Test_KeyColumnMaxOver_CallBackFunc(Const $handle)
	$AutoExcelRunnerValue[$AutoExcelRunnerCount] = _
			AutoExcelRunner_GetString($handle, $AutoExcelRunnerTestMaxOverKey)
	$AutoExcelRunnerCount += 1
EndFunc   ;==>AutoExcelRunner_GetString_Test_KeyColumnMaxOver_CallBackFunc
; 無効なハンドルを指定した場合.
; 期待する結果: 戻り値として 0 を返す.
Func AutoExcelRunner_GetString_Test_InvalidHandle()
	Local $hd = AutoExcelRunner_Open($AutoExcelRunnerTestFile2)
	AutoExcelRunner_Run($hd, $AutoExcelRunnerTestSheetName, _
			"AutoExcelRunner_GetString_Test_InvalidHandle_CallBackFunc")
	Return $AutoExcelRunnerValue[0]
EndFunc   ;==>AutoExcelRunner_GetString_Test_InvalidHandle
; コールバック関数.
Func AutoExcelRunner_GetString_Test_InvalidHandle_CallBackFunc(Const $handle)
	$AutoExcelRunnerValue[$AutoExcelRunnerCount] = _
			AutoExcelRunner_GetString($handle, $AutoExcelRunnerTestMaxOverKey)
	$AutoExcelRunnerCount += 1
EndFunc   ;==>AutoExcelRunner_GetString_Test_InvalidHandle_CallBackFunc
; 無効なキーを指定した場合.
; 期待する結果: 戻り値として 0 を返す.
Func AutoExcelRunner_GetString_Test_InvalidKey()
	Local $hd = AutoExcelRunner_Open($AutoExcelRunnerTestFile2)
	AutoExcelRunner_Run($hd, $AutoExcelRunnerTestSheetName, _
			"AutoExcelRunner_GetString_Test_InvalidKey_CallBackFunc")
	Return $AutoExcelRunnerValue[0]
EndFunc   ;==>AutoExcelRunner_GetString_Test_InvalidKey
; コールバック関数.
Func AutoExcelRunner_GetString_Test_InvalidKey_CallBackFunc(Const $handle)
	$AutoExcelRunnerValue[$AutoExcelRunnerCount] = _
			AutoExcelRunner_GetString($handle, $AutoExcelRunnerTestMaxOverKey)
	$AutoExcelRunnerCount += 1
EndFunc   ;==>AutoExcelRunner_GetString_Test_InvalidKey_CallBackFunc
#endregion AutoExcelRunner_GetString_Test