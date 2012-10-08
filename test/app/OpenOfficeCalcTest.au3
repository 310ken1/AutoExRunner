#include "..\AutoItTest.au3"
#include "..\..\app\OpenOfficeCalc.au3"
#include "..\..\utility\ConsoleUtility.au3"
#include "..\..\utility\FileUtility.au3"

#region Constant_Define
; 無効なファイル名.
Const $InvalidFile = FileUtility_ScriptDirFilePath("OpenOfficeCalcTest0.ods")
; 有効なファイル名.
Const $ValidFile1 = FileUtility_ScriptDirFilePath("OpenOfficeCalcTest1.ods")
; 有効なファイル名.
Const $ValidFile2 = FileUtility_ScriptDirFilePath("OpenOfficeCalcTest2.ods")
; テスト用の一時ファイル名.
Const $TempFile = FileUtility_ScriptDirFilePath("OpenOfficeCalcTemp.ods")
; 有効なシート名.
Const $ValidSheet = "シート１"
; 無効なシート名.
Const $InvalidSheet = "Sheet1"
; デフォルトのシート名.
Const $DefaltSheet = "Sheet1"
; セル(0,0)の値.
Const $Cell00 = "シート１_セル00"
; テスト用の書き込み文字列.
Const $TestString = "Test"
#endregion Constant_Define

Local $OpenOfficeCalcTest[20][5] = [ _
		["", "OpenOfficeCalc_Open_Test_ValidFile", "AutoItTest_IsObj", True, "OpenOfficeCalc_Close"], _
		["", "OpenOfficeCalc_Open_Test_InvalidFile", "AutoItTest_FileExists", $InvalidFile, "OpenOfficeCalc_Open_Test_InvalidFile_After"], _
		["", "OpenOfficeCalc_Close_Test_Document", "AutoItTest_WinNotExists", $OpenOfficeCalc_ClassName, ""], _
		["", "OpenOfficeCalc_Close_Test_NoDocument", "AutoItTest_WinNotExists", $OpenOfficeCalc_ClassName, ""], _
		["", "OpenOfficeCalc_Close_Test_InvalidDocument", "AutoItTest_WinNotExists", $OpenOfficeCalc_ClassName, ""], _
		["OpenOfficeCalc_Save_Test_Before", "OpenOfficeCalc_Save_Test_NoArgument", "AutoItTest_Assert", $TestString, "OpenOfficeCalc_Save_Test_After"], _
		["OpenOfficeCalc_Save_Test_Before", "OpenOfficeCalc_Save_Test_InvalidFile", "AutoItTest_Assert", $TestString, "OpenOfficeCalc_Save_Test_After"], _
		["OpenOfficeCalc_Save_Test_Before", "OpenOfficeCalc_Save_Test_ValidFile_Own", "AutoItTest_Assert", $TestString, "OpenOfficeCalc_Save_Test_After"], _
		["OpenOfficeCalc_Save_Test_Before", "OpenOfficeCalc_Save_Test_ValidFile_Other", "AutoItTest_Assert", $TestString, "OpenOfficeCalc_Save_Test_After"], _
		["", "OpenOfficeCalc_Save_Test_ValidDocument", "AutoItTest_Assert", $TestString, "OpenOfficeCalc_Save_Test_After"], _
		["OpenOfficeCalc_Save_Test_Before", "OpenOfficeCalc_Save_Test_InvalidDocument", "AutoItTest_Assert", $Cell00, "OpenOfficeCalc_Save_Test_After"], _
		["", "OpenOfficeCalc_GetSheet_Test_ValidSheet", "AutoItTest_IsObj", True, "OpenOfficeCalc_Close"], _
		["", "OpenOfficeCalc_GetSheet_Test_InvalidSheet", "AutoItTest_IsObj", False, "OpenOfficeCalc_Close"], _
		["", "OpenOfficeCalc_GetSheet_Test_ValidDocument", "AutoItTest_IsObj", True, "OpenOfficeCalc_Close"], _
		["", "OpenOfficeCalc_GetSheet_Test_InvalidDocument", "AutoItTest_IsObj", False, "OpenOfficeCalc_Close"], _
		["", "OpenOfficeCalc_GetCell_Test_NoSheet", "AutoItTest_IsObj", True, "OpenOfficeCalc_Close"], _
		["", "OpenOfficeCalc_GetCell_Test_ValidSheet", "AutoItTest_IsObj", True, "OpenOfficeCalc_Close"], _
		["", "OpenOfficeCalc_GetCell_Test_InvalidSheet", "AutoItTest_IsObj", False, "OpenOfficeCalc_Close"], _
		["", "OpenOfficeCalc_GetCell_Test_RowNegative", "AutoItTest_IsObj", False, "OpenOfficeCalc_Close"], _
		["", "OpenOfficeCalc_GetCell_Test_ColumnNegative", "AutoItTest_IsObj", False, "OpenOfficeCalc_Close"] _
		]

AutoItTest_Runner($OpenOfficeCalcTest)

#region OpenOfficeCalc_Open_Test
;
; OpenOfficeCalc_Open関数に存在するファイル名を指定した場合.
;
; 期待する結果: ドキュメントオブジェクトが戻り値として返る.
;
Func OpenOfficeCalc_Open_Test_ValidFile()
	Return OpenOfficeCalc_Open($ValidFile1)
EndFunc   ;==>OpenOfficeCalc_Open_Test_ValidFile
;
; OpenOfficeCalc_Open関数に存在しないファイルを指定した場合.
;
; 期待する結果: 指定したファイルが作成される.
;
Func OpenOfficeCalc_Open_Test_InvalidFile()
	Return OpenOfficeCalc_Open($InvalidFile)
EndFunc   ;==>OpenOfficeCalc_Open_Test_InvalidFile
;
; OpenOfficeCalc_Open_Test_InvalidFile関数の後処理.
;
Func OpenOfficeCalc_Open_Test_InvalidFile_After()
	OpenOfficeCalc_Close()
	FileDelete($InvalidFile)
EndFunc   ;==>OpenOfficeCalc_Open_Test_InvalidFile_After
#endregion OpenOfficeCalc_Open_Test

#region OpenOfficeCalc_Close_Test
;
; OpenOfficeCalc_Close関数に引数が有効な引数が指定された場合.
;
; 期待する結果: OpenOfficeCalcが閉じる.
;
Func OpenOfficeCalc_Close_Test_Document()
	Local $document = OpenOfficeCalc_Open($ValidFile1)
	OpenOfficeCalc_Close($document)
EndFunc   ;==>OpenOfficeCalc_Close_Test_Document
;
; OpenOfficeCalc_Close関数に引数を指定しない(グローバル変数値を利用した)場合.
;
; 期待する結果: OpenOfficeCalcが閉じる.
;
Func OpenOfficeCalc_Close_Test_NoDocument()
	Local $document = OpenOfficeCalc_Open($ValidFile1)
	OpenOfficeCalc_Close()
EndFunc   ;==>OpenOfficeCalc_Close_Test_NoDocument
;
; OpenOfficeCalc_Close関数に無効なドキュメントオブジェクトを指定した場合.
;
; 期待する結果: OpenOfficeCalcが閉じる.(ただし, 意図したウィンドが閉じるとは限らない)
;
Func OpenOfficeCalc_Close_Test_InvalidDocument()
	Local $document = OpenOfficeCalc_Open($ValidFile1)
	OpenOfficeCalc_Close(0)
EndFunc   ;==>OpenOfficeCalc_Close_Test_InvalidDocument
#endregion OpenOfficeCalc_Close_Test

#region OpenOfficeCalc_Save_Test
;
; OpenOfficeCalc_Save_Testの前処理.
;
Func OpenOfficeCalc_Save_Test_Before()
	FileCopy($ValidFile1, $TempFile, 1)
	OpenOfficeCalc_Open($ValidFile1)
	OpenOfficeCalc_GetSheet($ValidSheet)
	Local $cell = OpenOfficeCalc_GetCell(0, 0)
	$cell.setString($TestString)
EndFunc   ;==>OpenOfficeCalc_Save_Test_Before
;
; OpenOfficeCalc_Save_Testの後処理.
;
Func OpenOfficeCalc_Save_Test_After()
	OpenOfficeCalc_Close()
	FileMove($TempFile, $ValidFile1, 1)
	FileCopy($ValidFile1, $ValidFile2, 1)
	FileDelete($InvalidFile)
EndFunc   ;==>OpenOfficeCalc_Save_Test_After
;
; OpenOfficeCalc_Save関数に引数を指定しない場合.
;
; 期待する結果: OpenOfficeCalc_Open関数でオープンしたファイルを上書き保存する.
;
Func OpenOfficeCalc_Save_Test_NoArgument()
	OpenOfficeCalc_Save()
	OpenOfficeCalc_Close()

	OpenOfficeCalc_Open($ValidFile1)
	OpenOfficeCalc_GetSheet($ValidSheet)
	$cell = OpenOfficeCalc_GetCell(0, 0)
	Return $cell.getString()
EndFunc   ;==>OpenOfficeCalc_Save_Test_NoArgument
;
; OpenOfficeCalc_Save関数に存在しないファイルを指定した場合.
;
; 期待する結果: 指定したファイル名で保存する.
;
Func OpenOfficeCalc_Save_Test_InvalidFile()
	OpenOfficeCalc_Save($InvalidFile)
	OpenOfficeCalc_Close()

	OpenOfficeCalc_Open($InvalidFile)
	OpenOfficeCalc_GetSheet($ValidSheet)
	$cell = OpenOfficeCalc_GetCell(0, 0)
	Return $cell.getString()
EndFunc   ;==>OpenOfficeCalc_Save_Test_InvalidFile
;
; OpenOfficeCalc_Save関数に存在するファイル名(OpenOfficeCalc_Open関数でオープンしたファイル)を指定した場合.
;
; 期待する結果: OpenOfficeCalc_Open関数でオープンしたファイルを上書き保存する.
;
Func OpenOfficeCalc_Save_Test_ValidFile_Own()
	OpenOfficeCalc_Save($ValidFile1)
	OpenOfficeCalc_Close()

	OpenOfficeCalc_Open($ValidFile1)
	OpenOfficeCalc_GetSheet($ValidSheet)
	$cell = OpenOfficeCalc_GetCell(0, 0)
	Return $cell.getString()
EndFunc   ;==>OpenOfficeCalc_Save_Test_ValidFile_Own
;
; OpenOfficeCalc_Save関数に存在するファイル名(OpenOfficeCalc_Open関数でオープンしたファイル以外)を指定した場合.
;
; 期待する結果: 指定したファイル名で保存する.
;
Func OpenOfficeCalc_Save_Test_ValidFile_Other()
	OpenOfficeCalc_Save($ValidFile2)
	OpenOfficeCalc_Close()

	OpenOfficeCalc_Open($ValidFile2)
	OpenOfficeCalc_GetSheet($ValidSheet)
	$cell = OpenOfficeCalc_GetCell(0, 0)
	Return $cell.getString()
EndFunc   ;==>OpenOfficeCalc_Save_Test_ValidFile_Other
;
; OpenOfficeCalc_Save関数に有効なドキュメントオブジェクトを指定した場合.
;
; 期待する結果: 指定したドキュメントオブジェクトのファイルを指定したファイル名で保存する.
;
Func OpenOfficeCalc_Save_Test_ValidDocument()
	FileCopy($ValidFile1, $TempFile, 1)
	Local $document = OpenOfficeCalc_Open($ValidFile1)
	OpenOfficeCalc_GetSheet($ValidSheet)
	Local $cell = OpenOfficeCalc_GetCell(0, 0)
	$cell.setString($TestString)

	OpenOfficeCalc_Save($ValidFile1, 0, $document)
	OpenOfficeCalc_Close()

	OpenOfficeCalc_Open($ValidFile1)
	OpenOfficeCalc_GetSheet($ValidSheet)
	$cell = OpenOfficeCalc_GetCell(0, 0)
	Return $cell.getString()
EndFunc   ;==>OpenOfficeCalc_Save_Test_ValidDocument
;
; OpenOfficeCalc_Save関数に無効なドキュメントオブジェクトを指定した場合.
;
; 期待する結果: 保存されない.
;
Func OpenOfficeCalc_Save_Test_InvalidDocument()
	OpenOfficeCalc_Save($ValidFile1, 0, 0)
	OpenOfficeCalc_Close()

	OpenOfficeCalc_Open($ValidFile1)
	OpenOfficeCalc_GetSheet($ValidSheet)
	$cell = OpenOfficeCalc_GetCell(0, 0)
	Return $cell.getString()
EndFunc   ;==>OpenOfficeCalc_Save_Test_InvalidDocument
#endregion OpenOfficeCalc_Save_Test

#region OpenOfficeCalc_GetSheet_Test
;
; OpenOfficeCalc_GetSheet関数に有効なシート名を指定した場合.
;
; 期待する結果: シートオブジェクトが戻り値として返る.
;
Func OpenOfficeCalc_GetSheet_Test_ValidSheet()
	Local $document = OpenOfficeCalc_Open($ValidFile1)
	Return OpenOfficeCalc_GetSheet($ValidSheet)
EndFunc   ;==>OpenOfficeCalc_GetSheet_Test_ValidSheet
;
; OpenOfficeCalc_GetSheet関数に無効なシート名を指定した場合.
;
; 期待する結果: シートオブジェクトが戻り値として返らない.
;
Func OpenOfficeCalc_GetSheet_Test_InvalidSheet()
	Local $document = OpenOfficeCalc_Open($ValidFile1)
	Return OpenOfficeCalc_GetSheet($InvalidSheet)
EndFunc   ;==>OpenOfficeCalc_GetSheet_Test_InvalidSheet
;
;OpenOfficeCalc_GetSheet関数に有効なドキュメントオブジェクトを指定した場合.
;
; 期待する結果: シートオブジェクトが戻り値として返る.
;
Func OpenOfficeCalc_GetSheet_Test_ValidDocument()
	Local $document = OpenOfficeCalc_Open($ValidFile1)
	Return OpenOfficeCalc_GetSheet($ValidSheet, $document)
EndFunc   ;==>OpenOfficeCalc_GetSheet_Test_ValidDocument
;
; OpenOfficeCalc_GetSheet関数に無効なドキュメントオブジェクトを指定した場合.
;
; 期待する結果: シートオブジェクトが戻り値として返らない.
;
Func OpenOfficeCalc_GetSheet_Test_InvalidDocument()
	Local $document = OpenOfficeCalc_Open($ValidFile1)
	Return OpenOfficeCalc_GetSheet($ValidSheet, 0)
EndFunc   ;==>OpenOfficeCalc_GetSheet_Test_InvalidDocument
#endregion OpenOfficeCalc_GetSheet_Test

#region OpenOfficeCalc_GetCell_Test
;
; OpenOfficeCalc_GetCell関数にシートオブジェクトを指定しない場合.
;
; 期待する結果: セルオブジェクトが戻り値として返る.
;
Func OpenOfficeCalc_GetCell_Test_NoSheet()
	Local $document = OpenOfficeCalc_Open($ValidFile1)
	Local $sheet = OpenOfficeCalc_GetSheet($ValidSheet)
	Return OpenOfficeCalc_GetCell(0, 0)
EndFunc   ;==>OpenOfficeCalc_GetCell_Test_NoSheet
;
; OpenOfficeCalc_GetCell関数に有効なシートオブジェクトを指定した場合.
;
; 期待する結果: セルオブジェクトが戻り値として返る.
;
Func OpenOfficeCalc_GetCell_Test_ValidSheet()
	Local $document = OpenOfficeCalc_Open($ValidFile1)
	Local $sheet = OpenOfficeCalc_GetSheet($ValidSheet)
	Return OpenOfficeCalc_GetCell(0, 0, $sheet)
EndFunc   ;==>OpenOfficeCalc_GetCell_Test_ValidSheet
;
; OpenOfficeCalc_GetCell関数に有効なシートオブジェクトを指定した場合.
;
; 期待する結果: セルオブジェクトが戻り値として返らない.
;
Func OpenOfficeCalc_GetCell_Test_InvalidSheet()
	Local $document = OpenOfficeCalc_Open($ValidFile1)
	Local $sheet = OpenOfficeCalc_GetSheet($ValidSheet)
	Return OpenOfficeCalc_GetCell(0, 0, 0)
EndFunc   ;==>OpenOfficeCalc_GetCell_Test_InvalidSheet
;
; OpenOfficeCalc_GetCell関数に負数の行を指定した場合.
;
; 期待する結果: セルオブジェクトが戻り値として返らない.
;
Func OpenOfficeCalc_GetCell_Test_RowNegative()
	Local $document = OpenOfficeCalc_Open($ValidFile1)
	Local $sheet = OpenOfficeCalc_GetSheet($ValidSheet)
	Return OpenOfficeCalc_GetCell(-1, 0)
EndFunc   ;==>OpenOfficeCalc_GetCell_Test_RowNegative
;
; OpenOfficeCalc_GetCell関数に負数の列を指定した場合.
;
; 期待する結果: セルオブジェクトが戻り値として返らない.
;
Func OpenOfficeCalc_GetCell_Test_ColumnNegative()
	Local $document = OpenOfficeCalc_Open($ValidFile1)
	Local $sheet = OpenOfficeCalc_GetSheet($ValidSheet)
	Return OpenOfficeCalc_GetCell(0, -1)
EndFunc   ;==>OpenOfficeCalc_GetCell_Test_ColumnNegative
#endregion OpenOfficeCalc_GetCell_Test