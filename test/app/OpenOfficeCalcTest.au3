#include "..\AutoItTest.au3"
#include "..\..\app\OpenOfficeCalc.au3"
#include "..\..\utility\ConsoleUtility.au3"

#region Constant_Define
; 有効なファイル名.
Const $ValidFile = "OpenOfficeCalcTest.ods"
; 無効なファイル名.
Const $InvalidFile = "OpenOfficeCalcTest1.ods"
; 有効なシート名.
Const $ValidSheet = "シート１"
; 無効なシート名.
Const $InvalidSheet = "Sheet1"
#endregion Constant_Define

Local $OpenOfficeCalcTest[14][5] = [ _
		["", "OpenOfficeCalc_Open_Test_ValidFile", "AutoItTest_IsObj", True, "OpenOfficeCalc_Close"], _
		["", "OpenOfficeCalc_Open_Test_InvalidFile", "AutoItTest_FileExists", $InvalidFile, "OpenOfficeCalc_Open_Test_InvalidFile_After"], _
		["", "OpenOfficeCalc_Close_Test_Document", "AutoItTest_WinNotExists", $OpenOfficeCalc_ClassName, ""], _
		["", "OpenOfficeCalc_Close_Test_NoDocument", "AutoItTest_WinNotExists", $OpenOfficeCalc_ClassName, ""], _
		["", "OpenOfficeCalc_Close_Test_InvalidDocument", "AutoItTest_WinNotExists", $OpenOfficeCalc_ClassName, ""], _
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
	Local $file = FileUtility_ScriptDirFilePath($ValidFile)
	Return OpenOfficeCalc_Open($file)
EndFunc   ;==>OpenOfficeCalc_Open_Test_ValidFile

;
; OpenOfficeCalc_Open関数に存在しないファイルを指定した場合.
;
; 期待する結果: 指定したファイルが作成される.
;
Func OpenOfficeCalc_Open_Test_InvalidFile()
	Local $file = FileUtility_ScriptDirFilePath($InvalidFile)
	Return OpenOfficeCalc_Open($file)
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
	Local $file = FileUtility_ScriptDirFilePath($ValidFile)
	Local $document = OpenOfficeCalc_Open($file)
	OpenOfficeCalc_Close($document)
EndFunc   ;==>OpenOfficeCalc_Close_Test_Document

;
; OpenOfficeCalc_Close関数に引数を指定しない(グローバル変数値を利用した)場合.
;
; 期待する結果: OpenOfficeCalcが閉じる.
;
Func OpenOfficeCalc_Close_Test_NoDocument()
	Local $file = FileUtility_ScriptDirFilePath($ValidFile)
	Local $document = OpenOfficeCalc_Open($file)
	OpenOfficeCalc_Close()
EndFunc   ;==>OpenOfficeCalc_Close_Test_NoDocument

;
; OpenOfficeCalc_Close関数に無効なドキュメントオブジェクトを指定した場合.
;
; 期待する結果: OpenOfficeCalcが閉じる.(ただし, 意図したウィンドが閉じるとは限らない)
;
Func OpenOfficeCalc_Close_Test_InvalidDocument()
	Local $file = FileUtility_ScriptDirFilePath($ValidFile)
	Local $document = OpenOfficeCalc_Open($file)
	OpenOfficeCalc_Close(0)
EndFunc   ;==>OpenOfficeCalc_Close_Test_InvalidDocument
#endregion OpenOfficeCalc_Close_Test

#region OpenOfficeCalc_GetSheet_Test
;
; OpenOfficeCalc_GetSheet関数に有効なシート名を指定した場合.
;
; 期待する結果: シートオブジェクトが戻り値として返る.
;
Func OpenOfficeCalc_GetSheet_Test_ValidSheet()
	Local $file = FileUtility_ScriptDirFilePath($ValidFile)
	Local $document = OpenOfficeCalc_Open($file)
	Return OpenOfficeCalc_GetSheet($ValidSheet)
EndFunc   ;==>OpenOfficeCalc_GetSheet_Test_ValidSheet

;
; OpenOfficeCalc_GetSheet関数に無効なシート名を指定した場合.
;
; 期待する結果: シートオブジェクトが戻り値として返らない.
;
Func OpenOfficeCalc_GetSheet_Test_InvalidSheet()
	Local $file = FileUtility_ScriptDirFilePath($ValidFile)
	Local $document = OpenOfficeCalc_Open($file)
	Return OpenOfficeCalc_GetSheet($InvalidSheet)
EndFunc   ;==>OpenOfficeCalc_GetSheet_Test_InvalidSheet

;
;OpenOfficeCalc_GetSheet関数に有効なドキュメントオブジェクトを指定した場合.
;
; 期待する結果: シートオブジェクトが戻り値として返る.
;
Func OpenOfficeCalc_GetSheet_Test_ValidDocument()
	Local $file = FileUtility_ScriptDirFilePath($ValidFile)
	Local $document = OpenOfficeCalc_Open($file)
	Return OpenOfficeCalc_GetSheet($ValidSheet, $document)
EndFunc   ;==>OpenOfficeCalc_GetSheet_Test_ValidDocument

;
; OpenOfficeCalc_GetSheet関数に無効なドキュメントオブジェクトを指定した場合.
;
; 期待する結果: シートオブジェクトが戻り値として返らない.
;
Func OpenOfficeCalc_GetSheet_Test_InvalidDocument()
	Local $file = FileUtility_ScriptDirFilePath($ValidFile)
	Local $document = OpenOfficeCalc_Open($file)
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
	Local $file = FileUtility_ScriptDirFilePath($ValidFile)
	Local $document = OpenOfficeCalc_Open($file)
	Local $sheet = OpenOfficeCalc_GetSheet($ValidSheet)
	Return OpenOfficeCalc_GetCell(0, 0)
EndFunc   ;==>OpenOfficeCalc_GetCell_Test_NoSheet

;
; OpenOfficeCalc_GetCell関数に有効なシートオブジェクトを指定した場合.
;
; 期待する結果: セルオブジェクトが戻り値として返る.
;
Func OpenOfficeCalc_GetCell_Test_ValidSheet()
	Local $file = FileUtility_ScriptDirFilePath($ValidFile)
	Local $document = OpenOfficeCalc_Open($file)
	Local $sheet = OpenOfficeCalc_GetSheet($ValidSheet)
	Return OpenOfficeCalc_GetCell(0, 0, $sheet)
EndFunc   ;==>OpenOfficeCalc_GetCell_Test_ValidSheet

;
; OpenOfficeCalc_GetCell関数に有効なシートオブジェクトを指定した場合.
;
; 期待する結果: セルオブジェクトが戻り値として返らない.
;
Func OpenOfficeCalc_GetCell_Test_InvalidSheet()
	Local $file = FileUtility_ScriptDirFilePath($ValidFile)
	Local $document = OpenOfficeCalc_Open($file)
	Local $sheet = OpenOfficeCalc_GetSheet($ValidSheet)
	Return OpenOfficeCalc_GetCell(0, 0, 0)
EndFunc   ;==>OpenOfficeCalc_GetCell_Test_InvalidSheet

;
; OpenOfficeCalc_GetCell関数に負数の行を指定した場合.
;
; 期待する結果: セルオブジェクトが戻り値として返らない.
;
Func OpenOfficeCalc_GetCell_Test_RowNegative()
	Local $file = FileUtility_ScriptDirFilePath($ValidFile)
	Local $document = OpenOfficeCalc_Open($file)
	Local $sheet = OpenOfficeCalc_GetSheet($ValidSheet)
	Return OpenOfficeCalc_GetCell(-1, 0)
EndFunc   ;==>OpenOfficeCalc_GetCell_Test_RowNegative

;
; OpenOfficeCalc_GetCell関数に負数の列を指定した場合.
;
; 期待する結果: セルオブジェクトが戻り値として返らない.
;
Func OpenOfficeCalc_GetCell_Test_ColumnNegative()
	Local $file = FileUtility_ScriptDirFilePath($ValidFile)
	Local $document = OpenOfficeCalc_Open($file)
	Local $sheet = OpenOfficeCalc_GetSheet($ValidSheet)
	Return OpenOfficeCalc_GetCell(0, -1)
EndFunc   ;==>OpenOfficeCalc_GetCell_Test_ColumnNegative
#endregion OpenOfficeCalc_GetCell_Test