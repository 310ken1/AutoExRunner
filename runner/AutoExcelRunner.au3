#include-once
#include <Array.au3>
#include "..\utility\ArrayUtility.au3"
#include "..\utility\FileUtility.au3"

#region Globale_Argument_Define
; AutoExcelRunner の開始行.
Global $AutoExcelRunner_StartLine = 3
; No項目 の列.
Global $AutoExcelRunner_NoStartColumn = 1
; 項目名 の開始行.
Global $AutoExcelRunner_KeyStartLine = 2
; 項目名 の開始列.
Global $AutoExcelRunner_KeyStartColumn = 1
; 項目の最大個数(No項目を除く).
Global $AutoExcelRunner_KeyColumnMaxCount = 10
; Excel から一度に読み出す行数.
Global $AutoExcelRunner_ExecutableUnit = 25
; ハンドル(配列)のインデックス.
Global Enum _
		$AutoExcelRunnerDocument = 0, _
		$AutoExcelRunnerFileName = 1, _
		$AutoExcelRunnerSheet = 2, _
		$AutoExcelRunnerLine = 3, _
		$AutoExcelRunnerKeyCash = 4, _
		$AutoExcelRunnerValueArray = 5, _
		$AutoExcelRunnerUserData = 6, _
		$AutoExcelRunnerMax = 7
#endregion Globale_Argument_Define

#region Public_Method
;
; 指定した Excel ファイルを読込み, 行毎に指定した関数を呼出す.
; No 列 に数値がある場合のみ関数が読み出され, "End" もしくは "E"が記載されたセルまで, 実行し続ける.
;
; @param $file 読込む Excel ファイル名.
; @param $sheet_name 読込む Excel ファイルのシート名.
; @param $callback コールバックする関数名.
; @param $save Excelファイルの保存有無.
;
Func AutoExcelRunner(Const ByRef $file, Const ByRef $sheet_name, Const ByRef $callback, Const $save = False)
	Local $handle = AutoExcelRunner_Open($file)
	If IsArray($handle) Then
		AutoExcelRunner_Run($handle, $sheet_name, $callback)
		AutoExcelRunner_Close($handle, $save)
	EndIf
EndFunc   ;==>AutoExcelRunner

;
; AutoExcelRunnerを実行するためのハンドルをオープンする.
;
; @param $file 読込む Excel ファイル名.
; @return ハンドル.
;
Func AutoExcelRunner_Open(Const ByRef $file)
	Local $handle[$AutoExcelRunnerMax]
	$handle[$AutoExcelRunnerDocument] = ObjGet($file, "Excel.Application")
	If (Not @error) And IsObj($handle[$AutoExcelRunnerDocument]) Then
		$handle[$AutoExcelRunnerFileName] = FileUtility_FileBaseName($file)
		With $handle[$AutoExcelRunnerDocument]
			.Application.Visible = True
			.Windows($handle[$AutoExcelRunnerFileName]).Visible = True
		EndWith
		Return $handle
	EndIf
	Return 0
EndFunc   ;==>AutoExcelRunner_Open

;
; AutoExcelRunnerを実行する.
;
; @param $handle
; @param $sheet_name 読込むOpenOffice Calcファイルのシート名.
; @param $callback コールバック関数名.
; @return 成否.
;
Func AutoExcelRunner_Run(ByRef $handle, Const ByRef $sheet_name, Const ByRef $callback)
	Local $result = False
	If IsArray($handle) Then
		Local $document = $handle[$AutoExcelRunnerDocument]
		If AutoExcelRunner_SheetExist($document, $sheet_name) Then
			$handle[$AutoExcelRunnerSheet] = $document.Worksheets($sheet_name)
			If (Not @error) And IsObj($handle[$AutoExcelRunnerSheet]) Then
				Local $sheet = $handle[$AutoExcelRunnerSheet]
				$handle[$AutoExcelRunnerLine] = $AutoExcelRunner_StartLine
				$handle[$AutoExcelRunnerKeyCash] = AutoExcelRunner_CreateKeyArray($handle[$AutoExcelRunnerSheet])
				Local $run = True
				While $run
					Local $range = $sheet.range( _
							$sheet.cells($handle[$AutoExcelRunnerLine], $AutoExcelRunner_KeyStartColumn), _
							$sheet.cells($handle[$AutoExcelRunnerLine] + $AutoExcelRunner_ExecutableUnit, _
							$AutoExcelRunner_KeyStartColumn + $AutoExcelRunner_KeyColumnMaxCount) _
							)
					Local $two_dimensional_array = $range.value()
					For $i = 0 To $AutoExcelRunner_KeyColumnMaxCount
						Local $value = $two_dimensional_array[$AutoExcelRunner_KeyStartColumn - 1][$i]
						If AutoExcelRunner_IsNoEnd($value) Then
							$run = False
							ExitLoop
						ElseIf StringIsDigit($value) Then
							$handle[$AutoExcelRunnerValueArray] = ArrayUtility_ExtractionColumn($two_dimensional_array, $i)
							$ret = Call($callback, $handle)
						EndIf
						$handle[$AutoExcelRunnerLine] += 1
					Next
				WEnd
				$result = True
			EndIf
		EndIf
	EndIf
	Return $result
EndFunc   ;==>AutoExcelRunner_Run

;
; AutoExcelRunnerを実行するためのハンドルをクローズする.
;
; @param $handle ハンドル.
; @param $save OpenOffice Calcファイルの保存有無.
;
Func AutoExcelRunner_Close(ByRef $handle, $save = False)
	If IsArray($handle) Then
		Local $document = $handle[$AutoExcelRunnerDocument]
		If True = $save Then
			$document.Windows($handle[$AutoExcelRunnerFileName]).Visible = True
			$document.Save
		EndIf
		$document.saved = 1
		$document.Close
	EndIf
EndFunc   ;==>AutoExcelRunner_Close

;
; セルオブジェクトを取得する.
; 取得に失敗した場合は, 0(数値) が戻り値として返る.
;
; @param $handle ハンドル.
; @param $key 取得したいセルの項目名.
; @return セルオブジェクト.
;
Func AutoExcelRunner_GetCell(Const ByRef $handle, Const ByRef $key)
	Local $cell = 0
	If IsArray($handle) Then
		Local $column = _ArraySearch($handle[$AutoExcelRunnerKeyCash], $key, 0, 0, 0, 0, 1, 0)
		Local $sheet = $handle[$AutoExcelRunnerSheet]
		If - 1 < $column Then
			$cell = $sheet.cells($handle[$AutoExcelRunnerLine], $column + 1)
		EndIf
	EndIf
	Return $cell
EndFunc   ;==>AutoExcelRunner_GetCell

;
; セルの値(文字列)を取得する.
; 取得に失敗した場合は, 空文字列("")が戻り値として返る.
;
; @param $handle ハンドル.
; @param $key 取得したいセルの項目名.
; @return 値(文字列).
;
Func AutoExcelRunner_GetString(Const ByRef $handle, Const ByRef $key)
	Local $string = ""
	If IsArray($handle) Then
		Local $index = _ArraySearch($handle[$AutoExcelRunnerKeyCash], $key, 0, 0, 0, 0, 1, 0)
		If - 1 < $index Then
			Local $value = $handle[$AutoExcelRunnerValueArray]
			$string = $value[$index]
		EndIf
	EndIf
	Return $string
EndFunc   ;==>AutoExcelRunner_GetString
#endregion Public_Method

#region Private_Method
;
; 終端かチェックする.
;
; @param $value Noの値.
; @return 終端の有無.
;
Func AutoExcelRunner_IsNoEnd(Const ByRef $value)
	Local $ret = False
	If "End" = $value Or "E" = $value Then
		$ret = True
	EndIf
	Return $ret
EndFunc   ;==>AutoExcelRunner_IsNoEnd

;
; Excelファイルに指定した名前のシートがあるか確認する.
;
; @param $document ドキュメントオブジェクト.
; @param $sheet_name シート名.
; @return シートの有無(True/False).
;
Func AutoExcelRunner_SheetExist(Const ByRef $document, Const ByRef $sheet_name)
	Local $result = False
	Local $count = $document.Worksheets.Count
	For $i = 1 To $count
		If $sheet_name = $document.Worksheets($i).Name Then
			$result = True
			ExitLoop
		EndIf
	Next
	Return $result
EndFunc   ;==>AutoExcelRunner_SheetExist

;
; [項目名(キー) , 列]の２次元配列を取得する.
;
; @return 項目名(キー)と列の要素を持つ２次元配列.
;
Func AutoExcelRunner_CreateKeyArray(Const ByRef $sheet_object)
	Local $array[$AutoExcelRunner_KeyColumnMaxCount + 1][2]
	Local $range = $sheet_object.range( _
			$sheet_object.cells($AutoExcelRunner_KeyStartLine, $AutoExcelRunner_KeyStartColumn), _
			$sheet_object.cells($AutoExcelRunner_KeyStartLine, _
			$AutoExcelRunner_KeyStartColumn + $AutoExcelRunner_KeyColumnMaxCount))
	Local $key = $range.value
	For $i = 0 To $AutoExcelRunner_KeyColumnMaxCount
		$array[$i][0] = $key[$i][0]
		$array[$i][1] = $i + $AutoExcelRunner_KeyStartColumn
	Next
	Return $array
EndFunc   ;==>AutoExcelRunner_CreateKeyArray
#endregion Private_Method