#include-once
#include "..\app\OpenOfficeCalc.au3"
#include <Array.au3>

#region Globale_Argument_Define
; AutoOpenOfficeRunner の開始行.
Global $AutoOpenOfficeRunner_StartLine = 2
; No項目 の列.
Global $AutoOpenOfficeRunner_NoStartColumn = 0
; 項目名 の開始行.
Global $AutoOpenOfficeRunner_KeyStartLine = 1
; 項目名 の開始列.
Global $AutoOpenOfficeRunner_KeyStartColumn = 0
; 項目の最大個数(No項目を除く).
Global $AutoOpenOfficeRunner_KeyColumnMaxCount = 10
; OpenOfficeCalcから一度に読み出す行数.
Global $AutoOpenOfficeRunner_ExecutableUnit = 25
; ハンドル(配列)のインデックス.
Global Enum _
		$AutoOpenOfficeRunnerDocument = 0, _
		$AutoOpenOfficeRunnerSheet = 1, _
		$AutoOpenOfficeRunnerLine = 2, _
		$AutoOpenOfficeRunnerKeyCash = 3, _
		$AutoOpenOfficeRunnerValueArray = 4, _
		$AutoOpenOfficeRunnerUserData = 5, _
		$AutoOpenOfficeRunnerMax = 6
#endregion Globale_Argument_Define

#region Public_Method
;
; 指定したOpenOffice Calcファイルを読込み, 行毎に指定した関数を呼出す.
; No 列 に数値がある場合のみ関数が読み出され, "End" もしくは "E"が記載されたセルまで, 実行し続ける.
;
; @param $file 読込むOpenOffice Calcファイル名.
; @param $sheet_name 読込むOpenOffice Calcファイルのシート名.
; @param $callback コールバックする関数名.
; @param $save OpenOffice Calcファイルの保存有無.
; @return ドキュメントオブジェクト.
;
Func AutoOpenOfficeRunner(Const ByRef $file, Const ByRef $sheet_name, Const ByRef $callback, Const $save = False)
	Local $handle = AutoOpenOfficeRunner_Open($file)
	If IsArray($handle) Then
		AutoOpenOfficeRunner_Run($handle, $sheet_name, $callback)
		AutoOpenOfficeRunner_Close($handle, $save)
	EndIf
EndFunc   ;==>AutoOpenOfficeRunner

;
; AutoOpenOfficeRunnerを実行するためのハンドルをオープンする.
;
; @param $file 読込むOpenOffice Calcファイル名.
; @return ハンドル.
;
Func AutoOpenOfficeRunner_Open(Const ByRef $file)
	If FileExists($file) Then
		Local $handle[$AutoOpenOfficeRunnerMax]
		$handle[$AutoOpenOfficeRunnerDocument] = OpenOfficeCalc_Open($file)
		If IsObj($handle[$AutoOpenOfficeRunnerDocument]) Then
			Return $handle
		EndIf
	EndIf
	Return 0
EndFunc   ;==>AutoOpenOfficeRunner_Open

;
; AutoOpenOfficeRunnerを実行する.
;
; @param $handle
; @param $sheet_name 読込むOpenOffice Calcファイルのシート名.
; @param $callback コールバック関数名.
; @return 成否.
;
Func AutoOpenOfficeRunner_Run(ByRef $handle, Const ByRef $sheet_name, Const ByRef $callback)
	Local $result = False
	If IsArray($handle) Then
		$handle[$AutoOpenOfficeRunnerSheet] = OpenOfficeCalc_GetSheet($sheet_name)
		If IsObj($handle[$AutoOpenOfficeRunnerSheet]) Then
			$handle[$AutoOpenOfficeRunnerLine] = $AutoOpenOfficeRunner_StartLine
			$handle[$AutoOpenOfficeRunnerKeyCash] = AutoOpenOfficeRunner_CreateKeyArray($handle[$AutoOpenOfficeRunnerSheet])
			Local $run = True
			While $run
				Local $range = OpenOfficeCalc_GetRange( _
						$handle[$AutoOpenOfficeRunnerLine], _
						$AutoOpenOfficeRunner_KeyStartColumn, _
						$handle[$AutoOpenOfficeRunnerLine] + $AutoOpenOfficeRunner_ExecutableUnit, _
						$AutoOpenOfficeRunner_KeyStartColumn + $AutoOpenOfficeRunner_KeyColumnMaxCount, _
						$handle[$AutoOpenOfficeRunnerSheet])
				Local $line_array = $range.getDataArray()
				For $value_array In $line_array
					Local $value = $value_array[$AutoOpenOfficeRunner_NoStartColumn]
					If AutoOpenOfficeRunner_IsNoEnd($value) Then
						$run = False
						ExitLoop
					ElseIf StringIsDigit($value) Then
						$handle[$AutoOpenOfficeRunnerValueArray] = $value_array
						$ret = Call($callback, $handle)
					EndIf
					$handle[$AutoOpenOfficeRunnerLine] += 1
				Next
			WEnd
			$result = True
		EndIf
	EndIf
	Return $result
EndFunc   ;==>AutoOpenOfficeRunner_Run

;
; AutoOpenOfficeRunnerを実行するためのハンドルをクローズする.
;
; @param $handle ハンドル.
; @param $save OpenOffice Calcファイルの保存有無.
;
Func AutoOpenOfficeRunner_Close(ByRef $handle, $save = False)
	If IsArray($handle) Then
		If True = $save Then
			OpenOfficeCalc_Save(0, 0, $handle[$AutoOpenOfficeRunnerDocument])
		EndIf
		OpenOfficeCalc_Close($handle[$AutoOpenOfficeRunnerDocument])
	EndIf
EndFunc   ;==>AutoOpenOfficeRunner_Close

;
; セルオブジェクトを取得する.
; 取得に失敗した場合は, 0(数値) が戻り値として返る.
;
; @param $handle ハンドル.
; @param $key 取得したいセルの項目名.
; @return セルオブジェクト.
;
Func AutoOpenOfficeRunner_GetCell(Const ByRef $handle, Const ByRef $key)
	Local $cell = 0
	If IsArray($handle) Then
		Local $column = _ArraySearch($handle[$AutoOpenOfficeRunnerKeyCash], $key, 0, 0,0,0,1,0)
		If -1 < $column Then
			$cell = OpenOfficeCalc_GetCell( _
					$handle[$AutoOpenOfficeRunnerLine], _
					$column + $AutoOpenOfficeRunner_KeyStartColumn, _
					$handle[$AutoOpenOfficeRunnerSheet])
		EndIf
	EndIf
	Return $cell
EndFunc   ;==>AutoOpenOfficeRunner_GetCell

;
; セルの値(文字列)を取得する.
; 取得に失敗した場合は, 空文字列("")が戻り値として返る.
;
; @param $handle ハンドル.
; @param $key 取得したいセルの項目名.
;
Func AutoOpenOfficeRunner_GetString(Const ByRef $handle, Const ByRef $key)
	Local $string = ""
	If IsArray($handle) Then
		Local $index = _ArraySearch($handle[$AutoOpenOfficeRunnerKeyCash], $key, 0, 0,0,0,1,0)
		If - 1 < $index Then
			Local $value = $handle[$AutoOpenOfficeRunnerValueArray]
			$string = $value[$index]
		EndIf
	EndIf
	Return $string
EndFunc   ;==>AutoOpenOfficeRunner_GetString
#endregion Public_Method

#region Private_Method
;
; 終端かチェックする.
;
; @param $value Noの値.
; @return 終端の有無.
;
Func AutoOpenOfficeRunner_IsNoEnd(Const ByRef $value)
	Local $ret = False
	If "End" = $value Or "E" = $value Then
		$ret = True
	EndIf
	Return $ret
EndFunc   ;==>AutoOpenOfficeRunner_IsNoEnd

;
; [項目名(キー) , 列]の２次元配列を取得する.
;
; @return 項目名(キー)と列の要素を持つ２次元配列.
;
Func AutoOpenOfficeRunner_CreateKeyArray(Const ByRef $sheet_object)
	Local $array[$AutoOpenOfficeRunner_KeyColumnMaxCount+1][2]
	Local $range = OpenOfficeCalc_GetRange( _
			$AutoOpenOfficeRunner_KeyStartLine, _
			$AutoOpenOfficeRunner_KeyStartColumn, _
			$AutoOpenOfficeRunner_KeyStartLine, _
			$AutoOpenOfficeRunner_KeyStartColumn + $AutoOpenOfficeRunner_KeyColumnMaxCount, _
			$sheet_object)
	Local $line_array = $range.getDataArray()
	Local $value_array = $line_array[0]
	For $i = 0 To $AutoOpenOfficeRunner_KeyColumnMaxCount
		$array[$i][0] = $value_array[$i]
		$array[$i][1] = $i + $AutoOpenOfficeRunner_KeyStartColumn
	Next
	Return $array
EndFunc   ;==>AutoOpenOfficeRunner_CreateKeyArray
#endregion Private_Method