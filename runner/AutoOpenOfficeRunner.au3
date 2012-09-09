#include-once
#include "..\app\OpenOfficeCalc.au3"

#region Globale_Argument_Define
; No の開始行.
Global $NoStartLine = 2
; No の開始列.
Global $NoStartColumn = 0
; 項目名 の開始行.
Global $KeyStartLine = 1
; 項目名 の開始列.
Global $KeyStartColumn = 0
; 項目の最大個数.
Global $KeyColumnMaxCount = 10
#endregion Globale_Argument_Define

#region Static_Argument_Define
; 実行中の行.
Static Local $CurrentLine = 0
#endregion

#region Public_Method
;
; 指定したOpenOffice Calcファイルを読込み, 行毎に指定した関数を呼出す.
; No 列 に数値がある場合のみ関数が読み出され, "End" もしくは "E"が記載されたセルまで, 実行し続ける.
;
; @param $file 読込むOpenOffice Calcファイル名.
; @param $sheet_name 読込むOpenOffice Calcファイルのシート名.
; @param $callback_name コールバックする関数名.
; @return ドキュメントオブジェクト.
;
Func AutoOpenOfficeRunner($file, $sheet_name, $callback_name)
	Local $document = OpenOfficeCalc_Open($file)
	If IsObj($document) Then
		Local $sheet = OpenOfficeCalc_GetSheet($sheet_name)
		If IsObj($sheet) Then
			$CurrentLine = $NoStartLine
			While True
				Local $cell = OpenOfficeCalc_GetCell($CurrentLine, $NoStartColumn)
				Local $value = $cell.String
				If IsNoEnd($value) Then
					ExitLoop
				ElseIf StringIsDigit($value) Then
					Call($callback_name, $sheet, $CurrentLine)
				EndIf
				$CurrentLine += 1
			WEnd
			$sheet = 0
		EndIf
	EndIf
	Return $document
EndFunc   ;==>AutoOpenOfficeRunner

;
; セルを取得する.
;
; @param $key 取得したいセルの項目名.
; @param $line 実行中の行.
; @return セル.
;
Func GetCell($key, $line=$CurrentLine)
	Local $end = $KeyStartColumn + $KeyColumnMaxCount
	Local $cell = 0
	For $i = $KeyStartColumn To $end
		Local $c = OpenOfficeCalc_GetCell($KeyStartLine, $i)
		Local $value = $c.String
		If $key = $value Then
			$cell = OpenOfficeCalc_GetCell($line, $i)
			ExitLoop
		EndIf
	Next
	Return $cell
EndFunc   ;==>GetCell

;
; セルの値(文字列)を取得する.
;
; @param $key 取得したいセルの項目名.
; @param $line 実行中の行.
; @return セルの値(文字列).
;
Func GetString($key, $line=$CurrentLine)
	Local $value = ""
	Local $cell = GetCell($key, $line)
	If IsObj($cell) Then
		$value = $cell.String
	EndIf
	Return $value
EndFunc   ;==>GetString
#endregion Public_Method

#region Private_Method
;
; 終端かチェックする.
;
; @param $value Noの値.
; @return 終端の有無.
;
Func IsNoEnd($value)
	Local $ret = False
	If "End" = $value Or "E" = $value Then
		$ret = True
	EndIf
	Return $ret
EndFunc   ;==>IsNoEnd
#endregion Private_Method