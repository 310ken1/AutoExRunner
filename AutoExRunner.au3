#include "AutoItUtility.au3"

Global $AutoExRunnerConfig = @ScriptDir & "\AutoExRunner.ini"

;
; 指定したExcelファイルを読込み, 指定した関数を呼出す.
;
; @param $file 読込むExcelファイル名.
; @param $sheet_name 読込むExcelファイルのシート名.
; @param $callback_name コールバックする関数名.
;
Func AutoExRunner($file, $sheert_name, $callback_name)
	Local $excel = ObjGet($file)
	If (Not @error) And IsObj($excel) Then
		Local $sheet = $excel.Worksheets($sheert_name)
		If (Not @error) And IsObj($sheet) Then
			Local $line = Int(IniRead($AutoExRunnerConfig, "Setting", "StartLine", 3))
			Local $column = Int(IniRead($AutoExRunnerConfig, "Setting", "StartColumn", 1))
			While True
				Local $cell = $sheet.cells($line, $column)
				Local $value = $cell.value
				If IsNoEnd($value) Then
					ExitLoop
				ElseIf IsNumber($value) Then
					Call($callback_name, $sheet, $line)
				EndIf
				$line += 1
			WEnd
			$excel = 0
		EndIf
	Else
		MsgBox(0, "Error", "Could not open " & $file & " as an Excel Object.")
	EndIf
EndFunc   ;==>AutoExRunner

;
; セルを取得する.
;
; @param $sheet 実行中のシートオブジェクト.
; @param $line 実行中の行.
; @param $key 取得したいセルの項目名.
;
Func GetCell($sheet, $line, $key)
	Local $key_line = Int(IniRead($AutoExRunnerConfig, "Setting", "KeyLine", 2))
	Local $key_column = Int(IniRead($AutoExRunnerConfig, "Setting", "KeyColumn", 2))
	Local $key_column_max = Int(IniRead($AutoExRunnerConfig, "Setting", "KeyColumnMax", 10))
	Local $cell = 0
	While $key_column <= $key_column_max
		If $key = $sheet.cells($key_line, $key_column).value Then
			$cell = $sheet.cells($line, $key_column)
			ExitLoop
		EndIf
		$key_column += 1
	WEnd
	Return $cell
EndFunc   ;==>GetCell

;
; No を取得する.
;
; @param $sheet 実行中のシートオブジェクト.
; @param $line 実行中の行.
;
Func GetNo($sheet, $line)
	Local $column = Int(IniRead($AutoExRunnerConfig, "Setting", "StartColumn", 1))
	Return $sheet.cells($line, $column).value
EndFunc   ;==>GetNo
;
;  終端かチェックする.
;
;  @param $value Noの値.
;
Func IsNoEnd($value)
	Local $ret = False
	If "End" = $value Or "E" = $value Then
		$ret = True
	EndIf
	Return $ret
EndFunc   ;==>IsNoEnd


