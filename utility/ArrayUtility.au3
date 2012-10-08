#include-once

#region Public_Method
;
; ２次元配列から指定行の配列を取得する.
;
; @param $two_dimensional_array ２次元配列.
; @param $row 行.
; @return 行配列.
;
Func ArrayUtility_ExtractionRow($two_dimensional_array, $row)
	Local $column_count = UBound($two_dimensional_array, 2)
	Local $array[$column_count]
	For $i = 0 To $column_count - 1
		$array[$i] = $two_dimensional_array[$row][$i]
	Next
	Return $array
EndFunc   ;==>ArrayUtility_ExtractionRow
;
; ２次元配列から指定列の配列を取得する.
;
; @param $two_dimensional_array ２次元配列.
; @param $column 列.
; @return 列配列.
;
Func ArrayUtility_ExtractionColumn($two_dimensional_array, $column)
	Local $row_count = UBound($two_dimensional_array, 1)
	Local $array[$row_count]
	For $i = 0 To $row_count - 1
		$array[$i] = $two_dimensional_array[$i][$column]
	Next
	Return $array
EndFunc   ;==>ArrayUtility_ExtractionColumn
#endregion Public_Method