#include-once
#include "..\utility\FileUtility.au3"

#region Constant_Define
; ウィンドウのクラス名.
Const $OpenOfficeCalc_ClassName = "[CLASS:SALFRAME]"
; サービスマネージャ.名
Const $OpenOfficeCalc_ServiceManger = "com.sun.star.ServiceManager"
; デスクトップ名.
Const $OpenOfficeCalc_Desktop = "com.sun.star.frame.Desktop"
; プロパティ名.
Const $OpenOfficeCalc_PropertyValue = "com.sun.star.beans.PropertyValue"
#endregion Constant_Define

#region Static_Argument_Define
; ドキュメントオブジェクト.
Static Local $OpenOfficeCalc_Document = 0
; シートオブジェクト.
Static Local $OpenOfficeCalc_Sheet = 0
#endregion Static_Argument_Define

#region Public_Method
;
; OpenOffice  Calc が サポートするファイルを開く.
; 存在しないファイルを指定した場合は, ファイルが作成される.
;
; @param $file 開くファイル名(フルパスで指定する必要がある).
; @param $property ファイルを開く時の属性の配列.
; @return ドキュメントオブジェクト.
;
Func OpenOfficeCalc_Open(Const ByRef $file, $property = 0)
	Local $document = 0

	Local $url = ""
	If FileExists($file) Then
		$url = FileUtiilty_PathToUrl($file)
	Else
		$url = "private:factory/scalc"
	EndIf
	Local $array[1]
	If Not IsArray($property) Then
		$property = $array
	EndIf

	Local $manager = ObjCreate($OpenOfficeCalc_ServiceManger)
	If Not @error Then
		Local $desktop = $manager.createInstance($OpenOfficeCalc_Desktop)
		If Not @error Then
			$document = $desktop.loadComponentFromURL($url, "_default", 0, $property)
			WinWaitActive("[CLASS:SALFRAME]")
			If Not FileExists($file) Then
				$document.storeAsURL(FileUtiilty_PathToUrl($file), $array)
			EndIf
			$OpenOfficeCalc_Document = $document
		EndIf
	EndIf
	Return $document
EndFunc   ;==>OpenOfficeCalc_Open

;
; OpenOffice  Calc が サポートするファイルを閉じる.
; 保存していないデータがある場合も, 強制的に閉じてしまうため,
; 保存が必要な場合は, 本関数実行前に保存を行なっておくこと.
;
; @param $document ドキュメントオブジェクト(OpenOfficeCalc_Open関数の戻り値)を指定する.
;
Func OpenOfficeCalc_Close($document = $OpenOfficeCalc_Document)
	If IsObj($document) Then
		$document.close(True)
	Else
		WinClose($OpenOfficeCalc_ClassName)
	EndIf
	WinWaitClose($OpenOfficeCalc_ClassName)
	$OpenOfficeCalc_Document = 0
	$OpenOfficeCalc_Sheet = 0
EndFunc   ;==>OpenOfficeCalc_Close

;
; ファイルを保存する.
; 引数に何も指定しない場合は, 上書き保存する.
;
; @param $file 保存するファイル名(フルパスで指定する必要がある).
; @param $property ファイルを開く時の属性の配列.
; @param $document ドキュメントオブジェクト(OpenOfficeCalc_Open関数の戻り値)を指定する.
;
Func OpenOfficeCalc_Save($file = 0, $property = 0, $document = $OpenOfficeCalc_Document)
	Local $array[1]
	If Not IsArray($property) Then
		$property = $array
	EndIf
	If IsObj($document) Then
		If IsString($file) Then
			FileDelete($file)
			$document.storeAsURL(FileUtiilty_PathToUrl($file), $property)
		Else
			$document.store()
		EndIf
	EndIf
EndFunc   ;==>OpenOfficeCalc_Save

;
; シートオブジェクトを取得する.
;
; @param $sheet_name シート名.
; @param $document  ドキュメントオブジェクト.
; @return シートオブジェクト.
;
Func OpenOfficeCalc_GetSheet(Const ByRef $sheet_name, $document = $OpenOfficeCalc_Document)
	Local $sheet = 0
	If IsObj($document) Then
		Local $sheets = $document.getSheets()
		Local $is_sheet = $sheets.hasByName($sheet_name)
		If True = $is_sheet Then
			$sheet = $document.Sheets.getByName($sheet_name)
			$OpenOfficeCalc_Sheet = $sheet
		EndIf
	EndIf
	Return $sheet
EndFunc   ;==>OpenOfficeCalc_GetSheet

;
; セルオブジェクトを取得する.
;
; @param $row 行の位置.
; @param $column 列の位置.
; @param $sheet  シートオブジェクト.
; @return セルオブジェクト.
;
Func OpenOfficeCalc_GetCell($row, $column, $sheet = $OpenOfficeCalc_Sheet)
	Local $cell = 0
	If IsObj($sheet) And - 1 < $row And - 1 < $column Then
		$cell = $sheet.getCellByPosition($column, $row)
	EndIf
	Return $cell
EndFunc   ;==>OpenOfficeCalc_GetCell

;
; 範囲オブジェクトを取得する.
;
; @param $start_row 行の開始位置.
; @param $start_column 列の開始位置.
; @param $end_row 行の終了位置.
; @param $end_column 列の終了位置.
; @param $sheet  シートオブジェクト.
; @return 範囲オブジェクト.
;
Func OpenOfficeCalc_GetRange($start_row, $start_column, $end_row, $end_column, $sheet = $OpenOfficeCalc_Sheet)
	Local $range = 0
	If IsObj($sheet) And - 1 < $start_row And - 1 < $start_column And - 1 < $end_row And - 1 < $end_column Then
		$range = $sheet.getCellRangeByPosition($start_column, $start_row, $end_column, $end_row)
	EndIf
	Return $range
EndFunc   ;==>OpenOfficeCalc_GetRange

;
; プロパティを設定する.
; ToDo: 現状の実装では正常に動作しない.
;
; @param $name プロパティ名.
; @param $value 値.
; @return プロパティ.
;
Func OpenOfficeCalc_SetProperty($name, $value)
	Local $manager = OpenOfficeCalc_CreateServiceManager()
	Local $property = $manager.Bridge_GetStruct($OpenOfficeCalc_PropertyValue)
	$property.Name = $name
	$property.Value = $value
	Return $property
EndFunc   ;==>OpenOfficeCalc_SetProperty
#endregion Public_Method