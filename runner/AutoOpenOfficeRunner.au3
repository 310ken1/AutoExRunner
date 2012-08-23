#include-once

;
; == オプション ==
; AutoOpenOfficeRuuner.au3 の振る舞いを変更する場合は,
; $AutoOpenOfficeRuunerConfig グローバル変数 に 設定ファイルを設定すること.
;
Global $AutoOpenOfficeRuunerConfig = ""

Const $AutoOpenOfficeRuunerSettingTag = "AutoOpenOfficeRuunerSetting"

;===============================================================================
; Pubilc Method
;===============================================================================
;
; 指定したOpenOffice Calcファイルを読込み, 指定した関数を呼出す.
;
; @param $file 読込むOpenOffice Calcファイル名.
; @param $sheet_name 読込むOpenOffice Calcファイルのシート名.
; @param $callback_name コールバックする関数名.
;
Func AutoOpenOfficeRunner($file, $sheet_name, $callback_name)
	Local $server_manager = ObjCreate("com.sun.star.ServiceManager")
	Local $desktop = $server_manager.createInstance("com.sun.star.frame.Desktop")

	Local $property[1]
	$property[0] = SetProperty("Hidden", "1")
	Local $component = $desktop.loadComponentFromURL(PathToUrl($file), "_default", 0, $property)
	If (Not @error) And IsObj($component) Then
		Local $sheet = $component.Sheets.getByName($sheet_name)
		If (Not @error) And IsObj($sheet) Then
			Local $line = Int(IniRead($AutoOpenOfficeRuunerConfig, $AutoOpenOfficeRuunerSettingTag, "StartLine", 2))
			Local $column = Int(IniRead($AutoOpenOfficeRuunerConfig, $AutoOpenOfficeRuunerSettingTag, "StartColumn", 0))
			While True
				Local $cell = $sheet.getCellByPosition($column, $line)
				Local $value = $cell.String
				If IsNoEnd($value) Then
					ExitLoop
				ElseIf StringIsDigit($value) Then
					Call($callback_name, $sheet, $line)
				EndIf
				$line += 1
			WEnd
			$excel = 0
		EndIf
	Else
		MsgBox(0, "Error", "Could not open " & $file & " as an Excel Object.")
	EndIf

	$component.close(True)
EndFunc   ;==>AutoOpenOfficeRunner

;
; セルを取得する.
;
; @param $sheet 実行中のシートオブジェクト.
; @param $line 実行中の行.
; @param $key 取得したいセルの項目名.
; @return セル.
;
Func GetCell($sheet, $line, $key)
	Local $key_line = Int(IniRead($AutoOpenOfficeRuunerConfig, $AutoOpenOfficeRuunerSettingTag, "KeyLine", 1))
	Local $key_column = Int(IniRead($AutoOpenOfficeRuunerConfig, $AutoOpenOfficeRuunerSettingTag, "KeyColumn", 1))
	Local $key_column_max = Int(IniRead($AutoOpenOfficeRuunerConfig, $AutoOpenOfficeRuunerSettingTag, "KeyColumnMax", 9))
	Local $cell = 0
	While $key_column <= $key_column_max
		Local $value = $sheet.getCellByPosition($key_column, $key_line).String
		If $key = $value Then
			$cell = $sheet.getCellByPosition($key_column, $line)
			ExitLoop
		EndIf
		$key_column += 1
	WEnd
	Return $cell
EndFunc   ;==>GetCell

;
; セルの値(文字列)を取得する.
;
; @param $sheet 実行中のシートオブジェクト.
; @param $line 実行中の行.
; @param $key 取得したいセルの項目名.
; @return セルの値(文字列).
;
Func GetString($sheet, $line, $key)
	Local $value = ""
	Local $cell = GetCell($sheet, $line, $key)
	If IsObj($cell) Then
		$value = $cell.String
	EndIf
	Return $value
EndFunc   ;==>GetString

;
; No を取得する.
;
; @param $sheet 実行中のシートオブジェクト.
; @param $line 実行中の行.
; @return No.
;
Func GetNo($sheet, $line)
	Local $column = Int(IniRead($AutoOpenOfficeRuunerConfig, $AutoOpenOfficeRuunerSettingTag, "StartColumn", 0))
	Return $sheet.getCellByPosition($column, $line).getString
EndFunc   ;==>GetNo

#region Private Method
;===============================================================================
; Private Method
;===============================================================================
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

;
; プロパティを設定する.
;
; @param $name プロパティ名.
; @param $value 値.
; @return プロパティ.
;
Func SetProperty($name, $value)
	Local $service_manager = ObjCreate("com.sun.star.ServiceManager")
	Local $property = $service_manager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
	$property.Name = $name
	$property.Value = $value
	Return $property
EndFunc   ;==>SetProperty

;
; パスをURLに変換する.
;
; @param $path パス.
; @return  URL.
;
Func PathToUrl($path)
	Local $ret = "file:///" & StringRegExpReplace($path, "\\", "/")
	Return $ret
EndFunc   ;==>PathToUrl

#endregion Private Method