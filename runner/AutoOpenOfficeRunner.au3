#include-once

;
; == �I�v�V���� ==
; AutoOpenOfficeRuuner.au3 �̐U�镑����ύX����ꍇ��,
; $AutoOpenOfficeRuunerConfig �O���[�o���ϐ� �� �ݒ�t�@�C����ݒ肷�邱��.
;
Global $AutoOpenOfficeRuunerConfig = ""

Const $AutoOpenOfficeRuunerSettingTag = "AutoOpenOfficeRuunerSetting"

;===============================================================================
; Pubilc Method
;===============================================================================
;
; �w�肵��OpenOffice Calc�t�@�C����Ǎ���, �w�肵���֐����ďo��.
;
; @param $file �Ǎ���OpenOffice Calc�t�@�C����.
; @param $sheet_name �Ǎ���OpenOffice Calc�t�@�C���̃V�[�g��.
; @param $callback_name �R�[���o�b�N����֐���.
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
; �Z�����擾����.
;
; @param $sheet ���s���̃V�[�g�I�u�W�F�N�g.
; @param $line ���s���̍s.
; @param $key �擾�������Z���̍��ږ�.
; @return �Z��.
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
; �Z���̒l(������)���擾����.
;
; @param $sheet ���s���̃V�[�g�I�u�W�F�N�g.
; @param $line ���s���̍s.
; @param $key �擾�������Z���̍��ږ�.
; @return �Z���̒l(������).
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
; No ���擾����.
;
; @param $sheet ���s���̃V�[�g�I�u�W�F�N�g.
; @param $line ���s���̍s.
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
; �I�[���`�F�b�N����.
;
; @param $value No�̒l.
; @return �I�[�̗L��.
;
Func IsNoEnd($value)
	Local $ret = False
	If "End" = $value Or "E" = $value Then
		$ret = True
	EndIf
	Return $ret
EndFunc   ;==>IsNoEnd

;
; �v���p�e�B��ݒ肷��.
;
; @param $name �v���p�e�B��.
; @param $value �l.
; @return �v���p�e�B.
;
Func SetProperty($name, $value)
	Local $service_manager = ObjCreate("com.sun.star.ServiceManager")
	Local $property = $service_manager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
	$property.Name = $name
	$property.Value = $value
	Return $property
EndFunc   ;==>SetProperty

;
; �p�X��URL�ɕϊ�����.
;
; @param $path �p�X.
; @return  URL.
;
Func PathToUrl($path)
	Local $ret = "file:///" & StringRegExpReplace($path, "\\", "/")
	Return $ret
EndFunc   ;==>PathToUrl

#endregion Private Method