#include-once
#include "..\utility\Console.au3"

;
; == �I�v�V���� ==
; AutoExRuuner.au3 �̐U�镑����ύX����ꍇ��,
; $AutoExRunnerConfig �O���[�o���ϐ� �� �ݒ�t�@�C����ݒ肷�邱��.
;
Global $AutoExRunnerConfig = ""

;===============================================================================
; Pubilc Method
;===============================================================================
;
; �w�肵��Excel�t�@�C����Ǎ���, �w�肵���֐����ďo��.
;
; @param $file �Ǎ���Excel�t�@�C����.
; @param $sheet_name �Ǎ���Excel�t�@�C���̃V�[�g��.
; @param $callback_name �R�[���o�b�N����֐���.
;
Func AutoExRunner($file, $sheet_name, $callback_name)
	Local $excel = ObjGet($file)
	If (Not @error) And IsObj($excel) Then
		Local $sheet = $excel.Worksheets($sheet_name)
		If (Not @error) And IsObj($sheet) Then
			Local $line = Int(IniRead($AutoExRunnerConfig, "Setting", "StartLine", 3))
			Local $column = Int(IniRead($AutoExRunnerConfig, "Setting", "StartColumn", 1))
			While True
				Local $cell = $sheet.cells($line, $column)
				Local $value = $cell.value
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
EndFunc   ;==>AutoExRunner

;
; �Z�����擾����.
;
; @param $sheet ���s���̃V�[�g�I�u�W�F�N�g.
; @param $line ���s���̍s.
; @param $key �擾�������Z���̍��ږ�.
; @return �Z��.
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
; No ���擾����.
;
; @param $sheet ���s���̃V�[�g�I�u�W�F�N�g.
; @param $line ���s���̍s.
; @return No.
;
Func GetNo($sheet, $line)
	Local $column = Int(IniRead($AutoExRunnerConfig, "Setting", "StartColumn", 1))
	Return $sheet.cells($line, $column).value
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

#endregion
