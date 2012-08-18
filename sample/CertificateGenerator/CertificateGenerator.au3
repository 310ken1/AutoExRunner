#include "..\..\runner\AutoExRunner.au3"
#include "..\..\app\OpenSSL.au3"

$OpenSSLCmd = "..\..\bin\openssl.exe"

AutoExRunner(@ScriptDir & "\CertificateList.xls", "���[�g�ؖ���", "CreateRootCA")

Func CreateRootCA($sheet, $line)
	Local $factor[4] = ["����", "����", "���b�Z�[�W�_�C�W�F�X�g", "�L������"]
	Local $value[4]
	Local $i = 0
	For $f In $factor
		Local $cell = GetCell($sheet, $line, $f)
		$value[$i] = $cell.value
		$i += 1
	Next

	CreateRootCertificate($value[0], $value[1], $value[2], $value[3])
EndFunc   ;==>CreateRootCA
