#include "..\..\runner\AutoOpenOfficeRunner.au3"
#include "..\..\app\OpenSSL.au3"

$OpenSSLCmd = @ScriptDir & "\..\..\bin\openssl.exe"
$AutoOpenOfficeRuunerConfig = @ScriptDir & "\CertificateGenerator.ini"

;
; �o�̓t�H���_.
;
Const $OutputDir = @ScriptDir & "\out"

;
; ���̓t�@�C��.
;
Const $InputFile = @ScriptDir & "\CertificateList.ods"

;
; �T�u�W�F�N�g.
;
Local $subject[8][2] = [ _
		["C", ""], _
		["ST", ""], _
		["L", ""], _
		["O", ""], _
		["OU", ""], _
		["CN", ""], _
		["emailAddress", ""] _
		]

;
; ���C���֐��Ăяo��.
;
Main()

;
; ���C���֐�.
;
Func Main()
	DirCreate($OutputDir)
	FileChangeDir($OutputDir)

	AutoOpenOfficeRunner($InputFile, "���[�g�ؖ���", "CreateRootCrt")
	AutoOpenOfficeRunner($InputFile, "���ԏؖ���", "CreateIntermediateCrt")
	AutoOpenOfficeRunner($InputFile, "�T�[�o�ؖ���", "CreateServerCrt")
EndFunc   ;==>Main

;
; ���[�g�ؖ����̐�������.
;
; @param $sheet ���s���̃V�[�g�I�u�W�F�N�g.
; @param $line ���s���̍s.
;
Func CreateRootCrt($sheet, $line)
	InitSubject()
	SetSubject($sheet, $line)
	CreateRootCertificate( _
			GetString($sheet, $line, "����"), _
			GetString($sheet, $line, "����"), _
			GetString($sheet, $line, "���b�Z�[�W�_�C�W�F�X�g"), _
			GetString($sheet, $line, "�L������"), _
			"HookCreateRootCrt" _
			)
EndFunc   ;==>CreateRootCrt

;
; ���[�g�ؖ����̐ݒ�t�@�C�����������t�b�N�֐�.
;
; @param $config �ݒ�t�@�C����.
;
Func HookCreateRootCrt($config)
	WriteSubject($config)
EndFunc   ;==>HookCreateRootCrt

;
; ���ԏؖ����̐�������.
;
; @param $sheet ���s���̃V�[�g�I�u�W�F�N�g.
; @param $line ���s���̍s.
;
Func CreateIntermediateCrt($sheet, $line)
	InitSubject()
	SetSubject($sheet, $line)
	CreateIntermediateCertificate( _
			GetString($sheet, $line, "����"), _
			GetString($sheet, $line, "����"), _
			GetString($sheet, $line, "���b�Z�[�W�_�C�W�F�X�g"), _
			GetString($sheet, $line, "�L������"), _
			StringRegExpReplace($OutputDir & "\" & GetString($sheet, $line, "�F�؋�"), "\\", "/"), _
			"HookCreateIntermediateCrt" _
			)
EndFunc   ;==>CreateIntermediateCrt

;
; ���ԏؖ����̐ݒ�t�@�C�����������t�b�N�֐�.
;
; @param $config �ݒ�t�@�C����.
;
Func HookCreateIntermediateCrt($config)
	WriteSubject($config)
EndFunc   ;==>HookCreateIntermediateCrt

;
; �T�[�o�ؖ����̐�������.
;
; @param $sheet ���s���̃V�[�g�I�u�W�F�N�g.
; @param $line ���s���̍s.
;
Func CreateServerCrt($sheet, $line)
	InitSubject()
	SetSubject($sheet, $line)
	CreateServerCertificate( _
			GetString($sheet, $line, "����"), _
			GetString($sheet, $line, "����"), _
			GetString($sheet, $line, "���b�Z�[�W�_�C�W�F�X�g"), _
			GetString($sheet, $line, "�L������"), _
			StringRegExpReplace($OutputDir & "\" & GetString($sheet, $line, "�F�؋�"), "\\", "/"), _
			"HookCreateServerCrt" _
			)
EndFunc   ;==>CreateServerCrt

;
; �T�[�o�ؖ����̐ݒ�t�@�C�����������t�b�N�֐�.
;
; @param $config �ݒ�t�@�C����.
;
Func HookCreateServerCrt($config)
	WriteSubject($config)
EndFunc   ;==>HookCreateServerCrt

;
; �T�u�W�F�N�g������������.
;
Func InitSubject()
	Local $count = UBound($subject, 1)
	For $i = 0 To $count - 1
		$subject[$i][1] = ""
	Next
EndFunc   ;==>InitSubject

;
; ���̓t�@�C������l��Ǎ���, �T�u�W�F�N�g�ɐݒ肷��.
;
; @param $sheet ���s���̃V�[�g�I�u�W�F�N�g.
; @param $line ���s���̍s.
;
Func SetSubject($sheet, $line)
	Local $count = UBound($subject, 1)
	For $i = 0 To $count - 1
		Local $value = GetString($sheet, $line, $subject[$i][0])
		If Not "" = $value Then
			$subject[$i][1] = $value
		EndIf
	Next
EndFunc   ;==>SetSubject

;
; �T�u�W�F�N�g��ݒ�t�@�C���ɏ�������.
;
; @param $config �ݒ�t�@�C����.
;
Func WriteSubject($config)
	Local $count = UBound($subject, 1)
	For $i = 0 To $count - 1
		If Not "" = $subject[$i][1] Then
			IniWrite($config, "req_distinguished_name", $subject[$i][0], $subject[$i][1])
		EndIf
	Next
EndFunc   ;==>WriteSubject

