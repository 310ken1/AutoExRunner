#include-once
#include "..\utility\Console.au3"

;
; !! �K�{ !!
; �{�\�[�X�R�[�h��Œ�`���ꂽ���\�b�h�𗘗p����ꍇ��,
; $OpenSSLCmd �O���[�o���ϐ� �� openssl.exe �ւ̃p�X��ݒ肷�邱��.
;
Global $OpenSSLCmd = "openssl.exe"

Const $OpenSSLConsoleClass = "[CLASS:ConsoleWindowClass]"

Const $OpenSSLKeyName = "private.key"
Const $OpenSSLConfigName = "openssl.cnf"
Const $OpenSSLCertificateName = "certificate.crt"
Const $OpenSSLCertificationSigningRequest = "CertificationSigningRequest.csr"
Const $OpenSSLDataBaseName = "database"
Const $OpenSSLSerialName = "serial"
Const $OpenSSLRandName = ".rand"

;===============================================================================
; Pubilc Method
;===============================================================================
;
; ���[�g�ؖ����𐶐�����.
; $name �Ŏw�肵���t�H���_�𐶐���, ���̃t�H���_�Ɋe��t�@�C����ۑ�����.
;
; @param $name ����.
; @param $bits ����.
; @param $md �����̃��b�Z�[�W�_�C�W�F�X�g.
; @param $days �ؖ����̗L������.
; @param $hook �������̐ݒ�t�@�C�������������邽�߂̃t�b�N�֐���.
;
Func CreateRootCertificate($name, $bits, $md, $days, $hook = 0)
	DirCreate($name)

	Local $keyfile = $name & "\" & $OpenSSLKeyName
	CreateRsaPrivateKey($bits, $keyfile)

	Local $configfile = $name & "\" & $OpenSSLConfigName
	CreateConfigurationTemplate($configfile, $name)
	If IsString($hook) Then
		Call($hook, $configfile)
	EndIf

	Local $crtfile = $name & "\" & $OpenSSLCertificateName
	SelfSigned($keyfile, $days, $md, $configfile, $crtfile)

	Local $dbfile = $name & "\" & $OpenSSLDataBaseName
	CreateDataBase($dbfile)

	Local $serialfile = $name & "\" & $OpenSSLSerialName
	CreateSerialFile($crtfile, $serialfile)
EndFunc   ;==>CreateRootCertificate

;
; ���ԏؖ����𐶐�����.
;
; @param $name ����.
; @param $bits ����.
; @param $md �����̃��b�Z�[�W�_�C�W�F�X�g.
; @param $days �ؖ����̗L������.
; @param $ca �F�؋ǃt�H���_�ւ̃p�X.
; @param $hook �������̐ݒ�t�@�C�������������邽�߂̃t�b�N�֐���.
;
Func CreateIntermediateCertificate($name, $bits, $md, $days, $ca, $hook = 0)
	DirCreate($name)

	Local $keyfile = $name & "\" & $OpenSSLKeyName
	CreateRsaPrivateKey($bits, $keyfile)

	Local $configfile = $name & "\" & $OpenSSLConfigName
	CreateConfigurationTemplate($configfile, $name)
	IniWrite($configfile, "CA_default", "dir", $ca)
	If IsString($hook) Then
		Call($hook, $configfile)
	EndIf

	Local $csrfile = $name & "\" & $OpenSSLCertificationSigningRequest
	CreateCertificationSigningRequest($keyfile, $configfile, $csrfile)

	Local $crtfile = $name & "\" & $OpenSSLCertificateName
	Signed($csrfile, $days, $md, $configfile, $crtfile)

	Local $dbfile = $name & "\" & $OpenSSLDataBaseName
	CreateDataBase($dbfile)

	Local $serialfile = $name & "\" & $OpenSSLSerialName
	CreateSerialFile($crtfile, $serialfile)
EndFunc   ;==>CreateIntermediateCertificate

;
; �T�[�o�ؖ����𐶐�����.
;
; @param $name ����.
; @param $bits ����.
; @param $md �����̃��b�Z�[�W�_�C�W�F�X�g.
; @param $days �ؖ����̗L������.
; @param $ca �F�؋ǃt�H���_�ւ̃p�X.
; @param $hook �������̐ݒ�t�@�C�������������邽�߂̃t�b�N�֐���.
;
Func CreateServerCertificate($name, $bits, $md, $days, $ca, $hook = 0)
	DirCreate($name)

	Local $keyfile = $name & "\" & $OpenSSLKeyName
	CreateRsaPrivateKey($bits, $keyfile)

	Local $configfile = $name & "\" & $OpenSSLConfigName
	CreateConfigurationTemplate($configfile, $name)
	IniWrite($configfile, "CA_default", "dir", $ca)
	IniWrite($configfile, "usr_cert", "basicConstraints", "CA:FALSE")
	If IsString($hook) Then
		Call($hook, $configfile)
	EndIf

	Local $csrfile = $name & "\" & $OpenSSLCertificationSigningRequest
	CreateCertificationSigningRequest($keyfile, $configfile, $csrfile)

	Local $crtfile = $name & "\" & $OpenSSLCertificateName
	Signed($csrfile, $days, $md, $configfile, $crtfile)
EndFunc   ;==>CreateServerCertificate

;
; PEM�`���̏ؖ�����DER�`���ɕϊ�����.
;
; @param $pem PEM�`���ؖ���(����).
; @param $der DER�`���ؖ���(�o��).
;
Func PemToDer($pem, $der)
	Local $cmd = StringFormat("%s x509�@-outform DER -in %s -out %s", $OpenSSLCmd, $pem, $der)
	ConsoleWriteLn($cmd)
	RunWait($cmd)
EndFunc   ;==>PemToDer

;===============================================================================
; Private Method
;===============================================================================
;
; RSA�閧���𐶐�����.
;
; @param $bits �Í����̃r�b�g��.
; @param $output_file �o�̓t�@�C����(�閧���t�@�C��).
;
Func CreateRsaPrivateKey($bits, $output_file)
	Local $cmd = StringFormat("%s genrsa -out %s %s", $OpenSSLCmd, $output_file, $bits)
	ConsoleWriteLn($cmd)
	RunWait($cmd)
EndFunc   ;==>CreateRsaPrivateKey

;
; ���ȏ���X.509�ؖ����𐶐�����.
;
; @param $key_file �閧��.
; @param $days �ؖ����̗L������.
; @param $md ���b�Z�[�W�_�C�W�F�X�g.
; @param $config_file �ݒ�t�@�C����.
; @param $output_file �o�̓t�@�C����(���ȏ����ؖ���).
;
Func SelfSigned($key_file, $days, $md, $config_file, $output_file)
	Local $cmd = StringFormat("%s req -new -x509 -key %s -%s -out %s -config %s -days %s", $OpenSSLCmd, $key_file, $md, $output_file, $config_file, $days)
	ConsoleWriteLn($cmd)
	RunWait($cmd)
EndFunc   ;==>SelfSigned

;
; X.509�ؖ����𐶐�����.
;
; @param $csr_file �����v��(CSR)�t�@�C����.
; @param $days �ؖ����̗L������.
; @param $md ���b�Z�[�W�_�C�W�F�X�G�g.
; @param $config_file �ݒ�t�@�C����.
; @param $output_file �o�̓t�@�C����(X.509�ؖ���).
;
Func Signed($csr_file, $days, $md, $config_file, $output_file)
	Local $cmd = StringFormat("%s ca -md %s -in %s -out %s -config %s -days %s", $OpenSSLCmd, $md, $csr_file, $output_file, $config_file, $days)
	ConsoleWriteLn($cmd)
	Run($cmd)
	WinWaitActive($OpenSSLConsoleClass)
	ControlSend($OpenSSLConsoleClass, "", "", "y{ENTER}")
	ControlSend($OpenSSLConsoleClass, "", "", "y{ENTER}")
EndFunc   ;==>Signed

;
; �����v��(CSR)�𐶐�����.
;
; @param $key_file �閧��.
; @param $config_file �ݒ�t�@�C����.
; @param $output_file �o�̓t�@�C����(�����v��).
;
Func CreateCertificationSigningRequest($key_file, $config_file, $output_file)
	Local $cmd = StringFormat("%s req -new -key %s -out %s -config %s", $OpenSSLCmd, $key_file, $output_file, $config_file)
	ConsoleWriteLn($cmd)
	RunWait($cmd)
EndFunc   ;==>CreateCertificationSigningRequest

;
; �ݒ�t�@�C���̃e���v���[�g�𐶐�����.
;
; @param $file �ݒ�t�@�C��.
; @param $ca_dir �F�؋Ǐ�񂪊i�[���ꂽ�t�H���_.
;
Func CreateConfigurationTemplate($file, $ca_dir)
	Local $ca[1][2] = [ _
			["default_ca", "CA_default"] _
			]
	IniWriteSection($file, "ca", $ca, 0)

	Local $ca_default[9][2] = [ _
			["dir", $ca_dir], _
			["certificate", "$dir/" & $OpenSSLCertificateName], _
			["private_key", "$dir/" & $OpenSSLKeyName], _
			["database", "$dir/" & $OpenSSLDataBaseName], _
			["serial", "$dir/" & $OpenSSLSerialName], _
			["policy", "policy_anything"], _
			["x509_extensions", "usr_cert"], _
			["new_certs_dir", "$dir/"], _
			["RANDFILE", "$dir" & $OpenSSLRandName] _
			]
	IniWriteSection($file, "CA_default", $ca_default, 0)

	Local $policy_anything[7][2] = [ _
			["countryName", "optional"], _
			["stateOrProvinceName", "optional"], _
			["localityName", "optional"], _
			["organizationName", "optional"], _
			["organizationalUnitName", "optional"], _
			["commonName", "supplied"], _
			["emailAddress", "optional"] _
			]
	IniWriteSection($file, "policy_anything", $policy_anything, 0)

	Local $usr_cert[1][2] = [ _
			["basicConstraints", "CA:TRUE"] _
			]
	IniWriteSection($file, "usr_cert", $usr_cert, 0)

	Local $req[3][2] = [ _
			["distinguished_name", "req_distinguished_name"], _
			["x509_extensions", "v3_ca"], _
			["prompt", "no"] _
			]
	IniWriteSection($file, "req", $req, 0)

	Local $reqreq_distinguished_name[1][2] = [ _
			["CN", "example.com"] _
			]
	IniWriteSection($file, "req_distinguished_name", $reqreq_distinguished_name, 0)

	Local $v3_ca[3][2] = [ _
			["basicConstraints", "CA:TRUE"], _
			["subjectKeyIdentifier", "hash"], _
			["authorityKeyIdentifier", "keyid:always,issuer:always"] _
			]
	IniWriteSection($file, "v3_ca", $v3_ca, 0)
EndFunc   ;==>CreateConfigurationTemplate

;
; �f�[�^�x�[�X�𐶐�����.
;
; @param $file ��������t�@�C����.
;
Func CreateDataBase($file)
	FileOpen($file, 2 + 8)
EndFunc   ;==>CreateDataBase

;
; �V���A���t�@�C���𐶐�����.
;
; @param $crt_file �ؖ���.
; @param $serial_file ��������V���A���t�@�C����.
;
Func CreateSerialFile($crt_file, $serial_file)
	Local $cmd = StringFormat("%s x509 -in %s -noout -next_serial -out %s", $OpenSSLCmd, $crt_file, $serial_file)
	ConsoleWriteLn($cmd)
	RunWait($cmd)
EndFunc   ;==>CreateSerialFile