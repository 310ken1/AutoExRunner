#include-once
#include "..\utility\Console.au3"

;
; !! 必須 !!
; 本ソースコード上で定義されたメソッドを利用する場合は,
; $OpenSSLCmd グローバル変数 に openssl.exe へのパスを設定すること.
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
; ルート証明書を生成する.
; $name で指定したフォルダを生成し, そのフォルダに各種ファイルを保存する.
;
; @param $name 名称.
; @param $bits 鍵長.
; @param $md 署名のメッセージダイジェスト.
; @param $days 証明書の有効期限.
; @param $hook 生成時の設定ファイルを書き換えるためのフック関数名.
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
; 中間証明書を生成する.
;
; @param $name 名称.
; @param $bits 鍵長.
; @param $md 署名のメッセージダイジェスト.
; @param $days 証明書の有効期限.
; @param $ca 認証局フォルダへのパス.
; @param $hook 生成時の設定ファイルを書き換えるためのフック関数名.
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
; サーバ証明書を生成する.
;
; @param $name 名称.
; @param $bits 鍵長.
; @param $md 署名のメッセージダイジェスト.
; @param $days 証明書の有効期限.
; @param $ca 認証局フォルダへのパス.
; @param $hook 生成時の設定ファイルを書き換えるためのフック関数名.
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
; PEM形式の証明書をDER形式に変換する.
;
; @param $pem PEM形式証明書(入力).
; @param $der DER形式証明書(出力).
;
Func PemToDer($pem, $der)
	Local $cmd = StringFormat("%s x509　-outform DER -in %s -out %s", $OpenSSLCmd, $pem, $der)
	ConsoleWriteLn($cmd)
	RunWait($cmd)
EndFunc   ;==>PemToDer

;===============================================================================
; Private Method
;===============================================================================
;
; RSA秘密鍵を生成する.
;
; @param $bits 暗号鍵のビット数.
; @param $output_file 出力ファイル名(秘密鍵ファイル).
;
Func CreateRsaPrivateKey($bits, $output_file)
	Local $cmd = StringFormat("%s genrsa -out %s %s", $OpenSSLCmd, $output_file, $bits)
	ConsoleWriteLn($cmd)
	RunWait($cmd)
EndFunc   ;==>CreateRsaPrivateKey

;
; 自己署名X.509証明書を生成する.
;
; @param $key_file 秘密鍵.
; @param $days 証明書の有効期限.
; @param $md メッセージダイジェスト.
; @param $config_file 設定ファイル名.
; @param $output_file 出力ファイル名(自己署名証明書).
;
Func SelfSigned($key_file, $days, $md, $config_file, $output_file)
	Local $cmd = StringFormat("%s req -new -x509 -key %s -%s -out %s -config %s -days %s", $OpenSSLCmd, $key_file, $md, $output_file, $config_file, $days)
	ConsoleWriteLn($cmd)
	RunWait($cmd)
EndFunc   ;==>SelfSigned

;
; X.509証明書を生成する.
;
; @param $csr_file 署名要求(CSR)ファイル名.
; @param $days 証明書の有効期限.
; @param $md メッセージダイジェスエト.
; @param $config_file 設定ファイル名.
; @param $output_file 出力ファイル名(X.509証明書).
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
; 署名要求(CSR)を生成する.
;
; @param $key_file 秘密鍵.
; @param $config_file 設定ファイル名.
; @param $output_file 出力ファイル名(署名要求).
;
Func CreateCertificationSigningRequest($key_file, $config_file, $output_file)
	Local $cmd = StringFormat("%s req -new -key %s -out %s -config %s", $OpenSSLCmd, $key_file, $output_file, $config_file)
	ConsoleWriteLn($cmd)
	RunWait($cmd)
EndFunc   ;==>CreateCertificationSigningRequest

;
; 設定ファイルのテンプレートを生成する.
;
; @param $file 設定ファイル.
; @param $ca_dir 認証局情報が格納されたフォルダ.
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
; データベースを生成する.
;
; @param $file 生成するファイル名.
;
Func CreateDataBase($file)
	FileOpen($file, 2 + 8)
EndFunc   ;==>CreateDataBase

;
; シリアルファイルを生成する.
;
; @param $crt_file 証明書.
; @param $serial_file 生成するシリアルファイル名.
;
Func CreateSerialFile($crt_file, $serial_file)
	Local $cmd = StringFormat("%s x509 -in %s -noout -next_serial -out %s", $OpenSSLCmd, $crt_file, $serial_file)
	ConsoleWriteLn($cmd)
	RunWait($cmd)
EndFunc   ;==>CreateSerialFile