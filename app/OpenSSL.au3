#include-once
#include "..\utility\Console.au3"

;
; !! 必須 !!
; 本ソースコード上で定義されたメソッドを利用する場合は,
; $OpenSSLCmd グローバル変数 に openssl.exe へのパスを設定すること.
;
Global $OpenSSLCmd = ""

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

	Local $keyfile = $name & "\private.key"
	CreateRsaPrivateKey($bits, $keyfile)

	Local $confgfile = $name & "\config"
	CreateConfigurationTemplate($confgfile, $name)
	If IsString($hook) Then
		Call($hook, $confgfile)
	EndIf

	Local $crtfile = $name & "\root.crt"
	SelfSigned($keyfile, $days, $md, $confgfile, $crtfile)

	Local $dbfile = $name & "\database"
	CreateDataBase($dbfile)

	Local $serialfile = $name & "\serial"
	CreateSerialFile($crtfile, $serialfile)
EndFunc   ;==>CreateRootCertificate

;
; 中間証明書を生成する.
;
Func CreateIntermediateCertificate($name, $bits, $md, $days, $ca, $hook = 0)
	DirCreate($name)

	Local $keyfile = $name & "\private.key"
	CreateRsaPrivateKey($bits, $keyfile)

	Local $confgfile = $name & "\config"
	CreateConfigurationTemplate($confgfile, $name)
	IniWrite($confgfile, "CA_default", "dir", "./" & $ca)
	If IsString($hook) Then
		Call($hook, $confgfile)
	EndIf

	Local $csrfile = $name & "\req.csr"
	CreateCertificationSigningRequest($keyfile, $confgfile, $csrfile)

	Local $crtfile = $name & "\intermediate.crt"
	Signed($csrfile, $days, $md, $confgfile, $crtfile)
EndFunc   ;==>CreateIntermediateCertificate

;
; サーバ証明書を生成する.
;
Func CreateServerCertificate()
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
	WinWaitActive($OpenSSLCmd)
	ControlSend($OpenSSLCmd, "", "", "y{ENTER}")
	ControlSend($OpenSSLCmd, "", "", "y{ENTER}")
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
; @param $name 名称(フォルダ名).
;
Func CreateConfigurationTemplate($file, $name)
	Local $ca[1][2] = [["default_ca", "CA_default"]]
	IniWriteSection($file, "ca", $ca, 0)

	Local $ca_default[8][2] = [ _
			["dir", "./" & $name], _
			["certificate", "$dir/root.crt"], _
			["private_key", "$dir/private.key"], _
			["database", "$dir/database"], _
			["serial", "$dir/serial"], _
			["policy", "policy_anything"], _
			["x509_extensions", "usr_cert"], _
			["new_certs_dir", "$dir"] _
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
			["commonName", "example.com"] _
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