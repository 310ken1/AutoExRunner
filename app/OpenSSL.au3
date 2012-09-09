#include-once
#include "..\utility\ConsoleUtility.au3"

#region グローバル変数定義
;
; !! 必須 !!
; 本ソースコード上で定義されたメソッドを利用する場合は,
; $OpenSSL_CmdPath グローバル変数 に openssl.exe へのパスを設定すること.
;
Global $OpenSSL_CmdPath = "openssl.exe"

;
; デバッグログフラグ.
; 1 を指定することで, デバッグログが出力される.
;
Global $OpenSSL_DebugLog = 0
#endregion グローバル変数定義

#region 定数定義
Const $OpenSSL_ConsoleClass = "[CLASS:ConsoleWindowClass]"

Const $OpenSSL_KeyName = "private.key"
Const $OpenSSL_ConfigName = "openssl.cnf"
Const $OpenSSL_CertificatePemName = "certificate.crt"
Const $OpenSSL_CertificateDerName = "certificate.der"
Const $OpenSSL_CertificationSigningRequest = "CertificationSigningRequest.csr"
Const $OpenSSL_DataBaseName = "database"
Const $OpenSSL_SerialName = "serial"
#endregion 定数定義

#region パブリックメソッド
;
; ルート証明書を生成する.
; $name で指定したフォルダを生成し, そのフォルダに各種ファイルを保存する.
; $name で指定したフォルダが既に存在している場合は, 何もしない.
;
; @param $name 名称.
; @param $bits 鍵長.
; @param $md 署名のメッセージダイジェスト.
; @param $days 証明書の有効期限.
; @param $hook 生成時の設定ファイルを書き換えるためのフック関数名.
;
Func OpenSSL_CreateRootCertificate($name, $bits, $md, $days, $hook = 0)
	If Not FileExists($name) Then
		Local $path = FileUtiilty_ChangeDir($name)

		OpenSSL_CreateRsaPrivateKey($bits, $OpenSSL_KeyName)
		OpenSSL_CreateConfigurationTemplate($OpenSSL_ConfigName, $name)
		If IsString($hook) Then
			Call($hook, $OpenSSL_ConfigName)
		EndIf
		OpenSSL_SelfSigned($OpenSSL_KeyName, $days, $md, $OpenSSL_ConfigName, $OpenSSL_CertificatePemName)
		OpenSSL_PemToDer($OpenSSL_CertificatePemName, $OpenSSL_CertificateDerName)

		OpenSSL_CreateDataBase($OpenSSL_DataBaseName)
		OpenSSL_CreateSerialFile($OpenSSL_CertificatePemName, $OpenSSL_SerialName)

		FileChangeDir($path)
	EndIf
EndFunc   ;==>OpenSSL_CreateRootCertificate

;
; 中間証明書を生成する.
; $name で指定したフォルダを生成し, そのフォルダに各種ファイルを保存する.
; $name で指定したフォルダが既に存在している場合は, 何もしない.
;
; @param $name 名称.
; @param $bits 鍵長.
; @param $md 署名のメッセージダイジェスト.
; @param $days 証明書の有効期限.
; @param $ca 認証局フォルダへのフルパス.
; @param $hook 生成時の設定ファイルを書き換えるためのフック関数名.
;
Func OpenSSL_CreateIntermediateCertificate($name, $bits, $md, $days, $ca, $hook = 0)
	If Not FileExists($name) Then
		Local $path = FileUtiilty_ChangeDir($name)

		OpenSSL_CreateRsaPrivateKey($bits, $OpenSSL_KeyName)
		OpenSSL_CreateConfigurationTemplate($OpenSSL_ConfigName, $name)
		IniWrite($OpenSSL_ConfigName, "CA_default", "dir", $ca)
		If IsString($hook) Then
			Call($hook, $OpenSSL_ConfigName)
		EndIf
		OpenSSL_CreateCertificationSigningRequest($OpenSSL_KeyName, $OpenSSL_ConfigName, $OpenSSL_CertificationSigningRequest)
		OpenSSL_Signed($OpenSSL_CertificationSigningRequest, $days, $md, $OpenSSL_ConfigName, $OpenSSL_CertificatePemName)
		OpenSSL_PemToDer($OpenSSL_CertificatePemName, $OpenSSL_CertificateDerName)

		OpenSSL_CreateDataBase($OpenSSL_DataBaseName)
		OpenSSL_CreateSerialFile($OpenSSL_CertificatePemName, $OpenSSL_SerialName)

		FileChangeDir($path)
	EndIf
EndFunc   ;==>OpenSSL_CreateIntermediateCertificate

;
; サーバ証明書を生成する.
; $name で指定したフォルダを生成し, そのフォルダに各種ファイルを保存する.
; $name で指定したフォルダが既に存在している場合は, 何もしない.
;
; @param $name 名称.
; @param $bits 鍵長.
; @param $md 署名のメッセージダイジェスト.
; @param $days 証明書の有効期限.
; @param $ca 認証局フォルダへのフルパス.
; @param $hook 生成時の設定ファイルを書き換えるためのフック関数名.
;
Func OpenSSL_CreateServerCertificate($name, $bits, $md, $days, $ca, $hook = 0)
	If Not FileExists($name) Then
		Local $path = FileUtiilty_ChangeDir($name)

		OpenSSL_CreateRsaPrivateKey($bits, $OpenSSL_KeyName)
		OpenSSL_CreateConfigurationTemplate($OpenSSL_ConfigName, $name)
		IniWrite($OpenSSL_ConfigName, "CA_default", "dir", $ca)
		IniWrite($OpenSSL_ConfigName, "usr_cert", "basicConstraints", "CA:FALSE")
		If IsString($hook) Then
			Call($hook, $OpenSSL_ConfigName)
		EndIf
		OpenSSL_CreateCertificationSigningRequest($OpenSSL_KeyName, $OpenSSL_ConfigName, $OpenSSL_CertificationSigningRequest)
		OpenSSL_Signed($OpenSSL_CertificationSigningRequest, $days, $md, $OpenSSL_ConfigName, $OpenSSL_CertificatePemName)
		OpenSSL_PemToDer($OpenSSL_CertificatePemName, $OpenSSL_CertificateDerName)

		FileChangeDir($path)
	EndIf
EndFunc   ;==>OpenSSL_CreateServerCertificate

;
; PEM形式の証明書をDER形式に変換する.
;
; @param $pem PEM形式証明書(入力).
; @param $der DER形式証明書(出力).
;
Func OpenSSL_PemToDer($pem, $der)
	Local $cmd = StringFormat("%s x509 -inform PEM -outform DER -in %s -out %s", $OpenSSL_CmdPath, $pem, $der)
	ConsoleUtility_DebugLogLn($OpenSSL_DebugLog, $cmd)
	RunWait($cmd)
EndFunc   ;==>OpenSSL_PemToDer

;
; DER形式の証明書をPEM形式に変換する.
;
; @param $der DER形式証明書(入力).
; @param $pem PEM形式証明書(出力).
;
Func OpenSSL_DerToPem($der, $pem)
	Local $cmd = StringFormat("%s x509 -inform DER -outform PEM -in %s -out %s", $OpenSSL_CmdPath, $der, $pem)
	ConsoleUtility_DebugLogLn($OpenSSL_DebugLog, $cmd)
	RunWait($cmd)
EndFunc   ;==>OpenSSL_DerToPem
#endregion パブリックメソッド

#region プライベートメソッド
;
; RSA秘密鍵を生成する.
;
; @param $bits 暗号鍵のビット数.
; @param $output_file 出力ファイル名(秘密鍵ファイル).
;
Func OpenSSL_CreateRsaPrivateKey($bits, $output_file)
	Local $cmd = StringFormat("%s genrsa -out %s %s", $OpenSSL_CmdPath, $output_file, $bits)
	ConsoleUtility_DebugLogLn($OpenSSL_DebugLog, $cmd)
	RunWait($cmd)
EndFunc   ;==>OpenSSL_CreateRsaPrivateKey

;
; 自己署名X.509証明書を生成する.
;
; @param $key_file 秘密鍵.
; @param $days 証明書の有効期限.
; @param $md メッセージダイジェスト.
; @param $config_file 設定ファイル名.
; @param $output_file 出力ファイル名(自己署名証明書).
;
Func OpenSSL_SelfSigned($key_file, $days, $md, $config_file, $output_file)
	Local $cmd = StringFormat("%s req -new -x509 -key %s -%s -out %s -config %s -days %s", $OpenSSL_CmdPath, $key_file, $md, $output_file, $config_file, $days)
	ConsoleUtility_DebugLogLn($OpenSSL_DebugLog, $cmd)
	RunWait($cmd)
EndFunc   ;==>OpenSSL_SelfSigned

;
; X.509証明書を生成する.
;
; @param $csr_file 署名要求(CSR)ファイル名.
; @param $days 証明書の有効期限.
; @param $md メッセージダイジェスエト.
; @param $config_file 設定ファイル名.
; @param $output_file 出力ファイル名(X.509証明書).
;
Func OpenSSL_Signed($csr_file, $days, $md, $config_file, $output_file)
	Local $cmd = StringFormat("%s ca -md %s -in %s -out %s -config %s -days %s", $OpenSSL_CmdPath, $md, $csr_file, $output_file, $config_file, $days)
	ConsoleUtility_DebugLogLn($OpenSSL_DebugLog, $cmd)
	Run($cmd)
	WinWaitActive($OpenSSL_ConsoleClass)
	ControlSend($OpenSSL_ConsoleClass, "", "", "y{ENTER}")
	ControlSend($OpenSSL_ConsoleClass, "", "", "y{ENTER}")
	WinWaitClose($OpenSSL_ConsoleClass)
EndFunc   ;==>OpenSSL_Signed

;
; 署名要求(CSR)を生成する.
;
; @param $key_file 秘密鍵.
; @param $config_file 設定ファイル名.
; @param $output_file 出力ファイル名(署名要求).
;
Func OpenSSL_CreateCertificationSigningRequest($key_file, $config_file, $output_file)
	Local $cmd = StringFormat("%s req -new -key %s -out %s -config %s", $OpenSSL_CmdPath, $key_file, $output_file, $config_file)
	ConsoleUtility_DebugLogLn($OpenSSL_DebugLog, $cmd)
	RunWait($cmd)
EndFunc   ;==>OpenSSL_CreateCertificationSigningRequest

;
; 設定ファイルのテンプレートを生成する.
;
; @param $file 設定ファイル.
; @param $ca_dir 認証局情報が格納されたフォルダ.
;
Func OpenSSL_CreateConfigurationTemplate($file, $ca_dir)
	Local $ca[1][2] = [ _
			["default_ca", "CA_default"] _
			]
	IniWriteSection($file, "ca", $ca, 0)

	Local $ca_default[9][2] = [ _
			["dir", $ca_dir], _
			["certificate", "$dir/" & $OpenSSL_CertificatePemName], _
			["private_key", "$dir/" & $OpenSSL_KeyName], _
			["database", "$dir/" & $OpenSSL_DataBaseName], _
			["serial", "$dir/" & $OpenSSL_SerialName], _
			["policy", "policy_anything"], _
			["x509_extensions", "usr_cert"], _
			["new_certs_dir", "$dir/"] _
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
EndFunc   ;==>OpenSSL_CreateConfigurationTemplate

;
; データベースを生成する.
;
; @param $file 生成するファイル名.
;
Func OpenSSL_CreateDataBase($file)
	FileOpen($file, 2 + 8)
EndFunc   ;==>OpenSSL_CreateDataBase

;
; シリアルファイルを生成する.
;
; @param $crt_file 証明書.
; @param $serial_file 生成するシリアルファイル名.
;
Func OpenSSL_CreateSerialFile($crt_file, $serial_file)
	Local $cmd = StringFormat("%s x509 -in %s -noout -next_serial -out %s", $OpenSSL_CmdPath, $crt_file, $serial_file)
	ConsoleUtility_DebugLogLn($OpenSSL_DebugLog, $cmd)
	RunWait($cmd)
EndFunc   ;==>OpenSSL_CreateSerialFile
#endregion プライベートメソッド