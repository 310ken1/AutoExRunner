#include-once
#include "..\utility\ConsoleUtility.au3"
#include "..\utility\DateUtility.au3"

#region Globale_Argument_Define
; !! 必須 !!
; 本ソースコード上で定義されたメソッドを利用する場合は,
; $OpenSSL_CmdPath グローバル変数 に openssl.exe へのパスを設定すること.
Global $OpenSSL_CmdPath = "openssl.exe"
; デバッグログフラグ.
; 1 を指定することで, デバッグログが出力される.
Global $OpenSSL_DebugLog = 0
; 鍵の種類ID.
Global Enum $OpenSSL_KeyID_RSA, $OpenSSL_KeyID_DSA
#endregion Globale_Argument_Define

#region Constant_Define
; DSAパラメータファイル名.
Const $OpenSSL_DSAParam = "dasparam.txt"
; 秘密鍵のファイル名
Const $OpenSSL_KeyName = "private.key"
; openssl の設定ファイル名
Const $OpenSSL_ConfigName = "openssl.cnf"
; PEM形式の証明書ファイル名
Const $OpenSSL_CertificatePemName = "certificate.crt"
; DER形式の証明書ファイル名
Const $OpenSSL_CertificateDerName = "certificate.der"
; 証明書署名リクエストのファイル名
Const $OpenSSL_CertificationSigningRequest = "CertificationSigningRequest.csr"
; 発行した証明書情報のデータベースファイル名
Const $OpenSSL_DataBaseName = "database.txt"
; 次に発行する証明書のシリアル番号
Const $OpenSSL_SerialName = "serial.txt"
#endregion Constant_Define

#region Public_Method
;
; ルート証明書を生成する.
; $dir で指定したフォルダを生成し, そのフォルダに各種ファイルを保存する.
; $dir で指定したフォルダが既に存在している場合は, 何もしない.
;
; @param $dir 出力先フォルダ.
; @param $keyfile 鍵ファイル.
; @param $configfile 設定ファイル.
; @param $md 署名のメッセージダイジェスト.
; @param $days 証明書の有効期限.
; @param $start_date 証明書の有効期限の開始時刻文字列(2000年1月1日は2000/01/01と指定).
;
Func OpenSSL_CreateRootCertificate($dir, $keyfile, $configfile, $md, $days, $start_date = 0)
	DirCreate($dir)

	; 時刻の保存.
	Local $handle = 0
	If IsString($start_date) Then
		$handle = DateUtility_Save()
		If False = DateUtility_SetLocalTimeString($start_date) Then
			$handle = 0
		EndIf
	EndIf

	; 自己署名.
	Local $cerfile = FileUtility_MakePath($dir, $OpenSSL_CertificatePemName)
	OpenSSL_SelfSigned($keyfile, $days, $md, $configfile, $cerfile)

	; 時刻の復元.
	If 0 <> $handle Then
		DateUtility_Restore($handle)
	EndIf

	; DER形式証明書の生成.
	Local $derfile = FileUtility_MakePath($dir, $OpenSSL_CertificateDerName)
	OpenSSL_PemToDer($cerfile, $derfile)

	; データベースファイルの生成.
	Local $database = FileUtility_MakePath($dir, $OpenSSL_DataBaseName)
	OpenSSL_CreateDataBase($database)

	; シリアルファイルの生成.
	Local $serial = FileUtility_MakePath($dir, $OpenSSL_SerialName)
	OpenSSL_CreateSerialFile($cerfile, $serial)
EndFunc   ;==>OpenSSL_CreateRootCertificate
;
; 中間証明書を生成する.
; $dir で指定したフォルダを生成し, そのフォルダに各種ファイルを保存する.
; $dir で指定したフォルダが既に存在している場合は, 何もしない.
;
; @param $dir 出力先フォルダ.
; @param $keyfile 鍵ファイル.
; @param $configfile 設定ファイル.
; @param $md 署名のメッセージダイジェスト.
; @param $days 証明書の有効期限.
; @param $start_date 証明書の有効期限の開始時刻文字列(2000年1月1日は2000/01/01と指定).
;
Func OpenSSL_CreateIntermediateCertificate($dir, $keyfile, $configfile, $md, $days, $start_date = 0)
	DirCreate($dir)

	; 署名要求(CSR)の生成.
	Local $csr = FileUtility_MakePath($dir, $OpenSSL_CertificationSigningRequest)
	OpenSSL_CreateCertificationSigningRequest($keyfile, $configfile, $csr)

	; 署名
	Local $cerfile = FileUtility_MakePath($dir, $OpenSSL_CertificatePemName)
	Local $date = OpenSSL_ToDate($start_date)
	If "" <> $date Then
		OpenSSL_Signed($csr, $days, $md, $configfile, $cerfile, $date)
	Else
		OpenSSL_Signed($csr, $days, $md, $configfile, $cerfile)
	EndIf

	; DER形式証明書の生成.
	Local $derfile = FileUtility_MakePath($dir, $OpenSSL_CertificateDerName)
	OpenSSL_PemToDer($cerfile, $derfile)

	; データベースファイルの生成.
	Local $database = FileUtility_MakePath($dir, $OpenSSL_DataBaseName)
	OpenSSL_CreateDataBase($database)

	; シリアルファイルの生成.
	Local $serial = FileUtility_MakePath($dir, $OpenSSL_SerialName)
	OpenSSL_CreateSerialFile($cerfile, $serial)
EndFunc   ;==>OpenSSL_CreateIntermediateCertificate
;
; サーバ証明書を生成する.
; $name で指定したフォルダを生成し, そのフォルダに各種ファイルを保存する.
; $name で指定したフォルダが既に存在している場合は, 何もしない.
;
; @param $dir 出力先フォルダ.
; @param $keyfile 鍵ファイル.
; @param $configfile 設定ファイル.
; @param $md 署名のメッセージダイジェスト.
; @param $days 証明書の有効期限.
; @param $start_date 証明書の有効期限の開始時刻文字列(2000年1月1日は2000/01/01と指定).
;
Func OpenSSL_CreateServerCertificate($dir, $keyfile, $configfile, $md, $days, $start_date = 0)
	DirCreate($dir)

	; 署名要求(CSR)の生成.
	Local $csr = FileUtility_MakePath($dir, $OpenSSL_CertificationSigningRequest)
	OpenSSL_CreateCertificationSigningRequest($keyfile, $configfile, $csr)

	; 署名
	Local $cerfile = FileUtility_MakePath($dir, $OpenSSL_CertificatePemName)
	Local $date = OpenSSL_ToDate($start_date)
	If "" <> $date Then
		OpenSSL_Signed($csr, $days, $md, $configfile, $cerfile, $date)
	Else
		OpenSSL_Signed($csr, $days, $md, $configfile, $cerfile)
	EndIf

	; DER形式証明書の生成.
	Local $derfile = FileUtility_MakePath($dir, $OpenSSL_CertificateDerName)
	OpenSSL_PemToDer($cerfile, $derfile)
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
;
; 秘密鍵を生成する.
;
; @param $dir 出力先フォルダ.
; @param $keyID 鍵の種類ID($OpenSSL_KeyID_RSA もしくは $OpenSSL_KeyID_DSA を指定する).
; @param $bits 暗号鍵のビット数.
; @return 秘密鍵ファイル.
;
Func OpenSSL_CreatePrivateKey($dir, $keyID, $bits)
	Local $keyfile = FileUtility_MakePath($dir, $OpenSSL_KeyName)
	Switch $keyID
		Case $OpenSSL_KeyID_RSA
			OpenSSL_CreateRsaPrivateKey($bits, $keyfile)
		Case $OpenSSL_KeyID_DSA
			OpenSSL_CreateDsaPrivateKey($bits, $keyfile)
	EndSwitch
	Return $keyfile
EndFunc   ;==>OpenSSL_CreatePrivateKey
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
; DSA秘密鍵を生成する.
;
; @param $bits 暗号鍵のビット数.
; @param $output_file 出力ファイル名(秘密鍵ファイル).
;
Func OpenSSL_CreateDsaPrivateKey($bits, $output_file)
	Local $cmd = StringFormat("%s dsaparam -genkey -out %s %s", $OpenSSL_CmdPath, $output_file, $bits)
	ConsoleUtility_DebugLogLn($OpenSSL_DebugLog, $cmd)
	RunWait($cmd)
EndFunc   ;==>OpenSSL_CreateDsaPrivateKey
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
	Local $cmd = StringFormat("%s req -new -x509 -key %s -%s -out %s -config %s -days %s -batch", _
			$OpenSSL_CmdPath, $key_file, $md, $output_file, $config_file, $days)
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
; @param $start_date 証明書の有効期限の開始時刻文字列(YYMMDDHHMMSS形式).指定が無い場合は現在時刻.
;
Func OpenSSL_Signed($csr_file, $days, $md, $config_file, $output_file, $start_date = 0)
	Local $cmd = ""
	If IsString($start_date) Then
		$cmd = StringFormat("%s ca -md %s -in %s -out %s -config %s -startdate %s -days %s -notext -batch", _
				$OpenSSL_CmdPath, $md, $csr_file, $output_file, $config_file, $start_date, $days)
	Else
		$cmd = StringFormat("%s ca -md %s -in %s -out %s -config %s -days %s -notext -batch", _
				$OpenSSL_CmdPath, $md, $csr_file, $output_file, $config_file, $days)
	EndIf
	ConsoleUtility_DebugLogLn($OpenSSL_DebugLog, $cmd)
	RunWait($cmd)
EndFunc   ;==>OpenSSL_Signed
;
; 署名要求(CSR)を生成する.
;
; @param $key_file 秘密鍵.
; @param $config_file 設定ファイル名.
; @param $output_file 出力ファイル名(署名要求).
;
Func OpenSSL_CreateCertificationSigningRequest($key_file, $config_file, $output_file)
	Local $cmd = StringFormat("%s req -new -key %s -out %s -config %s", _
			$OpenSSL_CmdPath, $key_file, $output_file, $config_file)
	ConsoleUtility_DebugLogLn($OpenSSL_DebugLog, $cmd)
	RunWait($cmd)
EndFunc   ;==>OpenSSL_CreateCertificationSigningRequest
;
; 証明書のための設定ファイルのテンプレートを生成する.
;
; @param $dir 出力先フォルダ.
;
Func OpenSSL_CreateSignedConfig($dir)
	Local $config = FileUtility_MakePath($dir, $OpenSSL_ConfigName)
	Local $ca[1][2] = [ _
			["default_ca", "CA_default"] _
			]
	IniWriteSection($config, "ca", $ca, 0)

	Local $ca_default[7][2] = [ _
			["dir", $dir], _
			["certificate", "$dir/" & $OpenSSL_CertificatePemName], _
			["private_key", "$dir/" & $OpenSSL_KeyName], _
			["database", "$dir/" & $OpenSSL_DataBaseName], _
			["serial", "$dir/" & $OpenSSL_SerialName], _
			["new_certs_dir", "$dir/"], _
			["policy", "policy_anything"] _
			]
	IniWriteSection($config, "CA_default", $ca_default, 0)

	Local $policy_anything[7][2] = [ _
			["countryName", "optional"], _
			["stateOrProvinceName", "optional"], _
			["localityName", "optional"], _
			["organizationName", "optional"], _
			["organizationalUnitName", "optional"], _
			["commonName", "supplied"], _
			["emailAddress", "optional"] _
			]
	IniWriteSection($config, "policy_anything", $policy_anything, 0)

	Local $req[2][2] = [ _
			["distinguished_name", "req_distinguished_name"], _
			["prompt", "no"] _
			]
	IniWriteSection($config, "req", $req, 0)

	Local $reqreq_distinguished_name[1][2] = [ _
			["CN", "example.com"] _
			]
	IniWriteSection($config, "req_distinguished_name", $reqreq_distinguished_name, 0)

	Return $config
EndFunc   ;==>OpenSSL_CreateSignedConfig
;
; 設定ファイルに認証局(フォルダ)を設定する.
;
; @param $file 設定ファイル.
; @param $cadir 認証局(フォルダ)
;
Func OpenSSL_ConfigAddCertificateAuthority($file, $cadir)
	IniWrite($file, "CA_default", "dir", $cadir)
EndFunc   ;==>OpenSSL_ConfigAddCertificateAuthority
;
; 設定ファイルにサブジェクト(所有者情報)を設定する.
;
; @param $file 設定ファイル.
; @param $cn 一般名.
; @param $c 国名.
; @param $st 県名/州名.
; @param $l 都市名.
; @param $o0 団体名.
; @param $o1 団体名の拡張.
; @param $ou0 組織名.
; @param $ou1 組織名の拡張.
; @param $email メールアドレス.
;
Func OpenSSL_ConfigAddSubject($file, $cn, _
		$c = "", $st = "", $l = "", $o0 = "", $o1 = "", $ou0 = "", $ou1 = "", $email = "")
	Local $subject[9][2] = [ _
			["countryName", $c], _
			["stateOrProvinceName", $st], _
			["localityName", $l], _
			["0.organizationName", $o0], _
			["1.organizationName", $o1], _
			["0.organizationalUnitName", $ou0], _
			["1.organizationalUnitName", $ou1], _
			["commonName", $cn], _
			["emailAddress", $email] _
			]
	IniWriteSection($file, "req_distinguished_name", OpenSSL_ArrayVoidDelete($subject), 0)
EndFunc   ;==>OpenSSL_ConfigAddSubject
;
; 設定ファイルに X.509証明書 v3拡張 を設定する.
;
; @param $file 設定ファイル.
; @param $bc 基本的制約.
; @param $keyusage 鍵の用途.
; @param $exkeyusage 鍵用途の拡張.
; @parma $subject サブジェクト鍵の識別子.
; @param $auth 認証局鍵の識別子.
;
Func OpenSSL_ConfigAddExtensions($file, $bc, _
		$keyusage = "", $exkeyusage = "", $subject = "", $auth = "")
	Local $extensions[5][2] = [ _
			["basicConstraints", $bc], _
			["keyUsage", $keyusage], _
			["extendedKeyUsage", $exkeyusage], _
			["subjectKeyIdentifier", $subject], _
			["authorityKeyIdentifier", $auth] _
			]
	IniWriteSection($file, "usr_cert", OpenSSL_ArrayVoidDelete($extensions), 0)

	IniWrite($file, "req", "x509_extensions", "usr_cert")
EndFunc   ;==>OpenSSL_ConfigAddExtensions
#endregion Public_Method

#region Private_Method
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
	Local $cmd = StringFormat("%s x509 -in %s -noout -next_serial -out %s", _
			$OpenSSL_CmdPath, $crt_file, $serial_file)
	ConsoleUtility_DebugLogLn($OpenSSL_DebugLog, $cmd)
	RunWait($cmd)
EndFunc   ;==>OpenSSL_CreateSerialFile
;
; 設定配列の空設定を削除する.
;
; @param $array 設定配列.
; @return 空設定を削除した設定配列.
;
Func OpenSSL_ArrayVoidDelete(ByRef $array)
	Local $count = UBound($array, 1)
	Local $ret = 0
	If 0 < $count Then
		Local $tmp[$count][2]
		Local $index = 0
		For $i = 0 To $count - 1
			If "" <> $array[$i][1] Then
				$tmp[$index][0] = $array[$i][0]
				$tmp[$index][1] = $array[$i][1]
				$index += 1
			EndIf
		Next
		ReDim $tmp[$index][2]
		$ret = $tmp
	Else
		SetError(1)
	EndIf
	Return $ret
EndFunc   ;==>OpenSSL_ArrayVoidDelete
;
; YYYY/MM/DD HH:MM:SS 形式 を YYMMDDHHMMSSZ形式に変換する.
;
; @param $start_date YYYY/MM/DD HH:MM:SS形式の日時.
; @return YYMMDDHHMMSSZ形式の日時.
;
Func OpenSSL_ToDate($start_date)
	Local $ret = ""
	Local $date = DateUtility_DateStringToArray($start_date)
	If Not @error Then
		$ret = StringFormat("%04d%02d%02d%02d%02d%02dZ", _
				$date[$DateUtilityYear], _
				$date[$DateUtilityMonth], _
				$date[$DateUtilityDay], _
				$date[$DateUtilityHour], _
				$date[$DateUtilityMinute], _
				$date[$DateUtilitySecond])
	EndIf
	Return $ret
EndFunc   ;==>OpenSSL_ToDate
#endregion Private_Method