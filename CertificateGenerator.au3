#include "runner\AutoOpenOfficeRunner.au3"
#include "utility\FileUtility.au3"
#include "app\OpenSSL.au3"

#region Globale_Argument_Define
; OpenSSL設定
$OpenSSL_CmdPath = FileUtility_ScriptDirFilePath("bin\openssl.exe")
$OpenSSL_DebugLog = 1
; AutoOpenOfficeRunner設定
$AutoOpenOfficeRunner_StartLine = 2 ; AutoOpenOfficeRunner の開始行.
$AutoOpenOfficeRunner_NoStartColumn = 0 ; No項目 の列.
$AutoOpenOfficeRunner_KeyStartLine = 1 ; 項目名 の開始行.
$AutoOpenOfficeRunner_KeyStartColumn = 0 ; 項目名 の開始列.
$AutoOpenOfficeRunner_KeyColumnMaxCount = 25 ; 項目の最大個数(No項目を除く).
$AutoOpenOfficeRunner_ExecutableUnit = 25 ; OpenOfficeCalcから一度に読み出す行数.
; cat.exe コマンドへのパス.
$Cat_CmdPath = FileUtility_ScriptDirFilePath("bin\cat.exe")
#endregion Globale_Argument_Define

#region Constant_Define
; ワークスペース
Const $WorkSpace = FileUtility_ScriptDirFilePath("00_workspace")
; 出力フォルダ.
Const $OutputDir = FileUtility_ScriptDirFilePath("01_output")
; 入力ファイル.
Const $InputFile = FileUtility_ScriptDirFilePath("CertificateList.ods")
; 鍵の種類字列と鍵IDの変換テーブル.
Const $KeyKindTable[2][2] = [ _
		["RSA", $OpenSSL_KeyID_RSA], _
		["DSA", $OpenSSL_KeyID_DSA] _
		]
#endregion Constant_Define

;
; メイン関数呼び出し.
;
Main()

;
; メイン関数.
;
Func Main()
	DirCreate($WorkSpace)
	DirCreate($OutputDir)

	Local $handle = AutoOpenOfficeRunner_Open($InputFile)
	If IsArray($handle) Then
		AutoOpenOfficeRunner_Run($handle, "ルート証明書", "CreateRootCrt")
		AutoOpenOfficeRunner_Run($handle, "中間証明書", "CreateIntermediateCrt")
		AutoOpenOfficeRunner_Run($handle, "サーバ証明書", "CreateServerCrt")
		AutoOpenOfficeRunner_Run($handle, "証明書チェイン", "CreateChain")
		AutoOpenOfficeRunner_Close($handle)
	EndIf
EndFunc   ;==>Main

;
; ルート証明書の生成処理.
;
; @param $handle ハンドル.
;
Func CreateRootCrt(Const $handle)
	Local $name = AutoOpenOfficeRunner_GetString($handle, "名称")
	Local $dir = FileUtility_MakePath($WorkSpace, $name)

	TrayTip("ルート証明書", $name, 10)

	If Not FileExists($dir) Then
		DirCreate($dir)

		Local $keyfile = OpenSSL_CreatePrivateKey($dir, _
				KeyStringToID(AutoOpenOfficeRunner_GetString($handle, "鍵種別")), _
				AutoOpenOfficeRunner_GetString($handle, "鍵長"))

		Local $config = OpenSSL_CreateSignedConfig($dir)
		OpenSSL_ConfigAddCertificateAuthority($config, $dir)
		OpenSSL_ConfigAddSubject($config, _
				AutoOpenOfficeRunner_GetString($handle, "commonName"), _
				AutoOpenOfficeRunner_GetString($handle, "countryName"), _
				AutoOpenOfficeRunner_GetString($handle, "stateOrProvinceName"), _
				AutoOpenOfficeRunner_GetString($handle, "localityName"), _
				AutoOpenOfficeRunner_GetString($handle, "0.organizationName"), _
				AutoOpenOfficeRunner_GetString($handle, "1.organizationName"), _
				AutoOpenOfficeRunner_GetString($handle, "0.organizationalUnitName"), _
				AutoOpenOfficeRunner_GetString($handle, "1.organizationalUnitName"), _
				AutoOpenOfficeRunner_GetString($handle, "emailAddress"))
		If "有効" = AutoOpenOfficeRunner_GetString($handle, "v3拡張有無") Then
			OpenSSL_ConfigAddExtensions($config, _
					AutoOpenOfficeRunner_GetString($handle, "basicConstraints"), _
					AutoOpenOfficeRunner_GetString($handle, "keyUsage"), _
					AutoOpenOfficeRunner_GetString($handle, "extendedKeyUsage"), _
					AutoOpenOfficeRunner_GetString($handle, "subjectKeyIdentifier"), _
					AutoOpenOfficeRunner_GetString($handle, "authorityKeyIdentifier"))
		EndIf
		OpenSSL_CreateRootCertificate($dir, $keyfile, $config, _
				AutoOpenOfficeRunner_GetString($handle, "メッセージダイジェスト"), _
				AutoOpenOfficeRunner_GetString($handle, "日数"), _
				AutoOpenOfficeRunner_GetString($handle, "開始日"))
	EndIf
EndFunc   ;==>CreateRootCrt

;
; 中間証明書の生成処理.
;
; @param $handle ハンドル.
;
Func CreateIntermediateCrt(Const $handle)
	Local $name = AutoOpenOfficeRunner_GetString($handle, "名称")
	Local $dir = FileUtility_MakePath($WorkSpace, $name)

	TrayTip("中間証明書", $name, 10)

	If Not FileExists($dir) Then
		DirCreate($dir)
		Local $keyfile = OpenSSL_CreatePrivateKey($dir, _
				KeyStringToID(AutoOpenOfficeRunner_GetString($handle, "鍵種別")), _
				AutoOpenOfficeRunner_GetString($handle, "鍵長"))

		Local $config = OpenSSL_CreateSignedConfig($dir)
		Local $ca = StringRegExpReplace( _
				FileUtility_MakePath($WorkSpace, AutoOpenOfficeRunner_GetString($handle, "認証局")), '\\', '/')
		OpenSSL_ConfigAddCertificateAuthority($config, $ca)
		OpenSSL_ConfigAddSubject($config, _
				AutoOpenOfficeRunner_GetString($handle, "commonName"), _
				AutoOpenOfficeRunner_GetString($handle, "countryName"), _
				AutoOpenOfficeRunner_GetString($handle, "stateOrProvinceName"), _
				AutoOpenOfficeRunner_GetString($handle, "localityName"), _
				AutoOpenOfficeRunner_GetString($handle, "0.organizationName"), _
				AutoOpenOfficeRunner_GetString($handle, "1.organizationName"), _
				AutoOpenOfficeRunner_GetString($handle, "0.organizationalUnitName"), _
				AutoOpenOfficeRunner_GetString($handle, "1.organizationalUnitName"), _
				AutoOpenOfficeRunner_GetString($handle, "emailAddress"))
		If "有効" = AutoOpenOfficeRunner_GetString($handle, "v3拡張有無") Then
			OpenSSL_ConfigAddExtensions($config, _
					AutoOpenOfficeRunner_GetString($handle, "basicConstraints"), _
					AutoOpenOfficeRunner_GetString($handle, "keyUsage"), _
					AutoOpenOfficeRunner_GetString($handle, "extendedKeyUsage"), _
					AutoOpenOfficeRunner_GetString($handle, "subjectKeyIdentifier"), _
					AutoOpenOfficeRunner_GetString($handle, "authorityKeyIdentifier"))
		EndIf
		OpenSSL_CreateIntermediateCertificate($dir, $keyfile, $config, _
				AutoOpenOfficeRunner_GetString($handle, "メッセージダイジェスト"), _
				AutoOpenOfficeRunner_GetString($handle, "日数"), _
				AutoOpenOfficeRunner_GetString($handle, "開始日"))

	EndIf
EndFunc   ;==>CreateIntermediateCrt

;
; サーバ証明書の生成処理.
;
; @param $handle ハンドル.
;
Func CreateServerCrt(Const $handle)
	Local $name = AutoOpenOfficeRunner_GetString($handle, "名称")
	Local $dir = FileUtility_MakePath($WorkSpace, $name)

	TrayTip("サーバ証明書", $name, 10)

	If Not FileExists($dir) Then
		DirCreate($dir)
		Local $keyfile = OpenSSL_CreatePrivateKey($dir, _
				KeyStringToID(AutoOpenOfficeRunner_GetString($handle, "鍵種別")), _
				AutoOpenOfficeRunner_GetString($handle, "鍵長"))

		Local $config = OpenSSL_CreateSignedConfig($dir)
		Local $ca = StringRegExpReplace( _
				FileUtility_MakePath($WorkSpace, AutoOpenOfficeRunner_GetString($handle, "認証局")), '\\', '/')
		OpenSSL_ConfigAddCertificateAuthority($config, $ca)
		OpenSSL_ConfigAddSubject($config, _
				AutoOpenOfficeRunner_GetString($handle, "commonName"), _
				AutoOpenOfficeRunner_GetString($handle, "countryName"), _
				AutoOpenOfficeRunner_GetString($handle, "stateOrProvinceName"), _
				AutoOpenOfficeRunner_GetString($handle, "localityName"), _
				AutoOpenOfficeRunner_GetString($handle, "0.organizationName"), _
				AutoOpenOfficeRunner_GetString($handle, "1.organizationName"), _
				AutoOpenOfficeRunner_GetString($handle, "0.organizationalUnitName"), _
				AutoOpenOfficeRunner_GetString($handle, "1.organizationalUnitName"), _
				AutoOpenOfficeRunner_GetString($handle, "emailAddress"))
		If "有効" = AutoOpenOfficeRunner_GetString($handle, "v3拡張有無") Then
			OpenSSL_ConfigAddExtensions($config, _
					AutoOpenOfficeRunner_GetString($handle, "basicConstraints"), _
					AutoOpenOfficeRunner_GetString($handle, "keyUsage"), _
					AutoOpenOfficeRunner_GetString($handle, "extendedKeyUsage"), _
					AutoOpenOfficeRunner_GetString($handle, "subjectKeyIdentifier"), _
					AutoOpenOfficeRunner_GetString($handle, "authorityKeyIdentifier"))
		EndIf
		OpenSSL_CreateServerCertificate($dir, $keyfile, $config, _
				AutoOpenOfficeRunner_GetString($handle, "メッセージダイジェスト"), _
				AutoOpenOfficeRunner_GetString($handle, "日数"), _
				AutoOpenOfficeRunner_GetString($handle, "開始日"))

	EndIf
EndFunc   ;==>CreateServerCrt

;
; 証明書チェインの生成処理.
;
; @param $handle ハンドル.
;
Func CreateChain(Const $handle)
	Local $name = AutoOpenOfficeRunner_GetString($handle, "名称")

	Local $chainfile = FileUtility_MakePath($OutputDir, $name & ".cer")
	Local $cmd = "cmd.exe /k """ & $Cat_CmdPath & """"
	$cmd &= " " & GetFilePath($handle, "サーバ証明書", $OpenSSL_CertificatePemName)
	$cmd &= " " & GetFilePath($handle, "中間証明書1", $OpenSSL_CertificatePemName)
	$cmd &= " " & GetFilePath($handle, "中間証明書2", $OpenSSL_CertificatePemName)
	$cmd &= " " & GetFilePath($handle, "中間証明書3", $OpenSSL_CertificatePemName)
	$cmd &= " " & GetFilePath($handle, "ルート証明書", $OpenSSL_CertificatePemName)
	$cmd &= " > " & $chainfile & " & exit"
	$cmd = StringRegExpReplace($cmd, " +", " ")
	RunWait($cmd)

	Local $keyfile = FileUtility_MakePath($OutputDir, $name & ".key")
	Local $keysrc = GetFilePath($handle, "サーバ証明書", $OpenSSL_KeyName)
	FileCopy($keysrc, $keyfile)
EndFunc   ;==>CreateChain

;
; 各証明書に関するファイルのパスを取得する.
;
; @param $handle ハンドル.
; @param $key 証明書名.
; @param $file 取得したいファイル名.
; @return  各証明書に関するファイルのパス.
;
Func GetFilePath($handle, $key, $file)
	Local $result = ""
	Local $name = AutoOpenOfficeRunner_GetString($handle, $key)
	If "" <> $name Then
		$result = $WorkSpace & "\" & $name & "\" & $file
	EndIf
	Return $result
EndFunc   ;==>GetFilePath

#region Configuration
;
; 鍵の種類文字列から鍵の種類($OpenSSL_KeyID_RSA/$OpenSSL_KeyID_DSA)を取得する.
;
; @param $keyString 鍵の種類文字列.
; @return 鍵の種類($OpenSSL_RSA/$OpenSSL_DSA).
;
Func KeyStringToID($keyString)
	Local $index = _ArraySearch($KeyKindTable, $keyString, 0, 0, 0, 0, 1, 0)
	Return $KeyKindTable[$index][1]
EndFunc   ;==>KeyStringToID
#endregion Configuration