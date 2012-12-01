#include "runner\AutoOpenOfficeRunner.au3"
#include "utility\FileUtility.au3"
#include "app\OpenSSL.au3"

#region Globale_Argument_Define
$OpenSSL_CmdPath = FileUtility_ScriptDirFilePath("bin\openssl.exe")
$OpenSSL_DebugLog = 1
$AutoOpenOfficeRunner_KeyColumnMaxCount = 15
#endregion Globale_Argument_Define

#region Constant_Define
; 出力フォルダ.
Const $OutputDir = FileUtility_ScriptDirFilePath("out")
; 入力ファイル.
Const $InputFile = FileUtility_ScriptDirFilePath("CertificateList.ods")
#endregion Constant_Define

#region Static_Argument_Define
; サブジェクトの設定書き換え用配列.
Static $Subject[7][2] = [ _
		["C", ""], _
		["ST", ""], _
		["L", ""], _
		["O", ""], _
		["OU", ""], _
		["CN", ""], _
		["emailAddress", ""] _
		]
#endregion Static_Argument_Define

;
; メイン関数呼び出し.
;
Main()

;
; メイン関数.
;
Func Main()
	DirCreate($OutputDir)
	FileChangeDir($OutputDir)

	Local $handle = AutoOpenOfficeRunner_Open($InputFile)
	If IsArray($handle) Then
		AutoOpenOfficeRunner_Run($handle, "ルート証明書", "CreateRootCrt")
		AutoOpenOfficeRunner_Run($handle, "中間証明書", "CreateIntermediateCrt")
		AutoOpenOfficeRunner_Run($handle, "サーバ証明書", "CreateServerCrt")

		AutoOpenOfficeRunner_Close($handle)
	EndIf
EndFunc   ;==>Main

#region Root_Certificate
;
; ルート証明書の生成処理.
;
; @param $handle ハンドル.
;
Func CreateRootCrt(Const $handle)
	InitConfigurationArray($Subject)
	SetConfigurationArray($handle, $Subject)

	OpenSSL_CreateRootCertificate( _
			AutoOpenOfficeRunner_GetString($handle, "名称"), _
			AutoOpenOfficeRunner_GetString($handle, "鍵長"), _
			AutoOpenOfficeRunner_GetString($handle, "メッセージダイジェスト"), _
			AutoOpenOfficeRunner_GetString($handle, "有効期限"), _
			"HookCreateRootCrt" _
			)
EndFunc   ;==>CreateRootCrt

;
; ルート証明書の設定ファイル書き換えフック関数.
;
; @param $config 設定ファイル名.
;
Func HookCreateRootCrt($config)
	WriteConfigurationArray($Subject, "req_distinguished_name", $config)
EndFunc   ;==>HookCreateRootCrt
#endregion Root_Certificate

#region Intermediate_Certificate
;
; 中間証明書の生成処理.
;
; @param $handle ハンドル.
;
Func CreateIntermediateCrt(Const $handle)
	InitConfigurationArray($Subject)
	SetConfigurationArray($handle, $Subject)

	OpenSSL_CreateIntermediateCertificate( _
			AutoOpenOfficeRunner_GetString($handle, "名称"), _
			AutoOpenOfficeRunner_GetString($handle, "鍵長"), _
			AutoOpenOfficeRunner_GetString($handle, "メッセージダイジェスト"), _
			AutoOpenOfficeRunner_GetString($handle, "有効期限"), _
			StringRegExpReplace($OutputDir & "\" & AutoOpenOfficeRunner_GetString($handle, "認証局"), "\\", "/"), _
			"HookCreateIntermediateCrt" _
			)
EndFunc   ;==>CreateIntermediateCrt

;
; 中間証明書の設定ファイル書き換えフック関数.
;
; @param $config 設定ファイル名.
;
Func HookCreateIntermediateCrt($config)
	WriteConfigurationArray($Subject, "req_distinguished_name", $config)
EndFunc   ;==>HookCreateIntermediateCrt
#endregion Intermediate_Certificate

#region Server_Certificate
;
; サーバ証明書の生成処理.
;
; @param $handle ハンドル.
;
Func CreateServerCrt(Const $handle)
	InitConfigurationArray($Subject)
	SetConfigurationArray($handle, $Subject)

	OpenSSL_CreateServerCertificate( _
			AutoOpenOfficeRunner_GetString($handle, "名称"), _
			AutoOpenOfficeRunner_GetString($handle, "鍵長"), _
			AutoOpenOfficeRunner_GetString($handle, "メッセージダイジェスト"), _
			AutoOpenOfficeRunner_GetString($handle, "有効期限"), _
			StringRegExpReplace($OutputDir & "\" & AutoOpenOfficeRunner_GetString($handle, "認証局"), "\\", "/"), _
			"HookCreateServerCrt" _
			)
EndFunc   ;==>CreateServerCrt

;
; サーバ証明書の設定ファイル書き換えフック関数.
;
; @param $config 設定ファイル名.
;
Func HookCreateServerCrt($config)
	WriteConfigurationArray($Subject, "req_distinguished_name", $config)
EndFunc   ;==>HookCreateServerCrt
#endregion Server_Certificate

#region Configuration
;
; 設定書き換え用の配列を初期化する.
;
; @param $array 設定値配列.
;
Func InitConfigurationArray(ByRef $array)
	Local $count = UBound($array, 1)
	For $i = 0 To $count - 1
		$array[$i][1] = ""
	Next
EndFunc   ;==>InitConfigurationArray

;
; 入力ファイルから値を読込み, 設定書き換え用配列に設定する.
;
; @param $handle ハンドル.
; @param $array 設定値配列.
;
Func SetConfigurationArray(Const ByRef $handle, ByRef $array)
	Local $count = UBound($array, 1)
	For $i = 0 To $count - 1
		Local $value = AutoOpenOfficeRunner_GetString($handle, $array[$i][0])
		If "" <> $value Then
			$array[$i][1] = $value
		EndIf
	Next
EndFunc   ;==>SetConfigurationArray

;
; 設定書き換え用配列を設定ファイル(openssl.cnf)に書き込む.
;
; @param $array 設定値配列.
; @param $section 設定値を書き込むセクション.
; @param $config 設定ファイル名.
;
Func WriteConfigurationArray(ByRef $array, $section, $config)
	Local $count = UBound($array, 1)
	For $i = 0 To $count - 1
		If "" <> $array[$i][1] Then
			IniWrite($config, $section, $array[$i][0], $array[$i][1])
		EndIf
	Next
EndFunc   ;==>WriteConfigurationArray
#endregion Configuration