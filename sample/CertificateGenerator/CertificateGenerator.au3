#include "..\..\runner\AutoOpenOfficeRunner.au3"
#include "..\..\utility\FileUtility.au3"
#include "..\..\app\OpenSSL.au3"

#region グローバル変数設定
$OpenSSL_CmdPath = FileUtility_ScriptDirFilePath("..\..\bin\openssl.exe")
$OpenSSL_DebugLog = 1
$AutoOpenOfficeRuunerConfig = FileUtility_ScriptDirFilePath("CertificateGenerator.ini")
#endregion グローバル変数設定

#region 定数定義
;
; 出力フォルダ.
;
Const $OutputDir = FileUtility_ScriptDirFilePath("out")

;
; 入力ファイル.
;
Const $InputFile = FileUtility_ScriptDirFilePath("CertificateList.ods")
#endregion 定数定義

#region 静的変数定義
;
; サブジェクトの設定書き換え用配列.
;
Static $Subject[8][2] = [ _
		["C", ""], _
		["ST", ""], _
		["L", ""], _
		["O", ""], _
		["OU", ""], _
		["CN", ""], _
		["emailAddress", ""] _
		]
#endregion 静的変数定義

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

	AutoOpenOfficeRunner($InputFile, "ルート証明書", "CreateRootCrt")
	AutoOpenOfficeRunner($InputFile, "中間証明書", "CreateIntermediateCrt")
	AutoOpenOfficeRunner($InputFile, "サーバ証明書", "CreateServerCrt")
EndFunc   ;==>Main

#region ルート証明書
;
; ルート証明書の生成処理.
;
; @param $sheet 実行中のシートオブジェクト.
; @param $line 実行中の行.
;
Func CreateRootCrt($sheet, $line)
	InitConfigurationArray($Subject)
	SetConfigurationArray($sheet, $line, $Subject)

	OpenSSL_CreateRootCertificate( _
			GetString($sheet, $line, "名称"), _
			GetString($sheet, $line, "鍵長"), _
			GetString($sheet, $line, "メッセージダイジェスト"), _
			GetString($sheet, $line, "有効期限"), _
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
#endregion ルート証明書

#region 中間証明書
;
; 中間証明書の生成処理.
;
; @param $sheet 実行中のシートオブジェクト.
; @param $line 実行中の行.
;
Func CreateIntermediateCrt($sheet, $line)
	InitConfigurationArray($Subject)
	SetConfigurationArray($sheet, $line, $Subject)

	OpenSSL_CreateIntermediateCertificate( _
			GetString($sheet, $line, "名称"), _
			GetString($sheet, $line, "鍵長"), _
			GetString($sheet, $line, "メッセージダイジェスト"), _
			GetString($sheet, $line, "有効期限"), _
			StringRegExpReplace($OutputDir & "\" & GetString($sheet, $line, "認証局"), "\\", "/"), _
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
#endregion 中間証明書

#region サーバ証明書
;
; サーバ証明書の生成処理.
;
; @param $sheet 実行中のシートオブジェクト.
; @param $line 実行中の行.
;
Func CreateServerCrt($sheet, $line)
	InitConfigurationArray($Subject)
	SetConfigurationArray($sheet, $line, $Subject)

	OpenSSL_CreateServerCertificate( _
			GetString($sheet, $line, "名称"), _
			GetString($sheet, $line, "鍵長"), _
			GetString($sheet, $line, "メッセージダイジェスト"), _
			GetString($sheet, $line, "有効期限"), _
			StringRegExpReplace($OutputDir & "\" & GetString($sheet, $line, "認証局"), "\\", "/"), _
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
#endregion サーバ証明書

#region 設定ファイル書き換え
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
; @param $sheet 実行中のシートオブジェクト.
; @param $line 実行中の行.
; @param $array 設定値配列.
;
Func SetConfigurationArray($sheet, $line, ByRef $array)
	Local $count = UBound($array, 1)
	For $i = 0 To $count - 1
		Local $value = GetString($sheet, $line, $array[$i][0])
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
#endregion 設定ファイル書き換え