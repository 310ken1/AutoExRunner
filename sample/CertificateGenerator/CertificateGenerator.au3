#include "..\..\runner\AutoOpenOfficeRunner.au3"
#include "..\..\app\OpenSSL.au3"

$OpenSSLCmd = @ScriptDir & "\..\..\bin\openssl.exe"
$AutoOpenOfficeRuunerConfig = @ScriptDir & "\CertificateGenerator.ini"

;
; 出力フォルダ.
;
Const $OutputDir = @ScriptDir & "\out"

;
; 入力ファイル.
;
Const $InputFile = @ScriptDir & "\CertificateList.ods"

;
; サブジェクト.
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

;
; ルート証明書の生成処理.
;
; @param $sheet 実行中のシートオブジェクト.
; @param $line 実行中の行.
;
Func CreateRootCrt($sheet, $line)
	InitSubject()
	SetSubject($sheet, $line)
	CreateRootCertificate( _
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
	WriteSubject($config)
EndFunc   ;==>HookCreateRootCrt

;
; 中間証明書の生成処理.
;
; @param $sheet 実行中のシートオブジェクト.
; @param $line 実行中の行.
;
Func CreateIntermediateCrt($sheet, $line)
	InitSubject()
	SetSubject($sheet, $line)
	CreateIntermediateCertificate( _
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
	WriteSubject($config)
EndFunc   ;==>HookCreateIntermediateCrt

;
; サーバ証明書の生成処理.
;
; @param $sheet 実行中のシートオブジェクト.
; @param $line 実行中の行.
;
Func CreateServerCrt($sheet, $line)
	InitSubject()
	SetSubject($sheet, $line)
	CreateServerCertificate( _
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
	WriteSubject($config)
EndFunc   ;==>HookCreateServerCrt

;
; サブジェクトを初期化する.
;
Func InitSubject()
	Local $count = UBound($subject, 1)
	For $i = 0 To $count - 1
		$subject[$i][1] = ""
	Next
EndFunc   ;==>InitSubject

;
; 入力ファイルから値を読込み, サブジェクトに設定する.
;
; @param $sheet 実行中のシートオブジェクト.
; @param $line 実行中の行.
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
; サブジェクトを設定ファイルに書き込む.
;
; @param $config 設定ファイル名.
;
Func WriteSubject($config)
	Local $count = UBound($subject, 1)
	For $i = 0 To $count - 1
		If Not "" = $subject[$i][1] Then
			IniWrite($config, "req_distinguished_name", $subject[$i][0], $subject[$i][1])
		EndIf
	Next
EndFunc   ;==>WriteSubject

