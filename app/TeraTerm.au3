#include-once
#include "..\utility\ConsoleUtility.au3"

#region グローバル変数定義
;
; !! 必須 !!
; 本ソースコード上で定義されたメソッドを利用する場合は,
; $TeraTerm_CmdPath グローバル変数 に ttermpro.exe へのパスを設定する,
; もしくは, $TeraTerm_MacroCmdPath グローバル変数 に ttpmacro.exe へのパスを設定すること.
; どちらのグローバル変数を設定するかは, 実行する関数によって異なる.
;
Global $TeraTerm_CmdPath = "C:\usr\local\teraterm-4.74\ttermpro.exe"
Global $TeraTerm_MacroCmdPath = "C:\usr\local\teraterm-4.74\ttpmacro.exe"

;
; デバッグログフラグ.
; 1 を指定することで, デバッグログが出力される.
;
Global $TeraTerm_DebugLog = 0
#endregion グローバル変数定義

#region 定数定義
Const $TeraTerm_Class = "[CLASS:VTWin32]"
Const $TeraTerm_MacroTitle = "MACRO - "
#endregion 定数定義

#region パブリックメソッド
;
; TeraTerm ウィンドウを閉じる.
;
Func TeraTerm_Exit()
	WinKill($TeraTerm_Class)
	WinWaitClose($TeraTerm_Class)
EndFunc   ;==>TeraTerm_Exit

;
; TeraTerm Macroを実行する.
; Macro実行中はダイアログが表示され, Macro実行完了後にクローズされる.
;
; @param $ttl TeraTerm Macro ファイル.
;
Func TeraTerm_MacroRun($ttl)
	Local $cmd = StringFormat("%s %s", $TeraTerm_MacroCmdPath, $ttl)
	ConsoleUtility_DebugLogLn($TeraTerm_DebugLog, $cmd)
	Run($cmd)
	WinWaitActive($TeraTerm_MacroTitle)
EndFunc   ;==>TeraTerm_MacroRun

;
; TeraTerm Macroを実行し、実行が完了するまでスクリプト処理を一時停止する.
; Macro実行中はダイアログが表示され, Macro実行完了後にクローズされる.
;
; @param $ttl TeraTerm Macro ファイル.
;
Func TeraTerm_MacroRunWait($ttl)
	Local $cmd = StringFormat("%s %s", $TeraTerm_MacroCmdPath, $ttl)
	ConsoleUtility_DebugLogLn($TeraTerm_DebugLog, $cmd)
	RunWait($cmd)
EndFunc   ;==>TeraTerm_MacroRunWait

;
; TeraTerm Macroの実行が終了するまで待つ.
;
Func TeraTerm_MacroWaitClose()
	WinWaitClose($TeraTerm_MacroTitle)
EndFunc   ;==>TeraTerm_MacroWaitClose
#endregion パブリックメソッド