#include-once
#include "..\utility\ConsoleUtility.au3"

#region Globale_Argument_Define
;
; !! 必須 !!
; 本ソースコード上で定義されたメソッドを利用する場合は,
; $Wireshark_TsharkPath グローバル変数 に tshark.exe へのパスを設定すること.
;
Global $Wireshark_TsharkPath = "C:\Program Files\Wireshark\tshark.exe"
;
; デバッグログフラグ.
; 1 を指定することで, デバッグログが出力される.
;
Global $Wireshark_DebugLog = 0
#endregion Globale_Argument_Define

#region Public_Method
;
; パケットキャプチャーを開始する.
;
; @param $file パケットキャプチャー結果を保存するファイル.
;
Func Wireshark_StartCapture($file)
	Local $cmd = StringFormat("""%s"" -w ""%s""", $Wireshark_TsharkPath, $file)
	ConsoleUtility_DebugLogLn($Wireshark_DebugLog, $cmd)
	Run($cmd)
	WinWaitActive($Wireshark_TsharkPath)
EndFunc   ;==>Wireshark_StartCapture

;
; パケットキャプチャーを終了する.
;
Func Wireshark_EndCapture()
	WinClose($Wireshark_TsharkPath)
	WinWaitClose($Wireshark_TsharkPath)
EndFunc   ;==>Wireshark_EndCapture

;
; pcapng 形式のファイルを plain text 形式に変換する.
;
; @param $infile pcapng 形式のファイル(入力).
; @param $outfile plain text 形式のファイル(出力).
;
Func Wireshark_PcapngToText($infile, $outfile)
	Local $cmd = StringFormat("cmd.exe /k """"%s"" -r ""%s"" > ""%s"" & exit""", $Wireshark_TsharkPath, $infile, $outfile)
	ConsoleUtility_DebugLogLn($Wireshark_DebugLog, $cmd)
	RunWait($cmd)
EndFunc   ;==>Wireshark_PcapngToText
#endregion Public_Method
