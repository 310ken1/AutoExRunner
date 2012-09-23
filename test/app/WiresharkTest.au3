#include "..\AutoItTest.au3"
#include "..\..\app\Wireshark.au3"
#include "..\..\utility\FileUtility.au3"

#region Globale_Argument_Define
$Wireshark_TsharkPath = "C:\Program Files\Wireshark\tshark.exe"
#endregion Globale_Argument_Define

#region Constant_Define
; キャプチャーファイル(pcapng形式).
Const  $CaptureFile = "Wireshark.pcapng"
; キャプチャーファイル(text形式).
Const $CaptureTextFile = "Wireshark.txt"
#endregion Constant_Define

Local $WiresharkTest[2][5] = [ _
		["", "Wireshark_StartCapture_Test", "AutoItTest_FileExists", $CaptureFile, "Wireshark_StartCapture_Test_After"], _
		["", "Wireshark_PcapngToText_Test", "AutoItTest_FileExists", $CaptureTextFile, "Wireshark_PcapngToText_Test_After"] _
		]

AutoItTest_Runner($WiresharkTest)

#region Wireshark_StartCapture_Test
;
; Wireshark_StartCapture関数のテスト.
;
; 期待する結果: キャプチャーファイルが作成される.
;
Func Wireshark_StartCapture_Test()
	Wireshark_StartCapture($CaptureFile)
	Sleep(3000)
	Wireshark_EndCapture()
	Return True
EndFunc
;
; Wireshark_StartCapture関数テストの後処理.
;
Func Wireshark_StartCapture_Test_After()
	FileDelete($CaptureFile)
EndFunc
#endregion

#region Wireshark_PcapngToText_Test
;
; Wireshark_PcapngToText関数のテスト.
;
; 期待する結果: キャプチャーファイルをPlain Textにしたファイルが作成される.
;
Func Wireshark_PcapngToText_Test()
	Wireshark_StartCapture($CaptureFile)
	Sleep(3000)
	Wireshark_EndCapture()
	Wireshark_PcapngToText($CaptureFile, $CaptureTextFile)
	Return True
EndFunc
;
; Wireshark_PcapngToText関数テストの後処理.
;
Func Wireshark_PcapngToText_Test_After()
	FileDelete($CaptureFile)
	FileDelete($CaptureTextFile)
EndFunc
#endregion
