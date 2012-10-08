#include-once
#include <File.au3>

#region Public_Method
;
; スペースが含まれるパスを, コマンドラインで利用出来る様にダブルクォートで囲む.
; 戻り値のパスは, AutoIt API で利用出来ないため, 注意すること.
;
; @param $path パス.
; @return ダブルクォートで囲まれたパス.
;
Func FileUtility_PathSpaceEnable($path)
	Return """" & $path & """"
EndFunc   ;==>FileUtility_PathSpaceEnable

;
; フォルダ名とファイル名を結合する.
;
; @param $dir フォルダ名.
; @param $file ファイル名.
; @return フォルダ名とファイル名を結合したパス.
;
Func FileUtility_MakePath($dir, $file)
	Return $dir & "\" & $file
EndFunc   ;==>FileUtility_MakePath

;
; スクリプトフォルダにあるファイルのパスを取得する.
;
; @param $file ファイル名.
; @return スクリプトフォルダにあるファイルのパス.
;
Func FileUtility_ScriptDirFilePath($file)
	Return FileUtility_MakePath(@ScriptDir, $file)
EndFunc   ;==>FileUtility_ScriptDirFilePath

;
; テンポラリーファイルフォルダにあるファイルのパスを取得する.
;
; @param $file ファイル名.
; @return テンポラリーファイルフォルダにあるファイルのパス.
;
Func FileUtility_TempDirFilePath($file)
	Return FileUtility_MakePath(@TempDir, $file)
EndFunc   ;==>FileUtility_TempDirFilePath

;
; 現在の作業フォルダを移動する.
; 移動先のフォルダが無い場合は, フォルダを作成して移動する.
;
; @param $path 移動先フォルダ.
; @return 移動前のフォルダ.
;
Func FileUtiilty_ChangeDir($path)
	Local $before = @WorkingDir
	If Not FileExists($path) Then
		DirCreate($path)
	EndIf
	FileChangeDir($path)
	Return $before
EndFunc   ;==>FileUtiilty_ChangeDir

;
; パスをURLに変換する.
;
; @param $path パス.
; @return  URL.
;
Func FileUtiilty_PathToUrl($path)
	Local $ret = "file:///" & StringRegExpReplace(StringRegExpReplace($path, """", ""), "\\", "/")
	Return $ret
EndFunc   ;==>FileUtiilty_PathToUrl

;
; フルパスからファイル名を取得する.
;
; @param $fullpath フォルダ名.
; @return ファイル名.
;
Func FileUtility_FileBaseName($fullpath)
	Local $drive, $dir, $file, $ext
	_PathSplit($fullpath, $drive, $dir, $file, $ext)
	Return $file & $ext
EndFunc   ;==>FileUtility_FileBaseName
#endregion Public_Method