#include-once
#region パブリックメソッド
;
; スペースが含まれるパスを有効化(ダブルクォートで囲む)する.
;
; @param $path パス.
;
Func FileUtility_PathSpaceEnable($path)
	Return """" & $path & """"
EndFunc   ;==>FileUtility_PathSpaceEnable

;
; スクリプトフォルダにあるファイルのパスを取得する.
;
; @param $file ファイル名.
;
Func FileUtility_ScriptDirFilePath($file)
	Return FileUtility_PathSpaceEnable(@ScriptDir & "\" & $file)
EndFunc   ;==>FileUtility_ScriptDirFilePath
#endregion パブリックメソッド