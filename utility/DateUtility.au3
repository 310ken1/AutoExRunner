#include-once
#include <Date.au3>

#region Globale_Argument_Define
; ハンドル(配列)のインデックス.
Global Enum _
		$DateUtilityMonth = 0, _
		$DateUtilityDay = 1, _
		$DateUtilityYear = 2, _
		$DateUtilityHour = 3, _
		$DateUtilityMinute = 4, _
		$DateUtilitySecond = 5, _
		$DateUtilityMillisecond = 6
#endregion Globale_Argument_Define

#region Public_Method
;
; 日付けと時刻の入った文字列(yyyy/mm/dd[ hh:mm[:ss]])を配列に変換する.
;
; @param $date_string 日付けと時刻の入った文字列(yyyy/mm/dd[ hh:mm[:ss]]).
; @error 0:成功 1:フォーマットが不正.
; @return 日付配列([月,日,年,時,分,秒].
;
Func DateUtility_DateStringToArray($date_string)
	Local $result[6] = [0, 0, 0, 0, 0, 0]
	Local $reg = '\d{4}/\d{1,2}/\d{1,2}( \d{1,2}:\d{1,2}(:\d{1,2})?)?$'
	If StringRegExp($date_string, $reg, 0) Then
		Local $date
		Local $time
		_DateTimeSplit($date_string, $date, $time)
		If @error Then
			SetError(@error)
		Else
			$result[$DateUtilityMonth] = $date[2]
			$result[$DateUtilityDay] = $date[3]
			$result[$DateUtilityYear] = $date[1]
			If 0 < $time[0] Then
				$result[$DateUtilityHour] = $time[1]
				$result[$DateUtilityMinute] = $time[2]
				If 2 < $time[0] Then
					$result[$DateUtilitySecond] = $time[3]
				EndIf
			EndIf
		EndIf
	Else
		SetError(1)
	EndIf
	Return $result
EndFunc   ;==>DateUtility_DateStringToArray
;
; 日付けと時刻の入った文字列(yyyy/mm/dd[ hh:mm[:ss]])を指定して,
; 現在の地域時間と日付けを設定する.
;
; @param $date_string 日付けと時刻の入った文字列(yyyy/mm/dd[ hh:mm[:ss]]).
; @return 成否(True/False).
;
Func DateUtility_SetLocalTimeString($date_string)
	Local $result = False
	Local $time = DateUtility_DateStringToArray($date_string)
	If Not @error Then
		Local $new = _Date_Time_EncodeSystemTime( _
				$time[$DateUtilityMonth], _
				$time[$DateUtilityDay], _
				$time[$DateUtilityYear], _
				$time[$DateUtilityHour], _
				$time[$DateUtilityMinute], _
				$time[$DateUtilitySecond])
		$result = _Date_Time_SetLocalTime($new)
	EndIf
	Return $result
EndFunc   ;==>DateUtility_SetLocalTimeString
;
; 現在時刻を保存する.
;
; @return ハンドラ.
;
Func DateUtility_Save()
	Local $handle[2]
	$handle[0] = _Date_Time_GetLocalTime()
	$handle[1] = _Date_Time_GetTickCount()
	Return $handle
EndFunc   ;==>DateUtility_Save
;
; 保存した時刻に戻す.
; DateUtility_Save関数を実行した時刻 + 経過時間 に戻す.
; 戻す精度は, 秒単位.
; 経過時間の上限は, 日単位. 月単位の経過は戻せない.
; Windows Vista 以降のOSでは, 実行に管理者権限が必要となる.
; 管理者権限がない場合は, ユーザーアカウントコントロール(UAC)の承認ダイアログが表示される.
;
; @param $handle ハンドラ(DateUtility_Save関数の戻り値).
; @return 成否.
;
Func DateUtility_Restore(ByRef $handle)
	#RequireAdmin
	Local $old = _Date_Time_SystemTimeToArray($handle[0])
	Local $elapsed = _Date_Time_GetTickCount() - $handle[1]
	$old[$DateUtilityDay] += $elapsed / (24 * 60 * 60 * 1000)
	$old[$DateUtilityHour] += Mod($elapsed, (24 * 60 * 60 * 1000)) / (60 * 60 * 1000)
	$old[$DateUtilityMinute] += Mod($elapsed, (60 * 60 * 1000)) / (60 * 1000)
	$old[$DateUtilitySecond] += Mod($elapsed, (60 * 1000)) / 1000
	Local $new = _Date_Time_EncodeSystemTime( _
			$old[$DateUtilityMonth], _
			$old[$DateUtilityDay], _
			$old[$DateUtilityYear], _
			$old[$DateUtilityHour], _
			$old[$DateUtilityMinute], _
			$old[$DateUtilitySecond])
	Return _Date_Time_SetLocalTime($new)
EndFunc   ;==>DateUtility_Restore
#endregion Public_Method