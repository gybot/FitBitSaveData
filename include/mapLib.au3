
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Version=Beta
#AutoIt3Wrapper_Res_Language=1033
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include-once

#include <Array.au3>
#include "jsonMap.au3"

Global $blankMap[]
Global $mapBlank[]

Func excelHeaderCopyToMap()
	Local $clip = ClipGet()

	Local $keyIndex[]
	Local $map[]

	Local $header = regex($clip, '^(.+?)[\r\n].*')
	; (?s)(?:(?:\t|^)"?(.*?)"?(?:\t|$|[\r\n]))+

	Local $tab = StringRegExp($header, regexlib('tab'), 3)
	Local $headerItems = UBound($tab)
	For $io = 0 To $headerItems - 1
		$keyIndex[$io] = $tab[$io]
	Next

	If $headerItems == 0 Then
		ConsoleWrite('ERROR: no header provided' & @CRLF)
		Return
	EndIf

	Local $line = StringRegExp($clip, '(?s)((?<=\t|^|\r\n|\n])(?<!\\)".*?"(?=\t|$|\r\n|\n)|(?<=\t|^|\r\n|\n]).*?(?=\t|$|\r\n|\n))', 3)
	Local $ii = $headerItems
	While ($ii + $headerItems) < UBound($line)
		;ConsoleWrite('#### Set: ' & $ii & @CRLF)
		Local $tMap[]
		Local $io = 0
		While $io < $headerItems
			Local $item = $line[$ii + $io] ? $line[$ii + $io] : $line[$ii + $io + 1]
			Local $item = regex($line[$ii + $io], '(?s)^["\s\h\v]*(.*?)["\s\h\v]*$')
			;ConsoleWrite($ii + $io & ' - ' & $keyIndex[$io] & @TAB & $item & @CRLF)
			$tMap[$keyIndex[$io]] = $item

			$io += 1
		WEnd
		MapAppend($map, $tMap)
		$ii += $headerItems
	WEnd
	Return $map
EndFunc   ;==>excelHeaderCopyToMap

Func excelCopyToMap($headers = True)
	Local $clip = ClipGet()

	Local $keyIndex[]
	Local $map[]

	Local $line = StringRegExp($clip, regexlib('line'), 3)
	For $ii = 0 To UBound($line) - 1
		If $ii = 0 Then
			Local $tab = StringRegExp($line[$ii], regexlib('tab'), 3)
			For $io = 0 To UBound($tab) - 1
				$keyIndex[$io] = $tab[$io]
			Next
		Else
			Local $tMap[]
			Local $tab = StringRegExp($line[$ii], regexlib('tab'), 3)
			For $io = 0 To UBound($tab) - 1
				If $headers Then
					$tMap[$keyIndex[$io]] = $tab[$io]
				Else
					$tMap[$io] = $tab[$io]
				EndIf
			Next
			MapAppend($map, $tMap)
		EndIf
	Next
	Return $map
EndFunc   ;==>excelCopyToMap

Func mapJoinR(ByRef $map, ByRef Const $mapj)
	Local $keys = MapKeys($mapj)
	Local $size = UBound($keys)
	Local $ii
	For $ii = 0 To $size - 1
		$map[$keys[$ii]] = $mapj[$keys[$ii]]
	Next
EndFunc   ;==>mapJoinR

Func mapJoin($map, ByRef Const $mapj)
	If Not IsMap($map) And IsMap($mapj) Then Return SetError(1, 0, $mapj)
	Local $keys = MapKeys($mapj)
	Local $size = UBound($keys)
	Local $ii
	For $ii = 0 To $size - 1
		$map[$keys[$ii]] = $mapj[$keys[$ii]]
	Next
	Return $map
EndFunc   ;==>mapJoin

Func arrMapToMap(Const ByRef $arr, $keyName)
	Local $ii, $ta
	Local $map[]
	Local $max = UBound($arr) - 1
	For $ii = 0 To $max
		$ta = $arr[$ii]
		$map[$ta[$keyName]] = $ta
	Next
	Return $map
EndFunc   ;==>arrMapToMap

Func arrToMap($arr)
	Local $ii
	Local $map[]
	For $ii = 0 To UBound($arr) - 1
		$map[$arr[$ii]] = $ii
	Next
	Return $map
EndFunc   ;==>arrToMap

Func mapToArr(Const ByRef $map)
	Local $keys = MapKeys($map)
	Local $size = UBound($keys)
	Local $arr[$size]
	Local $ii
	For $ii = 0 To $size - 1
		$arr[$ii] = $map[$keys[$ii]]
	Next
	Return $arr
EndFunc   ;==>mapToArr

Func mapToCSV(Const ByRef $map, $sort = -1)

	Local $csv

	; GET keys
	Local $keymap[]
	Local $arrMap = MapKeys($map)
	If $sort > -1 Then
		_ArraySort($arrMap, $sort)
	EndIf

	Local $arrMapMax = UBound($arrMap) - 1
	Local $ii, $io
	Local $mm, $mmMax
	Local $csvHeader
	For $ii = 0 To $arrMapMax
		;ConsoleWrite($arrMap[$arrMap[$ii]] & @CRLF)
		$mm = MapKeys($map[$arrMap[$ii]])
		$mmMax = UBound($mm) - 1
		For $io = 0 To $mmMax
			;ConsoleWrite('   ' & $mm[$io] & @TAB & $io & '/' & $mmMax & @CRLF)
			If Not MapExists($keymap, $mm[$io]) Then
				$keymap[$mm[$io]] = $io
				$csvHeader &= csvEscape($mm[$io]) & ','
			EndIf
		Next
	Next
	$csvHeader = StringTrimRight($csvHeader, 1)

	$csv = $csvHeader & @CRLF

	Local $keyMapKeys = MapKeys($keymap)
	Local $keyMapMax = UBound($keymap) - 1
	Local $ii, $io
	Local $im
	For $ii = 0 To $arrMapMax
		$im = $map[$arrMap[$ii]]
		;ConsoleWrite('---' & $arrMap[$ii] & '---' & @CRLF)
		For $io = 0 To $keyMapMax
			;ConsoleWrite($keyMapKeys[$io] & @CRLF)
			If MapExists($im, $keyMapKeys[$io]) Then
				If $io = $keyMapMax Then
					$csv &= csvEscape($im[$keyMapKeys[$io]])
				Else
					$csv &= csvEscape($im[$keyMapKeys[$io]]) & ','
				EndIf
			Else
				If $io = $keyMapMax Then
					$csv &= '""'
				Else
					$csv &= '"",'
				EndIf
			EndIf
		Next
		$csv &= @CRLF
	Next
	Return $csv
EndFunc   ;==>mapToCSV

Func mapToCSVs($map, $sort = -1)

	if IsString($map) Then
		$map = json_decode($map)
	EndIf

	;jprint($map)

	If IsArray($map) Then
		Local $ii
		Local $mp[]
		For $ii = 0 To UBound($map) - 1
			$mp[$ii] = $map[$ii]
		Next
	EndIf
	$map = $mp
	;jprint($map)

	Local $csv

	; GET nomalized keys
	Local $keymap[]
	Local $arrMap = MapKeys($map)
	If $sort > -1 Then
		_ArraySort($arrMap, $sort)
	EndIf

	Local $arrMapMax = UBound($arrMap) - 1
	Local $ii, $io
	Local $mm, $mmMax
	Local $csvHeader
	For $ii = 0 To $arrMapMax
		;ConsoleWrite($arrMap[$arrMap[$ii]] & @CRLF)
		$mm = MapKeys($map[$arrMap[$ii]])
		$mmMax = UBound($mm) - 1
		For $io = 0 To $mmMax
			;ConsoleWrite('   ' & $mm[$io] & @TAB & $io & '/' & $mmMax & @CRLF)
			If Not MapExists($keymap, $mm[$io]) Then
				$keymap[$mm[$io]] = $io
				$csvHeader &= csvEscape($mm[$io]) & ','
			EndIf
		Next
	Next
	$csvHeader = StringTrimRight($csvHeader, 1)

	$csv = $csvHeader & @CRLF

	Local $keyMapKeys = MapKeys($keymap)
	Local $keyMapMax = UBound($keymap) - 1
	Local $ii, $io
	Local $im
	For $ii = 0 To $arrMapMax
		$im = $map[$arrMap[$ii]]
		;ConsoleWrite('---' & $arrMap[$ii] & '---' & @CRLF)
		For $io = 0 To $keyMapMax
			;ConsoleWrite($keyMapKeys[$io] & @CRLF)
			If MapExists($im, $keyMapKeys[$io]) Then
				If $io = $keyMapMax Then
					$csv &= csvEscape($im[$keyMapKeys[$io]])
				Else
					$csv &= csvEscape($im[$keyMapKeys[$io]]) & ','
				EndIf
			Else
				If $io = $keyMapMax Then
					$csv &= '""'
				Else
					$csv &= '"",'
				EndIf
			EndIf
		Next
		$csv &= @CRLF
	Next
	Return $csv
EndFunc   ;==>mapToCSVs



Func mapIndex(Const ByRef $map, $ii = 0)
	Local $arr = MapKeys($map)
	Local $max = UBound($map) - 1
	If $ii > -1 And $ii <= $max Then
		Return SetError(0, 0, $arr[$ii])
	Else
		Return SetError(1, 0, '')
	EndIf
EndFunc   ;==>mapIndex


Func jsonPairToMap(Const ByRef $json)
	Local $arr = StringRegExp($json, '(?:"(.+?)(?<!\\)"\s*:\s*(".+?(?<!\\)"|[\d\.]+|(?i)null))', 3)
	Local $ii
	Local $map[]
	Local $vv
	For $ii = 0 To UBound($arr) - 1 Step 2
		$vv = $arr[$ii + 1]
		If $vv = 'null' Then
			$vv = ''
		ElseIf StringRegExp($vv, '^[\d\.]+$') Then
			$vv = Number($vv)
		Else
			$vv = regex($arr[$ii + 1], '^"?(.+?)"?$')
			$vv = Json_StringDecode($vv)
		EndIf
		$map[$arr[$ii]] = $vv
	Next
	Return $map
EndFunc   ;==>jsonPairToMap


Func fileReadJson($fileName)
	Local $jmap = BinaryToString(FileRead($fileName))
	If Not $jmap Then
		$jmap = $blankMap
	Else
		$jmap = json_decode($jmap)
	EndIf
	Return $jmap
EndFunc   ;==>fileReadJson


Func fileWriteJson($fileName, $data)
	Local $er = 1
	Local $rc
	Local $fh = FileOpen($fileName, 2 + 8)
	$rc = FileWrite($fh, json_encode($data))
	If $rc = 1 Then $er = 0
	FileFlush($fh)
	FileClose($fh)
	Return SetError($er, 0, $rc)
EndFunc   ;==>fileWriteJson


Func mapRead(Const ByRef $map, $item, $item1 = '')
	If Not MapExists($map, $item) Then Return ''
	If Not $item1 Then
		Return $map[$item]
	Else
		If Not MapExists($map[$item], $item1) Then Return ''
		Return $map[$item][$item1]
	EndIf
EndFunc   ;==>mapRead

Func mapWrite(ByRef $map, $it1, $it2, $val = '')
	If Not MapExists($map, $it1) Then
		$map[$it1] = $blankMap
	EndIf
	If Not $val Then
		$map[$it1] = $it2
	Else
		If Not MapExists($map[$it1], $it2) Then
			$map[$it1][$it2] = $blankMap
		EndIf
		$map[$it1][$it2] = $val
	EndIf
EndFunc   ;==>mapWrite





















