#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Version=Beta
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****


; ============================================================================================================================
; File		: JsonMap.au3 (2017.02.09)
; Purpose	: A Non-Strict JavaScript Object Notation (JSON) Parser UDF implemented using Map
; Origianl Author	: Ward
; Dependency: BinaryCall.au3
; Website	: http://www.json.org/index.html
;
; Map adaptation: dexto
;
; Source	: jsmn.c
; Author	: zserge
; Website	: http://zserge.com/jsmn.html
;
; Source	: json_string_encode.c, json_string_decode.c
; Author	: Ward
;
; Example   : https://www.autoitscript.com/forum/topic/148114-a-non-strict-json-udf-jsmn/
; ============================================================================================================================

; ============================================================================================================================
; Public Functions:
;   Json_StringEncode(Const ByRef $String, $Option = 0)
;   Json_StringDecode($String)
;   Json_Encode(Const ByRef $Data, $Option = 0, $Indent = Default, $ArraySep = Default, $ObjectSep = Default, $ColonSep = Default)
;   Json_Decode(Const ByRef $Json, $InitTokenCount = 1000)
;	Json_Beautify($json)
; ============================================================================================================================

#include-once
#include "BinaryCall.au3"

; The following constants can be combined to form options for Json_Encode()
Global Const $JSON_UNESCAPED_UNICODE = 1 ; Encode multibyte Unicode characters literally
Global Const $JSON_UNESCAPED_SLASHES = 2 ; Don't escape /
Global Const $JSON_HEX_TAG = 4 ; All < and > are converted to \u003C and \u003E
Global Const $JSON_HEX_AMP = 8 ; All &s are converted to \u0026
Global Const $JSON_HEX_APOS = 16 ; All ' are converted to \u0027
Global Const $JSON_HEX_QUOT = 32 ; All " are converted to \u0022
Global Const $JSON_UNESCAPED_ASCII = 64 ; Don't escape ascii charcters between chr(1) ~ chr(0x1f)
Global Const $JSON_PRETTY_PRINT = 128 ; Use whitespace in returned data to format it
Global Const $JSON_STRICT_PRINT = 256 ; Make sure returned JSON string is RFC4627 compliant
Global Const $JSON_UNQUOTED_STRING = 512 ; Output unquoted string if possible (conflicting with $Json_STRICT_PRINT)

; Error value returnd by Json_Decode()
Global Const $JSMN_ERROR_NOMEM = -1 ; Not enough tokens were provided
Global Const $JSMN_ERROR_INVAL = -2 ; Invalid character inside JSON string
Global Const $JSMN_ERROR_PART = -3 ; The string is not a full JSON packet, more bytes expected

Func jsonWrite(Const ByRef $json, $filename)
	Local $fh = FileOpen($filename, 2 + 8)
	FileWrite($fh, Json_Encode($json))
	FileFlush($fh)
	FileClose($fh)
EndFunc   ;==>jsonWrite

Func jsonRead($filename)
	Local $fh = FileOpen($filename, 0)
	Local $dat = FileRead($fh)
	FileClose($fh)
	Return Json_Decode($dat)
EndFunc   ;==>jsonRead

Func Json_BeautifyHTML($json)
	If Not IsMap($json) Then
		$json = Json_Decode($json)
		Return StringReplace(StringReplace(Json_Encode($json, $JSON_PRETTY_PRINT, "  ", ",\r\n", ",\r\n", ":"), @CRLF, '<br>' & @CRLF), ' ', '&nbsp;')
	Else
		Return StringReplace(StringReplace(Json_Encode($json, $JSON_PRETTY_PRINT, "  ", ",\r\n", ",\r\n", ":"), @CRLF, '<br>' & @CRLF), ' ', '&nbsp;')
	EndIf
EndFunc   ;==>Json_BeautifyHTML

Func Json_Beautify($json)
	If Not IsMap($json) Then
		$json = Json_Decode($json)
	EndIf
	Return Json_Encode($json, $JSON_PRETTY_PRINT, "  ", ",\r\n", ",\r\n", ":")
EndFunc   ;==>Json_Beautify


Func jCopy($obj, $text = '')
	ClipPut(($text ? $text & @TAB : '') & Json_Encode($obj))
EndFunc   ;==>jPrint

Func jPrint($obj, $text = '')
	ConsoleWrite(($text ? $text & @TAB : '') & Json_Encode($obj) & @CRLF)
EndFunc   ;==>jPrint

Func jPrintI($obj, $text = '')
	ConsoleWrite(($text ? $text & @TAB : '') & Json_Encode($obj))
EndFunc   ;==>jPrint

Func jPrintP($obj, $text = '')
	$ll(($text ? $text & @TAB : '') & Json_Encode($obj, $JSON_PRETTY_PRINT, "  ", ",\r\n", ",\r\n", ":") & @CRLF)
EndFunc   ;==>jsonPrint

Func jsonPrint(Const ByRef $obj)
	$ll(Json_Encode($obj, $JSON_PRETTY_PRINT, "  ", ",\r\n", ",\r\n", ":") & @CRLF)
EndFunc   ;==>jsonPrint


Func __Jsmn_RuntimeLoader($ProcName = "")
	Static $SymbolList
	If Not IsDllStruct($SymbolList) Then
		Local $Code
		If @AutoItX64 Then
			$Code = 'AwAAAAQfCAAAAAAAAAA1HbEvgTNrvX54gCiWSTVmt5v7RCdoFJ/zhkKmwcm8yVqZPjJBoVhNHHAIzrHWKbZh1J0QAUaHB5zyQTilTmWa9O0OKeLrk/Jg+o7CmMzjEk74uPongdHv37nwYXvg97fiHvjP2bBzI9gxSkKq9Cqh/GxSHIlZPYyW76pXUt//25Aqs2Icfpyay/NFd50rW7eMliH5ynkrp16HM1afithVrO+LpSaz/IojowApmXnBHUncHliDqbkx6/AODUkyDm1hj+AiEZ9Me1Jy+hBQ1/wC/YnuuYSJvNAKp6XDnyc8Nwr54Uqx5SbUW2CezwQQ7aXX/HFiHSKpQcFW/gi8oSx5nsoxUXVjxeNI/L7z6GF2mfu3Tnpt7hliWEdA2r2VB+TIM7Pgwl9X3Ge0T3KJQUaRtLJZcPvVtOuKXr2Q9wy7hl80hVRrt9zYrbjBHXLrRx/HeIMkZwxhmKo/dD/vvaNgE+BdU8eeJqFBJK2alrK2rh2WkRynftyepm1WrdKrz/5KhQPp/4PqH+9IADDjoGBbfvJQXdT+yiO8DtfrVnd+JOEKsKEsdgeM3UXx5r6tEHO9rYWbzbnyEiX7WozZemry+vBZMMtHn1aA63+RcDQED73xOsnj00/9E5Z6hszM5Hi8vi6Hw3iOgf3cHwcXG44aau0JpuA2DlrUvnJOYkNnY+bECeSdAR1UQkFNyqRoH2xm4Y7gYMCPsFtPBlwwleEKI27SsUq1ZHVQvFCoef7DXgf/GwPCAvwDMIQfb3hJtIVubOkASRQZVNIJ/y4KPrn/gcASV7fvMjE34loltTVlyqprUWxpI51tN6vhTOLAp+CHseKxWaf9g1wdbVs0e/5xAiqgJbmKNi9OYbhV/blpp3SL63XKxGiHdxhK1aR+4rUY4eckNbaHfW7ob+q7aBoHSs6LVX9lWakb/xWxwQdwcX/7/C+TcQSOOg6rLoWZ8wur9qp+QwzoCbXkf04OYpvD5kqgEiwQnB90kLtcA+2XSbDRu+aq02eNNCzgkZujeL/HjVISjf2EuQKSsZkBhS15eiXoRgPaUoQ5586VS7t7rhM8ng5LiVzoUQIZ0pNKxWWqD+gXRBvOMIXY2yd0Ei4sE5KFIEhbs3u8vwP7nFLIpZ/RembPTuc0ZlguGJgJ2F5iApfia+C2tRYRNjVCqECCveWw6P2Btfaq9gw7cWWmJflIQbjxtccDqsn52cftLqXSna9zk05mYdJSV8z2W7vM1YJ5Rd82v0j3kau710A/kQrN41bdaxmKjL+gvSRlOLB1bpvkCtf9+h+eVA4XIkIXKFydr1OjMZ8wq2FIxPJXskAe4YMgwQmeWZXMK1KBbLB3yQR1YOYaaHk1fNea9KsXgs5YLbiP/noAusz76oEDo/DJh1aw7cUwdhboVPg1bNq88mRb5RGa13KDK9uEET7OA02KbSL+Q4HOtyasLUoVrZzVyd8iZPoGrV36vHnj+yvG4fq6F/fkug/sBRp186yVZQVmdAgFd+WiRLnUjxHUKJ6xBbpt4FTP42E/PzPw3JlDb0UQtXTDnIL0CWqbns2E7rZ5PBwrwQYwvBn/gaEeLVGDSh84DfW4zknIneGnYDXdVEHC+ITzejAnNxb1duB+w2aVTk64iXsKHETq53GMH6DuFi0oUeEFb/xp0HsRyNC8vBjOq3Kk7NZHxCQLh7UATFttG7sH+VIqGjjNwmraGJ0C92XhpQwSgfAb3KHucCHGTTti0sn6cgS3vb36BkjGKsRhXVuoQCFH96bvTYtl8paQQW9ufRfvxPqmU0sALdR0fIvZwd7Z8z0UoEec6b1Sul4e60REj/H4scb6N2ryHBR9ua5N1YxJu1uwgoLXUL2wT9ZPBjPjySUzeqXikUIKKYgNlWy+VlNIiWWTPtKpCTr508logA=='
		Else
			$Code = 'AwAAAASFBwAAAAAAAAA1HbEvgTNrvX54gCiqsa1mt5v7RCdoAFjCfVE40DZbE5UfabA9UKuHrjqOMbvjSoB2zBJTEYEQejBREnPrXL3VwpVOW+L9SSfo0rTfA8U2W+Veqo1uy0dOsPhl7vAHbBHrvJNfEUe8TT0q2eaTX2LeWpyrFEm4I3mhDJY/E9cpWf0A78e+y4c7NxewvcVvAakIHE8Xb8fgtqCTVQj3Q1eso7n1fKQj5YsQ20A86Gy9fz8dky78raeZnhYayn0b1riSUKxGVnWja2i02OvAVM3tCCvXwcbSkHTRjuIAbMu2mXF1UpKci3i/GzPmbxo9n/3aX/jpR6UvxMZuaEDEij4yzfZv7EyK9WCNBXxMmtTp3Uv6MZsK+nopXO3C0xFzZA/zQObwP3zhJ4sdatzMhFi9GAM70R4kgMzsxQDNArueXj+UFzbCCFZ89zXs22F7Ixi0FyFTk3jhH56dBaN65S+gtPztNGzEUmtk4M8IanhQSw8xCXr0x0MPDpDFDZs3aN5TtTPYmyk3psk7OrmofCQGG5cRcqEt9902qtxQDOHumfuCPMvU+oMjzLzBVEDnBbj+tY3y1jvgGbmEJguAgfB04tSeAt/2618ksnJJK+dbBkDLxjB4xrFr3uIFFadJQWUckl5vfh4MVXbsFA1hG49lqWDa7uSuPCnOhv8Yql376I4U4gfcF8LcgorkxS+64urv2nMUq6AkBEMQ8bdkI64oKLFfO7fGxh5iMNZuLoutDn2ll3nq4rPi4kOyAtfhW0UPyjvqNtXJ/h0Wik5Mi8z7BVxaURTDk81TP8y9+tzjySB/uGfHFAzjF8DUY1vqJCgn0GQ8ANtiiElX/+Wnc9HWi2bEEXItbm4yv97QrEPvJG9nPRBKWGiAQsIA5J+WryX5NrfEfRPk0QQwyl16lpHlw6l0UMuk7S21xjQgyWo0MywfzoBWW7+t4HH9sqavvP4dYAw81BxXqVHQhefUOS23en4bFUPWE98pAN6bul+kS767vDK34yTC3lA2a8wLrBEilmFhdB74fxbAl+db91PivhwF/CR4Igxr35uLdof7+jAYyACopQzmsbHpvAAwT2lapLix8H03nztAC3fBqFSPBVdIv12lsrrDw4dfhJEzq7AbL/Y7L/nIcBsQ/3UyVnZk4kZP1KzyPCBLLIQNpCVgOLJzQuyaQ6k2QCBy0eJ0ppUyfp54LjwVg0X7bwncYbAomG4ZcFwTQnC2AX3oYG5n6Bz4SLLjxrFsY+v/SVa+GqH8uePBh1TPkHVNmzjXXymEf5jROlnd+EjfQdRyitkjPrg2HiQxxDcVhCh5J2L5+6CY9eIaYgrbd8zJnzAD8KnowHwh2bi4JLgmt7ktJ1XGizox7cWf3/Dod56KAcaIrSVw9XzYybdJCf0YRA6yrwPWXbwnzc/4+UDkmegi+AoCEMoue+cC7vnYVdmlbq/YLE/DWJX383oz2Ryq8anFrZ8jYvdoh8WI+dIugYL2SwRjmBoSwn56XIaot/QpMo3pYJIa4o8aZIZrjvB7BXO5aCDeMuZdUMT6AXGAGF1AeAWxFd2XIo1coR+OplMNDuYia8YAtnSTJ9JwGYWi2dJz3xrxsTQpBONf3yn8LVf8eH+o5eXc7lzCtHlDB+YyI8V9PyMsUPOeyvpB3rr9fDfNy263Zx33zTi5jldgP2OetUqGfbwl+0+zNYnrg64bluyIN/Awt1doDCQkCKpKXxuPaem/SyCHrKjg'
		EndIf

		Local $Symbol[] = ["jsmn_parse", "jsmn_init", "json_string_decode", "json_string_encode"]
		Local $CodeBase = _BinaryCall_Create($Code)
		If @error Then Exit MsgBox(16, "Json", "Startup Failure!")

		$SymbolList = _BinaryCall_SymbolList($CodeBase, $Symbol)
		If @error Then Exit MsgBox(16, "Json", "Startup Failure!")
	EndIf
	If $ProcName Then Return DllStructGetData($SymbolList, $ProcName)
EndFunc   ;==>__Jsmn_RuntimeLoader

Func Json_StringEncode_original($String, $Option = 0)
	Static $Json_StringEncode = __Jsmn_RuntimeLoader("json_string_encode")
	Local $Length = StringLen($String) * 6 + 1
	Local $Buffer = DllStructCreate("wchar[" & $Length & "]")
	Local $Ret = DllCallAddress("int:cdecl", $Json_StringEncode, "wstr", $String, "ptr", DllStructGetPtr($Buffer), "uint", $Length, "int", $Option)
	Return SetError($Ret[0], 0, DllStructGetData($Buffer, 1))
EndFunc   ;==>Json_StringEncode_original

; dexto - changed encode to catch Unicode edge case
Func Json_StringEncode($s, $Option = 0)
	Local Const $e = '[^[:print:]]'
	Local $r, $c, $u
	$s = StringRegExpReplace($s, '([\"\\])', '\\\0')
	$s = StringRegExpReplace($s, '\r', '\\r')
	$s = StringRegExpReplace($s, '\n', '\\n')
	$s = StringRegExpReplace($s, '\t', '\\t')
	Local $r = ''
	Local $a = StringRegExp($s, '.', 3)
	If Not @error Then
		Local $i, $c, $m = UBound($a) - 1
		For $i = 0 To $m
			$c = $a[$i]
			If StringRegExp($c, $e) Then
				$r &= '\u' & Hex(AscW($c), 4)
			Else
				$r &= $c
			EndIf
		Next
	Else
		$r = $s
	EndIf
	Return $r
EndFunc   ;==>Json_StringEncode

Func Json_StringDecode(Const ByRef $String)
	Static $Json_StringDecode = __Jsmn_RuntimeLoader("json_string_decode")
	Local $Length = StringLen($String) + 1
	Local $Buffer = DllStructCreate("wchar[" & $Length & "]")
	Local $Ret = DllCallAddress("int:cdecl", $Json_StringDecode, "wstr", $String, "ptr", DllStructGetPtr($Buffer), "uint", $Length)
	Return SetError($Ret[0], 0, DllStructGetData($Buffer, 1))
EndFunc   ;==>Json_StringDecode

Func Json_Decode(Const ByRef $json, $InitTokenCount = 1000)
	Static $Jsmn_Init = __Jsmn_RuntimeLoader("jsmn_init"), $Jsmn_Parse = __Jsmn_RuntimeLoader("jsmn_parse")
	Local $default[]
	If $json = "" Then Return SetError(0, 1, $default)
	Local $TokenList, $Ret
	Local $Parser = DllStructCreate("uint pos;int toknext;int toksuper")
	Do
		DllCallAddress("none:cdecl", $Jsmn_Init, "ptr", DllStructGetPtr($Parser))
		$TokenList = DllStructCreate("byte[" & ($InitTokenCount * 20) & "]")
		$Ret = DllCallAddress("int:cdecl", $Jsmn_Parse, "ptr", DllStructGetPtr($Parser), "wstr", $json, "ptr", DllStructGetPtr($TokenList), "uint", $InitTokenCount)
		$InitTokenCount *= 2
	Until $Ret[0] <> $JSMN_ERROR_NOMEM

	Local $Next = 0
	Return SetError($Ret[0], 0, _Json_Token($json, DllStructGetPtr($TokenList), $Next))
EndFunc   ;==>Json_Decode

Func _Json_Token(Const ByRef $json, $Ptr, ByRef $Next)
	If $Next = -1 Then Return Null

	Local $Token = DllStructCreate("int;int;int;int", $Ptr + ($Next * 20))
	Local $Type = DllStructGetData($Token, 1)
	Local $Start = DllStructGetData($Token, 2)
	Local $End = DllStructGetData($Token, 3)
	Local $Size = DllStructGetData($Token, 4)
	$Next += 1

	;ConsoleWrite($Type & @TAB & $Start & @TAB & $End & @TAB & $Size & @CRLF)

	If $Type = 0 And $Start = 0 And $End = 0 And $Size = 0 Then ; Null Item
		$Next = -1
		Return Null
	EndIf

	Switch $Type
		Case 0 ; Json_PRIMITIVE
			Local $Primitive = StringMid($json, $Start + 1, $End - $Start)
			Switch $Primitive
				Case "true"
					Return True
				Case "false"
					Return False
				Case "null"
					Return Null
				Case Else
					If StringRegExp($Primitive, "^[+\-0-9]") Then
						Return Number($Primitive)
					Else
						Return Json_StringDecode($Primitive)
					EndIf
			EndSwitch

		Case 1 ; Json_OBJECT
			Local $map[]
			Local $ky, $vl
			For $i = 0 To $Size - 1 Step 2
				$ky = _Json_Token($json, $Ptr, $Next)
				$vl = _Json_Token($json, $Ptr, $Next)
				If Not IsString($ky) Then $ky = Json_Encode($ky)

				$map[$ky] = $vl
			Next
			Return $map

		Case 2 ; Json_ARRAY
			Local $Array[$Size]
			For $i = 0 To $Size - 1
				$Array[$i] = _Json_Token($json, $Ptr, $Next)
			Next
			Return $Array

		Case 3 ; Json_STRING
			Local $t = StringMid($json, $Start + 1, $End - $Start)
			Return Json_StringDecode($t)
	EndSwitch
EndFunc   ;==>_Json_Token

Func Json_Encode_Compact($Data, $Option = 0)
	Select
		Case IsString($Data)
			Return '"' & Json_StringEncode($Data, $Option) & '"'

		Case IsNumber($Data)
			Return $Data

		Case IsArray($Data) And UBound($Data, 0) = 1
			Local $json = "["
			For $i = 0 To UBound($Data) - 1
				$json &= Json_Encode_Compact($Data[$i], $Option) & ","
			Next
			If StringRight($json, 1) = "," Then $json = StringTrimRight($json, 1)
			Return $json & "]"

		Case IsMap($Data)
			Local $json = "{"
			Local $Keys = MapKeys($Data)
			For $i = 0 To UBound($Keys) - 1
				;ConsoleWrite(MapExists($Data,$Keys[$i])&@tab&$Keys[$i] & @CRLF)
				;ConsoleWrite(VarGetType($Data[$Keys[$i]])&' = '&$Data[$Keys[$i]] & @CRLF)
				$json &= '"' & Json_StringEncode($Keys[$i], $Option) & '":' & Json_Encode_Compact($Data[$Keys[$i]], $Option) & ","
			Next
			If StringRight($json, 1) = "," Then $json = StringTrimRight($json, 1)
			Return $json & "}"

		Case IsBool($Data)
			Return StringLower($Data)

		Case IsPtr($Data)
			Return Number($Data)

		Case IsBinary($Data)
			Return '"' & Json_StringEncode(BinaryToString($Data, 4), $Option) & '"'

		Case Else ; Keyword, DllStruct, Object
			Return "null"
	EndSelect
EndFunc   ;==>Json_Encode_Compact

Func Json_Encode_Pretty(Const ByRef $Data, $Option, $Indent, $ArraySep, $ObjectSep, $ColonSep, $ArrayCRLF = Default, $ObjectCRLF = Default, $NextIdent = "")
	Local $ThisIdent = $NextIdent, $json = ""
	Select
		Case IsString($Data)
			Local $String = Json_StringEncode($Data, $Option)
			If BitAND($Option, $JSON_UNQUOTED_STRING) And Not BitAND($Option, $JSON_STRICT_PRINT) And Not StringRegExp($String, "[\s,:]") And Not StringRegExp($String, "^[+\-0-9]") Then
				Return $String
			Else
				Return '"' & $String & '"'
			EndIf

		Case IsArray($Data) And UBound($Data, 0) = 1
			If UBound($Data) = 0 Then Return "[]"
			If IsKeyword($ArrayCRLF) Then
				$ArrayCRLF = ""
				Local $Match = StringRegExp($ArraySep, "[\r\n]+$", 3)
				If IsArray($Match) Then $ArrayCRLF = $Match[0]
			EndIf

			If $ArrayCRLF Then $NextIdent &= $Indent
			Local $Length = UBound($Data) - 1
			For $i = 0 To $Length
				If $ArrayCRLF Then $json &= $NextIdent
				$json &= Json_Encode_Pretty($Data[$i], $Option, $Indent, $ArraySep, $ObjectSep, $ColonSep, $ArrayCRLF, $ObjectCRLF, $NextIdent)
				If $i < $Length Then $json &= $ArraySep
			Next

			If $ArrayCRLF Then Return "[" & $ArrayCRLF & $json & $ArrayCRLF & $ThisIdent & "]"
			Return "[" & $json & "]"

		Case IsMap($Data)
			If IsKeyword($ObjectCRLF) Then
				$ObjectCRLF = ""
				Local $Match = StringRegExp($ObjectSep, "[\r\n]+$", 3)
				If IsArray($Match) Then $ObjectCRLF = $Match[0]
			EndIf

			If $ObjectCRLF Then $NextIdent &= $Indent
			Local $Keys = MapKeys($Data)
			Local $Length = UBound($Keys) - 1
			If $Length = -1 Then Return "{}"
			For $i = 0 To $Length
				If $ObjectCRLF Then $json &= $NextIdent
				$json &= Json_Encode_Pretty(String($Keys[$i]), $Option, $Indent, $ArraySep, $ObjectSep, $ColonSep) & $ColonSep _
						 & Json_Encode_Pretty($Data[$Keys[$i]], $Option, $Indent, $ArraySep, $ObjectSep, $ColonSep, $ArrayCRLF, $ObjectCRLF, $NextIdent)
				If $i < $Length Then $json &= $ObjectSep
			Next

			If $ObjectCRLF Then Return "{" & $ObjectCRLF & $json & $ObjectCRLF & $ThisIdent & "}"
			Return "{" & $json & "}"

		Case Else
			Return Json_Encode_Compact($Data, $Option)

	EndSelect
EndFunc   ;==>Json_Encode_Pretty

Func Json_Encode(Const ByRef $Data, $Option = 0, $Indent = Default, $ArraySep = Default, $ObjectSep = Default, $ColonSep = Default)
	If BitAND($Option, $JSON_PRETTY_PRINT) Then
		Local $Strict = BitAND($Option, $JSON_STRICT_PRINT)

		If IsKeyword($Indent) Then
			$Indent = @TAB
		Else
			$Indent = Json_StringDecode($Indent)
			If StringRegExp($Indent, "[^\t ]") Then $Indent = @TAB
		EndIf

		If IsKeyword($ArraySep) Then
			$ArraySep = "," & @CRLF
		Else
			$ArraySep = Json_StringDecode($ArraySep)
			If $ArraySep = "" Or StringRegExp($ArraySep, "[^\s,]|,.*,") Or ($Strict And Not StringRegExp($ArraySep, ",")) Then $ArraySep = "," & @CRLF
		EndIf

		If IsKeyword($ObjectSep) Then
			$ObjectSep = "," & @CRLF
		Else
			$ObjectSep = Json_StringDecode($ObjectSep)
			If $ObjectSep = "" Or StringRegExp($ObjectSep, "[^\s,]|,.*,") Or ($Strict And Not StringRegExp($ObjectSep, ",")) Then $ObjectSep = "," & @CRLF
		EndIf

		If IsKeyword($ColonSep) Then
			$ColonSep = ": "
		Else
			$ColonSep = Json_StringDecode($ColonSep)
			If $ColonSep = "" Or StringRegExp($ColonSep, "[^\s,:]|[,:].*[,:]") Or ($Strict And (StringRegExp($ColonSep, ",") Or Not StringRegExp($ColonSep, ":"))) Then $ColonSep = ": "
		EndIf

		Return Json_Encode_Pretty($Data, $Option, $Indent, $ArraySep, $ObjectSep, $ColonSep)

	ElseIf BitAND($Option, $JSON_UNQUOTED_STRING) Then
		Return Json_Encode_Pretty($Data, $Option, "", ",", ",", ":")
	Else
		Return Json_Encode_Compact($Data, $Option)
	EndIf
EndFunc   ;==>Json_Encode






