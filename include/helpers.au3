

global $_DEBUG = False
global $_DEBUG_REST = False


Func llc($tt)
	If StringRegExp($tt, '[\r\n]+$') Then
		ConsoleWrite($tt)
	Else
		ConsoleWrite($tt & @CRLF)
	EndIf
EndFunc   ;==>llc

Func RegEx(ByRef $string, $pattern, $default = '', $element = 0, $separator = ',')
	Local $aa = StringRegExp($string, $pattern, 3)
	If @error Then
		If $default = -1 Then
			Return SetError(@error, 1, $string)
		Else
			Return SetError(@error, 1, $default)
		EndIf
	EndIf
	;ConsoleWrite(UBound($a) & @CRLF)
	If Not UBound($aa) Then
		If $default = -1 Then
			Return $string
		Else
			Return SetError(@error, 2, $default)
		EndIf
	EndIf
	If $element == -1 Then
		Local $out, $elm, $count
		For $elm In $aa
			$out &= $elm & $separator
			$count += 1
		Next
		Return SetError(0, $count, StringTrimRight($out, StringLen($separator)))
	EndIf
	If $element < 0 And ((UBound($aa) + $element) > -1) Then Return SetError(0, 0, $aa[UBound($aa) + $element])
	If Not ($element < UBound($aa)) Then
		If $default = -1 Then
			Return $string
		Else
			Return SetError(@error, 3, $default)
		EndIf
	EndIf
	Return SetError(0, 0, $aa[$element])
EndFunc   ;==>RegEx



Func regexLib($name, $ext = '')
	If $name = 'oneperline' Then
		SetError(0)
		SetExtended(10)
		Return '(?<=[\r\n]|^)[^#].+?(?=[\r\n]|$)'
	ElseIf $name = 'line' Then
		SetError(0)
		SetExtended(10)
		Return '(?<=[\r\n]|^)([^\r\n]+?)(?=[\r\n]|$)'
	ElseIf $name = 'tab' Then
		SetError(0)
		SetExtended(10)
		Return '(?<=\t|^)([^\t]+?)(?=\t|$)'
	ElseIf $name = 'nows' Then ; no white spaces traling or leading
		SetError(0)
		SetExtended(16)
		Local $ts = '(?s)(?:(?<=\t|^)"?(.*?)"?(?=\t|$))+'
		Return $ts
	ElseIf $name = 'csv' Then ; no white spaces traling or leading
		SetError(0)
		SetExtended(16)
		Local $ts = '(?s)((?<=\t|^|\r\n|\n])(?<!\\)".*?"(?=\t|$|\r\n|\n)|(?<=\t|^|\r\n|\n]).*?(?=\t|$|\r\n|\n))'
		Return $ts
	ElseIf $name = 'json' Then ; no white spaces traling or leading
		SetError(0)
		SetExtended(17)
		If $ext = '' Then $ext = '.+?(?<!\\)'
		Local $ts = '"' & $ext & '": *(?:".*?(?<!\\)"|[^\W"]+|\d+)'
		Return $ts
	ElseIf $name = 'jsonString' Then ; no white spaces traling or leading
		SetError(0)
		SetExtended(18)
		If $ext = '' Then $ext = '.+?(?<!\\)'
		Local $ts = '"' & $ext & '": *"(.*?(?<!\\))"'
		Return $ts
	ElseIf $name = 'jsonNum' Then ; no white spaces traling or leading
		SetError(0)
		SetExtended(19)
		If $ext = '' Then $ext = '.+?(?<!\\)'
		Local $ts = '"' & $ext & '": *(\w+)'
		Return $ts
	ElseIf $name = 'jsonKeys' Then ; no white spaces traling or leading
		SetError(0)
		SetExtended(20)
		Local $ts = '"([^"]+?)":'
		Return $ts
	Else
		SetError(1)
		SetExtended(0)
		Return ''
	EndIf
EndFunc   ;==>regexLib




Func csvEscape($csv, $clean = 0)
	$csv = StringRegExpReplace($csv, '[\r\n]', '')
	$csv = StringRegExpReplace($csv, '"', '""')
	If Not $clean Or StringRegExp($csv, ',') Then
		Return '"' & $csv & '"'
	EndIf
	Return $csv
EndFunc   ;==>csvEscape


Func FileWriteW($filename, $data = '')
	Local $er = 1
	Local $rc
	Local $fh = FileOpen($filename, 2 + 8)
	If $fh == -1 Then
		ConsoleWrite('Can not open "' & $filename & '" for writing.' & @CRLF)
	EndIf
	If $data Then
		$rc = FileWrite($fh, $data)
		If $rc = 1 Then
			$er = 0
			FileFlush($fh)
		Else
			ConsoleWrite('Failed to write the file "' & $filename & '"' & @CRLF)
		EndIf
	EndIf
	FileClose($fh)
	Return SetError($er, 0, $rc)
EndFunc   ;==>FileWriteW



Func RESTW($server, $port, $verb, $path, $headers, $data, $getHeaders = False, $redirect = False, $ignorecert = False, $user = '', $pass = '', $isBinary = False, $forceSSL = Null)

	If IsDeclared('_RESTW_COUNTER') Then Assign('_RESTW_COUNTER', Eval('_RESTW_COUNTER') + 1, 2)

	If $_DEBUG Then $ll('> RESTW: ' & $server & ' ' & $port & ' ' & $verb & ' ' & $path)
	SetError(0)
	Local $hw_open

	$hw_open = _WinHttpOpen('Mozilla/5.0 (Windows NT 6.3; WOW64; Trident/7.0; rv:11.0) like Gecko', $WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY)

	If @error Then
		If $_DEBUG_REST Then $ll("ERROR: _WinHttpOpen()")
		_WinHttpCloseHandle($hw_open)
		Return SetError(1, 0, 0)
	EndIf

	_WinHttpSetTimeouts($hw_open, 2000, 10000, 120000, 180000)
	;_WinHttpSetTimeouts($hw_open, 1000, 1000, 1000, 1000)

	; disable passthough Auth
	;_WinHttpSetOption($hw_open, $WINHTTP_OPTION_DISABLE_FEATURE, $WINHTTP_DISABLE_AUTHENTICATION)
	If $user = 'pass' Then
		_WinHttpSetOption($hw_open, $WINHTTP_OPTION_AUTOLOGON_POLICY, $WINHTTP_AUTOLOGON_SECURITY_LEVEL_LOW)
	EndIf

	If $redirect Then
		_WinHttpSetOption($hw_open, $WINHTTP_OPTION_REDIRECT_POLICY, $WINHTTP_OPTION_REDIRECT_POLICY_ALWAYS)
	Else
		_WinHttpSetOption($hw_open, $WINHTTP_OPTION_REDIRECT_POLICY, $WINHTTP_OPTION_REDIRECT_POLICY_NEVER)
	EndIf
	If @error Then
		If $_DEBUG_REST Then $ll("ERROR: _WinHttpSetOption() WINHTTP_OPTION_REDIRECT_POLICY")
		_WinHttpCloseHandle($hw_open)
		Return SetError(2, 0, 0)
	EndIf

	Local $hw_connect = _WinHttpConnect($hw_open, $server, $port)
	If @error Then
		If $_DEBUG_REST Then $ll("ERROR: _WinHttpConnect() FAILED")
		_WinHttpCloseHandle($hw_connect)
		_WinHttpCloseHandle($hw_open)
		Return SetError(3, 0, 0)
	EndIf

	; @@@ OPEN REQUEST @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
	Local $h_openRequest

	Local $isSSL = False
	If $port == 443 Or $forceSSL == True Then
		$isSSL = True
	EndIf
	If Not $isSSL Then
		If $_DEBUG_REST Then $ll("INFO: _WinHttpConnect() not SSL")
		$h_openRequest = _WinHttpOpenRequest($hw_connect, $verb, $path, "HTTP/1.1", $WINHTTP_NO_REFERER, Default, $WINHTTP_FLAG_ESCAPE_DISABLE)
	Else
		If $_DEBUG_REST Then $ll("INFO: _WinHttpConnect() $WINHTTP_FLAG_SECURE")
		$h_openRequest = _WinHttpOpenRequest($hw_connect, $verb, $path, "HTTP/1.1", $WINHTTP_NO_REFERER, Default, $WINHTTP_FLAG_SECURE)
	EndIf
	If @error Then
		If $_DEBUG_REST Then $ll("ERROR: _WinHttpOpenRequest() FAILED")
		_WinHttpCloseHandle($h_openRequest)
		_WinHttpCloseHandle($hw_connect)
		_WinHttpCloseHandle($hw_open)
		Return SetError(4, 0, 0)
	EndIf

	_WinHttpSetOption($h_openRequest, $WINHTTP_OPTION_DECOMPRESSION, $WINHTTP_DECOMPRESSION_FLAG_ALL)
	_WinHttpSetOption($h_openRequest, $WINHTTP_OPTION_UNSAFE_HEADER_PARSING, 1)

	If $isSSL Then
		$WINHTTP_FLAG_SECURE_PROTOCOL_SSL2 = 0x00000008
		$WINHTTP_FLAG_SECURE_PROTOCOL_SSL3 = 0x00000020
		$WINHTTP_FLAG_SECURE_PROTOCOL_TLS1 = 0x00000080
		$WINHTTP_FLAG_SECURE_PROTOCOL_ALL = BitOR($WINHTTP_FLAG_SECURE_PROTOCOL_TLS1, $WINHTTP_FLAG_SECURE_PROTOCOL_SSL3, $WINHTTP_FLAG_SECURE_PROTOCOL_SSL2)

		_WinHttpSetOption($h_openRequest, $WINHTTP_OPTION_SECURE_PROTOCOLS, $WINHTTP_FLAG_SECURE_PROTOCOL_ALL)
	EndIf


	If $isSSL And $ignorecert Then
		If $_DEBUG_REST Then $ll('$WINHTTP_OPTION_SECURITY_FLAGS = 0x3300')
		;_WinHttpSetOption($hw_open, $WINHTTP_OPTION_SECURITY_FLAGS, BitOR($SECURITY_FLAG_IGNORE_CERT_CN_INVALID, $SECURITY_FLAG_IGNORE_UNKNOWN_CA))
		_WinHttpSetOption($h_openRequest, $WINHTTP_OPTION_SECURITY_FLAGS, BitOR($SECURITY_FLAG_IGNORE_UNKNOWN_CA, $SECURITY_FLAG_IGNORE_CERT_DATE_INVALID, $SECURITY_FLAG_IGNORE_CERT_CN_INVALID, $SECURITY_FLAG_IGNORE_CERT_WRONG_USAGE))
		If @error Then
			If $_DEBUG_REST Then $ll("ERROR " & @error & ": _WinHttpSetOption() WINHTTP_OPTION_SECURITY_FLAGS")
			_WinHttpCloseHandle($hw_open)
			Return SetError(2, 0, 0)
		EndIf
	EndIf

	If Eval('_WinHttpSetTimeouts_value') Then
		_WinHttpSetTimeouts($h_openRequest, Default, Default, Number(Eval('_WinHttpSetTimeouts_value')), Number(Eval('_WinHttpSetTimeouts_value')))
	Else
		;_WinHttpSetTimeouts($h_openRequest, Default, Default, Default, Default)
	EndIf

	If $user And $pass Then
		If $_DEBUG_REST Then $ll("INFO: SETTING AUTH NEGOTIATE CREDS")
		_WinHttpSetCredentials($h_openRequest, $WINHTTP_AUTH_TARGET_SERVER, $WINHTTP_AUTH_SCHEME_NEGOTIATE, $user, $pass)
	EndIf

	; @@@ SEND @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
	Local $bdata
	If IsBinary($data) Then
		$bdata = $data
		$data = ''
	EndIf

	If $_DEBUG_REST Then $ll('Sending...' & @CRLF)
	If $_DEBUG_REST Then $ll($headers & @CRLF)
	If $_DEBUG_REST Then $ll($data & @CRLF)
	_WinHttpSendRequest($h_openRequest, $headers, $data)
	Local $er = @error
	If $er Then

		#CS
			SECURITY_FLAG_IGNORE_UNKNOWN_CA |
			SECURITY_FLAG_IGNORE_CERT_WRONG_USAGE |
			SECURITY_FLAG_IGNORE_CERT_CN_INVALID |
			SECURITY_FLAG_IGNORE_CERT_DATE_INVALID
		#CE

		;_WinHttpSendRequest($h_openRequest, $headers, $data)

		If $_DEBUG_REST Then $ll('Error: _WinHttpSendRequest() [' & $er & '] ' & Eval('ERROR_WINHTTP_' & $er))
		_WinHttpCloseHandle($h_openRequest)
		_WinHttpCloseHandle($hw_connect)
		_WinHttpCloseHandle($hw_open)
		Return SetError(5, 0, 0)
	EndIf

	If $bdata Then
		_WinHttpWriteData($h_openRequest, $bdata, 1)
		Local $er = @error
		If $er Then
			If $_DEBUG_REST Then $ll('Error: _WinHttpWriteData() [' & $er & '] ' & Eval('ERROR_WINHTTP_' & $er))
			_WinHttpCloseHandle($h_openRequest)
			_WinHttpCloseHandle($hw_connect)
			_WinHttpCloseHandle($hw_open)
			Return SetError(6, 0, 0)
		EndIf
	EndIf

	If $_DEBUG_REST Then $ll('Receiving...')
	Local $retries = 3
	Local $try = 0
	While $try < $retries
		_WinHttpReceiveResponse($h_openRequest)
		Local $er = @error
		If $er Then
			If $_DEBUG_REST Then $ll('Error: _WinHttpReceiveResponse() [' & $er & '] ' & Eval('ERROR_WINHTTP_' & $er))
		Else
			ExitLoop
		EndIf
		$try += 1
	WEnd
	If Not ($try < $retries) Then
		If $_DEBUG_REST Then $ll("ERROR: _WinHttpReceiveResponse(): Closing handle.")
		_WinHttpCloseHandle($h_openRequest)
		_WinHttpCloseHandle($hw_connect)
		_WinHttpCloseHandle($hw_open)
		Return SetError(7, 0, 0)
	EndIf

	If $_DEBUG_REST Then
		$header = _WinHttpQueryHeaders($h_openRequest, $WINHTTP_QUERY_FLAG_REQUEST_HEADERS + $WINHTTP_QUERY_RAW_HEADERS_CRLF)
		$ll('------------ REQUEST HEADERS -----------' & @CRLF & '; ' & $server & ':' & $port & @CRLF & _WinHttpQueryHeaders($h_openRequest, $WINHTTP_QUERY_FLAG_REQUEST_HEADERS + $WINHTTP_QUERY_RAW_HEADERS_CRLF))
		$ll('------------ SENT DATA -----------------' & @CRLF & $data)
		$ll('------------ RESPONSE HEADERS ----------' & @CRLF & _WinHttpQueryHeaders($h_openRequest))
	EndIf

	Local $dataAvailable = _WinHttpQueryDataAvailable($h_openRequest)
	Local $er = @error
	Local $ex = @extended
	If $er Then
		If $_DEBUG_REST Then $ll("ERROR: _WinHttpQueryDataAvailable() FAILED")
		_WinHttpCloseHandle($h_openRequest)
		_WinHttpCloseHandle($hw_connect)
		_WinHttpCloseHandle($hw_open)
		Return SetError(8, 0, 0)
	EndIf
	If $_DEBUG Then $ll('> RESTW: $dataAvailable: ' & $dataAvailable)
	Local $got
	If $dataAvailable Then
		If Not $isBinary Then
			Do
				$got &= _WinHttpReadData($h_openRequest, 1)
			Until @error <> 0
		Else
			Local $chunk
			While True
				$chunk = _WinHttpReadData($h_openRequest, 2, 250 * 1024)
				If @error Then ExitLoop
				$got = _WinHttpSimpleBinaryConcat($got, $chunk)
			WEnd
		EndIf
	EndIf
	If $_DEBUG_REST Then $ll('------------ RESPONSE DATA -------------' & @CRLF & $got & @CRLF & '----------------------------------------')

	If $getHeaders Then
		$REST_headers = _WinHttpQueryHeaders($h_openRequest)
	EndIf

	Local $status = Int(_WinHttpQueryHeaders($h_openRequest, $WINHTTP_QUERY_STATUS_CODE))
	If $_DEBUG Then $ll('> RESTW: Status: ' & $status)

	_WinHttpCloseHandle($h_openRequest)
	_WinHttpCloseHandle($hw_connect)
	_WinHttpCloseHandle($hw_open)
	Return SetError(0, $status, $got)
EndFunc   ;==>RESTW



Func _Base64Encode($binary, $iFlags = 0x00000001)
	$binary = Binary($binary)
	Local $tByteArray = DllStructCreate('byte[' & BinaryLen($binary) & ']')
	DllStructSetData($tByteArray, 1, $binary)
	Local $aSize = DllCall("Crypt32.dll", "bool", 'CryptBinaryToString', 'struct*', $tByteArray, 'dword', BinaryLen($binary), 'dword', $iFlags, 'str', Null, 'dword*', Null)
	Local $tOutput = DllStructCreate('char[' & $aSize[5] & ']')
	Local $aEncode = DllCall("Crypt32.dll", "bool", 'CryptBinaryToString', 'struct*', $tByteArray, 'dword', $aSize[2], 'dword', $iFlags, 'struct*', $tOutput, 'dword*', $aSize[5])
	If @error Or (Not $aEncode[0]) Then Return SetError(1, 0, 0)
	$tOutput = DllStructGetData($tOutput, 1)
	$tOutput = StringRegExpReplace($tOutput, '\s', '')
	Return $tOutput
EndFunc   ;==>_Base64Encode


Func _URIEncode($sData, $space = 0, $safe = '')
	Local $aData = StringSplit(BinaryToString(StringToBinary($sData, 4), 1), "")
	Local $aSafe = StringSplit($safe, "")
	Local $nChar
	$sData = ""
	For $i = 1 To $aData[0]

		$nChar = Asc($aData[$i])
		Local $done = 0
		For $sf = 1 To $aSafe[0]
			;ConsoleWrite($aData[$i]&@tab&Asc($aData[$i])&@tab&Asc($aSafe[$sf]) & @CRLF)
			If Asc($aSafe[$sf]) == Asc($aData[$i]) Then
				$sData &= $aData[$i]
				$done = 1
			EndIf
		Next

		If $done Then ContinueLoop

		Switch $nChar
			Case 45, 46, 48 To 57, 65 To 90, 95, 97 To 122, 126
				$sData &= $aData[$i]
			Case 32
				If $space Then
					$sData &= "+"
				Else
					$sData &= "%" & Hex($nChar, 2)
				EndIf
			Case Else
				$sData &= "%" & Hex($nChar, 2)
		EndSwitch
	Next
	Return $sData
EndFunc   ;==>_URIEncode