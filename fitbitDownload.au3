#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Version=Beta
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Change2CUI=y
#AutoIt3Wrapper_Res_Comment=fitbitSave
#AutoIt3Wrapper_Res_Description=fitbitSave
#AutoIt3Wrapper_Res_Fileversion=0.0.0.78
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_LegalCopyright=fitbitSave
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_Run_After=del "%scriptdir%\%scriptfile%_stripped.au3"
#AutoIt3Wrapper_Run_After="%scriptdir%\Public_RepoPush.exe"
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/sf /mi 8
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#CS
	FitBit API Docs:
	https://dev.fitbit.com/reference/web-api/explore/#/Devices/alarms2
	https://dev.fitbit.com/reference/web-api/activity/
#CE


Global $ll = llc

#include ".\include\WinHTTP\WinHttp.au3"
#include ".\include\mapLib.au3"
#include ".\include\helpers.au3"
#include <Date.au3>


Global $_DEBUG = False
Global $_DEBUG_REST = False
Global $_DEBUG_WRITE = False
Global $REST_headers

Global $blankMap[]

Global $ini = @ScriptDir & '\FitBit.ini'
Global $jfile = @ScriptDir & '\fitbit.json'
Global $dataFolder = IniRead($ini, 'main', 'FitBitDataFolder', @ScriptDir & '\FitBitData')

Global $jmap = fileReadJson($jfile)
Global $clientID = IniRead($ini, 'FitBitPrivateAppInfo', 'clientID', '')
If Not $clientID Then
	MsgBox(0, @ScriptName, 'clientID was not found in [privateApp] ' & $ini)
	Exit
EndIf
Global $clientSecret = IniRead($ini, 'FitBitPrivateAppInfo', 'clientSecret', '')
If Not $clientID Then
	MsgBox(0, @ScriptName, 'clientSecret was not found in [privateApp] ' & $ini)
	Exit
EndIf
Global $callbackURL = IniRead($ini, 'FitBitPrivateAppInfo', 'callbackURL', '')
If Not $clientID Then
	MsgBox(0, @ScriptName, 'callbackURL was not found in [privateApp] ' & $ini)
	Exit
EndIf

Global $authURL = 'https://www.fitbit.com/oauth2/authorize'
Global $tokenURL = 'https://api.fitbit.com/oauth2/token'


If Not (MapExists($jmap, 'token') And $jmap['token']) Then
	tokenGet()
EndIf

Global $DownloadOnceADay = (Int(IniRead($ini, 'main', 'DownloadOnceADay', 1)) <> 0)

If Int(IniRead($ini, 'main', 'DownloadGeneralData', 1)) Then updateGeneral()
If Int(IniRead($ini, 'main', 'DownloadSleepLog', 1)) Then syncSleep()
If Int(IniRead($ini, 'main', 'DownloadActivityLog', 1)) Then syncActivites()
If Int(IniRead($ini, 'main', 'DownloadTimeSeries', 1)) Then updateTimeSeries()

ConsoleWrite(@CRLF)
ConsoleWrite('Done' & @CRLF)
Sleep(10000)
Exit






Func updateGeneral()
	Local $todaysDate = _DateAdd('D', 0, _NowCalcDate())
	If $DownloadOnceADay And mapRead($jmap, 'lastDateGeneralUpdate') = $todaysDate Then
		ConsoleWrite('General data was already downloaded today.' & @CRLF)
		ConsoleWrite(@CRLF)
		Return
	EndIf

	ConsoleWrite('Refreshing general stats...' & @CRLF)
	Local $FitBitGeneralDataFile = $dataFolder & '\FitBitGeneralData.json'
	Local $map[]

	Local $arr = IniReadSection($ini, 'GeneralDataToDownload')
	If @error Then
		ConsoleWrite('Did not find section "GeneralDataToDownload" in "' & $ini & '"' & @CRLF)
		Return
	EndIf
	For $i = 1 To $arr[0][0]
		Local $key = $arr[$i][0]
		Local $url = StringReplace($arr[$i][1], '%date%', $todaysDate)
		ConsoleWrite('Downloading ' & $key & @CRLF)
		FitBitW($map, $url, 'general', $key)
	Next

	fileWriteJson($FitBitGeneralDataFile, $map)

	$jmap['lastDateGeneralUpdate'] = $todaysDate
	fileWriteJson($jfile, $jmap)
	ConsoleWrite(@CRLF)
EndFunc   ;==>updateGeneral


Func syncSleep($iter = 0)
	Local $todaysDate = _DateAdd('D', 0, _NowCalcDate())
	If $iter = 0 And $DownloadOnceADay And mapRead($jmap, 'lastDateSleepUpdate') = $todaysDate Then
		ConsoleWrite('Sleep data was already downloaded today.' & @CRLF)
		ConsoleWrite(@CRLF)
		Return
	EndIf

	ConsoleWrite('Reading downloaded list of sleep entries...' & @CRLF)
	Local $afile = $dataFolder & '\FitBitSleepLog.json'
	Local $amap = fileReadJson($afile)
	Local $arr

	Local $offset = 0
	If MapExists($amap, 'sleep') Then
		$offset = UBound($amap['sleep'])
		ConsoleWrite('Number of stored sleep entries: ' & $offset & @CRLF)
	Else
		ConsoleWrite('Creating new sleep data file.' & @CRLF)
		Local $tarr[0]
		$amap['sleep'] = $tarr
	EndIf
	ConsoleWrite('Downloading a list of sleep entries...' & @CRLF)
	Local $beforeDate = StringRegExpReplace(_DateAdd('d', 1, _NowCalcDate()), '/', '-')
	Local $tmap = FitBitGet('https://api.fitbit.com/1.2/user/-/sleep/list.json?beforeDate=' & $beforeDate & '&sort=asc&offset=' & $offset & '&limit=100')
	If Not MapExists($tmap, 'sleep') Then
		ConsoleWrite('Could not get a list of sleep entries.' & @CRLF)
	Else
		$arr = $tmap['sleep']
		If UBound($arr) < 1 Then
			Local $amapLen = UBound($amap['sleep'])
			If Not $amapLen Then
				ConsoleWrite('No sleep entries found.' & @CRLF)
			Else
				ConsoleWrite("There are " & $amapLen & " sleep entries downloaded." & @CRLF)
				$amapLen -= 1
				ConsoleWrite('No new sleep entries found since ' & (($amap['sleep'])[$amapLen])['lastModified'] & @CRLF)
			EndIf
		Else
			ConsoleWrite('Downloaded additional ' & UBound($arr) & ' sleep entries.' & @CRLF)
			_ArrayConcatenate($amap['sleep'], $arr)
			fileWriteJson($afile, $amap)
			syncSleep($iter + 1)
		EndIf
	EndIf


	#CS
		$arr = $amap['sleep']
		For $ii = 0 To UBound($arr) - 1
		Local $item = $arr[$ii]
		Local $acted = False
		Local $titleLong = StringReplace(StringReplace(StringRegExpReplace($item['originalStartTime'], '\.\d\d\d-.+', ''), 'T', ' '), ':', '-') & ' - ' & $item['activityName']
		Local $titleShort = $item['activityName'] & ' from ' & regex($item['originalStartTime'], '^(.{10})')
		ConsoleWrite('Checking the data for a ' & $titleShort & @CRLF)
		If Not MapExists($item, 'caloriesData') Then
		ConsoleWrite('   Getting calorie data...' & @CRLF)
		$item['caloriesData'] = FitBitGet($item['caloriesLink'])
		$arr[$ii] = $item
		$acted = True
		EndIf
		If Not MapExists($item, 'heartRateData') Then
		ConsoleWrite('   Getting heart rate data...' & @CRLF)
		$item['heartRateData'] = FitBitGet($item['heartRateLink'])
		$arr[$ii] = $item
		$acted = True
		EndIf
		If Not MapExists($item, 'tcxData') Then
		ConsoleWrite('   Getting TCX data...' & @CRLF)
		$item['tcxData'] = FitBitGet($item['tcxLink'])
		$arr[$ii] = $item
		$acted = True
		EndIf
		If StringLen($item['tcxData']) > 600 Then
		Local $tcxFile = $dataFolder & '\TCXfiles\' & $titleLong & '.tcx'
		If Not FileExists($tcxFile) Then
		FileWriteW($tcxFile, $item['tcxData'])
		EndIf
		EndIf
		If $acted Then
		ConsoleWrite(@CRLF)
		$amap['activities'] = $arr
		fileWriteJson($afile, $amap)
		EndIf
		Next
	#CE


	$jmap['lastDateSleepUpdate'] = $todaysDate
	fileWriteJson($jfile, $jmap)
	ConsoleWrite(@CRLF)
	Return
EndFunc   ;==>syncSleep

Func syncActivites()
	Local $todaysDate = _DateAdd('D', 0, _NowCalcDate())
	If $DownloadOnceADay And mapRead($jmap, 'lastDateActivitesUpdate') = $todaysDate Then
		ConsoleWrite('Activity data was already downloaded today.' & @CRLF)
		ConsoleWrite(@CRLF)
		Return
	EndIf

	ConsoleWrite('Reading downloaded list of activities...' & @CRLF)
	Local $afile = $dataFolder & '\FitBitActivityLog.json'
	Local $amap = fileReadJson($afile)
	Local $arr

	Local $offset = 0
	If MapExists($amap, 'activities') Then
		$offset = UBound($amap['activities'])
		ConsoleWrite('Number of stored activities: ' & $offset & @CRLF)
	Else
		ConsoleWrite('Creating new activities data file.' & @CRLF)
		Local $tarr[0]
		$amap['activities'] = $tarr
	EndIf
	ConsoleWrite('Downloading a list of activities...' & @CRLF)
	Local $beforeDate = StringRegExpReplace(_DateAdd('d', 1, _NowCalcDate()), '/', '-')
	Local $tmap = FitBitGet('https://api.fitbit.com/1/user/-/activities/list.json?beforeDate=' & $beforeDate & '&sort=asc&offset=' & $offset & '&limit=100')
	If Not MapExists($tmap, 'activities') Then
		ConsoleWrite('Could not get a list of activities.' & @CRLF)
	Else
		$arr = $tmap['activities']
		If UBound($arr) < 1 Then
			Local $amapLen = UBound($amap['activities'])
			If Not $amapLen Then
				ConsoleWrite('No activities found.' & @CRLF)
			Else
				ConsoleWrite("There are " & $amapLen & " activities downloaded." & @CRLF)
				$amapLen -= 1
				ConsoleWrite('No new activities found since ' & (($amap['activities'])[$amapLen])['lastModified'] & @CRLF)
			EndIf
		Else
			ConsoleWrite('Downloaded additional ' & UBound($arr) & ' activities.' & @CRLF)
			_ArrayConcatenate($amap['activities'], $arr)
			fileWriteJson($afile, $amap)
		EndIf
	EndIf

	$arr = $amap['activities']
	For $ii = 0 To UBound($arr) - 1
		Local $item = $arr[$ii]
		Local $acted = False
		Local $titleLong = StringReplace(StringReplace(StringRegExpReplace($item['originalStartTime'], '\.\d\d\d-.+', ''), 'T', ' '), ':', '-') & ' - ' & $item['activityName']
		Local $titleShort = $item['activityName'] & ' from ' & regex($item['originalStartTime'], '^(.{10})')
		ConsoleWrite('Checking the data for a ' & $titleShort & @CRLF)
		If Not MapExists($item, 'caloriesData') Then
			ConsoleWrite('   Getting calorie data...' & @CRLF)
			$item['caloriesData'] = FitBitGet($item['caloriesLink'])
			$arr[$ii] = $item
			$acted = True
		EndIf
		If Not MapExists($item, 'heartRateData') Then
			ConsoleWrite('   Getting heart rate data...' & @CRLF)
			$item['heartRateData'] = FitBitGet($item['heartRateLink'])
			$arr[$ii] = $item
			$acted = True
		EndIf
		If Not MapExists($item, 'tcxData') Then
			ConsoleWrite('   Getting TCX data...' & @CRLF)
			$item['tcxData'] = FitBitGet($item['tcxLink'])
			$arr[$ii] = $item
			$acted = True
		EndIf
		If StringLen($item['tcxData']) > 600 Then
			Local $tcxFile = $dataFolder & '\TCXfiles\' & $titleLong & '.tcx'
			If Not FileExists($tcxFile) Then
				FileWriteW($tcxFile, $item['tcxData'])
			EndIf
		EndIf
		If $acted Then
			ConsoleWrite(@CRLF)
			$amap['activities'] = $arr
			fileWriteJson($afile, $amap)
		EndIf
	Next

	$jmap['lastDateActivitesUpdate'] = $todaysDate
	fileWriteJson($jfile, $jmap)
	ConsoleWrite(@CRLF)
EndFunc   ;==>syncActivites


Func updateTimeSeries()

	ConsoleWrite('Updating FitBit time series data...' & @CRLF)
	Local $jsonDataFolder = $dataFolder & '\TimeSeriesData'
	Local $lastDate = mapRead($jmap, 'lastDateTimeSeries')
	If Not $lastDate Then
		Local $firstDate = StringRegExpReplace(IniRead($ini, 'main', 'lastDateTimeSeries', ''), '-', '/')
		If Not $firstDate Then
			$firstDate = FitBitGet('https://api.fitbit.com/1/user/-/profile.json')['user']['memberSince']
			ConsoleWrite('FitBit member since: ' & $firstDate & @CRLF)
		EndIf
		$lastDate = $firstDate
	EndIf
	ConsoleWrite('Date of the last completed session: ' & $lastDate & @CRLF)

	Local $dateDiff = -1 * _DateDiff('D', $lastDate, _NowCalcDate())
	ConsoleWrite('Days to process: ' & Abs($dateDiff) & @CRLF)

	IniReadSection($ini, 'TimeSeriesToDownload')
	If Not @error Then
		For $ii = $dateDiff To 0
			If FitBitDayW($ii, $jsonDataFolder) Then
				$jmap['lastDateTimeSeries'] = _DateAdd('D', $ii, _NowCalcDate())
				fileWriteJson($jfile, $jmap)
			EndIf
			ConsoleWrite(@CRLF)
		Next
	Else
		ConsoleWrite('Did not find section TimeSeriesToDownload in ' & $ini & @CRLF)
	EndIf
	ConsoleWrite(@CRLF)
EndFunc   ;==>updateTimeSeries

Func FitBitDayW($dayOffset, $jsonDataFolder)
	Local $date = StringReplace(_DateAdd('D', $dayOffset, _NowCalcDate()), '/', '-')
	ConsoleWrite('Loading data for ' & $date & @CRLF)
	Local $jdata = $jsonDataFolder & '\' & $date & '.json'

	Local $lmap = fileReadJson($jdata)
	Local $verifyOnly = (mapRead($lmap, $date, 'complete') = 'True')
	Local $updated = False

	If $verifyOnly Then
		ConsoleWrite('Data for ' & $date & ' was already downloaded.' & @CRLF)
		ConsoleWrite('Checking the data...' & @CRLF)
	Else
		ConsoleWrite('Downloading the data for the date: ' & $date & @CRLF)
	EndIf

	Local $arr = IniReadSection($ini, 'TimeSeriesToDownload')
	If @error Then
		ConsoleWrite('Did not find section "TimeSeriesToDownload" in "' & $ini & '"' & @CRLF)
		Return SetError(1, 0, 0)
	EndIf
	For $i = 1 To $arr[0][0]
		Local $key = $arr[$i][0]
		If $verifyOnly And MapExists($lmap[$date], $key) Then
			ConsoleWrite('Verified ' & $date & ' - ' & $key & @CRLF)
		Else
			ConsoleWrite('Downloading ' & $date & ' - ' & $key & @CRLF)
			Local $url = StringReplace($arr[$i][1], '%date%', $date)
			FitBitW($lmap, $url, $date, $key)
			$updated = True
		EndIf
	Next

	If ($dayOffset < 0) Then
		mapWrite($lmap, $date, 'complete', 'True')
	EndIf

	If $updated Then
		ConsoleWrite('Saving the ' & $date & ' data to: ' & $jdata & @CRLF)
		fileWriteJson($jdata, $lmap)
		Return SetError(0, 0, 1)
	Else
		Return SetError(0, 1, 1)
	EndIf

EndFunc   ;==>FitBitDayW







Func FitBitW(ByRef $map, $url, $date, $act)
	Local $complete = (mapRead($map, $date, 'complete') = 'True')
	If $complete And MapExists($map[$date], $act) Then
		Return SetError(0, 200, $map[$date][$act])
	EndIf
	Local $got = FitBitGet($url)
	Local $rc = @extended
	If Not MapExists($map, $date) Then $map[$date] = $blankMap
	If Not MapExists($map[$date], $act) Then $map[$date][$act] = $blankMap
	$map[$date][$act] = $got
	Return SetError(0, $rc, $got)
EndFunc   ;==>FitBitW

Func FitBitGet($url)
	Local $dat = REST($url, 'GET')
	Local $rc = @extended
	;ConsoleWrite('DATA [' & $rc & ']: ' & regex($dat, '([^\r\n]{0,120})') & '...' & @CRLF)
	If ($rc = 429) Then
		Local $retryAfterSec = Int(regex($REST_headers, 'Fitbit-Rate-Limit-Reset: (\d+)', 3600)) + 61
		ConsoleWrite(@CRLF)
		ConsoleWrite('Too many requests. Have to wait till: ' & _DateAdd('s', $retryAfterSec, _NowCalc()) & @CRLF)
		If Int(IniRead($ini, 'main', 'waitOutLimit', 0)) Then
			ConsoleWrite('Since you have waitOutLimit enabled, wating...' & @CRLF)
			Sleep($retryAfterSec * 1000)
			ConsoleWrite('Resuming API calls...' & @CRLF)
			ConsoleWrite(@CRLF)
			Local $tmp = FitBitGet($url)
			Return SetError(@error, @extended, $tmp)
		EndIf
	ElseIf StringLeft($dat, 24) = '{"errors":[{"errorType":' Then
		ConsoleWrite('ERROR [' & $rc & ']:' & @CRLF)
		jsonPrint(json_decode($dat))
		MsgBox(0, @ScriptName, 'ERROR [' & $rc & ']:' & @CRLF & $dat, 5)
		Exit
	EndIf
	If StringRegExp($REST_headers, '(?i)Content-Type: application/json') Then
		Return SetError(0, $rc, json_decode($dat))
	Else
		Return SetError(0, $rc, $dat)
	EndIf
EndFunc   ;==>FitBitGet




Func REST($url, $type = '', $data = '', $headers = '')
	If $_DEBUG Then $ll('REST() ' & $headers)
	If $type == '' Then
		If $data == '' Then
			$type = 'GET'
		Else
			$type = 'POST'
		EndIf
	EndIf
	Local $heads
	If Not $headers Then
		$heads = "User-Agent: Mozilla/5.0 (Windows NT 6.3; WOW64; Trident/7.0; rv:11.0) like Gecko"
		$heads &= @CRLF & 'Authorization: Bearer ' & $jmap['token'] & @CRLF & "Host: api.fitbit.com" & @CRLF & "X-Target-URI: https://api.fitbit.com"
	Else
		$heads = $headers
	EndIf
	Local $port = 80
	If StringRegExp($url, '^https://') Then $port = 443
	Local $server = RegEx($url, 'https?://([^/]+)', '')
	Local $path = RegEx($url, 'https?://[^/]+(.+)', '')
	Local $got = RESTW($server, $port, $type, $path, $heads, $data, True, True, False)
	Local $rc = @extended

	If $rc = 401 Then
		tokenGet()
		Local $tmp = REST($url, $type, $data, $headers)
		Return SetError(@extended, @extended, $tmp)
	EndIf

	Return SetError($rc, $rc, $got)
EndFunc   ;==>REST






Func tokenGet($iter = 0)
	ConsoleWrite('tokenGet($iter=' & $iter & ')' & @CRLF)
	Sleep(3000)
	If Int(IniRead($ini, 'main', 'getAuthorizationCode', 0)) Then
		If $iter > 2 Then
			ConsoleWrite('Failed getting a token...' & @CRLF)
			MsgBox(0, @ScriptName, 'Failed getting a token.', 5)
			Exit
		EndIf

		If Not (MapExists($jmap, 'code') And $jmap['code']) Then
			Local $code = codeGet()
			If $code Then
				ConsoleWrite('Got the code.' & @CRLF)
				$jmap['code'] = $code
			Else
				MsgBox(0, @ScriptName, 'Could not get a code, exiting.', 5)
				Exit
			EndIf
		EndIf

		Local $heads, $data, $dat
		$heads &= 'Authorization: Basic ' & _Base64Encode($clientID & ':' & $clientSecret) & @CRLF
		$heads &= 'Content-Type: application/x-www-form-urlencoded'
		If Not MapExists($jmap, 'refresh_token') Then
			ConsoleWrite('Getting a token...' & @CRLF)
			$data &= 'clientId=' & $clientID & '&'
			$data &= 'grant_type=authorization_code' & '&'
			$data &= 'redirect_uri=' & _URIEncode($callbackURL) & '&'
			$data &= 'code=' & $jmap['code']
			$dat = REST($tokenURL, 'POST', $data, $heads)
			Local $rc = @extended
			If ($rc > 299 Or $rc < 200) Then
				ConsoleWrite('Error [' & $rc & '] getting the token.' & @CRLF)
				MapRemove($jmap, 'code')
				MapRemove($jmap, 'token')
				MapRemove($jmap, 'refresh_token')
				fileWriteJson($jfile, $jmap)
				Return tokenGet($iter + 1)
			EndIf
		Else
			ConsoleWrite('Refreshing the token...' & @CRLF)
			$data &= 'clientId=' & $clientID & '&'
			$data &= 'grant_type=refresh_token' & '&'
			$data &= 'refresh_token=' & $jmap['refresh_token']
			$dat = REST($tokenURL, 'POST', $data, $heads)
			Local $rc = @extended
			If ($rc > 299 Or $rc < 200) Then
				ConsoleWrite('Error [' & $rc & '] refreshing the token.' & @CRLF)
				MapRemove($jmap, 'token')
				MapRemove($jmap, 'refresh_token')
				fileWriteJson($jfile, $jmap)
				Return tokenGet($iter + 1)
			EndIf
		EndIf

		Local $token = regex($dat, '"access_token":"([^"]+)"', '')
		Local $refreshToken = regex($dat, '"refresh_token":"([^"]+)"', '')
		If $token Then
			ConsoleWrite('Got a token.' & @CRLF)
			$jmap['token'] = $token
			$jmap['refresh_token'] = $refreshToken
			fileWriteJson($jfile, $jmap)
		Else
			ConsoleWrite('Could not find "access_token".' & @CRLF)
			MapRemove($jmap, 'token')
			MapRemove($jmap, 'refresh_token')
			fileWriteJson($jfile, $jmap)
			MsgBox(0, @ScriptName, 'Could not get a token, exiting.', 5)
			Exit
		EndIf
	Else
		MapRemove($jmap, 'code')
		MapRemove($jmap, 'refresh_token')
		Local $token = codeGet(True)
		If $token Then
			ConsoleWrite('Got the token.' & @CRLF)
			$jmap['token'] = $token
		Else
			MsgBox(0, @ScriptName, 'Could not get a code, exiting.', 5)
			Exit
		EndIf
		fileWriteJson($jfile, $jmap)
	EndIf
EndFunc   ;==>tokenGet


Func codeGet($token = False)
	Local $type = 'code'
	If $token Then $type = 'token'
	ConsoleWrite('Getting a authorisation for a ' & $type & '...' & @CRLF)
	Local $validHost = StringRegExp($callbackURL, '^http://(127.0.0.1|localhost)')
	If Not $validHost Then
		ConsoleWrite('Invalid Callback URL: ' & $callbackURL & @CRLF)
		ConsoleWrite('Should be something like: http://127.0.0.1:1271/' & @CRLF)
		Exit
	EndIf
	Local $sIPAddress = "127.0.0.1"
	Local $iPort = Int(regex($callbackURL, 'http://[^:/]+:(\d+)', 80))

	TCPStartup()
	Local $iListenSocket = TCPListen($sIPAddress, $iPort, 1)
	Local $iError = 0
	If @error Then
		$iError = @error
		MsgBox(BitOR($MB_SYSTEMMODAL, $MB_ICONHAND), "", "Server:" & @CRLF & "Could not listen, Error code: " & $iError)
		Return ''
	EndIf

	ConsoleWrite('Waiting for authorization from FitBit...' & @CRLF)
	ShellExecute($authURL & '?response_type=' & $type & '&client_id=' & $clientID & '&redirect_uri=' & _URIEncode($callbackURL) & '&scope=activity%20heartrate%20location%20nutrition%20profile%20settings%20sleep%20social%20weight&expires_in=86400')

	Local $iSocket = 0
	Local $ti = TimerInit()
	Local $expired = False
	Do
		$iSocket = TCPAccept($iListenSocket)
		If @error Then
			$iError = @error
			MsgBox(BitOR($MB_SYSTEMMODAL, $MB_ICONHAND), "", "Server:" & @CRLF & "Could not accept the incoming connection, Error code: " & $iError)
			Return ''
		EndIf
		$expired = (TimerDiff($ti) > 60000)
	Until (($iSocket <> -1) Or $expired)
	If $expired Then
		ConsoleWrite('Wated for a min., did not get anything from FitBit...' & @CRLF)
		TCPCloseSocket($iSocket)
		TCPShutdown()
		Return SetError(1, 0, 0)
	EndIf
	Local $sReceived = TCPRecv($iSocket, 2048)

	If $token Then
		ConsoleWrite($sReceived & @CRLF)
		;_SendHTML("<script>setTimeout(function(){window.location.href=window.location.href.split('#').join('?')}, 100)</script>", $iSocket)
		_SendHTML("<script>window.location.href=window.location.href.split('#').join('?')</script>", $iSocket)
		TCPCloseSocket($iSocket)
		Local $ti = TimerInit()
		Local $expired = False
		Do
			$iSocket = TCPAccept($iListenSocket)
			If @error Then
				$iError = @error
				MsgBox(BitOR($MB_SYSTEMMODAL, $MB_ICONHAND), "", "Server:" & @CRLF & "Could not accept the incoming connection, Error code: " & $iError)
				Return ''
			EndIf
			$expired = (TimerDiff($ti) > 10000)
		Until (($iSocket <> -1) Or $expired)
		If $expired Then
			ConsoleWrite('Wated for a redirect, did not get anything from the browser. Make sure to enable javascript.' & @CRLF)
			TCPCloseSocket($iSocket)
			TCPShutdown()
			Return SetError(1, 0, 0)
		EndIf
		$sReceived = TCPRecv($iSocket, 2048)
		_SendHTML('<center><button type="button" onclick="window.open('''', ''_self'', '''');window.close();">Successfully logged in, you can close this window.</button></center><script>window.open('''', ''_self'', '''');window.close();</script>', $iSocket)
		TCPCloseSocket($iSocket)
		TCPShutdown()
		;ConsoleWrite($sReceived & @CRLF)
		$sReceived = regex($sReceived, 'GET .*?access_token=([^& ]+)', '')
		Local $len = StringLen($sReceived)
		Return SetError((Not $len), $len, $sReceived)
	Else
		_SendHTML('<center><button type="button" onclick="window.open('''', ''_self'', '''');window.close();">Successfully logged in, you can close this window.</button></center><script>window.open('''', ''_self'', '''');window.close();</script>', $iSocket)
		TCPCloseSocket($iSocket)
		TCPShutdown()
		;ConsoleWrite($sReceived & @CRLF)
		$sReceived = regex($sReceived, 'GET .*?code=([^& ]+)', '')
		Local $len = StringLen($sReceived)
		Return SetError((Not $len), $len, $sReceived)
	EndIf
EndFunc   ;==>codeGet

Func _SendHTML($sHTML, $sSocket)
	Local $iLen, $sPacket, $sSplit
	$iLen = StringLen($sHTML)
	$sPacket = Binary("HTTP/1.1 200 OK" & @CRLF & _
			"Connection: close" & @CRLF & _
			"Content-Lenght: " & $iLen & @CRLF & _
			"Content-Type: text/html" & @CRLF & _
			@CRLF & _
			$sHTML)
	$sSplit = StringSplit($sPacket, "")
	$sPacket = ""
	For $i = 1 To $sSplit[0]
		If Asc($sSplit[$i]) <> 0 Then
			$sPacket = $sPacket & $sSplit[$i]
		EndIf
	Next
	TCPSend($sSocket, $sPacket)
EndFunc   ;==>_SendHTML






