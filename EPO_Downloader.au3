#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=..\..\Program\AutoIt\Icons\D.ico
#AutoIt3Wrapper_Run_Tidy=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
; #RequireAdmin

;
; AutoIt Version: 3.3.14
; Author:         Cloud Zhang
;
; Script Function:
;   EPO original document downloader
;

#include <IE.au3>
#include <Clipboard.au3>
#include <GDIPlus.au3>
#include <JY_OCR.au3>	; Thanks for http://www.jianyiit.com/post-136.html
#include <Json.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>

#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("EPO文件下载工具 1.3", 340, 120, 300, 300)
$Input1 = GUICtrlCreateInput("公开号...", 15, 15, 230, 25)
$Button1 = GUICtrlCreateButton("开始下载", 270, 15, 60, 25)
GUICtrlSetState(-1, $GUI_DISABLE)
$Input2 = GUICtrlCreateInput("下载到...", 15, 50, 230, 25)
$Button2 = GUICtrlCreateButton("路径浏览", 270, 50, 60, 25)
$Lable1 = GUICtrlCreateLabel("", 120, 85, 100, 25, $SS_CENTER)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $Button2
			Local $sFolderPath = FileSelectFolder('请选择文件夹', '')
			GUICtrlSetData($Input2, $sFolderPath)
			If FileExists($sFolderPath) Then GUICtrlSetState($Button1, $GUI_ENABLE)
		Case $Button1
			Local $sPubNum = StringStripWS(GUICtrlRead($Input1), 8)
			GUICtrlSetData($Input1, $sPubNum)
			GUICtrlSetState($Input1, $GUI_DISABLE)
			GUICtrlSetState($Input2, $GUI_DISABLE)
			GUICtrlSetState($Button1, $GUI_DISABLE)
			GUICtrlSetState($Button2, $GUI_DISABLE)
			_Main($sPubNum, $sFolderPath)
	EndSwitch
WEnd

Func _Main($PubNum, $FolderPath)
	RegWrite('HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main', 'Default Download Directory', 'REG_SZ', $FolderPath)

	GUICtrlSetColor($Lable1, 0x0000ff)
	GUICtrlSetData($Lable1, '开始检索公开号')

	Global $oIE = _IECreate("about:blank")
	While @error
		Sleep(1000)
		_IEQuit($oIE)
		Sleep(1000)
		Global $oIE = _IECreate("about:blank")
	WEnd
	_IEPropertySet($oIE, 'height', 300)
	_IEPropertySet($oIE, 'width', 300)
	Local $sIE_Title = _IEPropertyGet($oIE, 'title')
	WinSetState($sIE_Title, '', @SW_HIDE)
	Global $TCodeFile = @TempDir & "\EPO-Code.jpg"
	Local $sStartLink = 'https://worldwide.espacenet.com/searchResults?ST=singleline&locale=cn_EP&submitted=true&DB=&query=' & $PubNum
	Local $sCaseLink = _GetLink($sStartLink, 'https://worldwide.espacenet.com/publicationDetails/')
	If $sCaseLink <> '' Then
		GUICtrlSetData($Lable1, '公开号已确认')
		Local $sDocLink = _GetLink($sCaseLink, 'https://worldwide.espacenet.com/publicationDetails/originalDocument')
		Local $sDownLink = _GetLink($sDocLink, 'https://worldwide.espacenet.com/data/espacenetDocument.pdf?ND=4&flavour=trueFull')

		_IENavigate($oIE, $sDownLink)
		_IELoadWait($oIE)
		Sleep(2000)
		Local $sIE_Title = _IEPropertyGet($oIE, 'title')
		WinSetState($sIE_Title, '', @SW_SHOW)
		If FileExists($FolderPath & '\' & $PubNum & '.PDF') Then FileDelete($FolderPath & '\' & $PubNum & '.PDF')

		GUICtrlSetData($Lable1, '开始破解验证码')
		SplashTextOn('提示', '请停止使用鼠标及键盘，直至该提示消失！', 400, 50)
		_GetVerCodeImg()
		Local $sVerCode = _AuCodeOCR()
		_SubmitVerCode($sVerCode)

		While 1
			Sleep(2000)
			If __DirectUIHWND_Exist($sIE_Title) Then
				ExitLoop
			Else
				_GetVerCodeImg()
				Local $sVerCode = _AuCodeOCR()
				_SubmitVerCode($sVerCode)
			EndIf
		WEnd

		Sleep(1000)
		ControlSend($sIE_Title, '', '[CLASS:DirectUIHWND; INSTANCE:1]', "!s")
		Local $sIE_Title = _IEPropertyGet($oIE, 'title')
		WinSetState($sIE_Title, '', @SW_MINIMIZE)
		GUICtrlSetData($Lable1, '开始文件下载')
		SplashOff()

		While 1
			Sleep(500)
			If FileExists($FolderPath & '\' & $PubNum & '.PDF') Then ExitLoop
		WEnd
		GUICtrlSetColor($Lable1, 0x00ff00)
		GUICtrlSetData($Lable1, '文件下载完毕')
	Else
		GUICtrlSetColor($Lable1, 0xff0000)
		GUICtrlSetData($Lable1, '公开号检索失败')
	EndIf
	_IEQuit($oIE)
	RegDelete('HKCU64\Software\Microsoft\Internet Explorer\Main', 'Default Download Directory')

	GUICtrlSetState($Input1, $GUI_ENABLE)
	GUICtrlSetState($Input2, $GUI_ENABLE)
	GUICtrlSetState($Button1, $GUI_ENABLE)
	GUICtrlSetState($Button2, $GUI_ENABLE)
EndFunc   ;==>_Main

Func _GetLink($NowLink, $NextLink) ; Goto $NowLink then get $NextLink
	_IENavigate($oIE, $NowLink)
	_IELoadWait($oIE)
	Sleep(2000)

	Local $oLinks = _IELinkGetCollection($oIE)
	Local $iNumLinks = @extended

	For $oLink In $oLinks
		If StringInStr($oLink.href, $NextLink) Then
			Return $oLink.href
		EndIf
	Next
EndFunc   ;==>_GetLink

Func _GetVerCodeImg() ; Catch verification code image and save to file
	$oImg = _IEImgGetCollection($oIE, 0)
	$oPic = $oIE.Document.body.createControlRange()
	$oPic.Add($oImg)
	$oPic.execCommand("Copy")
	$bmp = ClipGet()

	_GDIPlus_Startup()
	_ClipBoard_Open(0)
	$iVerifyPics = _ClipBoard_GetDataEx($CF_BITMAP)
	$iVerifyPics = _GDIPlus_BitmapCreateFromHBITMAP($iVerifyPics)
	_ClipBoard_Close()
	FileDelete($TCodeFile)
	_GDIPlus_ImageSaveToFile($iVerifyPics, $TCodeFile)
EndFunc   ;==>_GetVerCodeImg

Func _AuCodeOCR() ; Get OCR result from Baidu API through verification code image file
	; Get Access token from Baidu API
	Local $sACodeUrl = 'https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id=xxxxxxx&client_secret=yyyyyyyyyyyy'
	Local $sACHeader = 'application/json; charset=UTF-8'
	Local $jAC = _HTTPPost($sACodeUrl, $sACHeader, '')
	Local $oACode = Json_Decode($jAC)
	Local $sACode = Json_objGet($oACode, 'access_token')

	; Get OCR result from Baidu API
	Local $File = FileOpen($TCodeFile, 16)
	Local $sImage = JY_Base64Encode(FileRead($File))
	FileClose($File)

	Local $sOCRUrl = 'https://aip.baidubce.com/rest/2.0/ocr/v1/general_basic?access_token=' & $sACode
	Local $sOCRHeader = 'application/x-www-form-urlencoded'
	Local $sOCRContent = 'image=' & JY_StringToURLEncode($sImage)
	Local $jOCR = _HTTPPost($sOCRUrl, $sOCRHeader, $sOCRContent)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $jOCR = ' & $jOCR & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	; Get verification code through Json.au3 (often get error)
;~ 	Local $oOCR = Json_Decode($jOCR)
;~ 	If Json_ObjExists($oOCR, 'words_result') Then
;~ 		Local $aOCR1 = Json_objGet($oOCR, 'words_result')
;~ 		Local $aOCR2 = $aOCR1[0]
;~ 		Local $sOCR = Json_objGet($aOCR2, 'words')
;~ 		If @error Then
;~ 			Return 88888
;~ 		Else
;~ 			Return $sOCR
;~ 		EndIf
;~ 	Else
;~ 		Return 99999
;~ 	EndIf

	; Get verification code through string analysis
	If StringInStr($jOCR, 'words_result') Then
		Local $sFOCR = StringRegExpReplace($jOCR, '\W', '')
		Local $nPos = StringInStr($sFOCR, 'words', 0, -1) + 4
		Local $sOCR = StringRight($sFOCR, StringLen($sFOCR) - $nPos)
		Return $sOCR
	Else
		Return 99999
	EndIf
EndFunc   ;==>_AuCodeOCR

Func _HTTPPost($PostUrl, $ReHeader, $Content) ; POST content and receive Json info
	$oHTTP = ObjCreate("winhttp.winhttprequest.5.1")
	$oHTTP.Open("POST", $PostUrl, False)
	$oHTTP.SetRequestHeader("Content-Type", $ReHeader)
	$oHTTP.Send($Content)
	$oReceived = $oHTTP.ResponseText
	Return $oReceived
EndFunc   ;==>_HTTPPost

Func _SubmitVerCode($VerCode) ; Input verification code and submit
	Local $oForm = _IEFormGetCollection($oIE, 0)

	Local $oQuery = _IEFormElementGetObjByName($oForm, 'response')
	_IEFormElementSetValue($oQuery, $VerCode)
	Sleep(2000)
	Local $oSubmit = _IEFormElementGetCollection($oForm, 1)
	_IEAction($oSubmit, 'click')
EndFunc   ;==>_SubmitVerCode

Func __DirectUIHWND_Exist($sIE_Title) ; Check file download diag show up
	Local $aWinGetPos = WinGetPos($sIE_Title, '')
	Local $aControlGetPos = ControlGetPos($sIE_Title, '', '[CLASS:DirectUIHWND; INSTANCE:1]')

;~     ConsoleWrite($aWinGetPos[3] & @CRLF)
;~     ConsoleWrite($aControlGetPos[1] & @CRLF)
;~     ConsoleWrite($aControlGetPos[3] & @CRLF)
;~     ConsoleWrite(($aWinGetPos[3] - ($aControlGetPos[1] + $aControlGetPos[3])) & @CRLF)
;~     ConsoleWrite("> " & @CRLF)
	If $aWinGetPos[3] - ($aControlGetPos[1] + $aControlGetPos[3]) < 10 Then
		If ControlCommand($sIE_Title, '', '[CLASS:DirectUIHWND; INSTANCE:1]', "IsVisible", "") Then Return True
	EndIf
	Return False
EndFunc   ;==>__DirectUIHWND_Exist
