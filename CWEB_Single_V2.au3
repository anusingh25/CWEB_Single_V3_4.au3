;open new explorer window with JED and search and open the drawing
;save file to folder
#include <IE.au3>
#include <MsgBoxConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <Array.au3>; For _ArrayDisplay()
Opt("SendKeyDelay", 0)
HotKeySet("+!d", "ShowMessage") ;Shift-Alt-d
;HotKeySet("+!1", "layer_on") ;Shift-Alt-1
;HotKeySet("+!2", "layer_off") ;Shift-Alt-2
HotKeySet("+!s", "bin") ;Shift-Alt-2
HotKeySet("+!f", "bin_input") ;Shift-Alt-2
HotKeySet("+!e", "lib") ;Shift-Alt-r
HotKeySet("+!{RIGHT}", "m_x") ;Shift-Alt-2
HotKeySet("+!{LEFT}", "m_x_n") ;Shift-Alt-2
HotKeySet("+!{UP}", "m_y") ;Shift-Alt-2
HotKeySet("+!{DOWN}", "m_y_n") ;Shift-Alt-2
HotKeySet("+!{END}", "m_z") ;Shift-Alt-2
HotKeySet("+!{HOME}", "m_z_n") ;Shift-Alt-2
HotKeySet("^`", "modify") ;Shift-Alt-2
HotKeySet("+!{PGUP}", "hide") ;Shift-Alt-PGUP
Global $value = 20

Local $mGUI = GUICreate("Move", 30, 60, 1830, 948, $WS_SIZEBOX)

Local $idInput = GUICtrlCreateInput($value, 20, 60, 10, 10)
GUICtrlCreateUpdown($idInput)


; Attempt to resize input control
GUICtrlSetPos($idInput, 10, 10, 60, 40)
;	 GUISetState(@SW_SHOW)
;	 GUISetState(@SW_SHOW)



Global $proj = "nissan_released"


While 1
	Sleep(100)
	$idMsg = GUIGetMsg()

	Switch $idMsg
		Case $GUI_EVENT_CLOSE
			$value = GUICtrlRead($idInput)
			GUISetState(@SW_HIDE)
	EndSwitch
WEnd


Func ShowMessage()
	Const $navOpenInNewTab = 0x0800
	Global $line = StringStripWS(StringReplace(StringReplace(InputBox("CWEB", "Enter Part Number", StringStripWS(StringReplace(StringReplace(StringStripWS(ClipGet(), $STR_STRIPTRAILING), "-", "_"), " ", "_"), $STR_STRIPTRAILING), "", 150, 130, 1990, 948), "-", "_"), " ", "_"), $STR_STRIPALL)
	$sUrl = "http://cweb/cweb/cgi-bin/user_login.cgi"
	$o_IE = _IEAttach($sUrl, "url")
	Local $oFrames = _IEFrameGetCollection($o_IE)
	Local $oFrame = _IEFrameGetObjByName($o_IE, "Top")
	Local $sTxt = _IEPropertyGet($oFrame, "innertext") & @CRLF
	;MsgBox($MB_SYSTEMMODAL, "Frames Info", $sTxt)
	If Not IsObj($o_IE) Then
		$o_IE = _IEAttach("CAD Centric Systems")
		_IEAction($o_IE, "back")
		;$o_IE = _IECreate()
		_IELoadWait($o_IE)
		; _IENavigate($o_IE, $sUrl)
	EndIf
	Local $oForms = _IEFormGetCollection($oFrame)
	Local $oForm = _IEFormGetCollection($oFrame, 1)
	Local $oQuery = _IEFormElementGetObjByName($oForm, "searchtext")
	_IEFormElementSetValue($oQuery, $line)
	_IEFormSubmit($oForm)
	_IELoadWait($o_IE)
	_IELoadWait($oForm)
	Sleep(2000)
	Local $oFrame_m = ""
	Local $oFrame_m = _IEFrameGetObjByName($o_IE, "Middle")
	Local $oTable = _IETableGetCollection($oFrame_m, 0)
	Local $aTableData = _IETableWriteToArray($oTable, True)
	;_ArrayDisplay($aTableData)
	;msgbox(4096,"Hi",$aTableData[0][0])
	;_IELinkClickByText ($oFrame_m, $aTableData[0][0])
	;msgbox(4096,"Hi","http://cweb"& stringreplace($aTableData[0][8],"'",""))
	;msgbox(4096,"Hi",ubound($aTableData,$UBOUND_COLUMNS ))
	If UBound($aTableData, $UBOUND_COLUMNS) > 1 Then
		ShellExecute("http://cweb" & StringReplace(StringLeft($aTableData[0][8], StringLen($aTableData[0][8]) - 1), "'", ""))
	Else
		MsgBox(4096, "Info", "DWG not found.", 2)
	EndIf
	ClipPut($line)
EndFunc   ;==>ShowMessage
#comments-start
Func layer_on()
	ControlSetText("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "/F PR R FIL C:\Temp\NXI6\n.prg OKAY")
	ControlSend("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "{ENTER}")

EndFunc   ;==>layer_on

Func layer_off()
	ControlSetText("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "/F PR R FIL C:\Temp\NXI6\f.prg OKAY")
	ControlSend("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "{ENTER}")
EndFunc   ;==>layer_off
#ce
Func m_x()
	GUISetState(@SW_SHOW,$mGUI)
	ControlSetText("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "/or mo " & move())
	ControlSend("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "{ENTER}")
	WinActivate("Move")
EndFunc   ;==>m_x

Func m_y()
	GUISetState(@SW_SHOW,$mGUI)
	ControlSetText("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "/or mo 0," & move())
	ControlSend("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "{ENTER}")
	WinActivate("Move")
EndFunc   ;==>m_y

Func m_z()
	GUISetState(@SW_SHOW,$mGUI)
	ControlSetText("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "/or mo 0,0," & move())
	ControlSend("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "{ENTER}")
	WinActivate("Move")
EndFunc   ;==>m_z

Func m_x_n()
	GUISetState(@SW_SHOW,$mGUI)
	ControlSetText("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "/or mo -" & move())
	ControlSend("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "{ENTER}")
	WinActivate("Move")
EndFunc   ;==>m_x_n

Func m_y_n()
	GUISetState(@SW_SHOW,$mGUI)
	ControlSetText("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "/or mo 0,-" & move())
	ControlSend("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "{ENTER}")
	WinActivate("Move")
EndFunc   ;==>m_y_n

Func modify()
	ControlSetText("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "/mo e")
	ControlSend("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "{ENTER}")
EndFunc   ;==>modify

Func m_z_n()
	GUISetState(@SW_SHOW,$mGUI)
	ControlSetText("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "/or mo 0,0,-" & move())
	ControlSend("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "{ENTER}")
	WinActivate("Move")
EndFunc   ;==>m_z_n

Func bin() ;shift alt s
	WinActivate("NX I-deas")
	Local $aPos = WinGetPos("NX I-deas")
	Opt("MouseCoordMode", 0)
	MouseClickDrag("left", 28,$aPos[3]-66, 400, $aPos[3]-66, 0)
	ControlSend("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:3]", "^c")
	$text = StringRegExp(Clipget(), '(?:Hi\d*-)(.*)(?:_\d*")', $STR_REGEXPARRAYMATCH)
	If Not @error Then
		 ;MsgBox($MB_OK, "SRE Example 5 Result", $text[0])
    bincommon("I *" & $text[0])
EndIf

EndFunc   ;==>bin


Func bin_input() ;shift alt f
	Global $libst = StringStripWS(StringReplace(StringReplace(InputBox("CWEB", "Enter Part Number", StringStripWS(StringReplace(StringReplace(StringStripWS(ClipGet(), $STR_STRIPTRAILING), "-", "_"), " ", "_"), $STR_STRIPTRAILING), "", 150, 130, 1990, 948), "-", "_"), " ", "_"), $STR_STRIPALL)
	bincommon("P *" & $libst)
EndFunc   ;==>bin_input

Func bincommon($val)
	ControlSetText("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "/ma mb FE C")
	ControlSend("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "{ENTER}")
	Sleep(200)
	ControlSetText("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "OKAY")
	ControlSend("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "{ENTER}")
	Sleep(200)
	ControlSetText("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "Canc")
	ControlSend("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "{ENTER}")
	Sleep(100)
	ControlSetText("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "/ma mb FE QG " & $val & "*")
	ControlSend("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "{ENTER}")
	;msgbox(4096,"Hi",$text)
	ControlSetText("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "OKAY")
	ControlSend("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "{ENTER}")
	Sleep(300)
	ControlSetText("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "CANC")
	ControlSend("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "{ENTER}")
	ControlSetText("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "/ma mb")
	ControlSend("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "{ENTER}")
EndFunc   ;==>bincommon

Func lib() ;shift alt e
	Global $libs = StringStripWS(StringReplace(StringReplace(InputBox("CWEB", "Enter Part Number", StringStripWS(StringReplace(StringReplace(StringStripWS(ClipGet(), $STR_STRIPTRAILING), "-", "_"), " ", "_"), $STR_STRIPTRAILING), "", 150, 130, 1990, 948), "-", "_"), " ", "_"), $STR_STRIPALL)
	;Global $proj = StringStripWS(InputBox("Project", "Enter Project Name", "nissan_released", "", 150, 130, 1990, 948), $STR_STRIPALL)
	;GUISetState(@SW_SHOW,$lGUI) ; will display an  dialog box with 1 checkbox
	;msgbox(4096,"Title",@error)
	;If @error=1 Then $proj="hardware"
	;msgbox(4096,"Title",$proj)
	Example()
	ControlSetText("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "/ma gfl FE QG PN " & $proj & " !QG P *" & $libs & "*")
	ControlSend("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "{ENTER}")
	Sleep(400)
	ControlSetText("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "OKAY")
	ControlSend("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "{ENTER}")
	Sleep(200)
	ControlSetText("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "OKAY")
	ControlSend("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "{ENTER}")
	;msgbox(4096,"Hi",$text)
	Sleep(300)
	ControlSetText("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "/ma gfl")
	ControlSend("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "{ENTER}")
	WinWaitActive("Get from Project Library")
	ControlClick("Get from Project Library", "", "[CLASS:Button; INSTANCE:20]", "left")
	WinWaitActive("Filter for Project Libraries")
	ControlClick("Filter for Project Libraries", "", "[CLASS:Button; INSTANCE:3]", "left")
	ControlClick("Filter for Project Libraries", "", "[CLASS:Button; INSTANCE:12]", "left")
	ClipPut($libs)
EndFunc   ;==>lib

Func hide()
	ControlSetText("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "/do hi")
	ControlSend("NX I-deas 6.2", "", "[CLASS:Edit; INSTANCE:4]", "{ENTER}")
EndFunc   ;==>m_y_n

Func Example()
	Local $lGUI = GUICreate("My GUI radio", 170, 130, 1995, 948) ; will create a dialog box that when displayed is centered

	Local $idRadio1 = GUICtrlCreateRadio("nissan_released", 10, 10, 120, 20)
	Local $idRadio2 = GUICtrlCreateRadio("hardware", 10, 40, 120, 20)
	Local $idRadio3 = GUICtrlCreateRadio("users.", 10, 70, 50, 20)
	Local $idMyedit = GUICtrlCreateEdit("patelv", 60, 70, 100, 20,$ES_AUTOVSCROLL + $WS_VSCROLL)
	GUICtrlSetState($idRadio1, $GUI_CHECKED)

	GUISetState(@SW_SHOW,$lGUI) ; will display an  dialog box with 1 checkbox

	Local $idMsg
	While 1
		$idMsg = GUIGetMsg()
		Select
			Case $idMsg = $GUI_EVENT_CLOSE
				ExitLoop
			Case $idMsg = $idRadio1 And BitAND(GUICtrlRead($idRadio1), $GUI_CHECKED) = $GUI_CHECKED
				Global $proj = "nissan_released"
				;MsgBox($MB_SYSTEMMODAL, 'Info:', 'You clicked the Radio 1 and it is Checked.')
				GUISetState(@SW_HIDE,$lGUI)
				ExitLoop
			Case $idMsg = $idRadio2 And BitAND(GUICtrlRead($idRadio2), $GUI_CHECKED) = $GUI_CHECKED
				Global $proj = "hardware"
				;MsgBox($MB_SYSTEMMODAL, 'Info:', 'You clicked on Radio 2 and it is Checked.')
				GUISetState(@SW_HIDE,$lGUI)
				ExitLoop
			Case $idMsg = $idRadio3 And BitAND(GUICtrlRead($idRadio3), $GUI_CHECKED) = $GUI_CHECKED
				Global $proj = "users." & GUICtrlRead($idMyedit)
				;MsgBox($MB_SYSTEMMODAL, 'Info:', 'You clicked on Radio 2 and it is Checked.')
				GUISetState(@SW_HIDE,$lGUI)
				ExitLoop
		EndSelect
	WEnd
EndFunc   ;==>Example

Func move()
	$value = GUICtrlRead($idInput)
	ControlFocus("Move", "", "[CLASS:Edit; INSTANCE:1]")
	;$EM_SETSEL = 0x00B1
	GUICtrlSendMsg($idInput, $EM_SETSEL, 0, 5)
	Return $value;
EndFunc   ;==>move
