#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Res_Description=Vparse application to process eCFR files
#AutoIt3Wrapper_Res_Fileversion=0.0.1.2
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_LegalCopyright=U.S. GPO
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Field=OriginalFilename|Vparse.exe
#AutoIt3Wrapper_Res_ProductVersion=0.1
#AutoIt3Wrapper_Res_Field=ProductName|Vparse
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Run_Au3Stripper=y
#AutoIt3Wrapper_icon=vparse_icon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;~ #include <Dbug.au3>
#include <Array.au3>
#include <Date.au3>
#include <GUIConstantsEx.au3>
#include <FileConstants.au3>
; #include <EditConstants.au3>
#include <ColorConstants.au3>
#include <WindowsConstants.au3>
#include <GUIConstants.au3>
#include <FontConstants.au3>
#include <File.au3>
#include <Color.au3>
#include <GuiRichEdit.au3>
#include <GuiTab.au3>


Global $sInFileDirDefault = "\\hqnapdcm0734\OFR\e_cfr"
Global $sOutFileDirDefault = "\\hqnapdcm0734\OFR\e_cfr\Apps\ECFRDATE"

Global $sInFileDir, $sOutFileDir

Global Const $COLOR_GPOTEAL = 0x3b80a1

Dim $hGUI, $hTab, $hInFolder, $hInTitle, $hInTitleLabel, $hInVolume, $hInVolumeLabel, $hDefault_Button, $hApply_Button, $hLoadFileButton, $hInRemarksList, _
		$hOutLabel, $hOut, $hOutFile, $hCreateAllOutsButton

fuMainGUI()
; create GUI and tabs
Func fuMainGUI()

	$hGUI = GUICreate("Vparse v" & _GetVersion(), 650, 600, Default, Default)
	GUISetState()

	$hTab = _GUICtrlTab_Create($hGUI, 5, 5, 642, 590)

	_GUICtrlTab_InsertItem($hTab, 0, "Main")
	_GUICtrlTab_InsertItem($hTab, 1, "Settings")

	; tab 0

	GUISetFont(15, $FW_NORMAL)
	$hInTitleLabel = GUICtrlCreateLabel("Title:", 17, 40)
	GUICtrlSetBkColor($hInTitleLabel, $GUI_BKCOLOR_TRANSPARENT)
	$hInTitle = GUICtrlCreateInput("1", 62, 37, 60, 33, BitOR($ES_NUMBER, $ES_CENTER))
	$hTitleUpDown = GUICtrlCreateUpdown($hInTitle)
	GUICtrlSetLimit(-1, 50, 1)

	$hInVolumeLabel = GUICtrlCreateLabel("Volume:", 125, 40)
	GUICtrlSetBkColor($hInVolumeLabel, $GUI_BKCOLOR_TRANSPARENT)
	$hInVolume = GUICtrlCreateInput("1", 200, 37, 60, 33, BitOR($ES_NUMBER, $ES_CENTER))
	$hVolumeUpDown = GUICtrlCreateUpdown($hInVolume)
	GUICtrlSetLimit(-1, 37, 1)

	$hOutLabel = GUICtrlCreateLabel("eCFR Date:", 270, 40)
	GUICtrlSetBkColor($hOutLabel, $GUI_BKCOLOR_TRANSPARENT)
	GUISetFont(15, $FW_BOLD)
	$hOut = GUICtrlCreateDate("", 380, 37, 140, 33, $DTS_SHORTDATEFORMAT)

	$hLoadFileButton = GUICtrlCreateButton("LOAD", 550, 37, 80, 33)
	GUICtrlSetBkColor($hLoadFileButton, $COLOR_GPOTEAL)
	GUICtrlSetColor($hLoadFileButton, $COLOR_WHITE)
	GUISetFont(8.5, $FW_NORMAL)


	$hInRemarksList = _GUICtrlRichEdit_Create($hGUI, "", 14, 80, 623, 475, BitOR($ES_MULTILINE, $WS_VSCROLL, $WS_HSCROLL))

	$hCreateAllOutsButton = GUICtrlCreateButton("PROCESS FILES", 270, 565, 120, 22)
	GUICtrlSetBkColor($hCreateAllOutsButton, $COLOR_GPOTEAL)
	GUICtrlSetColor($hCreateAllOutsButton, $COLOR_WHITE)
	GUICtrlSetState($hCreateAllOutsButton, $GUI_DISABLE)

	; tab 1

	$defDir_label = GUICtrlCreateLabel("Default Directory", 35, 45)
	GUICtrlSetBkColor($defDir_label, $GUI_BKCOLOR_TRANSPARENT)

	$hInFolder = GUICtrlCreateInput("", 35, 65, 320, 20)
	$sInFileDir = fuGetRegValsForSettings("Dir", $sInFileDirDefault)
	GUICtrlSetData($hInFolder, $sInFileDir)

	$efFileLoc_label = GUICtrlCreateLabel("Effective Date File Location", 35, 105)
	GUICtrlSetBkColor($efFileLoc_label, $GUI_BKCOLOR_TRANSPARENT)
	$hOutFile = GUICtrlCreateInput("", 35, 125, 320, 20)
	$sOutFileDir = fuGetRegValsForSettings("Date", $sOutFileDirDefault)
	GUICtrlSetData($hOutFile, $sOutFileDir)

	$hDefault_Button = GUICtrlCreateButton("Default", 400, 225, 75)
	GUICtrlSetBkColor($hDefault_Button, $COLOR_GPOTEAL)
	GUICtrlSetColor($hDefault_Button, $COLOR_WHITE)
	$hApply_Button = GUICtrlCreateButton("Apply", 485, 225, 75)
	GUICtrlSetBkColor($hApply_Button, $COLOR_GPOTEAL)
	GUICtrlSetColor($hApply_Button, $COLOR_WHITE)
	GUICtrlCreateTabItem("") ; end tabitem definition

	GUICtrlSetState($defDir_label, $GUI_HIDE)
	GUICtrlSetState($hInFolder, $GUI_HIDE)
	GUICtrlSetState($efFileLoc_label, $GUI_HIDE)
	GUICtrlSetState($hOutFile, $GUI_HIDE)
	GUICtrlSetState($hDefault_Button, $GUI_HIDE)
	GUICtrlSetState($hApply_Button, $GUI_HIDE)

	GUISetState()

	; This is the current active tab
	$iLastTab = 0

	; Run the GUI until the dialog is closed
	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				GUIDelete()
				Exit

			Case $hLoadFileButton
				If fuLoadFile() > 0 Then
					GUICtrlSetState($hCreateAllOutsButton, $GUI_ENABLE)
				EndIf

			Case $hDefault_Button
				$sInFileDir = $sInFileDirDefault
				GUICtrlSetData($hInFolder, $sInFileDir)
				$sOutFileDir = $sOutFileDirDefault
				GUICtrlSetData($hOutFile, $sOutFileDir)
				ContinueCase

			Case $hApply_Button
				fuApplySettingsValue($hInFolder, "Dir")
				fuApplySettingsValue($hOutFile, "Date")

			Case $hCreateAllOutsButton
				MsgBox($MB_ICONINFORMATION + $MB_OK, "Beta Version", "Sorry, this functionality has not been implemented yet!")
		EndSwitch

		; Check which Tab is active
		$iCurrTab = _GUICtrlTab_GetCurFocus($hTab)
		; If the Tab has changed
		If $iCurrTab <> $iLastTab Then
			; Store the value for future comparisons
			$iLastTab = $iCurrTab
			; Show/Hide controls as required
			Switch $iCurrTab
				Case 0
					GUICtrlSetState($defDir_label, $GUI_HIDE)
					GUICtrlSetState($hInFolder, $GUI_HIDE)
					GUICtrlSetState($efFileLoc_label, $GUI_HIDE)
					GUICtrlSetState($hOutFile, $GUI_HIDE)
					GUICtrlSetState($hDefault_Button, $GUI_HIDE)
					GUICtrlSetState($hApply_Button, $GUI_HIDE)

					GUICtrlSetState($hInTitleLabel, $GUI_SHOW)
					GUICtrlSetState($hInTitle, $GUI_SHOW)
					GUICtrlSetState($hTitleUpDown, $GUI_SHOW)
					GUICtrlSetState($hInVolumeLabel, $GUI_SHOW)
					GUICtrlSetState($hInVolume, $GUI_SHOW)
					GUICtrlSetState($hVolumeUpDown, $GUI_SHOW)
					GUICtrlSetState($hOutLabel, $GUI_SHOW)
					GUICtrlSetState($hOut, $GUI_SHOW)
					GUICtrlSetState($hLoadFileButton, $GUI_SHOW)
					ControlShow($hGUI, "", $hInRemarksList)
				Case 1
					GUICtrlSetState($hInTitleLabel, $GUI_HIDE)
					GUICtrlSetState($hInTitle, $GUI_HIDE)
					GUICtrlSetState($hTitleUpDown, $GUI_HIDE)
					GUICtrlSetState($hInVolumeLabel, $GUI_HIDE)
					GUICtrlSetState($hInVolume, $GUI_HIDE)
					GUICtrlSetState($hVolumeUpDown, $GUI_HIDE)
					GUICtrlSetState($hOutLabel, $GUI_HIDE)
					GUICtrlSetState($hOut, $GUI_HIDE)
					GUICtrlSetState($hLoadFileButton, $GUI_HIDE)
					ControlHide($hGUI, "", $hInRemarksList)

					GUICtrlSetState($defDir_label, $GUI_SHOW)
					GUICtrlSetState($hInFolder, $GUI_SHOW)
					GUICtrlSetState($efFileLoc_label, $GUI_SHOW)
					GUICtrlSetState($hOutFile, $GUI_SHOW)
					GUICtrlSetState($hDefault_Button, $GUI_SHOW)
					GUICtrlSetState($hApply_Button, $GUI_SHOW)
			EndSwitch
		EndIf
	WEnd
EndFunc   ;==>fuMainGUI

; function to get input or output values from registry if they exist
Func fuGetRegValsForSettings($sFolder, $DefaultFolder)

	Local $sRegValue

	$sRegValue = RegRead("HKEY_CURRENT_USER\Software\USGPO\PED\Vparse", $sFolder)
	If $sRegValue = "" Then
		RegWrite("HKEY_CURRENT_USER\Software\USGPO\PED\Vparse", $sFolder, "REG_SZ", $DefaultFolder)
		Return $DefaultFolder
	Else
		Return $sRegValue
	EndIf

EndFunc   ;==>fuGetRegValsForSettings

Func fuApplySettingsValue($hGUI, $sFolder)
	Local $cInputVal = GUICtrlRead($hGUI)
	$cInputVal = StringRegExpReplace($cInputVal, '\\* *$', '') ; strip trailing \ and spaces
	If Not FileExists($cInputVal) Then
		MsgBox(16, "Location invalid", $sFolder & " location does not exists. Enter a valid path to it.")
	Else
		If Not RegWrite("HKEY_CURRENT_USER\Software\USGPO\PED\Vparse", $sFolder, "REG_SZ", $cInputVal) Then
			MsgBox(16, "Could not be saved", $sFolder & " location could not be saved, Error #" & @error)
		EndIf
	EndIf
	GUICtrlSetData($hGUI, $cInputVal)
	Return
EndFunc   ;==>fuApplySettingsValue

Func _GetVersion()
	If @Compiled Then
		Return FileGetVersion(@AutoItExe)
	Else
		Return IniRead(@ScriptFullPath, "FileVersion", "#AutoIt3Wrapper_Res_Fileversion", "0.0.0.0")
	EndIf
EndFunc   ;==>_GetVersion

Func fuLoadFile()
	GUICtrlSetBkColor($hInTitle, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetBkColor($hInVolume, $GUI_BKCOLOR_TRANSPARENT)
	_GUICtrlRichEdit_SetText($hInRemarksList, "")

	Local $iTitleNum = GUICtrlRead($hInTitle)
	Local $iVolumeNum = GUICtrlRead($hInVolume)
	If $iTitleNum < 1 Or $iTitleNum > 50 Then
		GUICtrlSetBkColor($hInTitle, $COLOR_RED)
		MsgBox($MB_ICONERROR, "Title Number Out of Range", "Title Number Should Be Between 1 and 50 !!!")
		Return 0
	EndIf
	If $iVolumeNum < 1 Or $iVolumeNum > 37 Then
		GUICtrlSetBkColor($hInVolume, $COLOR_RED)
		MsgBox($MB_ICONERROR, "Volume Number Out of Range", "Volume Number Should Be Between 1 and 37 !!!")
		Return 0
	EndIf

	Local $aEffDate = FileGetTime($sOutFileDir, $FT_MODIFIED)
	GUICtrlSetData($hOut, $aEffDate[0] & "/" & $aEffDate[1] & "/" & $aEffDate[2])

	Local $sFilePath = $sInFileDir & '\' & Number($iTitleNum) & '\' & Number($iTitleNum) & 'V' & Number($iVolumeNum) & ".TXT"
	If FileExists($sFilePath) Then
		Local $hFileOpen = FileOpen($sFilePath, $FO_READ)
		Local $sFileRead = FileRead($hFileOpen)
		FileClose($hFileOpen)
		ControlHide($hGUI, "", $hInRemarksList)
		_GUICtrlRichEdit_SetText($hInRemarksList, $sFileRead)
		ControlShow($hGUI, "", $hInRemarksList)

		_GUICtrlRichEdit_SetSel($hInRemarksList, _GUICtrlRichEdit_FindText($hInRemarksList, "<AMDDATE>") + 9, _GUICtrlRichEdit_FindText($hInRemarksList, "<FMTR>") - 1)
		_GUICtrlRichEdit_SetCharBkColor($hInRemarksList, Dec('8888FF'))
		_GUICtrlRichEdit_SetSel($hInRemarksList, _GUICtrlRichEdit_FindText($hInRemarksList, "<TITLENUM>") + 10, _GUICtrlRichEdit_FindText($hInRemarksList, "<SUBJECT>") - 1)
		_GUICtrlRichEdit_SetCharBkColor($hInRemarksList, Dec('8888FF'))
		_GUICtrlRichEdit_Deselect($hInRemarksList)
		Return 1
	Else
		MsgBox($MB_ICONERROR, "Error Reading File", "Selected Title/Volume File is Not Ready for Parsing!")
		Return 0
	EndIf
EndFunc   ;==>fuLoadFile
