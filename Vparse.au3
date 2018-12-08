#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Res_Description=Vparse application to process eCFR files
#AutoIt3Wrapper_Res_Fileversion=1.0.1.8
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_LegalCopyright=U.S. GPO
#AutoIt3Wrapper_UseX64=Y
#AutoIt3Wrapper_Res_Field=OriginalFilename|Vparse.exe
#AutoIt3Wrapper_Res_ProductVersion=0.1
#AutoIt3Wrapper_Res_Field=ProductName|Vparse
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Run_Au3Stripper=y
#AutoIt3Wrapper_icon=vparse_icon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <Array.au3>
#include <Date.au3>
#include <GUIConstantsEx.au3>
#include <FileConstants.au3>
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
Global $iMaxVolNumDefault = 37

Global $sInFileDir, $sOutFileDir, $iMaxVolNum

Global Const $COLOR_GPOTEAL = 0x3b80a1

Dim $hGUI, $idTab, $idInFolder, $idInTitle, $idInTitleLabel, $idInVolume, $idInVolumeLabel, $idDefault_Button, $idApply_Button, $idLoadFileButton, $idInRemarksList, _
		$idOutLabel, $idOut, $idOutFile, $idCreateAllOutsButton, $idTitleUpDown, $idVolumeUpDown, $id_defDir_label, $idEfFileLoc_label, $idMaxVolume, $idMaxVolume_label

fuMainGUI()
; create GUI and tabs
Func fuMainGUI()

	$hGUI = GUICreate("Vparse v" & _GetVersion(), 650, 600, Default, Default)
	GUISetState()

	$idTab = _GUICtrlTab_Create($hGUI, 5, 5, 642, 590)

	_GUICtrlTab_InsertItem($idTab, 0, "Main")
	_GUICtrlTab_InsertItem($idTab, 1, "Settings")

	; tab 0

	GUISetFont(15, $FW_NORMAL)
	$idInTitleLabel = GUICtrlCreateLabel("Title:", 17, 40)
	GUICtrlSetBkColor($idInTitleLabel, $GUI_BKCOLOR_TRANSPARENT)
	$idInTitle = GUICtrlCreateInput("", 62, 37, 60, 33, BitOR($ES_NUMBER, $ES_CENTER))
	$idTitleUpDown = GUICtrlCreateUpdown($idInTitle)
	GUICtrlSetLimit($idTitleUpDown, 50, 1)

	$idInVolumeLabel = GUICtrlCreateLabel("Volume:", 125, 40)
	GUICtrlSetBkColor($idInVolumeLabel, $GUI_BKCOLOR_TRANSPARENT)
	$idInVolume = GUICtrlCreateInput("", 200, 37, 60, 33, BitOR($ES_NUMBER, $ES_CENTER))
	$idVolumeUpDown = GUICtrlCreateUpdown($idInVolume)

	$idOutLabel = GUICtrlCreateLabel("eCFR Date:", 270, 40)
	GUICtrlSetBkColor($idOutLabel, $GUI_BKCOLOR_TRANSPARENT)
	GUISetFont(15, $FW_BOLD)
	$idOut = GUICtrlCreateDate("1800/01/01", 380, 37, 140, 33, $DTS_SHORTDATEFORMAT)
	GUICtrlSetState($idOut, $GUI_DISABLE)

	$idLoadFileButton = GUICtrlCreateButton("LOAD", 550, 37, 80, 33)
	GUICtrlSetBkColor($idLoadFileButton, $COLOR_GPOTEAL)
	GUICtrlSetColor($idLoadFileButton, $COLOR_WHITE)
	GUISetFont(8.5, $FW_NORMAL)


	$idInRemarksList = _GUICtrlRichEdit_Create($hGUI, "", 14, 80, 623, 475, BitOR($ES_MULTILINE, $WS_VSCROLL, $WS_HSCROLL))

	GUISetFont(15, $FW_BOLD)
	$idCreateAllOutsButton = GUICtrlCreateButton("PROCESS FILES", 230, 557, 200, 33)
	GUICtrlSetBkColor($idCreateAllOutsButton, $COLOR_GPOTEAL)
	GUICtrlSetColor($idCreateAllOutsButton, $COLOR_WHITE)
	GUICtrlSetState($idCreateAllOutsButton, $GUI_DISABLE)
	GUISetFont(8.5, $FW_NORMAL)

	; tab 1

	$id_defDir_label = GUICtrlCreateLabel("Default Directory", 35, 45)
	GUICtrlSetBkColor($id_defDir_label, $GUI_BKCOLOR_TRANSPARENT)

	$idInFolder = GUICtrlCreateInput("", 35, 65, 320, 20)
	$sInFileDir = fuGetRegValsForSettings("Dir", $sInFileDirDefault, "REG_SZ")
	GUICtrlSetData($idInFolder, $sInFileDir)

	$idEfFileLoc_label = GUICtrlCreateLabel("Effective Date File Location", 35, 105)
	GUICtrlSetBkColor($idEfFileLoc_label, $GUI_BKCOLOR_TRANSPARENT)
	$idOutFile = GUICtrlCreateInput("", 35, 125, 320, 20)
	$sOutFileDir = fuGetRegValsForSettings("Date", $sOutFileDirDefault, "REG_SZ")
	GUICtrlSetData($idOutFile, $sOutFileDir)

	$idMaxVolume_label = GUICtrlCreateLabel("Largest Volume Number", 35, 165)
	GUICtrlSetBkColor($idMaxVolume_label, $GUI_BKCOLOR_TRANSPARENT)
	$idMaxVolume = GUICtrlCreateInput("", 35, 185, 20, 20)
	$iMaxVolNum = fuGetRegValsForSettings("Max_Volume", $iMaxVolNumDefault, "REG_DWORD")
	GUICtrlSetData($idMaxVolume, $iMaxVolNum)
	GUICtrlSetLimit($idVolumeUpDown, $iMaxVolNum, 1)

	$idDefault_Button = GUICtrlCreateButton("Default", 400, 225, 75)
	GUICtrlSetBkColor($idDefault_Button, $COLOR_GPOTEAL)
	GUICtrlSetColor($idDefault_Button, $COLOR_WHITE)
	$idApply_Button = GUICtrlCreateButton("Apply", 485, 225, 75)
	GUICtrlSetBkColor($idApply_Button, $COLOR_GPOTEAL)
	GUICtrlSetColor($idApply_Button, $COLOR_WHITE)
	GUICtrlCreateTabItem("") ; end tabitem definition

	GUICtrlSetState($id_defDir_label, $GUI_HIDE)
	GUICtrlSetState($idInFolder, $GUI_HIDE)
	GUICtrlSetState($idEfFileLoc_label, $GUI_HIDE)
	GUICtrlSetState($idMaxVolume_label, $GUI_HIDE)
	GUICtrlSetState($idOutFile, $GUI_HIDE)
	GUICtrlSetState($idMaxVolume, $GUI_HIDE)
	GUICtrlSetState($idDefault_Button, $GUI_HIDE)
	GUICtrlSetState($idApply_Button, $GUI_HIDE)

	GUISetState()

	; This is the current active tab
	$iLastTab = 0

	; Run the GUI until the dialog is closed
	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				GUIDelete()
				Exit

			Case $idLoadFileButton
				If fuLoadFile() > 0 Then
					GUICtrlSetState($idCreateAllOutsButton, $GUI_ENABLE)
					GUICtrlSetState($idOut, $GUI_ENABLE)
				EndIf

			Case $idDefault_Button
				$sInFileDir = $sInFileDirDefault
				GUICtrlSetData($idInFolder, $sInFileDir)
				$sOutFileDir = $sOutFileDirDefault
				GUICtrlSetData($idOutFile, $sOutFileDir)
				$iMaxVolNum = $iMaxVolNumDefault
				GUICtrlSetData($idMaxVolume, $iMaxVolNum)
				ContinueCase

			Case $idApply_Button
				fuApplySettingsValue($idInFolder, "Dir", "REG_SZ")
				fuApplySettingsValue($idOutFile, "Date", "REG_SZ")
				fuApplySettingsValue($idMaxVolume, "Max_Volume", "REG_DWORD")
				GUICtrlSetLimit($idVolumeUpDown, GUICtrlRead($idMaxVolume), 1)

			Case $idCreateAllOutsButton
				fuProcessFiles()
		EndSwitch

		; Check which Tab is active
		$iCurrTab = _GUICtrlTab_GetCurFocus($idTab)
		; If the Tab has changed
		If $iCurrTab <> $iLastTab Then
			; Store the value for future comparisons
			$iLastTab = $iCurrTab
			; Show/Hide controls as required
			Switch $iCurrTab
				Case 0
					GUICtrlSetState($id_defDir_label, $GUI_HIDE)
					GUICtrlSetState($idInFolder, $GUI_HIDE)
					GUICtrlSetState($idEfFileLoc_label, $GUI_HIDE)
					GUICtrlSetState($idMaxVolume_label, $GUI_HIDE)
					GUICtrlSetState($idOutFile, $GUI_HIDE)
					GUICtrlSetState($idMaxVolume, $GUI_HIDE)
					GUICtrlSetState($idDefault_Button, $GUI_HIDE)
					GUICtrlSetState($idApply_Button, $GUI_HIDE)

					GUICtrlSetState($idInTitleLabel, $GUI_SHOW)
					GUICtrlSetState($idInTitle, $GUI_SHOW)
					GUICtrlSetState($idTitleUpDown, $GUI_SHOW)
					GUICtrlSetState($idInVolumeLabel, $GUI_SHOW)
					GUICtrlSetState($idInVolume, $GUI_SHOW)
					GUICtrlSetState($idVolumeUpDown, $GUI_SHOW)
					GUICtrlSetState($idOutLabel, $GUI_SHOW)
					GUICtrlSetState($idOut, $GUI_SHOW)
					GUICtrlSetState($idLoadFileButton, $GUI_SHOW)
					ControlShow($hGUI, "", $idInRemarksList)
					GUICtrlSetState($idCreateAllOutsButton, $GUI_SHOW)
				Case 1
					GUICtrlSetState($idInTitleLabel, $GUI_HIDE)
					GUICtrlSetState($idInTitle, $GUI_HIDE)
					GUICtrlSetState($idTitleUpDown, $GUI_HIDE)
					GUICtrlSetState($idInVolumeLabel, $GUI_HIDE)
					GUICtrlSetState($idInVolume, $GUI_HIDE)
					GUICtrlSetState($idVolumeUpDown, $GUI_HIDE)
					GUICtrlSetState($idOutLabel, $GUI_HIDE)
					GUICtrlSetState($idOut, $GUI_HIDE)
					GUICtrlSetState($idLoadFileButton, $GUI_HIDE)
					ControlHide($hGUI, "", $idInRemarksList)
					GUICtrlSetState($idCreateAllOutsButton, $GUI_HIDE)

					GUICtrlSetState($id_defDir_label, $GUI_SHOW)
					GUICtrlSetState($idInFolder, $GUI_SHOW)
					GUICtrlSetState($idEfFileLoc_label, $GUI_SHOW)
					GUICtrlSetState($idMaxVolume_label, $GUI_SHOW)
					GUICtrlSetState($idOutFile, $GUI_SHOW)
					GUICtrlSetState($idMaxVolume, $GUI_SHOW)
					GUICtrlSetState($idDefault_Button, $GUI_SHOW)
					GUICtrlSetState($idApply_Button, $GUI_SHOW)
			EndSwitch
		EndIf
	WEnd
EndFunc   ;==>fuMainGUI

; function to get input or output values from registry if they exist
Func fuGetRegValsForSettings($sFolder, $DefaultFolder, $sRegValType)

	Local $sRegValue

	$sRegValue = RegRead("HKEY_CURRENT_USER\Software\USGPO\PED\Vparse", $sFolder)
	If $sRegValue = "" Then
		RegWrite("HKEY_CURRENT_USER\Software\USGPO\PED\Vparse", $sFolder, $sRegValType, $DefaultFolder)
		Return $DefaultFolder
	Else
		Return $sRegValue
	EndIf

EndFunc   ;==>fuGetRegValsForSettings

Func fuApplySettingsValue($hGUI, $sFolder, $sRegValType)
	Local $cInputVal = GUICtrlRead($hGUI)
	$cInputVal = StringRegExpReplace($cInputVal, '\\* *$', '') ; strip trailing \ and spaces
	If Not FileExists($cInputVal) And $sRegValType = "REG_SZ" Then
		MsgBox(16, "Location invalid", $sFolder & " location is not accessible. Enter a valid path to it.")
	Else
		If Not RegWrite("HKEY_CURRENT_USER\Software\USGPO\PED\Vparse", $sFolder, $sRegValType, $cInputVal) Then
			MsgBox(16, "Could not be saved", $sFolder & " location could not be saved in registry, Error #" & @error)
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
	GUICtrlSetBkColor($idInTitle, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetBkColor($idInVolume, $GUI_BKCOLOR_TRANSPARENT)
	_GUICtrlRichEdit_SetText($idInRemarksList, "")

	Local $iTitleNum = GUICtrlRead($idInTitle)
	Local $iVolumeNum = GUICtrlRead($idInVolume)
	If $iTitleNum < 1 Or $iTitleNum > 50 Then
		GUICtrlSetBkColor($idInTitle, $COLOR_RED)
		MsgBox($MB_ICONERROR, "Title Number Out of Range", "Title Number Should Be Between 1 and 50 !!!")
		Return 0
	EndIf
	If $iVolumeNum < 1 Or $iVolumeNum > 37 Then
		GUICtrlSetBkColor($idInVolume, $COLOR_RED)
		MsgBox($MB_ICONERROR, "Volume Number Out of Range", "Volume Number Should Be Between 1 and " & $iMaxVolNum & " !!!")
		Return 0
	EndIf

	Local $aEffDate = FileGetTime($sOutFileDir, $FT_MODIFIED)
	GUICtrlSetData($idOut, $aEffDate[0] & "/" & $aEffDate[1] & "/" & $aEffDate[2])

	Local $sFilePath = $sInFileDir & '\' & StringFormat("%02d", $iTitleNum) & '\' & $iTitleNum & 'V' & $iVolumeNum & ".TXT"
	If FileExists($sFilePath) Then
		Local $hFileOpen = FileOpen($sFilePath, $FO_READ)
		Local $sFileRead = FileRead($hFileOpen)
		FileClose($hFileOpen)
		ControlHide($hGUI, "", $idInRemarksList)
		_GUICtrlRichEdit_SetText($idInRemarksList, $sFileRead)
		ControlShow($hGUI, "", $idInRemarksList)

		_GUICtrlRichEdit_SetSel($idInRemarksList, _GUICtrlRichEdit_FindText($idInRemarksList, "<AMDDATE>") + 9, _GUICtrlRichEdit_FindText($idInRemarksList, "<FMTR>") - 1)
		_GUICtrlRichEdit_SetCharBkColor($idInRemarksList, Dec('8888FF'))
		_GUICtrlRichEdit_SetSel($idInRemarksList, _GUICtrlRichEdit_FindText($idInRemarksList, "<TITLENUM>") + 10, _GUICtrlRichEdit_FindText($idInRemarksList, "<SUBJECT>") - 1)
		_GUICtrlRichEdit_SetCharBkColor($idInRemarksList, Dec('8888FF'))
		_GUICtrlRichEdit_Deselect($idInRemarksList)
		Return 1
	Else
		MsgBox($MB_ICONERROR, "Error Reading File", "Selected Title/Volume File is Not Ready for Parsing!")
		Return 0
	EndIf
EndFunc   ;==>fuLoadFile

Func fuProcessFiles()
	Local $sNoExtFile, $sDocFile, $sTxtFile, $bNoExtFileStatus = "Fail", $bDocFileStatus = "Fail", $bTxtFileStatus = "Fail"
	Local $iTitleNum = GUICtrlRead($idInTitle)
	Local $iVolumeNum = GUICtrlRead($idInVolume)
	Local $contEffDate = StringSplit(GUICtrlRead($idOut), "/")
	Local $sHoldingFileName = $sInFileDir & '\Holding\' & StringFormat("%02d", $iTitleNum) & StringFormat("%02d", $iVolumeNum) & StringFormat("%02d", $contEffDate[1]) & StringFormat("%02d", $contEffDate[2]) & "." & StringRight($contEffDate[3], 3)
	Local $sXttDir = $sInFileDir & '\' & StringFormat("%02d", $iTitleNum) & '\x' & StringFormat("%02d", $iTitleNum) & "\"

	$sNoExtFile = $sInFileDir & '\' & StringFormat("%02d", $iTitleNum) & '\' & $iTitleNum & 'V' & $iVolumeNum
	$sDocFile = $sInFileDir & '\' & StringFormat("%02d", $iTitleNum) & '\' & $iTitleNum & 'V' & $iVolumeNum & ".doc"
	$sTxtFile = $sInFileDir & '\' & StringFormat("%02d", $iTitleNum) & '\' & $iTitleNum & 'V' & $iVolumeNum & ".TXT"

	;Test if any of the files are opened by other app
	If _FileIsUses($sNoExtFile) > 0 Or _FileIsUses($sDocFile) > 0 Or _FileIsUses($sTxtFile) > 0 Then
		MsgBox($MB_ICONERROR, "Files Locked", "Some files are locked by other applications!!! Please close all apps that use these files and restart application!!!")
		Return 0
	EndIf

;~ 	No extension file is deleted
	If FileDelete($sNoExtFile) Then
		$bNoExtFileStatus = "OK"
	EndIf


;~ 	TXT file is stripped off of extension and copied to /Holding/ directory with ttvvmmdd.0yy filename
	If Not _FileIsUses($sTxtFile) And FileMove($sTxtFile, $sNoExtFile, $FC_OVERWRITE) And FileCopy($sNoExtFile, $sHoldingFileName, $FC_OVERWRITE) Then
		$bTxtFileStatus = "OK"
	EndIf

;~ 	Move DOC file to xTT directory and copy /Holding/ dir version of the file there as well
	If Not _FileIsUses($sDocFile) And FileMove($sDocFile, $sXttDir & StringFormat("%02d", $iTitleNum) & StringFormat("%02d", $iVolumeNum) & StringFormat("%02d", $contEffDate[1]) & StringFormat("%02d", $contEffDate[2]) & ".doc", $FC_OVERWRITE) And FileCopy($sHoldingFileName, $sXttDir, $FC_OVERWRITE) Then
		$bDocFileStatus = "OK"
	EndIf

	GUICtrlSetData($idInVolume, "")
	GUICtrlSetData($idInTitle, "")
	GUICtrlSetData($idOut, "1800/01/01/")
	ControlHide($hGUI, "", $idInRemarksList)
	_GUICtrlRichEdit_SetText($idInRemarksList, "")
	ControlShow($hGUI, "", $idInRemarksList)
	GUICtrlSetState($idCreateAllOutsButton, $GUI_DISABLE)
	GUICtrlSetState($idOut, $GUI_DISABLE)
	MsgBox($MB_ICONINFORMATION + $MB_OK, "File Operations Status", "No Ext File: " & $bNoExtFileStatus & @CRLF & "TXT File: " & $bTxtFileStatus & @CRLF & "Doc File: " & $bDocFileStatus)

	Return 1

EndFunc   ;==>fuProcessFiles

Func _FileIsUses($sFile)

	Local $hFile = _WinAPI_CreateFile($sFile, 2, 2, 0)

	If $hFile Then
		_WinAPI_CloseHandle($hFile)
		Return 0
	EndIf

	Local $Error = _WinAPI_GetLastError()

	Switch $Error
		Case 32 ; ERROR_SHARING_VIOLATION
			Return 1
		Case Else
			Return SetError($Error, 0, 0)
	EndSwitch
EndFunc   ;==>_FileIsUses
