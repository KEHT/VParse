;~ #include <Dbug.au3>
#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <FileConstants.au3>
#include <EditConstants.au3>
#include <GuiListView.au3>
#include <ColorConstants.au3>
#include <WindowsConstants.au3>
#include <GUIConstants.au3>
#include <FontConstants.au3>
#include <File.au3>

Opt("GUIOnEventMode", 1)

Global $sInFileDirDefault = "\\alpha3\E\RiosBay\Ins"
Global $sOutFileDirDefault = "\\alpha3\E\RiosBay\Outs"

Global $sInFileDir, $sOutFileDir

Dim $hGUI, $hTab, $hInFolder, $hInFile, $hInFileLabel, $hDefault_Button, $hApply_Button, $hChooseFileButton, $hInRemarksList, _
		$hOutLabel, $hOut, $hOutFile, $hCreateAllOutsButton

fuMainGUI()
; create GUI and tabs
Func fuMainGUI()

	$hGUI = GUICreate("Vparse v" & _GetVersion(), 600, 500, Default, Default, BitOR($GUI_SS_DEFAULT_GUI, $WS_MAXIMIZEBOX, $WS_SIZEBOX))
	GUISetOnEvent($GUI_EVENT_CLOSE, "On_Close") ; Run this function when the main GUI [X] is clicked

	$hTab = GUICtrlCreateTab(5, 5, 592, 490)
	GUICtrlSetResizing($hTab, $GUI_DOCKBORDERS)
	; tab 0
	GUICtrlCreateTabItem("Main")

	$hInFileLabel = GUICtrlCreateLabel("'In' File Location:", 14, 37)
	$hInFile = GUICtrlCreateInput("", 134, 35, 360, 20, $ES_READONLY)
	GUICtrlSetBkColor($hInFile, 0xFFFFFF)
	$hChooseFileButton = GUICtrlCreateButton("CHOOSE", 515, 35, 70, 20)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	$hOutLabel = GUICtrlCreateLabel("Number of Outs:", 14, 57)
	GUISetFont(10, $FW_BOLD)
	$hOut = GUICtrlCreateLabel("", 134, 57, 140, 22)
	GUISetFont(8.5, $FW_NORMAL)

	GUICtrlSetResizing($hInFileLabel, $GUI_DOCKMENUBAR)
	GUICtrlSetResizing($hInFile, $GUI_DOCKMENUBAR)
	GUICtrlSetResizing($hChooseFileButton, $GUI_DOCKMENUBAR)
	GUICtrlSetResizing($hOutLabel, $GUI_DOCKMENUBAR)
	GUICtrlSetResizing($hOut, $GUI_DOCKMENUBAR)

	$hInRemarksList = GUICtrlCreateListView("", 14, 80, 573, 380, BitOR($LVS_SHOWSELALWAYS, $LVS_REPORT, $LVS_NOSORTHEADER, $LVS_NOLABELWRAP))
	GUICtrlSetState($hInRemarksList, $GUI_DISABLE)

	GUICtrlSetResizing($hInRemarksList, $GUI_DOCKBORDERS)
	_GUICtrlListView_SetExtendedListViewStyle($hInRemarksList, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES))
	_GUICtrlListView_AddColumn($hInRemarksList, "Outs", @DesktopWidth)

	$hCreateAllOutsButton = GUICtrlCreateButton("SPLIT && SAVE Outs", 245, 465, 120, 22)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	GUICtrlSetState($hCreateAllOutsButton, $GUI_DISABLE)
	GUICtrlSetResizing($hCreateAllOutsButton, $GUI_DOCKSTATEBAR)

	; tab 1
	GUICtrlCreateTabItem("Settings")

	GUICtrlCreateLabel("Default In Directory", 35, 45)
	GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR)
	$hInFolder = GUICtrlCreateInput("", 35, 65, 320, 20)
	GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR)
	$sInFileDir = fuGetRegValsForSettings("In", $sInFileDirDefault)
	GUICtrlSetData($hInFolder, $sInFileDir)

	GUICtrlCreateLabel("Directory to Save Out Files", 35, 105)
	GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR)
	$hOutFile = GUICtrlCreateInput("", 35, 125, 320, 20)
	GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR)
	$sOutFileDir = fuGetRegValsForSettings("Out", $sOutFileDirDefault)
	GUICtrlSetData($hOutFile, $sOutFileDir)

	$hDefault_Button = GUICtrlCreateButton("Default", 400, 225, 75)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR)
	$hApply_Button = GUICtrlCreateButton("Apply", 485, 225, 75)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR)
	GUICtrlCreateTabItem("") ; end tabitem definition

	GUISetState()

	; Run the GUI until the dialog is closed
	While 1
		Sleep(10)
	WEnd
EndFunc   ;==>fuMainGUI

Func On_Close()
	Switch @GUI_WinHandle ; See which GUI sent the CLOSE message
		Case $hGUI
			Exit ; If it was this GUI - we exit <<<<<<<<<<<<<<<
	EndSwitch
EndFunc   ;==>On_Close

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

Func On_Click()
	Switch @GUI_CtrlId ; See wich item sent a message
		Case $hChooseFileButton
			Local $sFileOpenDialog = FileOpenDialog("Select In File", $sInFileDir & "\", "In (*.In)", $FD_FILEMUSTEXIST + $FD_PATHMUSTEXIST, "", $hGUI)
			GUICtrlSetData($hInFile, $sFileOpenDialog)
;			Local $aInRemarksData = fuReadInDoc($sFileOpenDialog)
;			If UBound($aInRemarksData) > 0 Then fuPopulateListView($aInRemarksData)
		Case $hDefault_Button
			$sInFileDir = $sInFileDirDefault
			GUICtrlSetData($hInFolder, $sInFileDir)
			$sOutFileDir = $sOutFileDirDefault
			GUICtrlSetData($hOutFile, $sOutFileDir)
			ContinueCase
		Case $hApply_Button
			fuApplySettingsValue($hInFolder, "In")
			fuApplySettingsValue($hOutFile, "Out")
		Case $hCreateAllOutsButton
			Local $aAllRemarks = _GUICtrlListView_CreateArray($hInRemarksList)
;			fuCreateOuts($aAllRemarks)
	EndSwitch
EndFunc   ;==>On_Click

; #FUNCTION# ====================================================================================================================
; Name ..........: _GUICtrlListView_CreateArray
; Description ...: Creates a 2-dimensional array from a listview.
; Syntax ........: _GUICtrlListView_CreateArray($hListView[, $sDelimeter = '|'])
; Parameters ....: $hListView           - Control ID/Handle to the control
;                  $sDelimeter          - [optional] One or more characters to use as delimiters (case sensitive). Default is '|'.
;				   $bAllItems			- [optional]
; Return values .: Success - The array returned is two-dimensional and is made up of the following:
;                                $aArray[0][0] = Number of rows
;                                $aArray[0][1] = Number of columns
;                                $aArray[0][2] = Delimited string of the column name(s) e.g. Column 1|Column 2|Column 3|Column nth

;                                $aArray[1][0] = 1st row, 1st column
;                                $aArray[1][1] = 1st row, 2nd column
;                                $aArray[1][2] = 1st row, 3rd column
;                                $aArray[n][0] = nth row, 1st column
;                                $aArray[n][1] = nth row, 2nd column
;                                $aArray[n][2] = nth row, 3rd column
; Author ........: guinness, sjohnson
; Remarks .......: GUICtrlListView.au3 should be included.
; ===============================================================================================================================
Func _GUICtrlListView_CreateArray($hListView, $sDelimeter = '|', $bAllItems = True)
	Local $iColumnCount = _GUICtrlListView_GetColumnCount($hListView), $iDim = 0, $iItemCount = 0
	Local $aiListIndices[1]
	$iItemCount = ($bAllItems) ? (_GUICtrlListView_GetItemCount($hListView)) : (_GUICtrlListView_GetSelectedCount($hListView))
	If $bAllItems Then
		$aiListIndices[0] = $iItemCount
		For $a = 0 To $iItemCount - 1
			_ArrayAdd($aiListIndices, $a)
		Next
	Else
		$aiListIndices = _GUICtrlListView_GetSelectedIndices($hListView, True)
	EndIf

	If $iColumnCount < 3 Then
		$iDim = 3 - $iColumnCount
	EndIf
	If $sDelimeter = Default Then
		$sDelimeter = '|'
	EndIf

	Local $aColumns = 0, $aReturn[$iItemCount + 1][$iColumnCount + $iDim] = [[$iItemCount, $iColumnCount, '']]
	For $i = 0 To $iColumnCount - 1
		$aColumns = _GUICtrlListView_GetColumn($hListView, $i)
		$aReturn[0][2] &= $aColumns[5] & $sDelimeter
	Next
	$aReturn[0][2] = StringTrimRight($aReturn[0][2], StringLen($sDelimeter))

	For $i = 1 To $iItemCount
		For $j = 0 To $iColumnCount - 1
			$aReturn[$i][$j] = _GUICtrlListView_GetItemText($hListView, $aiListIndices[$i], $j)
		Next
	Next
	Return SetError(Number($aReturn[0][0] = 0), 0, $aReturn)
EndFunc   ;==>_GUICtrlListView_CreateArray
