Global Const $MB_OK = 0
Global Const $MB_ICONERROR = 16
Global Const $MB_ICONINFORMATION = 64
Global Const $tagRECT = "struct;long Left;long Top;long Right;long Bottom;endstruct"
Global Const $tagREBARBANDINFO = "uint cbSize;uint fMask;uint fStyle;dword clrFore;dword clrBack;ptr lpText;uint cch;" & "int iImage;hwnd hwndChild;uint cxMinChild;uint cyMinChild;uint cx;handle hbmBack;uint wID;uint cyChild;uint cyMaxChild;" & "uint cyIntegral;uint cxIdeal;lparam lParam;uint cxHeader" &((@OSVersion = "WIN_XP") ? "" : ";" & $tagRECT & ";uint uChevronState")
Global Const $FT_MODIFIED = 0
Global Const $FO_READ = 0
Global Const $HGDI_ERROR = Ptr(-1)
Global Const $INVALID_HANDLE_VALUE = Ptr(-1)
Global Const $KF_EXTENDED = 0x0100
Global Const $KF_ALTDOWN = 0x2000
Global Const $KF_UP = 0x8000
Global Const $LLKHF_EXTENDED = BitShift($KF_EXTENDED, 8)
Global Const $LLKHF_ALTDOWN = BitShift($KF_ALTDOWN, 8)
Global Const $LLKHF_UP = BitShift($KF_UP, 8)
Global Const $tagOSVERSIONINFO = 'struct;dword OSVersionInfoSize;dword MajorVersion;dword MinorVersion;dword BuildNumber;dword PlatformId;wchar CSDVersion[128];endstruct'
Global Const $__WINVER = __WINVER()
Func __WINVER()
Local $tOSVI = DllStructCreate($tagOSVERSIONINFO)
DllStructSetData($tOSVI, 1, DllStructGetSize($tOSVI))
Local $aRet = DllCall('kernel32.dll', 'bool', 'GetVersionExW', 'struct*', $tOSVI)
If @error Or Not $aRet[0] Then Return SetError(@error, @extended, 0)
Return BitOR(BitShift(DllStructGetData($tOSVI, 2), -8), DllStructGetData($tOSVI, 3))
EndFunc
Global Const $GUI_EVENT_CLOSE = -3
Global Const $GUI_ENABLE = 64
Global Const $GUI_DISABLE = 128
Global Const $GUI_DOCKMENUBAR = 0x0220
Global Const $GUI_DOCKSTATEBAR = 0x0240
Global Const $GUI_DOCKBORDERS = 0x0066
Global Const $GUI_BKCOLOR_TRANSPARENT = -2
Global Const $ES_CENTER = 1
Global Const $ES_MULTILINE = 4
Global Const $ES_NUMBER = 8192
Global Const $COLOR_RED = 0xFF0000
Global Const $COLOR_WHITE = 0xFFFFFF
Global Const $WS_MAXIMIZEBOX = 0x00010000
Global Const $WS_MINIMIZEBOX = 0x00020000
Global Const $WS_SIZEBOX = 0x00040000
Global Const $WS_SYSMENU = 0x00080000
Global Const $WS_HSCROLL = 0x00100000
Global Const $WS_VSCROLL = 0x00200000
Global Const $WS_CAPTION = 0x00C00000
Global Const $WS_POPUP = 0x80000000
Global Const $GUI_SS_DEFAULT_GUI = BitOR($WS_MINIMIZEBOX, $WS_CAPTION, $WS_POPUP, $WS_SYSMENU)
Global Const $FW_NORMAL = 400
Global Const $FW_BOLD = 700
Opt("GUIOnEventMode", 1)
Global $sInFileDirDefault = "\\hqnapdcm0734\OFR\e_cfr"
Global $sOutFileDirDefault = "\\hqnapdcm0734\OFR\e_cfr\Apps\ECFRDATE"
Global $sInFileDir, $sOutFileDir
Global Const $COLOR_GPOTEAL = 0x3b80a1
Dim $hGUI, $hTab, $hInFolder, $hInTitle, $hInTitleLabel, $hInVolume, $hInVolumeLabel, $hDefault_Button, $hApply_Button, $hLoadFileButton, $hInRemarksList, $hOutLabel, $hOut, $hOutFile, $hCreateAllOutsButton
fuMainGUI()
Func fuMainGUI()
$hGUI = GUICreate("Vparse v" & _GetVersion(), 600, 500, Default, Default, BitOR($GUI_SS_DEFAULT_GUI, $WS_MAXIMIZEBOX, $WS_SIZEBOX))
GUISetOnEvent($GUI_EVENT_CLOSE, "On_Close")
$hTab = GUICtrlCreateTab(5, 5, 592, 490)
GUICtrlSetResizing($hTab, $GUI_DOCKBORDERS)
$cTab_0 = GUICtrlCreateTabItem("Main")
GUISetFont(15, $FW_NORMAL)
$hInTitleLabel = GUICtrlCreateLabel("Title:", 17, 40)
$hInTitle = GUICtrlCreateInput("", 62, 37, 60, 33, BitOR($ES_NUMBER, $ES_CENTER))
GUICtrlCreateUpdown($hInTitle)
GUICtrlSetLimit(-1, 50, 1)
$hInVolumeLabel = GUICtrlCreateLabel("Volume:", 125, 40)
$hInVolume = GUICtrlCreateInput("", 200, 37, 60, 33, BitOR($ES_NUMBER, $ES_CENTER))
GUICtrlCreateUpdown($hInVolume)
GUICtrlSetLimit(-1, 37, 1)
$hOutLabel = GUICtrlCreateLabel("eCFR Date:", 270, 40)
GUISetFont(15, $FW_BOLD)
$hOut = GUICtrlCreateLabel("", 380, 40, 140, 35)
$hLoadFileButton = GUICtrlCreateButton("LOAD", 500, 37, 80, 33)
GUICtrlSetBkColor($hLoadFileButton, $COLOR_GPOTEAL)
GUICtrlSetColor($hLoadFileButton, $COLOR_WHITE)
GUICtrlSetOnEvent(-1, "On_Click")
GUISetFont(8.5, $FW_NORMAL)
GUICtrlSetResizing($hInTitleLabel, $GUI_DOCKMENUBAR)
GUICtrlSetResizing($hInTitle, $GUI_DOCKMENUBAR)
GUICtrlSetResizing($hInVolumeLabel, $GUI_DOCKMENUBAR)
GUICtrlSetResizing($hInVolume, $GUI_DOCKMENUBAR)
GUICtrlSetResizing($hLoadFileButton, $GUI_DOCKMENUBAR)
GUICtrlSetResizing($hOutLabel, $GUI_DOCKMENUBAR)
GUICtrlSetResizing($hOut, $GUI_DOCKMENUBAR)
$hInRemarksList = GUICtrlCreateEdit("", 14, 80, 573, 380, BitOR($ES_MULTILINE, $WS_HSCROLL, $WS_VSCROLL))
GUICtrlSetResizing($hInRemarksList, $GUI_DOCKBORDERS)
$hCreateAllOutsButton = GUICtrlCreateButton("PROCESS FILES", 245, 465, 120, 22)
GUICtrlSetBkColor($hCreateAllOutsButton, $COLOR_GPOTEAL)
GUICtrlSetColor($hCreateAllOutsButton, $COLOR_WHITE)
GUICtrlSetOnEvent(-1, "On_Click")
GUICtrlSetState($hCreateAllOutsButton, $GUI_DISABLE)
GUICtrlSetResizing($hCreateAllOutsButton, $GUI_DOCKSTATEBAR)
$cTab_1 = GUICtrlCreateTabItem("Settings")
GUICtrlCreateLabel("Default Directory", 35, 45)
GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR)
$hInFolder = GUICtrlCreateInput("", 35, 65, 320, 20)
GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR)
$sInFileDir = fuGetRegValsForSettings("Dir", $sInFileDirDefault)
GUICtrlSetData($hInFolder, $sInFileDir)
GUICtrlCreateLabel("Effective Date File Location", 35, 105)
GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR)
$hOutFile = GUICtrlCreateInput("", 35, 125, 320, 20)
GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR)
$sOutFileDir = fuGetRegValsForSettings("Date", $sOutFileDirDefault)
GUICtrlSetData($hOutFile, $sOutFileDir)
$hDefault_Button = GUICtrlCreateButton("Default", 400, 225, 75)
GUICtrlSetBkColor($hDefault_Button, $COLOR_GPOTEAL)
GUICtrlSetColor($hDefault_Button, $COLOR_WHITE)
GUICtrlSetOnEvent(-1, "On_Click")
GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR)
$hApply_Button = GUICtrlCreateButton("Apply", 485, 225, 75)
GUICtrlSetBkColor($hApply_Button, $COLOR_GPOTEAL)
GUICtrlSetColor($hApply_Button, $COLOR_WHITE)
GUICtrlSetOnEvent(-1, "On_Click")
GUICtrlSetResizing(-1, $GUI_DOCKMENUBAR)
GUICtrlCreateTabItem("")
GUISetState()
While 1
Sleep(5)
WEnd
EndFunc
Func On_Close()
Switch @GUI_WinHandle
Case $hGUI
Exit
EndSwitch
EndFunc
Func fuGetRegValsForSettings($sFolder, $DefaultFolder)
Local $sRegValue
$sRegValue = RegRead("HKEY_CURRENT_USER\Software\USGPO\PED\Vparse", $sFolder)
If $sRegValue = "" Then
RegWrite("HKEY_CURRENT_USER\Software\USGPO\PED\Vparse", $sFolder, "REG_SZ", $DefaultFolder)
Return $DefaultFolder
Else
Return $sRegValue
EndIf
EndFunc
Func fuApplySettingsValue($hGUI, $sFolder)
Local $cInputVal = GUICtrlRead($hGUI)
$cInputVal = StringRegExpReplace($cInputVal, '\\* *$', '')
If Not FileExists($cInputVal) Then
MsgBox(16, "Location invalid", $sFolder & " location does not exists. Enter a valid path to it.")
Else
If Not RegWrite("HKEY_CURRENT_USER\Software\USGPO\PED\Vparse", $sFolder, "REG_SZ", $cInputVal) Then
MsgBox(16, "Could not be saved", $sFolder & " location could not be saved, Error #" & @error)
EndIf
EndIf
GUICtrlSetData($hGUI, $cInputVal)
Return
EndFunc
Func _GetVersion()
If @Compiled Then
Return FileGetVersion(@AutoItExe)
Else
Return IniRead(@ScriptFullPath, "FileVersion", "#AutoIt3Wrapper_Res_Fileversion", "0.0.0.0")
EndIf
EndFunc
Func On_Click()
Switch @GUI_CtrlId
Case $hLoadFileButton
GUICtrlSetBkColor($hInTitle, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetBkColor($hInVolume, $GUI_BKCOLOR_TRANSPARENT)
Local $iTitleNum = GUICtrlRead($hInTitle)
Local $iVolumeNum = GUICtrlRead($hInVolume)
If $iTitleNum < 1 Or $iTitleNum > 50 Then
GUICtrlSetBkColor($hInTitle, $COLOR_RED)
MsgBox($MB_ICONERROR, "Title Number Out of Range", "Title Number Should Be Between 1 and 50 !!!")
EndIf
If $iVolumeNum < 1 Or $iVolumeNum > 37 Then
GUICtrlSetBkColor($hInVolume, $COLOR_RED)
MsgBox($MB_ICONERROR, "Volume Number Out of Range", "Volume Number Should Be Between 1 and 37 !!!")
EndIf
Local $aEffDate = FileGetTime($sOutFileDir, $FT_MODIFIED)
GUICtrlSetData($hOut, $aEffDate[1] & "/" & $aEffDate[2] & "/" & $aEffDate[0])
Local $sFilePath = $sInFileDir & '\' & Number($iTitleNum) & '\' & Number($iTitleNum) & 'V' & Number($iVolumeNum) & ".TXT"
Local $hFileOpen = FileOpen($sFilePath, $FO_READ)
If $hFileOpen = -1 Then
MsgBox($MB_ICONERROR, "Error Reading File", "An error occurred when reading the file.")
Return False
EndIf
Local $sFileRead = FileRead($hFileOpen)
FileClose($hFileOpen)
GUICtrlSetData($hInRemarksList, $sFileRead)
GUICtrlSetState($hCreateAllOutsButton, $GUI_ENABLE)
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
EndFunc
