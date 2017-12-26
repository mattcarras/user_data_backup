#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         Matt C

 Script Function:
   Backup user files using Robocopy optionally with multiple passes.

 Changelog
   12-16-2017 - File created
   12-22-2017 - Added automatic drive mapping to save log file (default: disabled) and prompting at startup for WO ID #
			  - Also made message boxes a bit more descriptive.
#ce ----------------------------------------------------------------------------

#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#include <GuiConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ButtonConstants.au3>
#include <FontConstants.au3>
#include <StaticConstants.au3>
#include <ColorConstants.au3>
#include <FileConstants.au3>
;#include "_GetUserDataSizes.au3" ; My user-defined script
#include "_DriveMapEx.au3" ; My user-defined script

; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; ~         Debugging          ~
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Opt('MustDeclareVars', 1)
#RequireAdmin

Local Const $Title = 'Copy User Data with Logs'
Local Const $Version = '0.2'
Local Const $Author = 'by MattC'

; Sizing constants
Local Const $MainWidthBase = 550
Local Const $MainHeightBase = 125
Local Const $LabelWidthBase = 350
Local Const $LabelHeightBase = 30
Local Const $InputWidthBase = 300
Local Const $InputHeightBase = 20
Local Const $LabelDefaultFontWidth = 6
Local Const $ButtonWidthBase = 65
Local Const $ButtonHeightBase = 20
Local Const $ButtonDefaultFontWidth = 6

; Program logic constants
Local Const $RobocopyExe = 'robocopy.exe'
Local Const $RobocopyWorkingDir = @WindowsDir & '\system32'
Local Const $RobocopyParameters = '/E /ZB /XJ /XJD /SL'
Local Const $RobocopyParametersPass1 = '/R:1 /W:5'
Local Const $RobocopyParametersPass2 = '/R:10 /W:10'
Local Const $RobocopyParametersPass3 = ''
Local Const $RobocopyParametersNetworkPath = '/IPG:5 /FFT'
Local Const $RobocopyLogParametersPass1 = '/TEE /LOG+:"' ; quote string must be terminated
Local Const $RobocopyLogParametersPass2 = '/TEE /LOG+:"' ; " " "
Local Const $RobocopyLogParametersPass3 = '/TEE /LOG+:"' ; " " "

Local Const $ResultsLogPrefix = 'WO'
Local Const $ResultsLogSuffix = '_robocopy.log'
Local Const $ResultsLogDirDefault = @DesktopDir
Local Const $LogFileProg = 'notepad.exe'

; If given, maps drive using given username and password
; Set to Null or False to disable
Local Const $StrDriveMapHost = Null
Local Const $StrDriveMapRelPath = Null ; requires at least '\'
Local Const $StrDriveMapPath = Null ; Example: '\\' & $StrDriveMapHost & $StrDriveMapRelPath
Local Const $StrDriveMapUsername = Null
Local Const $StrDriveMapPassword = Null ; If Null or False, but username is not, then will be requested

; Program text
Local Const $TextEnterSource = "Enter source folder..."
Local Const $TextEnterDest = "Enter destination..."
Local Const $TextStartCopying = "Start Copying"

; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
; ~    >>> MAIN START <<<      ~
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


Local $MainWidth = $MainWidthBase
Local $MainHeight = $MainHeightBase + ( 1 * $LabelHeightBase )

Opt("GUICoordMode", 1) ; Absolute
Local $hGUI = GUICreate ($Title&' v'&$Version, $MainWidth, $MainHeight )

; Title
Local $hLabel = GUICtrlCreateLabel($Title&' v'&$Version&' '&$Author, 5, -1, $MainWidth, $LabelHeightBase)
GUICtrlSetFont($hLabel, 10, $FW_BOLD)

If Not IsAdmin() Then
   MsgBox($MB_SYSTEMMODAL, $Title & ' - ERROR', "This program requires Administrator-level access.")
   Exit
EndIf

; GUICtrlCreateButton ( "text", left, top [, width [, height [, style = -1 [, exStyle = -1]]]] )

Local Const $LabelWidthCol1 = 45

Local $hLabel2 = GUICtrlCreateLabel("WO#: ", 5, $LabelHeightBase, $LabelWidthCol1, $LabelHeightBase, $SS_RIGHT)

Opt("GUICoordMode", 2) ; Relative
Local $hInputID = GUICtrlCreateInput( InputBox( @ScriptName, "Enter WO ID #", "", "", -1, 140), 3, -1, 150, $InputHeightBase)

Opt("GUICoordMode", 1) ; Absolute
Local $hLabel3 = GUICtrlCreateLabel("Source: ", 5, $LabelHeightBase * 2, $LabelWidthCol1, $LabelHeightBase, $SS_RIGHT)

Opt("GUICoordMode", 2) ; Relative
Local $hInputSrc = GUICtrlCreateInput ( $TextEnterSource, 3, -1, $InputWidthBase, $InputHeightBase)
Local $idBrowseSrc = GUICtrlCreateButton("Browse", 3, -1, $ButtonWidthBase, $ButtonHeightBase)
; SMART errors or other problems
Local $idWarningSrc = GUICtrlCreateButton("**WARNING**", 3, -1, 100, $ButtonHeightBase, $BS_FLAT)
GUICtrlSetFont($idWarningSrc, 10, $FW_BOLD)
GUICtrlSetColor($idWarningSrc, $COLOR_RED)
GUICtrlSetState($idWarningSrc, $GUI_HIDE)

Opt("GUICoordMode", 1) ; Absolute
Local $hLabel4 = GUICtrlCreateLabel("Dest: ", 5, $LabelHeightBase * 3, $LabelWidthCol1, $LabelHeightBase, $SS_RIGHT)

Opt("GUICoordMode", 2) ; Relative
Local $hInputDst = GUICtrlCreateInput ( $TextEnterDest, 3, -1, $InputWidthBase, $InputHeightBase)
Local $idBrowseDst = GUICtrlCreateButton("Browse", 3, -1, $ButtonWidthBase, $ButtonHeightBase)
; SMART errors or other problems
Local $idWarningDst = GUICtrlCreateButton("**WARNING**", 3, -1, 100, $ButtonHeightBase, $BS_FLAT)
GUICtrlSetFont($idWarningDst, 10, $FW_BOLD)
GUICtrlSetColor($idWarningDst, $COLOR_RED)
GUICtrlSetState($idWarningDst, $GUI_HIDE)

Opt("GUICoordMode", 1) ; Absolute
Local $idStart = GUICtrlCreateButton( $TextStartCopying, 15, $LabelHeightBase * 4, $ButtonWidthBase *2, $ButtonHeightBase + 2)
GUICtrlSetFont($idStart, 10, $FW_BOLD)
GUICtrlSetState($idStart, $GUI_DISABLE) ; enabled when both source and dest are valid

; Show GUI
GUISetState(@SW_SHOW, $hGUI)

; --Event Looping--
; Loop until the user exits.
; GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
While 1
	Local $msg = GUIGetMsg()
	Switch $msg
	  Case $GUI_EVENT_CLOSE
		 ExitLoop
	  Case $idBrowseSrc
		 Local $path = FileSelectFolder( "Select source folder for copying", @ScriptDir )
		 If FileExists($path) Then
			GUICtrlSetData($hInputSrc, $path)
		 EndIf
		 Local $src = GUICtrlRead($hInputSrc), $dst = GUICtrlRead($hInputDst)
		 If StringLen($src) > 1 And StringLen($dst) > 1 And $src <> $TextEnterSource And $dst <> $TextEnterDest Then
			GUICtrlSetState($idStart, $GUI_ENABLE)
		 EndIf
	  Case $idBrowseDst
		 Local $path = FileSelectFolder( "Select destination", @ScriptDir )
		 If FileExists($path) Then
			GUICtrlSetData($hInputDst, $path)
		 EndIf
		 Local $src = GUICtrlRead($hInputSrc), $dst = GUICtrlRead($hInputDst)
		 If StringLen($src) > 1 And StringLen($dst) > 1 And $src <> $TextEnterSource And $dst <> $TextEnterDest And FileExists($src) Then
			GUICtrlSetState($idStart, $GUI_ENABLE)
		 EndIf
	  Case $idStart
		 Local $srcpath = GUICtrlRead($hInputSrc)
		 Local $dstpath = GUICtrlRead($hInputDst)
		 ; If we are given an ID #, prefix logfile with this along with datestamp. Example will look like:
		 Local $logfile = @MON & @MDAY & @YEAR & '_' & @HOUR & $ResultsLogSuffix
		 Local $inputID = GUICtrlRead($hInputID)
		 If IsString($inputID) Then
			$logfile = $ResultsLogPrefix & $inputID & '_' & $logfile
		 EndIf
		 Local $logfilefullpath = $ResultsLogDirDefault & '\' & $logfile
		 Local $logfilepath = $logfilefullpath ; in case it's on a mapped network share

		 ; Source must exist, but dest can be created
		 If Not FileExists($srcpath) Or DirGetSize($srcpath) < 0 Then
			MsgBox($MB_SYSTEMMODAL, @ScriptName & ' - ERROR', "Source [" & $srcpath & "] does not seem to exist or is not a directory.")
			; Return / Cancel / Abort
		 Else
			Local $dstOK = False
			If Not FileExists($dstpath) Then
			   Local $idResult = MsgBox($MB_SYSTEMMODAL+$MB_OKCANCEL, @ScriptName, "Destination [" & $dstpath & "] does not seem to exist. Create it?")
			   If $idResult<>$IDCANCEL Then
				  DirCreate($dstpath)
				  Sleep(500)
				  If Not FileExists($dstpath) Then
					 MsgBox($MB_SYSTEMMODAL, @ScriptName & ' - ERROR', "Cannot create destination [" & $dstpath & "]. Do you have the right permissions?")
					 ; Return / Cancel / Abort
				  Else
					 $dstOK = True
				  EndIf
			   EndIf
			ElseIf DirGetSize($dstpath) < 0 Then
			   MsgBox($MB_SYSTEMMODAL, @ScriptName & ' - ERROR', "Destination [" & $srcpath & "] is not a directory.")
			   ; Return / Cancel / Abort
			Else ; Everything OK with destination
			   $dstOK = True
			EndIf
			If $dstOK Then
			   ; Clear Start state, since we're copying
			   GUICtrlSetState($idStart, $GUI_DISABLE)
			   GUICtrlSetData($idStart, "Copying...")

			   Local $mainparameters = $RobocopyParameters
			   ; Check to see if source or destination is a network path. If so, append $RobocopyParametersNetworkPath to help decongest network and deal with non-Windows hosted shares
			   If DriveGetType( $srcpath ) == "Network" Or DriveGetType( $dstpath ) == "Network" Then
				  $mainparameters = $mainparameters & '  ' & $RobocopyParametersNetworkPath
			   EndIf

			   ; Try to map log file location, if it's a network path and we need to map it
			   Local $destDrive
			   If $StrDriveMapHost And $StrDriveMapPath Then
				  ; Try to map drive. Assume we can ping server to check network.
				  $destDrive = _DriveMapEx( $StrDriveMapHost, $StrDriveMapPath, $StrDriveMapUsername, $StrDriveMapPassword )
				  If $destDrive Then
					  ; Change the results destination if successfully mapped
					 $logfilefullpath = $destDrive & '\' & $logfile
					 $logfilepath = $StrDriveMapPath & '\' & $logfile
				  Else
					 MsgBox( $MB_SYSTEMMODAL, @ScriptName, "Using default log file destination [" & $logfilefullpath & "]")
				  EndIf
			   EndIf ; If we need to map a drive

			   ; Write datestamp and ID # to file
			   Local $hFile = FileOpen( $logfilefullpath, BitOr($FO_OVERWRITE, $FO_CREATEPATH) )
			   If $hFile <> -1 Then
				  FileWrite( $hFile, "Log File: " & $logfile & @CRLF & '*****************************************************' & @CRLF)
			   EndIf
			   FileClose( $hFile )

			   ; Get starting size of source and destination
			   Local $iSrcSizeBytes = DirGetSize( $srcpath )
			   Local $iDstSizeBytesStart = DirGetSize( $dstpath )

			   ; Start Robocopy
			   Local $sFullCmd = '"' & $RobocopyExe & '" "' & $srcpath & '" "' & $dstpath & '" ' & $mainparameters & ' ' & $RobocopyParametersPass1 & ' ' & $RobocopyLogParametersPass1 & $logfilefullpath & '"'
			   RunWait($sFullCmd, $RobocopyWorkingDir)

			   Sleep(100)

			   ; Run second pass for any missed files
			   ;RunWait('"' & $RobocopyExe & '" "' & $srcpath & '" "' & $dstpath & '" ' & $mainparameters & ' ' & $RobocopyParametersPass2 & ' ' & $RobocopyLogParametersPass2 & $logfilefullpath & '"', $RobocopyWorkingDir)

			   ; Get final size of dest
			   Local $iDstSizeBytesEnd = DirGetSize( $dstpath )

			   ; Write user data sizes to log file then open log file
			   Local $bLogWriteSuccess = False
			   If FileExists( $logfilefullpath ) Then
				   $hFile = FileOpen( $logfilefullpath, $FO_APPEND )
				   If $hFile <> -1 Then
					  If FileWrite( $hFile, @CRLF & '*****************************************************' & @CRLF & 'Source Folder Size in Bytes: ' & $iSrcSizeBytes & @CRLF & 'Destination Folder Size in Bytes (Before Copy): ' & $iDstSizeBytesStart & @CRLF & 'Destination Folder Size in Bytes (After Copy): ' & $iDstSizeBytesEnd ) Then
						 $bLogWriteSuccess = True
					  EndIf
				   EndIf
				   FileClose( $hFile )
			   Else
				  MsgBox( $MB_SYSTEMMODAL, @ScriptName & ' - ERROR', "Log file [" & $logfilepath & "] could not be written to." )
			   EndIf

			   ; Release drive mapping if it exists
			   _DriveMapDelEx( $destDrive )

			   ; Display log file or results (using UNC path if saved to a network drive)
			   If $bLogWriteSuccess Or FileExists( $logfilepath ) Then
				  If Not Run('"' & $LogFileProg & '" "' & $logfilepath & '"') Then
					 MsgBox( $MB_SYSTEMMODAL, @ScriptName, "Log file saved to [" & $logfilepath & "]" )
				  EndIf
			   EndIf

			   ; In case destination is an existing user directory with files in it
			   Local $iDifferenceBytes = $iSrcSizeBytes - ($iDstSizeBytesEnd - $iDstSizeBytesStart)
			   If $iDifferenceBytes > 0 And $iSrcSizeBytes <> $iDstSizeBytesEnd Then
				  MsgBox( $MB_SYSTEMMODAL, @ScriptName & ' - WARNING', "WARNING: Destination is smaller than Source. Not all files may have been copied. (" & $iDifferenceBytes & " bytes size difference)" & @CRLF & "Check logfile for more information." )
			   ElseIf $iSrcSizeBytes == $iDstSizeBytesEnd And $iSrcSizeBytes <> $iDstSizeBytesStart Then
				  MsgBox( $MB_SYSTEMMODAL, @ScriptName & ' - WARNING', "WARNING: Destination and source are the same size before and after the copy attempt. Files were probably not copied." & @CRLF & "Check logfile for more information." )
			   Else
				  Local $sRoundedResults = $iSrcSizeBytes & " bytes"
				  If $iSrcSizeBytes >= 1073741824 Then
					 $sRoundedResults = $sRoundedResults & " (" & Round($iSrcSizeBytes / 1073741824, 2) & 'GB)'
				  ElseIf $iSrcSizeBytes >= 1048576 Then
					$sRoundedResults = $sRoundedResults & " (" & Round($iSrcSizeBytes / 1048576, 2) & 'MB)'
				  ElseIf $iSrcSizeBytes >= 1024 Then
					 $sRoundedResults = $sRoundedResults & " (" & Round($iSrcSizeBytes / 1024, 2) & 'KB)'
				  EndIf
				  MsgBox( $MB_SYSTEMMODAL, @ScriptName & ' - Success', "Files taking up " & $sRoundedResults & " successfully copied to [" & $dstpath & "]" & @CRLF & "Check [" & $logfilepath & "] for more information." )
			   EndIf

			   GUICtrlSetData($idStart, $TextStartCopying)
			EndIf ; if destination OK
		 EndIf ; if paths exist

    EndSwitch ; $msg
 WEnd

 Exit

