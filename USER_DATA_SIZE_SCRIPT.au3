#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         Matt C

 Script Function:
   Determine user data size of non-hidden folders and non-system files in Users directory.

 Changelog
   12-19-2017 - File created
   12-20-2017 - Moved main function into _GetUserDataSizes.au3 for another script
   12-22-2017 - Added automatic drive mapping to save log file (default: disabled)
   12-23-2017 - Fixed drive mapping code
#ce ----------------------------------------------------------------------------

#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <WinAPIShPath.au3> ; _WinAPI_PathIsRoot ( $sFilePath )
#include "_GetUserDataSizes.au3" ; My user-defined script
#include "_DriveMapEx.au3" ; My user-defined script

Local Const $StrResultsFileNamePrefix = "WO"
Local Const $StrResultsFileNameBase = 'USER_DATA_FOLDER_SIZES.txt'
Local Const $StrResultsDestDefault = @DesktopDir

; If given, maps drive using given username and password
; Set to Null or False to disable
Local Const $StrDriveMapHost = Null
Local Const $StrDriveMapRelPath = Null ; requires at least '\'
Local Const $StrDriveMapPath = Null ; Example: '\\' & $StrDriveMapHost & $StrDriveMapRelPath
Local Const $StrDriveMapUsername = Null
Local Const $StrDriveMapPassword = Null ; If Null or False, but username is not, then will be requested

; Require Admin just to be safe.
#RequireAdmin

If Not IsAdmin() Then
   MsgBox($MB_SYSTEMMODAL, @ScriptName, "Error: This program requires Administrator-level access.")
   Exit -1
EndIf

Local $resultsFilename = $StrResultsFileNamePrefix & InputBox( @ScriptName, "Enter WO ID #", "", "", -1, 140) & "_" & @MON & @MDAY & @YEAR & '_' & @HOUR & '_' & $StrResultsFileNameBase
; In case we successfully map
Local $resultsDest = $StrResultsDestDefault & '\' & $resultsFilename
Local $destDrive
If $StrDriveMapHost And $StrDriveMapPath Then
   ; Try to map drive. Assume we can ping server to check network.
   $destDrive = _DriveMapEx( $StrDriveMapHost, $StrDriveMapPath, $StrDriveMapUsername, $StrDriveMapPassword )
   If $destDrive Then
	   $resultsDest = $destDrive & '\' & $resultsFilename
   ;Else
	  ;MsgBox( $MB_SYSTEMMODAL, @ScriptName, "Using default log file destination [" & $StrResultsDestDefault & "]")
   EndIf
EndIf ; If we need to map a drive

Local $iReturn = 1

; Loop in case user selects wrong folder and then "No"
While 1
   Local $sSelectedPath = FileSelectFolder("Select drive or folder (IE, C:)...", @HomeDrive)
   Local $sResults = Null
   If Not $sSelectedPath Or StringLen($sSelectedPath) <= 0 Or Not FileExists( $sSelectedPath ) Then
	  ; User likely selected "Cancel"
	  $iReturn = 0
	  ExitLoop
   EndIf
   If Not _WinAPI_PathIsRoot( $sSelectedPath ) Then
	  Local $idResult = MsgBox( $MB_SYSTEMMODAL+$MB_YESNOCANCEL, @ScriptName, "[" & $sSelectedPath & "] is not the root of a drive or volume. Continue anyway?" )
	  If $idResult == $IDYES Then
		 $sResults =  _GetUserDataSizes( $sSelectedPath )
		 If Not $sResults Then $iReturn = -1
		 ExitLoop
	  ElseIf $idResult == $IDCANCEL Then
		 $iReturn = 0
		 ExitLoop
	  EndIf
   Else ; is a root path, so append
	  $sSelectedPath = $sSelectedPath & 'Users'
	  If FileExists( $sSelectedPath ) Then
		 $sResults =  _GetUserDataSizes( $sSelectedPath )
		 If Not $sResults Then $iReturn = -1
	  Else ; in case <drive>:\Users does not exist
		  MsgBox( $MB_SYSTEMMODAL, @ScriptName, "Error: [" & $sSelectedPath & "] does not exist." )
	  EndIf
	  ExitLoop
   EndIf ; path is root or not
WEnd

; If we have results, write to log file
Local $bWriteSuccess = False
If $sResults And $iReturn Then
   Local $hFile = FileOpen( $resultsDest, BitOr($FO_OVERWRITE, $FO_CREATEPATH) )
   If $hFile <> -1 Then
	  If FileWrite( $hFile, $sResults ) Then
		 $bWriteSuccess = True
	  Else
		 $iReturn = -1
	  EndIf
   EndIf
   FileClose( $hFile ) ; just in case
EndIf

; Release drive mapping if it exists
If $destDrive And StringLen($destDrive) > 0 Then
   _DriveMapDelEx( $destDrive )
   $resultsDest = $StrDriveMapPath & '\' & $resultsFilename ; for logfile location unmapped
EndIf

; open logfile and display success message
If $iReturn Then
   If Not Run( 'notepad.exe "' & $resultsDest & '"', @SystemDir ) Then
	  MsgBox($MB_SYSTEMMODAL, "", $sResults)
   EndIf
   MsgBox( $MB_SYSTEMMODAL, @ScriptName, "User data sizes log saved to [" & $resultsDest & "]")
EndIf

; Exit with given return error code (-1 on error, 0 on cancel, 1 otherwise)
Exit $iReturn