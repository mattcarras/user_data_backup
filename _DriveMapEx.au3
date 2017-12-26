#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         Matt C

 Script Function:
   User defined functions to map given remote share with username and password and error-checking.

 Changelog
   12-23-2017 - File created.
#ce ----------------------------------------------------------------------------

#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>

; User defined function to map given remote share with username and password and error-checking.
; $sHost - Hostname or IP for remote share
; $sPath - Absolute UNC path, including hostname or IP
; $sUsername - Username for share (optional)
; $sPassword - Password for share (optional). If username is given, will prompt.
; $sDriveLetter - Drive letter to map to (default: whatever drive letter is free)
Func _DriveMapEx($sHost, $sPath, $sUsername = Null, $sPassword = Null, $sDriveLetter = "*")
   Local $sMappedDrive = False

   If Not $sHost Or Not $sPath Then
	  MsgBox( $MB_SYSTEMMODAL, Default, "Invalid host or path given." )
	  Return False
   EndIf

   ; Assume we can ping server to check network.
   If Not Ping( $sHost ) Then
	  MsgBox( $MB_SYSTEMMODAL, Default, "Cannot reach server [" & $sHost & "]")
	  Return False
   EndIf

   If $sUsername And Not $sPassword Then
	  $sPassword = InputBox( Default, "Enter password for " & $sUsername & '@' & $sHost & " (or hit enter for blank):", "", '*', -1, 140)
   EndIf

   $sMappedDrive = DriveMapAdd($sDriveLetter, $sPath, $DMA_DEFAULT, $sUsername, $sPassword)
   Local $error = @error
   Local $extended = @extended
   If $error <> 0 Or StringLen($sMappedDrive) < 1 Then
	  Local $sError
	  Switch ($error)
		 Case 2
			$sError = "Access to the remote share was denied."
		 Case 3
			$sError = "The device is already assigned."
		 Case 4
			$sError = "Invalid device name."
		 Case 5
			$sError = "Invalid remote share."
		 Case 6
			$sError = "Invalid password."
		 Case Else
			$sError = "Undefined error. WinAPI Error=" & $extended
	  EndSwitch

	  MsgBox( $MB_SYSTEMMODAL, Default, "Cannot map [" & $sPath & "] (Error: " & $sError & ")")
	  Return False
   EndIf ; if error

   ; Return mapped drive letter or False if unsuccessful
   Return $sMappedDrive
EndFunc

; Try to gracefully unmap / delete a mapped drive letter.
Func _DriveMapDelEx( $sDrive )
   ; Sanity check. No return value.
   If Not $sDrive Or Not IsString($sDrive) Or Not StringLen($sDrive) > 0 Then
	  Return
   EndIf

   ; Try multiple times to unmap
   For $i = 1 To 5
	  Sleep(1000)
	  If DriveMapDel( $sDrive ) <> 0 Then ExitLoop
   Next
   Sleep(1000)

   DriveMapGet( $sDrive )
   If @error <> 1 Then
	  MsgBox( $MB_SYSTEMMODAL, Default, "Error: Could not unmap " & $sDrive)
	  Return False
   EndIf

   Return True
EndFunc