#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         Matt C

 Script Function:
   User-defined function to return string of user data sizes of non-hidden folders and non-system files in Users directory.

 Changelog
   12-20-2017 - File created
#ce ----------------------------------------------------------------------------

#include <MsgBoxConstants.au3>
#include <FileConstants.au3>

; Returns string of total size of all user folders, excluding system and hidden files and folders 1 folder deep
; If there's an error returns False
Func _GetUserDataSizes( $sPath )
   Local $isInRootUsersDir = False ; if true then we nest 1 level deep
   Local $sResults = "*********************************", $iTotalSize = 0
   Local $hSearch = FileFindFirstFile ( $sPath & '\*' )
   If $hSearch == -1 Then
	   MsgBox($MB_SYSTEMMODAL, "", "Error: Invalid or empty path [" & $sPath & "]")
	   Return False
   EndIf

   ; Determine if this is the Users directory or any 'ol directory by splitting the path, looking for Users or Users\
   Local $aStrings = StringSplit($sPath, '\' )
   If Not @error And IsArray( $aStrings ) And $aStrings[0] > 0 And ( ($aStrings[ $aStrings[0] ] == '\' And $aStrings[0] > 1 And StringUpper($aStrings[ $aStrings[0] -1 ]) == 'USERS' ) Or StringUpper($aStrings[ $aStrings[0] ]) == 'USERS') Then
	  $isInRootUsersDir = True
   EndIf

   ProgressOn("Tabulating user data size", "Calculating folder sizes...")
   Local $hTimer = TimerInit()
   Local Const $iTimerMaxMs = 240000 ; 4 minutes (estimate)
   While 1
	  Local $sFileName = FileFindNextFile($hSearch, 1) ; get file attributes in @extended
	  Local $sFileAttribs = @extended
	  If @error Then ExitLoop
	  Local $sFullPath = $sPath & '\' & $sFileName
      Local $isDir = $sFileAttribs And StringInStr($sFileAttribs, "D")
	  If $isInRootUsersDir And $isDir Then
		  ; Nest into this folder, but only if we are in the root Users folder/directory
		  Local $iTotalSizeThisFolder = 0
		  Local $hSearch2 = FileFindFirstFile ( $sFullPath & '\*' )
		  If $hSearch2 <> -1 Then
			While 1
			   Local $sFileName2 = FileFindNextFile($hSearch2, 1) ; get file attributes in @extended
			   If @error Then ExitLoop
			   Local $sFileAttribs2 = @extended
			   ProgressSet( TimerDiff( $hTimer ) / $iTimerMaxMs * 100 ) ; update progress meter
			   ; Skip hidden and system directories and files
			   If $sFileAttribs2 And Not StringInStr($sFileAttribs2, "H") And Not StringInStr($sFileAttribs2, "S") Then
				  Local $sFullPath2 = $sFullPath & '\' & $sFileName2
				  If StringInStr($sFileAttribs2, "D") Then ; is directory
					 $iTotalSizeThisFolder = $iTotalSizeThisFolder + DirGetSize( $sFullPath2 )
				  Else
					 $iTotalSizeThisFolder = $iTotalSizeThisFolder + FileGetSize( $sFullPath2 )
				  EndIf
			   EndIf ; if is not hidden or system folder or file
		   WEnd ; Endless loop, must be broken out of
		   FileClose($hSearch2) ; close previous search handle
	    EndIf ; if valid search handle
	    ; Add this to total and compute result string
	    $sResults = $sResults & @CRLF & $sFullPath & ', Size: '
	    If $iTotalSizeThisFolder >= 1073741824 Then
		 $sResults = $sResults & Round($iTotalSizeThisFolder / 1073741824, 2) & 'GB'
	    ElseIf $iTotalSizeThisFolder >= 1048576 Then
		 $sResults = $sResults & Round($iTotalSizeThisFolder / 1048576, 2) & 'MB'
	    ElseIf $iTotalSizeThisFolder >= 1024 Then
		 $sResults = $sResults & Round($iTotalSizeThisFolder / 1024, 2) & 'KB'
	    Else
		 $sResults = $sResults & $iTotalSizeThisFolder & 'B'
	    EndIf
		$iTotalSize = $iTotalSize + $iTotalSizeThisFolder
	  Else ; if we are NOT in the root Users directory, don't nest, but still only check non-hidden, non-system folders and files
		 Local $iTotalSizeThisFolder = 0
		 ; Skip hidden and system directories and files
		 If $sFileAttribs And Not StringInStr($sFileAttribs, "H") And Not StringInStr($sFileAttribs, "S") Then
			If $isDir Then ; is directory
			   $iTotalSizeThisFolder = $iTotalSizeThisFolder + DirGetSize( $sFullPath )
			Else
			   $iTotalSizeThisFolder = $iTotalSizeThisFolder + FileGetSize( $sFullPath )
			EndIf
			$sResults = $sResults & @CRLF & $sFullPath & ', Size: '
			If $iTotalSizeThisFolder >= 1073741824 Then
			   $sResults = $sResults & Round($iTotalSizeThisFolder / 1073741824, 2) & 'GB'
			ElseIf $iTotalSizeThisFolder >= 1048576 Then
			   $sResults = $sResults & Round($iTotalSizeThisFolder / 1048576, 2) & 'MB'
			ElseIf $iTotalSizeThisFolder >= 1024 Then
			   $sResults = $sResults & Round($iTotalSizeThisFolder / 1024, 2) & 'KB'
			Else
			   $sResults = $sResults & $iTotalSizeThisFolder & 'B'
			EndIf
			$iTotalSize = $iTotalSize + $iTotalSizeThisFolder
		 EndIf ; if is not hidden or system folder or file
	  EndIf ; If we start in the root Users folder or not
      ProgressSet( TimerDiff( $hTimer ) / $iTimerMaxMs * 100 ) ; update progress meter
   WEnd ; endless loop broken by @error
   FileClose($hSearch) ; close previous search handle
   ProgressSet(100, "Done", "Complete")
   Sleep(250)
   ProgressOff()

   ; Generate results string with date stamp
   $sResults = @MON & '/' & @MDAY & '/' & @YEAR & ' ' & @HOUR & ':' & @MIN & @CRLF & $sResults & @CRLF & "*********************************" & @CRLF & @CRLF & "Total Estimated User Data Size: " & Round($iTotalSize / 1073741824, 2) & 'GB' & @CRLF & @CRLF & 'Note: Excludes hidden and/or system folders.'

   Return $sResults
EndFunc