#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>

Local $driveArray = DriveGetDrive($DT_ALL)
If @error Then
    ; An error occurred when retrieving the drives.
    MsgBox($MB_SYSTEMMODAL, "", "It appears an error occurred.")
Else
    For $i = 1 To ($driveArray[0] - 1)
		$driveType = DriveGetType($driveArray[$i+1], $DT_DRIVETYPE)
		$driveLetter = StringUpper($driveArray[$i+1])
		If $driveType == "Removable" Then
			MsgBox($MB_SYSTEMMODAL, "", "Drive " & $i & "/" & ($driveArray[0] - 1) & ":" & @CRLF & $driveLetter & " " & DriveGetLabel($driveLetter) & " (" & $driveType & ")")
		EndIf
    Next
 EndIf

#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("Form1", 615, 437, 192, 124)
$Checkbox1 = GUICtrlCreateCheckbox("Checkbox1", 64, 80, 97, 17)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
	EndSwitch
WEnd
