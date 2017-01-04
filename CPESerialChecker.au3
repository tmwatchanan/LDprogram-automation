#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

Const $Version = "1.0"

Global $driveGroup;
Global $driveRadio[5];
Global $driveStatus[5];
Global $driveName[5];
Global $prg1Status[5];
Global $prg2Status[5];
Global $prg3Status[5];
Global $prg4Status[5];
Global $LDprogram[5];
Global $excelFileName[5];
Global $textFileName[5];
Global $fixButton[5];
Global $serialExcelValue[5][5];
Global $serialTextValue[5][5];

#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#include <GuiListView.au3>
#include <Excel.au3>
#include <file.au3>

Local $testFile = FileOpen("C:\Users\Tommie\Desktop\wow.txt", $FO_OVERWRITE) ; $FO_OVERWRITE = 2
FileWriteLine($testFile, "If you see this, OK2")
FileClose($testFile)

$LDprogram[1] = "ThaiSpellChecker2.1";
$LDprogram[2] = "ThaiWordPrediction2.2";
$LDprogram[3] = "ThaiWordProcessor2.1";
$LDprogram[4] = "ThaiWordSearch1.3";
;~ $LDprogram[5] = "Vaja 7 for OBEC. ( 1.10.1) 32-bit (20140905)";

$excelFileName[1] = "tsc.xls"
$textFileName[1] = "ThSp_Starter.txt"

$excelFileName[2] = "twPre.xls"
$textFileName[2] = "wp_Starter.txt"

$excelFileName[3] = "twPro.xls"
$textFileName[3] = "ThWp_Starter.txt"

$excelFileName[4] = "tws.xls"
$textFileName[4] = "ws_Starter.txt"

$yesPic = "images\success.jpg"
$noPic = "images\error.jpg"

$excelPATH = "\NECTEC\serial_Starter\excelStarter\"
$txtPATH = "\NECTEC\serial_Starter\txtStarter\"

CreateDriveList()

Global $MainGUI = GUICreate("CPE LD's Software Serial Checker" & $Version, 340, 220, -1, -1)
GUISetBkColor(0xFFFFFF)
GUISetIcon("images\cpecmu.ico")
GUISetState(@SW_SHOW)

#include <GuiStatusBar.au3>
$StatusBar1 = _GUICtrlStatusBar_Create($MainGUI)
Dim $StatusBar1_PartsWidth[2] = [110, -1]
_GUICtrlStatusBar_SetParts($StatusBar1, $StatusBar1_PartsWidth)
_GUICtrlStatusBar_SetText($StatusBar1, @TAB & @TAB & "นายวัชนันท์ จันทาภากุล", 0)
_GUICtrlStatusBar_SetText($StatusBar1, @TAB & @TAB & "ภาควิชาวิศวกรรมคอมพิวเตอร์ มหาวิทยาลัยเชียงใหม่", 1)

GUICtrlCreateTab(10, 10, 320, 200)

Local $oExcel = _Excel_Open()
For $i = 1 To $driveCount
   GUICtrlCreateTabItem($driveName[$i] & " (" & $driveStatus[$i] & ")")
   GUICtrlCreateLabel($LDprogram[1], 70, 40, 150, 20)
   GUICtrlCreateLabel($LDprogram[2], 70, 75, 150, 20)
   GUICtrlCreateLabel($LDprogram[3], 70, 105, 150, 20)
   GUICtrlCreateLabel($LDprogram[4], 70, 135, 150, 20)
   $fixButton[$i] = GUICtrlCreateButton("Fix it!", 70, 165, 60, 25)

   Local $errorCount = 0
   For $j = 1 To 4
	  Local $sWorkbook = $driveStatus[$i] & $excelPATH & $excelFileName[$j]
;~ 	  MsgBox($MB_SYSTEMMODAL, "", "$sWorkbook = " & $sWorkbook)
	  Local $oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook, Default, Default, True)
;~ 	  $oExcel.WorkBooks.Open($sWorkbook, Default, False)
	  $serialExcelValue[$i][$j] = _Excel_RangeRead($oWorkbook, Default, "B5")
	  _Excel_BookClose($oWorkbook, False)
;~ 	  _Excel_Close($oExcel, False, True)

	  $file = FileOpen($driveStatus[$i] & $txtPATH & $textFileName[$j], 0)
	  $serialTextValue[$i][$j] = FileReadLine($file)
;~ 	  MsgBox($MB_SYSTEMMODAL, "", "Serial = " & $serialTextValue)
	  FileClose($file)

	  If $serialExcelValue[$i][$j] == $serialTextValue[$i][$j] Then
;~ 		 MsgBox($MB_SYSTEMMODAL, "", "OK " & $j)
		 GUICtrlCreatePic($yesPic, 25, 32*$j, 30, 30)
	  Else
		 GUICtrlCreatePic($noPic, 25, 32*$j, 30, 30)
		 $errorCount += 1
	  EndIf

	  If $errorCount == 0 Then
		 GUICtrlSetState ($fixButton[$i],$GUI_DISABLE)
	  Else
		 GUICtrlSetState ($fixButton[$i],$GUI_ENABLE)
	  EndIf
   Next
;~    Local $oWorkbook = _Excel_BookNew($oExcel)
;~    _Excel_BookClose($oWorkbook, False)
;~    MsgBox($MB_SYSTEMMODAL, "", $serialExcelValue[$i][1] & @CRLF & $serialTextValue[$i][1] & @CRLF & $serialExcelValue[$i][2] & @CRLF & $serialTextValue[$i][2] & @CRLF & $serialExcelValue[$i][3] & @CRLF & $serialTextValue[$i][3] & @CRLF & $serialExcelValue[$i][4] & @CRLF & $serialTextValue[$i][4])
Next
_Excel_Close($oExcel, False, True)
GUICtrlCreateTabItem("") ; end tabitem definition

;~ CreateDataList()

While 1
   $choice = GUIGetMsg()
   Switch $choice
;~    Case $choice = $driveRadio[1] And BitAND(GUICtrlRead($driveRadio[1]), $GUI_CHECKED) = $GUI_CHECKED
;~ 	  MsgBox($MB_SYSTEMMODAL, 'Info:', 'You clicked the Radio 1 and it is Checked.')
   Case $GUI_EVENT_CLOSE
	  ExitLoop
   Case $fixButton[1]
	  If $driveCount == 1 Then
		 FixDrive1()
	  EndIf
   Case $fixButton[2]
	  If $driveCount == 2 Then
		 FixDrive2()
	  EndIf
   Case $fixButton[3]
	  If $driveCount == 3 Then
		 FixDrive3()
	  EndIf
   Case $fixButton[4]
	  If $driveCount == 4 Then
		 FixDrive4()
	  EndIf
   EndSwitch
WEnd

Func CreateDriveList()
   Global $driveArray = DriveGetDrive($DT_ALL)
   Global $driveCount = 0
   Local $yCheckbox = 30;
   If @error Then
	   ; An error occurred when retrieving the drives.
	   MsgBox($MB_SYSTEMMODAL, "", "It appears an error occurred.")
	Else
	   For $i = 1 To ($driveArray[0] - 1)
		   $driveType = DriveGetType($driveArray[$i+1], $DT_DRIVETYPE)
		   $driveLetter = StringUpper($driveArray[$i+1])
		   $driveLabel = DriveGetLabel($driveLetter)
		   If $driveType == "Removable" Then
;~ 			   MsgBox($MB_SYSTEMMODAL, "", "Drive " & $i & "/" & ($driveArray[0] - 1) & ":" & @CRLF & $driveLetter & " " & $driveLabel & " (" & $driveType & ")")
;~ 			   $driveCheckbox[$driveCount + 1] = GUICtrlCreateCheckbox($driveLabel & " (" & $driveLetter & ")", 550, $yCheckbox, 125, 17)
			   $driveStatus[$driveCount + 1] = $driveLetter
			   $driveName[$driveCount + 1] = $driveLabel
;~ 			   If $driveName[$driveCount + 1]  == "NECTEC LDSW" Then
;~ 				  MsgBox($MB_SYSTEMMODAL, "", "OK")
;~ 			   EndIf
;~ 			   MsgBox($MB_SYSTEMMODAL, "", "$driveCount = " & $driveCount)
			   $yCheckbox += 25
			   $driveCount += 1
		   EndIf
		Next
;~ 		For $i = 1 To $driveCount
;~ 		   MsgBox($MB_SYSTEMMODAL, "", "$driveStatus[" & $i & "] = " & $driveStatus[$i])
;~ 		 Next
   EndIf

   If $driveCount == 0 Then
;~ 	  $driveGroup = GUICtrlCreateGroup("Removable Drives", 100, 5, 150, 60)
	  $noDriveLabel = GUICtrlCreateLabel("* Please insert USB drive.", 545, $yCheckbox, 125)
	  $yCheckbox += 25
	  $yDriveGroup = 1
   Else
	  Local $yDriveGroup = $driveCount
   EndIf
EndFunc

;~ Func CreateDataList()
;~    $driveGroup = GUICtrlCreateGroup("Removable Drives", 15, 10, 140, 120)
;~    Local $USBlist = GUICtrlCreateList("------ USB drives ------", 27, 32, 115, 75)
;~    GUICtrlSetLimit($USBlist, 200) ; to limit horizontal scrolling
;~    _GUICtrlListView_DeleteAllItems($USBlist)

;~    GUIStartGroup()

;~    Local $yPos = 32
;~    For $i = 1 To $driveCount
;~ 	  GUICtrlSetData($USBlist, $driveName[$i] & " (" & $driveStatus[$i] & ")")
;~ 	  $driveRadio[$i] = GUICtrlCreateRadio($driveName[$i] & " (" & $driveStatus[$i] & ")", 27, $yPos, 120, 20)
;~ 	  $yPos += 20
;~    Next
;~    If $driveCount > 0 Then
;~ 	  GUICtrlSetState($driveRadio[1], $GUI_CHECKED)
;~    EndIf
;~ EndFunc

Func FixDrive1()
   Local $drive1File1 = FileOpen($driveStatus[1] & $txtPATH & $textFileName[1], $FO_OVERWRITE) ; $FO_OVERWRITE = 2
   FileWriteLine($drive1File1, $serialExcelValue[1][1])
   FileClose($drive1File1)

   Local $drive1File2 = FileOpen($driveStatus[1] & $txtPATH & $textFileName[2], $FO_OVERWRITE) ; $FO_OVERWRITE = 2
   FileWriteLine($drive1File2, $serialExcelValue[1][2])
   FileClose($drive1File2)

   Local $drive1File3 = FileOpen($driveStatus[1] & $txtPATH & $textFileName[3], $FO_OVERWRITE) ; $FO_OVERWRITE = 2
   FileWriteLine($drive1File3, $serialExcelValue[1][3])
   FileClose($drive1File3)

   Local $drive1File4 = FileOpen($driveStatus[1] & $txtPATH & $textFileName[4], $FO_OVERWRITE) ; $FO_OVERWRITE = 2
   FileWriteLine($drive1File4, $serialExcelValue[1][4])
   FileClose($drive1File4)

;~    MsgBox($MB_SYSTEMMODAL, $driveName[1] & "(" & $driveStatus[1] & ")", "Serial has been fixed.")
   Run( FileGetShortName(@ScriptFullPath))
;~    Run( FileGetShortName(@AutoItExe) & " " & FileGetShortName(@ScriptFullPath))
   Exit
EndFunc

Func FixDrive2()
   Local $drive2File1 = FileOpen($driveStatus[2] & $txtPATH & $textFileName[1], $FO_OVERWRITE) ; $FO_OVERWRITE = 2
   FileWriteLine($drive2File1, $serialExcelValue[2][1])
   FileClose($drive2File1)

   Local $drive2File2 = FileOpen($driveStatus[2] & $txtPATH & $textFileName[2], $FO_OVERWRITE) ; $FO_OVERWRITE = 2
   FileWriteLine($drive2File2, $serialExcelValue[2][2])
   FileClose($drive2File2)

   Local $drive2File3 = FileOpen($driveStatus[2] & $txtPATH & $textFileName[3], $FO_OVERWRITE) ; $FO_OVERWRITE = 2
   FileWriteLine($drive2File3, $serialExcelValue[2][3])
   FileClose($drive2File3)

   Local $drive2File4 = FileOpen($driveStatus[2] & $txtPATH & $textFileName[4], $FO_OVERWRITE) ; $FO_OVERWRITE = 2
   FileWriteLine($drive2File4, $serialExcelValue[2][4])
   FileClose($drive2File4)

;~    MsgBox($MB_SYSTEMMODAL, $driveName[1] & "(" & $driveStatus[1] & ")", "Serial has been fixed.")
   Run( FileGetShortName(@ScriptFullPath))
;~    Run( FileGetShortName(@AutoItExe) & " " & FileGetShortName(@ScriptFullPath))
   Exit
EndFunc

Func FixDrive3()
   Local $drive3File1 = FileOpen($driveStatus[3] & $txtPATH & $textFileName[1], $FO_OVERWRITE) ; $FO_OVERWRITE = 2
   FileWriteLine($drive3File1, $serialExcelValue[3][1])
   FileClose($drive3File1)

   Local $drive3File2 = FileOpen($driveStatus[3] & $txtPATH & $textFileName[2], $FO_OVERWRITE) ; $FO_OVERWRITE = 2
   FileWriteLine($drive3File2, $serialExcelValue[3][2])
   FileClose($drive3File2)

   Local $drive3File3 = FileOpen($driveStatus[3] & $txtPATH & $textFileName[3], $FO_OVERWRITE) ; $FO_OVERWRITE = 2
   FileWriteLine($drive3File3, $serialExcelValue[3][3])
   FileClose($drive3File3)

   Local $drive3File4 = FileOpen($driveStatus[3] & $txtPATH & $textFileName[4], $FO_OVERWRITE) ; $FO_OVERWRITE = 2
   FileWriteLine($drive3File4, $serialExcelValue[3][4])
   FileClose($drive3File4)

;~    MsgBox($MB_SYSTEMMODAL, $driveName[1] & "(" & $driveStatus[1] & ")", "Serial has been fixed.")
   Run( FileGetShortName(@ScriptFullPath))
   Exit
EndFunc

Func FixDrive4()
   Local $drive1File4 = FileOpen($driveStatus[4] & $txtPATH & $textFileName[1], $FO_OVERWRITE) ; $FO_OVERWRITE = 2
   FileWriteLine($drive4File1, $serialExcelValue[4][1])
   FileClose($drive4File1)

   Local $drive4File2 = FileOpen($driveStatus[4] & $txtPATH & $textFileName[2], $FO_OVERWRITE) ; $FO_OVERWRITE = 2
   FileWriteLine($drive4File2, $serialExcelValue[4][2])
   FileClose($drive4File2)

   Local $drive4File3 = FileOpen($driveStatus[4] & $txtPATH & $textFileName[3], $FO_OVERWRITE) ; $FO_OVERWRITE = 2
   FileWriteLine($drive4File3, $serialExcelValue[4][3])
   FileClose($drive4File3)

   Local $drive4File4 = FileOpen($driveStatus[4] & $txtPATH & $textFileName[4], $FO_OVERWRITE) ; $FO_OVERWRITE = 2
   FileWriteLine($drive4File4, $serialExcelValue[4][4])
   FileClose($drive4File4)

;~    MsgBox($MB_SYSTEMMODAL, $driveName[1] & "(" & $driveStatus[1] & ")", "Serial has been fixed.")
   Run( FileGetShortName(@ScriptFullPath))
   Exit
EndFunc