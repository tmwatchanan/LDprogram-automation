;~ #AutoIt3Wrapper_icon=C:\Users\Tommie\Desktop\cpecmu.ico
;Program by Tommy InwZaa

$LDprogram1 = "ThaiSpellChecker2.1";
$LDprogram2 = "ThaiWordPrediction2.2";
$LDprogram3 = "ThaiWordProcessor2.1";
$LDprogram4 = "ThaiWordSearch1.3";
$LDprogram5 = "Vaja 7 for OBEC. ( 1.10.1) 32-bit (20140905)";

$dataPATH = "C:\Users\" & @UserName & "\AppData\Local\";
$LDdataName1 = "SetTSpC20.dat";
$LDdataName2 = "SetTWPro32.dat";
$LDdataName3 = "SetTWPv21-Sapi.dat";
$LDdataName4 = "SetTWSv12-sapi.dat";

$databasePATH = "C:\DB_LDprograms\";
$LDdatabase1 = "dbuser_thaiwordprediction.db";
$LDdatabase2 = "dbuser_thaiwordsearch.db";

$installedPic = "images\installed.jpg"
$notinstalledPic = "images\not-installed.jpg"

$NECTECDir = "C:\NECTEC\"
$VajaDir = "C:\Program Files (x86)\Vaja7\"

Global $installedLDprogram1 = 0;
Global $installedLDprogram2 = 0;
Global $installedLDprogram3 = 0;
Global $installedLDprogram4 = 0;
Global $installedLDprogram5 = 0;

#include <ColorConstants.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include <GuiStatusBar.au3>
#include <MsgBoxConstants.au3>
#include <GuiListView.au3>

Const $Version = "1.3"

Global $MainGUI = GUICreate("CPE LD's Manager " & $Version, 700, 350, -1, -1)
GUISetBkColor(0xFFFFFF)
GUISetIcon("images\cpecmu.ico")
;~ $CPELogoPic = GUICtrlCreatePic("LogoCPE.gif", 150, 28, 350, 75)

Global $status1Pic;
Global $status2Pic;
Global $status3Pic;
Global $status4Pic;
Global $status5Pic;
CreateProgramList()

GUICtrlCreateGroup("Uninstall Programs", 15, 5, 300, 230)
CreateProgramButton()
Global $RefreshButton = GUICtrlCreateButton("Refresh List", 210, 28, 80, 30)

GUICtrlCreateGroup("Vaja Serial Number", 15, 238, 300, 72)
Global $SerialButton = GUICtrlCreateButton("Auto enter serial number", 45, 260, 230, 30)

GUICtrlCreateGroup("Delete Data", 330, 5, 180, 150)
Global $copyDBStatus = 0
CreateOpenDataButton()
CreateDeleteDataButton()
CreateDataList()

GUICtrlCreateGroup("Backup Database", 330, 160, 180, 150)
CreateDatabaseButton()
CreateDatabaseList()

Global $driveGroup;
Global $driveCheckbox[5];
Global $driveStatus[5];
Global $driveName[5];
Global $driveSelect[5];
Global $noDriveLabel;
CreateDriveList()

$StatusBar1 = _GUICtrlStatusBar_Create($MainGUI)
Dim $StatusBar1_PartsWidth[3] = [150, 640, -1]
_GUICtrlStatusBar_SetParts($StatusBar1, $StatusBar1_PartsWidth)
_GUICtrlStatusBar_SetText($StatusBar1, @TAB & @TAB & "ผู้พัฒนา นายวัชนันท์ จันทาภากุล", 0)
_GUICtrlStatusBar_SetText($StatusBar1, @TAB & @TAB & "ภาควิชาวิศวกรรมคอมพิวเตอร์ คณะวิศวกรรมศาสตร์ มหาวิทยาลัยเชียงใหม่", 1)
_GUICtrlStatusBar_SetText($StatusBar1, @TAB & "CPE CMU", 2)

GUISetState(@SW_SHOW)

While 1
   $choice = GUIGetMsg()
   Switch $choice
	  Case $GUI_EVENT_CLOSE
		 Exit
	  Case $ProgramButton1
		 UninstallProgram1()
	  Case $ProgramButton2
		 UninstallProgram2()
	  Case $ProgramButton3
		 UninstallProgram3()
	  Case $ProgramButton4
		 UninstallProgram4()
	  Case $ProgramButton5
		 UninstallProgram5()
	  Case $RefreshButton
		 RemoveProgramList()
		 CreateProgramList()
		 RemoveProgramButton()
		 CreateProgramButton()
	  Case $SerialButton
		 EnterSerial()
	  Case $OpenDataButton
		 OpenData()
	  Case $DeleteDataButton
		 DeleteData()
	  Case $MoveDBButton
		 MoveDB()
	  Case $CopyDBButton
		 CopyDB()
	  Case $DeleteDBButton
		 DeleteDB()
	  Case $OpenDBButton
		 OpenDB()
	  Case $RefreshDriveButton
		 RefreshDrive()
;~ 	  Case Else
	  Case $GUI_EVENT_CLOSE
		 ExitLoop
   EndSwitch
WEnd

Func UninstallProgram1()
   If WinExists($LDprogram1) Then
	  WinActivate("Programs and Features")
   Else
	  $PID1 = Run("C:\Windows\System32\control.exe appwiz.cpl")
	  WinWait("Programs and Features")
   EndIf
   WinWaitActive("Programs and Features", "", 10)
   WinActivate("Programs and Features")
   Send($LDprogram1)
   Send("{Enter}")
   WinActivate("Programs and Features")
   Send("{Y down}")
   ProcessClose($PID1)
EndFunc

Func UninstallProgram2()
   If WinExists($LDprogram2) Then
	  WinActivate("Programs and Features")
   Else
	  $PID2 = Run("C:\Windows\System32\control.exe appwiz.cpl")
	  WinWait("Programs and Features")
   EndIf
   WinWaitActive("Programs and Features", "", 10)
   WinActivate("Programs and Features")
   Send($LDprogram2)
   Send("{Enter}")
   WinActivate("Programs and Features")
   Send("{Y down}")
   ProcessClose($PID2)
EndFunc

Func UninstallProgram3()
   If WinExists($LDprogram3) Then
	  WinActivate("Programs and Features")
   Else
	  $PID3 = Run("C:\Windows\System32\control.exe appwiz.cpl")
	  WinWait("Programs and Features")
   EndIf
   WinWaitActive("Programs and Features", "", 10)
   WinActivate("Programs and Features")
   Send($LDprogram3)
   Send("{Enter}")
   WinActivate("Programs and Features")
   Send("{Y down}")
   ProcessClose($PID3)
EndFunc

Func UninstallProgram4()
   If WinExists($LDprogram4) Then
	  WinActivate("Programs and Features")
   Else
	  $PID4 = Run("C:\Windows\System32\control.exe appwiz.cpl")
	  WinWait("Programs and Features")
   EndIf
   WinWaitActive("Programs and Features", "", 10)
   WinActivate("Programs and Features")
   Send($LDprogram4)
   Send("{Enter}")
   WinActivate("Programs and Features")
   Send("{Y down}")
   ProcessClose($PID4)
EndFunc

Func UninstallProgram5()
   If WinExists($LDprogram5) Then
	  WinActivate("Programs and Features")
   Else
	  $PID5 = Run("C:\Windows\System32\control.exe appwiz.cpl")
	  WinWait("Programs and Features")
    EndIf
   WinWaitActive("Programs and Features", "", 10)
   WinActivate("Programs and Features")
   Send($LDprogram5)
   Send("{Enter}")
   WinActivate("Programs and Features")
   Send("{Y down}")
   ProcessClose($PID5)
EndFunc

Func EnterSerial()
   WinSetState("Vaja 7 for OBEC. ( 1.10.1) 32-bit (20140905)", "", @SW_HIDE)
   WinSetState("Vaja 7 for OBEC. ( 1.10.1) 32-bit (20140905)", "", @SW_SHOW)
   ControlClick("Vaja 7 for OBEC. ( 1.10.1) 32-bit (20140905)","","[CLASS:Edit; INSTANCE:1]","Primary")
   Send("AB12CFZGYXD613ZJ85NHRX59W")
EndFunc

Func DeleteData()

;~    Run("C:\WINDOWS\EXPLORER.EXE /n,/e," & $dataPATH)

   $LDdataMSG1 = $LDdataName1 & @CRLF;
   $LDdataMSG2 = $LDdataName2 & @CRLF;
   $LDdataMSG3 = $LDdataName3 & @CRLF;
   $LDdataMSG4 = $LDdataName4 & @CRLF;

   $DeletingMSG = "Deleted files:" & @CRLF;

   Local $dataCount = 0

   If FileExists ($dataPATH & $LDdataName1) Then
	  $DeletingMSG &= $LDdataMSG1;
	  FileDelete($dataPATH & $LDdataName1);
	  $dataCount += 1;
   EndIf
   If FileExists ($dataPATH & $LDdataName2) Then
	  $DeletingMSG &= $LDdataMSG2;
	  FileDelete($dataPATH & $LDdataName2);
	  $dataCount += 1;
   EndIf
   If FileExists ($dataPATH & $LDdataName3) Then
	  $DeletingMSG &= $LDdataMSG3;
	  FileDelete($dataPATH & $LDdataName3);
	  $dataCount += 1;
   EndIf
   If FileExists ($dataPATH & $LDdataName4) Then
	  $DeletingMSG &= $LDdataMSG4;
	  FileDelete($dataPATH & $LDdataName4);
	  $dataCount += 1;
   EndIf

   If Not $dataCount = 0 Then
	  $StatusMSG = "-> Done!"
	  $DeletingMSG &= @CRLF & @CRLF & $StatusMSG;
   Else
	  $StatusMSG = "-x Didn't do anything."
	  $DeletingMSG = "No such file" & @CRLF & @CRLF & $StatusMSG;
   EndIf

   MsgBox($MB_SYSTEMMODAL, "CPE LD's program", $DeletingMSG, 10)
   $dataCount = 0;
   GUICtrlSetState ($DeleteDataButton, $GUI_DISABLE)
   CreateDataList()
EndFunc

Func CreateProgramButton()
   Global $ProgramButton1 = GUICtrlCreateButton($LDprogram1, 60, 28, 120, 30)
   If $installedLDprogram1 == 0 Then
	  GUICtrlSetState ($ProgramButton1, $GUI_DISABLE)
   Else
	  GUICtrlSetState ($ProgramButton1, $GUI_ENABLE)
   EndIf
   Global $ProgramButton2 = GUICtrlCreateButton($LDprogram2, 60, 68, 120, 30)
   If $installedLDprogram2 == 0 Then
	  GUICtrlSetState ($ProgramButton2, $GUI_DISABLE)
   Else
	  GUICtrlSetState ($ProgramButton2, $GUI_ENABLE)
   EndIf
   Global $ProgramButton3 = GUICtrlCreateButton($LDprogram3, 60, 108, 120, 30)
   If $installedLDprogram3 == 0 Then
	  GUICtrlSetState ($ProgramButton3, $GUI_DISABLE)
   Else
	  GUICtrlSetState ($ProgramButton3, $GUI_ENABLE)
   EndIf
   Global $ProgramButton4 = GUICtrlCreateButton($LDprogram4, 60, 148, 120, 30)
   If $installedLDprogram4 == 0 Then
	  GUICtrlSetState ($ProgramButton4, $GUI_DISABLE)
   Else
	  GUICtrlSetState ($ProgramButton4, $GUI_ENABLE)
   EndIf
   Global $ProgramButton5 = GUICtrlCreateButton($LDprogram5, 60, 188, 230, 30)
   If $installedLDprogram5 == 0 Then
	  GUICtrlSetState ($ProgramButton5, $GUI_DISABLE)
   Else
	  GUICtrlSetState ($ProgramButton5, $GUI_ENABLE)
   EndIf
EndFunc

Func RemoveProgramButton()
   GUICtrlDelete($ProgramButton1)
   GUICtrlDelete($ProgramButton2)
   GUICtrlDelete($ProgramButton3)
   GUICtrlDelete($ProgramButton4)
   GUICtrlDelete($ProgramButton5)
EndFunc

Func RemoveProgramList()
   GUICtrlDelete($status1Pic)
   GUICtrlDelete($status2Pic)
   GUICtrlDelete($status3Pic)
   GUICtrlDelete($status4Pic)
   GUICtrlDelete($status5Pic)
EndFunc

Func CreateProgramList()
   If FileExists($NECTECDir & $LDprogram1 & "\" & "ThaiSpellChecker.exe") Then
	  $status1Pic = GUICtrlCreatePic($installedPic, 25, 28, 30, 30)
	  $installedLDprogram1 = 1;
   Else
	  $status1Pic = GUICtrlCreatePic($notinstalledPic, 25, 28, 30, 30)
	  $installedLDprogram1 = 0;
   EndIf
   If FileExists($NECTECDir & $LDprogram2 & "\" & "ThaiWordPrediction.exe") Then
	  $status2Pic = GUICtrlCreatePic($installedPic, 25, 68, 30, 30)
	  $installedLDprogram2 = 1;
   Else
	  $status2Pic = GUICtrlCreatePic($notinstalledPic, 25, 68, 30, 30)
	  $installedLDprogram2 = 0;
   EndIf
   If FileExists($NECTECDir & $LDprogram3 & "\" & "ThaiWordProcessor.exe") Then
	  $status3Pic = GUICtrlCreatePic($installedPic, 25, 108, 30, 30)
	  $installedLDprogram3 = 1;
   Else
	  $status3Pic = GUICtrlCreatePic($notinstalledPic, 25, 108, 30, 30)
	  $installedLDprogram3 = 0;
   EndIf
   If FileExists($NECTECDir & $LDprogram4 & "\" & "ThaiWordSearch.exe") Then
	  $status4Pic = GUICtrlCreatePic($installedPic, 25, 148, 30, 30)
	  $installedLDprogram4 = 1;
   Else
	  $status4Pic = GUICtrlCreatePic($notinstalledPic, 25, 148, 30, 30)
	  $installedLDprogram4 = 0;
   EndIf
   If FileExists($VajaDir & "Vaja7_32.dll") Then
	  $status5Pic = GUICtrlCreatePic($installedPic, 25, 188, 30, 30)
	  $installedLDprogram5 = 1;
   Else
	  $status5Pic = GUICtrlCreatePic($notinstalledPic, 25, 188, 30, 30)
	  $installedLDprogram5 = 0;
   EndIf
EndFunc

Func OpenData()
   Run("C:\WINDOWS\EXPLORER.EXE /n,/e," & $dataPATH)
EndFunc

Func CreateOpenDataButton()
   Global $OpenDataButton = GUICtrlCreateButton("Open Data", 350, 110, 60, 30)
EndFunc

Func CreateDeleteDataButton()
   Global $DeleteDataButton = GUICtrlCreateButton("Delete Data", 420, 110, 80, 30)
EndFunc

Func CreateDataList()
   Local $Datalist = GUICtrlCreateList("------ DATA FILES ------", 362, 32, 115, 75)
   GUICtrlSetLimit($Datalist, 200) ; to limit horizontal scrolling
   _GUICtrlListView_DeleteAllItems($Datalist)
   Global $dataCount = 0;
   If FileExists ($dataPATH & $LDdataName1) Then
	  GUICtrlSetData($Datalist, $LDdataName1)
	  $dataCount += 1;
   EndIf
   If FileExists ($dataPATH & $LDdataName2) Then
	  GUICtrlSetData($Datalist, $LDdataName2)
	  $dataCount += 1;
   EndIf
   If FileExists ($dataPATH & $LDdataName3) Then
	  GUICtrlSetData($Datalist, $LDdataName3)
	  $dataCount += 1;
   EndIf
   If FileExists ($dataPATH & $LDdataName4) Then
	  GUICtrlSetData($Datalist, $LDdataName4)
	  $dataCount += 1;
   EndIf

   If $dataCount == 0 Then
	  GUICtrlSetState ($DeleteDataButton, $GUI_DISABLE)
	  GUICtrlSetState ($OpenDataButton, $GUI_DISABLE)
   Else
	  GUICtrlSetState ($DeleteDataButton, $GUI_ENABLE)
	  GUICtrlSetState ($OpenDataButton, $GUI_ENABLE)
   EndIf
EndFunc

Func CreateDatabaseButton()
   Global $OpenDBButton = GUICtrlCreateButton("Open Directory", 340, 245, 90, 30)
   Global $CopyDBButton = GUICtrlCreateButton("Copy", 440, 245, 60, 30)
   Global $DeleteDBButton = GUICtrlCreateButton("Delete", 350, 285, 140, 15)
   GUICtrlSetState ($DeleteDBButton, $GUI_DISABLE)
EndFunc

Func CreateMoveDatabaseButton($yPos)
   Global $MoveDBButton = GUICtrlCreateButton("Move", 610, $yPos, 60, 30)
   SetStatMoveDatabaseButton()
EndFunc

Func SetStatMoveDatabaseButton()
   If $copyDBStatus = 0 Then
	  GUICtrlSetState ($MoveDBButton, $GUI_DISABLE)
   Else
	  GUICtrlSetState ($MoveDBButton, $GUI_ENABLE)
   EndIf
EndFunc

Func CreateDatabaseList()
   Local $DatabaseList = GUICtrlCreateList("--------- DB_LDprograms ---------", 345, 190, 150, 55)
   GUICtrlSetLimit($DatabaseList, 200) ; to limit horizontal scrolling
   Global $databaseCount = 0;
   If FileExists ($databasePATH & $LDdatabase1) Then
	  GUICtrlSetData($DatabaseList, $LDdatabase1)
	  $databaseCount += 1;
   EndIf
   If FileExists ($databasePATH & $LDdatabase2) Then
	  GUICtrlSetData($DatabaseList, $LDdatabase2)
	  $databaseCount += 1;
   EndIf

   If $databaseCount == 0 Then
	  GUICtrlSetState ($CopyDBButton, $GUI_DISABLE)
	  GUICtrlSetState ($OpenDBButton, $GUI_DISABLE)
   Else
	  GUICtrlSetState ($CopyDBButton, $GUI_ENABLE)
	  GUICtrlSetState ($OpenDBButton, $GUI_ENABLE)
   EndIf
EndFunc

Func MoveDB()
   MoveDatabaseToUSB()
   CreateDatabaseList()
EndFunc

Func CopyDB()
;~    If not FileExists("C:\Users\" & @UserName & "\Desktop\" & @UserName) Then
;~ 	  DirCreate("C:\Users\" & @UserName & "\Desktop\" & @UserName)
;~    EndIf

   Local $destinationDesktopPath = "C:\Users\" & @UserName & "\Desktop\" & @UserName
   Local $destinationDatabasePath = "C:\Users\" & @UserName & "\Desktop\" & @UserName & "\DB_LDprograms"
   Local $sourceDatabasePath = "C:\DB_LDprograms"
   If not FileExists($sourceDatabasePath) Then
	  MsgBox($MB_SYSTEMMODAL, "Database Operation", "-x No such file in source path.", 10)
   EndIf

   If not FileExists($destinationDesktopPath) Then
	  DirCreate("C:\Users\" & @UserName & "\Desktop\" & @UserName)
;~ 	  MsgBox($MB_SYSTEMMODAL, "Database Operation", "-x Destination path not found.", 10)
   EndIf
   If not FileExists($destinationDesktopPath & "DB_LDprograms") Then
	  DirCreate("C:\Users\" & @UserName & "\Desktop\" & @UserName & "\DB_LDprograms")
   EndIf
   Runwait(@ComSpec & " /c " & "xcopy " & '"' & $sourceDatabasePath & '"' & ' "' & $destinationDatabasePath & '"' & " /E /C /D /Y","",@SW_HIDE)
;~ 	  DirCopy("C:\NECTEC", "C:\Users\" & @UserName & "\Desktop\" & @UserName & "\NECTEC");

;~ 	  If FileExists("D:\") Then
;~ 		 DirCopy("C:\Users\" & @UserName & "\Desktop\" & @UserName, "D:\" & @UserName);
;~ 	  ElseIf FileExists("E:\") Then
;~ 		 DirCopy("C:\Users\" & @UserName & "\Desktop\" & @UserName, "E:\" & @UserName);
;~ 	  EndIf

   If $driveCount > 0 Then
	  $copyDBStatus = 1
   EndIf
   SetStatMoveDatabaseButton()
   EnableDeleteDatabase()
EndFunc

Func EnableDeleteDatabase()
   GUICtrlSetState ($DeleteDBButton, $GUI_ENABLE)
   GUICtrlSetBkColor($DeleteDBButton, $COLOR_RED)
EndFunc

Func DeleteDB()
   If FileExists ("C:\DB_LDprograms") Then
	  DirRemove("C:\DB_LDprograms", $DIR_REMOVE); % $DIR_REMOVE = 1
	  MsgBox($MB_SYSTEMMODAL, "Database Operation", "-> Sucessfully deleted DB_LDprograms database directory.", 10)
	  EnableDeleteDatabase()
   Else
	  MsgBox($MB_SYSTEMMODAL, "Database Operation", "-x Directory not found!", 10)
   EndIf
EndFunc

Func OpenDB()
   Run("C:\WINDOWS\EXPLORER.EXE /n,/e," & "C:\DB_LDprograms")
EndFunc

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
			   $driveCheckbox[$driveCount + 1] = GUICtrlCreateCheckbox($driveLabel & " (" & $driveLetter & ")", 550, $yCheckbox, 125, 17)
			   $driveStatus[$driveCount + 1] = $driveLetter
			   $driveName[$driveCount + 1] = $driveLabel
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
	  $driveGroup = GUICtrlCreateGroup("Removable Drives", 530, 5, 150, 60)
	  $noDriveLabel = GUICtrlCreateLabel("* Please insert USB drive.", 545, $yCheckbox, 125)
	  $yCheckbox += 25
   EndIf

   GUICtrlDelete($driveGroup)
   If $driveCount == 0 Then
	  $yDriveGroup = 1
   Else
	  Local $yDriveGroup = $driveCount
   EndIf
   $driveGroup = GUICtrlCreateGroup("Removable Drives", 530, 5, 150, 60 + ($yDriveGroup * 30))

   CreateRefreshDriveButton($yCheckbox)
   CreateMoveDatabaseButton($yCheckbox)
EndFunc

Func CreateRefreshDriveButton($yPos)
   Global $RefreshDriveButton = GUICtrlCreateButton("Refresh", 540, $yPos, 60, 30)
EndFunc

Func RefreshDrive()
   For $i = 1 To $driveCount
	  GUICtrlDelete($driveCheckbox[$i])
   Next
   GUICtrlDelete($driveGroup)
   GUICtrlDelete($noDriveLabel)
   GUICtrlDelete($RefreshDriveButton)
   GUICtrlDelete($MoveDBButton)
   CreateDriveList()
EndFunc

Func MoveDatabaseToUSB()
   Local $movedDrive = 0
   If $driveCount > 0 Then
	  Local $MoveDatabaseToUSBMSG = "Moved database to: " & @CRLF
	  For $i = 1 To $driveCount
		 If GuiCtrlRead($driveCheckbox[$i]) = $GUI_CHECKED Then
			Local $sourceDatabasePath = "C:\Users\" & @UserName & "\Desktop\" & @UserName
			Local $destinationDatabasePath = $driveStatus[$i] & "\" & @UserName
			Local $destinationDrivePath = $driveStatus[$i] & "\"
;~ 			MsgBox($MB_SYSTEMMODAL, "Database Operation", $sourceDatabasePath, 10)
   ;~ 		 MsgBox($MB_SYSTEMMODAL, "Database Operation", $destinationDatabasePath, 10)
			If FileExists($sourceDatabasePath) and FileExists ($destinationDrivePath) Then
;~ 			   Runwait(@ComSpec & " /c " & "xcopy " & '"' & $sourceDatabasePath & '"' & ' "' & $destinationDatabasePath & '"' & " /E /C /D /Y","",@SW_HIDE)
			   DirCopy($sourceDatabasePath, $destinationDatabasePath, $FC_OVERWRITE)
			   $movedDrive += 1
			Else
			   MsgBox($MB_SYSTEMMODAL, "Database Operation", "Problem found!", 10)
			EndIf
			Run("C:\WINDOWS\EXPLORER.EXE /n,/e," & $destinationDatabasePath)
			$MoveDatabaseToUSBMSG &= $driveName[$i] & " (" & $driveStatus[$i] & ") " & @CRLF

			If FileExists ($driveStatus[$i] & "\" & @UserName & "\NECTEC") Then
			   DirRemove($driveStatus[$i] & "\" & @UserName & "\NECTEC", $DIR_REMOVE); % $DIR_REMOVE = 1
			EndIf
			If FileExists ("C:\NECTEC") Then
			   DirRemove("C:\NECTEC", $DIR_REMOVE); % $DIR_REMOVE = 1
			EndIf
		 EndIf
	  Next

	  If $movedDrive > 0 Then
		 $MoveDatabaseToUSBMSG &= @CRLF & "-> Sucessfully!"
		 MsgBox($MB_SYSTEMMODAL, "CPE LD's program", $MoveDatabaseToUSBMSG)
	  Else
		 MsgBox($MB_SYSTEMMODAL, "CPE LD's program", "Found problems!" & @CRLF & "- Please select the drives" & @CRLF & "- No move operation over any selected drives.")
	  EndIf
   Else
	  MsgBox($MB_SYSTEMMODAL, "CPE LD's program", "No drive found!")
   EndIf
EndFunc