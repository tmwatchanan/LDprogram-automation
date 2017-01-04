#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

$dataPATH = "C:\Users\" & @UserName & "\AppData\Local\";

Run("C:\WINDOWS\EXPLORER.EXE /n,/e," & $dataPATH)

$LDdataName1 = "SetTSpC20.dat";
$LDdataName2 = "SetTWPro32.dat";
$LDdataName3 = "SetTWPv21-Sapi.dat";
$LDdataName4 = "SetTWSv12-sapi.dat";

$LDdataMSG1 = $LDdataName1 & @CR;
$LDdataMSG2 = $LDdataName2 & @CR;
$LDdataMSG3 = $LDdataName3 & @CR;
$LDdataMSG4 = $LDdataName4 & @CR;

$DeletingMSG = "Deleted files:" & @CR;

$count = 0

If FileExists ($dataPATH & $LDdataName1) Then
   $DeletingMSG &= $LDdataMSG1;
   FileDelete($dataPATH & $LDdataName1);
   $count += 1;
EndIf
If FileExists ($dataPATH & $LDdataName2) Then
   $DeletingMSG &= $LDdataMSG2;
   FileDelete($dataPATH & $LDdataName2);
   $count += 1;
EndIf
If FileExists ($dataPATH & $LDdataName3) Then
   $DeletingMSG &= $LDdataMSG3;
   FileDelete($dataPATH & $LDdataName3);
   $count += 1;
EndIf
If FileExists ($dataPATH & $LDdataName4) Then
   $DeletingMSG &= $LDdataMSG4;
   FileDelete($dataPATH & $LDdataName4);
   $count += 1;
EndIf

;~ $count += 1;
ConsoleWrite($count);

If Not $count = 0 Then
   $StatusMSG = "-> Done!"
   $DeletingMSG &= @CR & $StatusMSG & @CR & "From: Tommie eiei";
Else
   $StatusMSG = "-x Didn't do anything."
   $DeletingMSG = "No such existed files" & @CR & @CR & $StatusMSG & @CR & "From: Tommie eiei";
EndIf

#include <MsgBoxConstants.au3>
MsgBox($MB_SYSTEMMODAL, "CPE LD's program", $DeletingMSG, 10)