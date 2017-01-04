#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

;~ ConsoleWrite('User:   ' & @UserName)

;~ C:\Users\Tommie\AppData\Local
;~ ConsoleWrite("C:\Users\" & @UserName & "\AppData\Local")

if not FileExists("C:\Users\" & @UserName & "\Desktop\" & @UserName) Then
   DirCreate("C:\Users\" & @UserName & "\Desktop\" & @UserName)
endif

DirMove("C:\DB_LDprograms", "C:\Users\" & @UserName & "\Desktop\" & @UserName & "\DB_LDprograms");
DirMove("C:\NECTEC", "C:\Users\" & @UserName & "\Desktop\" & @UserName & "\NECTEC");
;~ FileDelete("C:\DB_LDprograms");
;~ DirRemove("C:\NECTEC");

If FileExists("D:\") Then
   DirCopy("C:\Users\" & @UserName & "\Desktop\" & @UserName, "D:\" & @UserName);
ElseIf FileExists("E:\") Then
   DirCopy("C:\Users\" & @UserName & "\Desktop\" & @UserName, "E:\" & @UserName);
EndIf

$dataPATH = "C:\Users\" & @UserName & "\AppData\Local\";
FileDelete($dataPATH & "SetTSpC20.dat");
FileDelete($dataPATH & "SetTWPro32.dat");
FileDelete($dataPATH & "SetTWPv21-Sapi.dat");
FileDelete($dataPATH & "SetTWSv12-sapi.dat");

;~ DirCopy("C:\DB_LDprograms", "C:\Users\" & @UserName & "\AppData\Local");

$RUNDLL32 = @SystemDir & "\rundll32.exe"
$cpl = 'appwiz.cpl'
Run($RUNDLL32 & " shell32.dll,Control_RunDLL " & $cpl)
$cp2 = 'wupdmgr'
Run($RUNDLL32 & " shell32.dll,Control_RunDLL " & $cp2)