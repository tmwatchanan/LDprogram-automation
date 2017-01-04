@echo off

set FOLDER_CURRENT=%cd%    

set VERSION=v0.x.x.x
set STATUS=prod
set NAME_PROJECT="CPE LD Program %VERSION%"
REM _%STATUS%"
set EXE_OUT=%NAME_PROJECT%.exe

REM set FOLDER_SRC=%FOLDER_CURRENT%\..\src\
set FOLDER_OUT=%FOLDER_CURRENT%\%VERSION%\%NAME_PROJECT%
set ICON_PATH=%FOLDER_SRC%\images\cpecmu.ico
set AUT2EXE_ARGS=/in "%FOLDER_SRC%\UninstallPrograms.au3" /out "%FOLDER_OUT%\%EXE_OUT%" /icon "%ICON_PATH%"

echo ----[Compilation with aut2exe]----
aut2exe %AUT2EXE_ARGS%
echo -----------------------------------