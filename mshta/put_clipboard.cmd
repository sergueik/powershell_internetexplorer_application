@echo OFF

REM set DEBUG to true to print additional innformation to the console
if "%DEBUG%" equ "" set DEBUG=false

REM origin https://stackoverflow.com/questions/128463/use-clipboard-from-vbscript
REM converted from chimera VB Script / JScript example
REM CreateObject("WScript.Shell").Run "mshta.exe ""javascript:clipboardData.setData('text','" & Replace(Replace(Replace(sText, "\", "\\"), """", """"""), "'", "\'") & "'.replace('""""',String.fromCharCode(34)));close();""", 0, True
REM the trivial alternatives are
REM echo %* > %DATAFILE%
REM C:\Windows\System32\clip.exe < %DATAFILE%
REM echo %* | C:\Windows\System32\clip.exe

set "SCRIPT=javascript:{"
set "SCRIPT=%SCRIPT% var _in = new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(0).ReadLine();"
set "SCRIPT=%SCRIPT% clipboardData.setData('text', _in);"
set "SCRIPT=%SCRIPT% close();}"

if /i "%DEBUG%"=="true" echo mshta.exe "%SCRIPT%"

REM the next line demonstrates how to collect the response from mstha.exe
echo %* |mshta.exe "%SCRIPT%"
ENDLOCAL
exit /b
