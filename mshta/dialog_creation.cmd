@echo off

SETLOCAL
set TITLE=%~1
REM No space in title allowed
if "%TITLE%" equ "" set TITLE=the_title

set MESSAGE=%~2
REM No space in message allowed
if "%MESSAGE%" equ "" set MESSAGE=the_message

REM set DEBUG to true to print script content and other information to the console
if "%DEBUG%" equ "" set DEBUG=false

call :CALL_JAVASCRIPT artifactId
set ARTIFACTID=%VALUE%

:FINISH

ENDLOCAL
exit /b

:CALL_JAVASCRIPT
REM This script illustrates the Run method
set SCRIPT_NAME=vb_input.ps1
set SCRIPT_PATH=%~dp0%SCRIPT_NAME%
set SCRIPT_PATH=.\%SCRIPT_NAME%

set SCRIPT_PATH=%SCRIPT_PATH:\=\\%
if /I "%DEBUG%"=="true" echo SCRIPT_PATH=%SCRIPT_PATH%

set "SCRIPT=javascript:{"
set "SCRIPT=%SCRIPT% var w = new ActiveXObject('WScript.shell');"
REM Case sensitive
set PATH=%PATH%;C:\Windows\System32\WindowsPowerShell\v1.0
if /I "%DEBUG%"=="true" set "SCRIPT=%SCRIPT% w.run('powershell.exe -ExecutionPolicy Bypass -windowstyle hidden -file "%SCRIPT_PATH%" "%TITLE%" "%MESSAGE%" -debug',0, 1);"
if /I  NOT "%DEBUG%"=="true" set "SCRIPT=%SCRIPT% w.run('powershell.exe -ExecutionPolicy Bypass -windowstyle hidden -file "%SCRIPT_PATH%" "%TITLE%" "%MESSAGE%"',0, 1);"
set "SCRIPT=%SCRIPT% close();}"
REM NOTE: No simple way to get the user entered input passed to MSHTA script
REM workarounds possible, through file
if /I "%DEBUG%"=="true" echo mshta.exe "%SCRIPT%"
if /I "%DEBUG%"=="true" for /F "delims=" %%_ in ('mshta.exe "%SCRIPT%" 1 ^| more') do echo %%_
if /I  NOT "%DEBUG%"=="true" for /F "tokens=2 delims==" %%_ in ('mshta.exe "%SCRIPT%" 1 ^| more') do set VALUE=%%_
ENDLOCAL
exit /b



REM http://www.cyberforum.ru/powershell/thread2477120-page2.html
REM https://ss64.com/vb/run.html
REM mshta.exe vbscript:Execute("CreateObject(""WScript.Shell"").Run ""C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass & 'vb_input.ps1' 'test' 'message'"", 0, True:close :")
REM mshta.exe vbscript:Execute("CreateObject(""WScript.Shell"").Run ""C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -file 'vb_input.ps1'"", 2, True:")


REM Set objShell = WScript.CreateObject("WScript.Shell")
REM Set FSO = CreateObject("Scripting.FileSystemObject")
REM Set F = FSO.GetFile(Wscript.ScriptFullName)
REM path = FSO.GetParentFolderName(F)
REM objShell.Run(CHR(34) & "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe "" -ExecutionPolicy Bypass & ""'"  & path & "\ИМЯСКРИПТА.ps1'" & CHR(34)), 0, True
