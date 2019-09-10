@echo off

SETLOCAL

REM set DEBUG to true to print additional innformation to the console
if "%DEBUG%" equ "" set DEBUG=false

set VERBOSE=true
REM TODO: add toggle
REM call :SETUP
REM goto :EOF
REM
set WINDOW_TITLE=Personalization
call :EXEC_MSHTA1_REDIRECT %WINDOW_TITLE%
set HWND=%VALUE%
echo HWND="%HWND%"
call :EXEC_MSHTA2_REDIRECT %WINDOW_TITLE%
set HWND=%VALUE%
echo HWND="%HWND%"
call :CLEANUP

ENDLOCAL
exit /b
goto :EOF

:EXEC_MSHTA1_REDIRECT

REM This script illustrates the CreateTextFile method
REM mshta.exe "javascript:{var fso = new ActiveXObject('Scripting.FileSystemObject');var f = fso.CreateTextFile('c:\\temp\\dummy.txt', true); f.Write('x');f.close();close();}"
REM NOTE: the GetStandardStream method is less stable:
REM mshta.exe "javascript:{var fso = new ActiveXObject('Scripting.FileSystemObject');var f = fso.GetStandardStream(1);f.Write('x');f.close();close();}"
REM Does not work:
REM error varies with OS release:
REM Windows 8.1
REM The data necessary to complete this operation is not yet available.
REM Windows 7
REM The handle is invalid.

set LOG=c:\temp\dummy%RANDOM%.txt
set LOG_TRANSLATED=%LOG:\=\\%
set "SCRIPT=javascript:{"
set "SCRIPT=%SCRIPT%o=new ActiveXObject('Shell.Application');"
set "SCRIPT=%SCRIPT%x=o.Windows();"
set "SCRIPT=%SCRIPT%f=new ActiveXObject('Scripting.FileSystemObject');"
set "SCRIPT=%SCRIPT%c=f.CreateTextFile('%LOG_TRANSLATED%', true);"
set "SCRIPT=%SCRIPT%c.Write(x.Count);c.Close();"
set "SCRIPT=%SCRIPT%close();}"

if /i "%DEBUG%"=="true" echo Script:
if /i "%DEBUG%"=="true" echo mshta.exe "%SCRIPT%"
call mshta.exe "%SCRIPT%"

for /F "tokens=1 delims==" %%_ in ('type %LOG% ^| more') do set VALUE=%%_
del /q %LOG%
ENDLOCAL

exit /b
goto :EOF


:EXEC_MSHTA2_REDIRECT
REM This script illustrates the Windows and LoctionName methods
REM C:\Users\sergueik>mshta.exe "javascript:{ var o = new ActiveXObject('Shell.Application'); var x =o.Windows(); var fso = new ActiveXObject('Scripting.FileSystemObject');var f = fso.CreateTextFile('c:\\temp\\dummy.txt', true); f.Write(x.Count); for (var cnt = 0; cnt < x.Count; cnt++) { w = x.item(cnt);  if (w.LocationName.match('Personalization')) {   f.Write('HWND:\t' + w.HWND); w.Quit();}} close();}"

set LOG=c:\temp\dummy%RANDOM%.txt
set LOG_TRANSLATED=%LOG:\=\\%

set "SCRIPT=javascript:{"
set "SCRIPT=%SCRIPT%o=new ActiveXObject('Shell.Application');"
set "SCRIPT=%SCRIPT%x=o.Windows();"
set "SCRIPT=%SCRIPT%f=new ActiveXObject('Scripting.FileSystemObject');"
set "SCRIPT=%SCRIPT%c=f.CreateTextFile('%LOG_TRANSLATED%', true);"
set "SCRIPT=%SCRIPT%c.Write(x.Count);"
set "SCRIPT=%SCRIPT%if(x.Count==0){close();}for(cnt=0;cnt!=x.Count;cnt++){w =x.item(cnt);"
REM set "SCRIPT=%SCRIPT% if (x.Count == 0) { close(); } for (var cnt = 0; cnt ^< x.Count; cnt++) { w = x.item(cnt);"
set "SCRIPT=%SCRIPT% if (w.LocationName.match('Personalization')) {"
set "SCRIPT=%SCRIPT%c.Write('HWND:\t' + w.HWND);"
set "SCRIPT=%SCRIPT%w.Quit();}}"
set "SCRIPT=%SCRIPT%close();}"

if /i "%DEBUG%"=="true" echo Script:
if /i "%DEBUG%"=="true" echo mshta.exe "%SCRIPT%"
call mshta.exe "%SCRIPT%"

for /F "tokens=1 delims==" %%_ in ('type %LOG% ^| more') do set VALUE=%%_
del /q %LOG%
ENDLOCAL
exit /b
goto :EOF

:SETUP
set THEME=C:\Windows\Resources\Themes\nature.theme
rundll32.exe %SystemRoot%\system32\shell32.dll,Control_RunDLL %SystemRoot%\system32\desk.cpl desk,@Themes /Action:OpenTheme /file:"%THEME%"

goto :EOF

:CLEANUP

1>NUL 2>NUL tasklist.exe /fi "IMAGENAME eq mshta.exe"
REM NOTE errorlevel is not set
1>NUL 2>NUL timeout.exe /T 10 /NOBREAK
1>NUL 2>NUL taskkill.exe /IM mshta.exe /T
goto :EOF
