@echo OFF
REM origin: http://forum.oszone.net/thread-283213-4.html
REM see also: https://www.dostips.com/forum/viewtopic.php?t=5311
REM see also: https://www.sparxsystems.com/forums/smf/index.php?topic=40060.0

REM set DEBUG to true to print additional innformation to the console
if "%DEBUG%" equ "" set DEBUG=false

set "SCRIPT=javascript:{"
set "SCRIPT=%SCRIPT% var s=clipboardData.getData('text');"
set "SCRIPT=%SCRIPT% if(s) new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).Write(s);"
set "SCRIPT=%SCRIPT% close();}"

REM TODO: remove program delimeters
REM MSHTA.EXE "javascript:var s=clipboardData.getData('text');if(s)new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).Write(s);close();"

if /i "%DEBUG%"=="true" echo mshta.exe "%SCRIPT%"

REM the next line demonstrates how to collect the response from mstha.exe
for /F "delims=" %%_ in ('mshta.exe "%SCRIPT%" 1 ^| more') do echo %%_
ENDLOCAL
exit /b

