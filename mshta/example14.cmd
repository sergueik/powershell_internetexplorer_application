@echo off
SETLOCAL
set DAYS=%1
if "%DAYS%" equ "" set DAYS=10
call :CALL_JAVASCRIPT %DAYS%
set NEW_DATE=%VALUE%

echo NEW_DATE=%NEW_DATE%
goto :EOF
:CALL_JAVASCRIPT


REM This script illustrates javascript functions. 
REM NOTE: limited size
set "SCRIPT=javascript:{"
set "SCRIPT=%SCRIPT% var o=new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1);"
REM https://stackoverflow.com/questions/563406/how-to-add-days-to-date
set "SCRIPT=%SCRIPT%Date.prototype.addDays=function(x) {"
set "SCRIPT=%SCRIPT%var d=new Date(this.valueOf());"
set "SCRIPT=%SCRIPT%d.setDate(d.getDate()+x);return d};"
REM https://stackoverflow.com/questions/23593052/format-javascript-date-as-yyyy-mm-dd
set "SCRIPT=%SCRIPT%var d=new Date(Date.now());"
set "SCRIPT=%SCRIPT%var n=%~1;"
REM Locale specific
REM set "SCRIPT=%SCRIPT%o.Write(d.addDays(n).toLocaleString());"
set "SCRIPT=%SCRIPT%function p(n){var t='0'+ n.toString();return t.substring(t.length-2)}"
set "SCRIPT=%SCRIPT%function f(d){return [d.getFullYear(),p(d.getMonth()+1),p(d.getDate())].join('.')}"
set "SCRIPT=%SCRIPT%o.Write(f(d.addDays(n)));"
set "SCRIPT=%SCRIPT%close();}"
REM https://stackoverflow.com/questions/23593052/format-javascript-date-as-yyyy-mm-dd
if /i "%DEBUG%"=="true" echo mshta.exe "%SCRIPT%"
if /i "%DEBUG%"=="true" for /F "delims=" %%_ in ('mshta.exe "%SCRIPT%" 1 ^| more') do echo %%_

for /F "tokens=*" %%_ in ('mshta.exe "%SCRIPT%" 1 ^| more') do set VALUE=%%_



ENDLOCAL
exit /b
