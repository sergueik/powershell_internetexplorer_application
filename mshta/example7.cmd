@echo off
REM This example exercises XML processing in mshta.exe
SETLOCAL

REM set DEBUG to true to print additional innformation to the console
if "%DEBUG%" equ "" set DEBUG=false

for %%. in (groupId artifactId version ) do (
call :CALL_JAVASCRIPT %%.
)
ENDLOCAL
exit /b

:CALL_JAVASCRIPT

REM NOTE: mshta.exe fail when inline script exceeds certain size:
REM 495 chars is OK
REM 519 chars is not OK
REM Script below tries to save on var declaration, variable names etc. and whitespace, and this sacrifices redability

set "SCRIPT=javascript:{"
set "SCRIPT=%SCRIPT%f=new ActiveXObject('Scripting.FileSystemObject');"
set "SCRIPT=%SCRIPT%c=f.GetStandardStream(1);"
set "SCRIPT=%SCRIPT%h=f.OpenTextFile('pom.xml',1,1);"
set "SCRIPT=%SCRIPT%x=new ActiveXObject('Msxml2.DOMDocument.6.0');"
set "SCRIPT=%SCRIPT%x.async=false;"
set "SCRIPT=%SCRIPT%x.loadXML(h.ReadAll());"
set "SCRIPT=%SCRIPT%r=x.documentElement;"
set "SCRIPT=%SCRIPT%t='%~1';"
set "SCRIPT=%SCRIPT%n=r.childNodes;"
set "SCRIPT=%SCRIPT%for(i=0;i!=n.length;i++){"
set "SCRIPT=%SCRIPT%if (n.item(i).nodeName.match(RegExp(t, 'g'))) {"
set "SCRIPT=%SCRIPT%c.Write(t+'='+n.item(i).text+'\n');"
set "SCRIPT=%SCRIPT%}"
set "SCRIPT=%SCRIPT%}"
set "SCRIPT=%SCRIPT%close();}"

if /i "%DEBUG%"=="true" echo mshta.exe "%SCRIPT%"

REM the next line demonstrates how to consume the response from mstha.exe
for /F "delims=" %%_ in ('mshta.exe "%SCRIPT%" 1 ^| more') do echo %%_
ENDLOCAL
exit /b
