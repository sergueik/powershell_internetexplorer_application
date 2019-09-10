@echo off
REM This example exercises XML processing in mshta.exe


SETLOCAL

REM set DEBUG to true to print additional innformation to the console
if "%DEBUG%" equ "" set DEBUG=false

REM This script illustrates the selectSingleNode method.
REM NOTE: mshta.exe fail when inline script exceeds certain size:
REM 495 chars is OK
REM 519 chars is not OK
REM note that the earlier revision has 
REM set "SCRIPT=mshta.exe "javascript:{"

set "SCRIPT=javascript:{"
set "SCRIPT=%SCRIPT% var fso = new ActiveXObject('Scripting.FileSystemObject');"
set "SCRIPT=%SCRIPT% var out = fso.GetStandardStream(1);"
set "SCRIPT=%SCRIPT% var fh = fso.OpenTextFile('pom.xml', 1, true);"
REM the COM object selection matters here. Uncoment the next line and it will fail
REM set "SCRIPT=%SCRIPT% var xd = new ActiveXObject('Msxml2.DOMDocument.6.0');"
set "SCRIPT=%SCRIPT% var xd = new ActiveXObject('Msxml2.DOMDocument');"
set "SCRIPT=%SCRIPT% xd.async = false;"
set "SCRIPT=%SCRIPT% data = fh.ReadAll();"
set "SCRIPT=%SCRIPT% xd.loadXML(data);"
set "SCRIPT=%SCRIPT% root = xd.documentElement;"
set "SCRIPT=%SCRIPT% var x = '/project/artifactId';"
set "SCRIPT=%SCRIPT% var xmlnode = root.selectSingleNode(x);"
REM set "SCRIPT=%SCRIPT% var xmlnode = root.selectSingleNode('/project/artifactId');"
set "SCRIPT=%SCRIPT% if (xmlnode != null) {"
set "SCRIPT=%SCRIPT%   out.Write(xmlnode.text);"
set "SCRIPT=%SCRIPT% } else {"
set "SCRIPT=%SCRIPT%   out.Write('ERROR');"
set "SCRIPT=%SCRIPT% }"
set "SCRIPT=%SCRIPT% close();}"

if /i "%DEBUG%"=="true" echo mshta.exe "%SCRIPT%"

REM the next line demonstrates how to collect the response from mstha.exe
for /F "delims=" %%_ in ('mshta.exe "%SCRIPT%" 1 ^| more') do echo %%_
ENDLOCAL
exit /b
