@echo off

SETLOCAL


REM set DEBUG to true to print additional innformation to the console
if "%DEBUG%" equ "" set DEBUG=false

call :CALL_JAVASCRIPT artifactId
set ARTIFACTID=%VALUE%

call :CALL_JAVASCRIPT groupId
set GROUPID=%VALUE%

call :CALL_JAVASCRIPT version
set VERSION=%VALUE%

call :CALL_JAVASCRIPT properties/mainClass
set DEFAULT_MAIN_CLASS=%VALUE%

if /i NOT "%VERBOSE%"=="true" goto :FINISH

echo VERSION="%VERSION%"
echo ARTIFACTID="%ARTIFACTID%"
echo GROUPID="%GROUPID%"
echo DEFAULT_MAIN_CLASS="%DEFAULT_MAIN_CLASS%"
:FINISH

ENDLOCAL
exit /b

:CALL_JAVASCRIPT
REM This script illustrates the selectSingleNode method
set "SCRIPT=javascript:{"
set "SCRIPT=%SCRIPT% var fso = new ActiveXObject('Scripting.FileSystemObject');"
set "SCRIPT=%SCRIPT% var out = fso.GetStandardStream(1);"
set "SCRIPT=%SCRIPT% var fh = fso.OpenTextFile('pom.xml', 1, true);"
set "SCRIPT=%SCRIPT% var xd = new ActiveXObject('Msxml2.DOMDocument');"
set "SCRIPT=%SCRIPT% xd.async = false;"
set "SCRIPT=%SCRIPT% data = fh.ReadAll();"
set "SCRIPT=%SCRIPT% xd.loadXML(data);"
set "SCRIPT=%SCRIPT% root = xd.documentElement;"
set "SCRIPT=%SCRIPT% var xpath = '/project/' + '%~1';"
set "SCRIPT=%SCRIPT% var xmlnode = root.selectSingleNode( xpath);"
set "SCRIPT=%SCRIPT% if (xmlnode != null) {"
set "SCRIPT=%SCRIPT%   out.Write(xpath + '=' + xmlnode.text);"
set "SCRIPT=%SCRIPT% } else {"
set "SCRIPT=%SCRIPT%   out.Write('ERR');"
set "SCRIPT=%SCRIPT% }"
set "SCRIPT=%SCRIPT% close();}"

if /i "%DEBUG%"=="true" echo mshta.exe "%SCRIPT%"
if /i "%DEBUG%"=="true" for /F "delims=" %%_ in ('mshta.exe "%SCRIPT%" 1 ^| more') do echo %%_

for /F "tokens=2 delims==" %%_ in ('mshta.exe "%SCRIPT%" 1 ^| more') do set VALUE=%%_
ENDLOCAL
exit /b
