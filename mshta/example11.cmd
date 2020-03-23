@echo off
REM This example exercises XML processing in mshta.exe

SETLOCAL
REM set DEBUG to true to print additional innformation to the console
if "%DEBUG%" equ "" set DEBUG=false

REM NOTE: mshta.exe fail with obscure error when inline script exceeds certain size
REM between 495 chars and 519 chars

set "SCRIPT="javascript:{"
set "SCRIPT=%SCRIPT% var f=new ActiveXObject('Scripting.FileSystemObject');"
set "SCRIPT=%SCRIPT% var o=f.GetStandardStream(1);"
set "SCRIPT=%SCRIPT% var fh=f.OpenTextFile('pom.xml', 1, true);"
set "SCRIPT=%SCRIPT% var x=new ActiveXObject('Msxml2.DOMDocument.6.0');"
set "SCRIPT=%SCRIPT% x.async=false;"
set "SCRIPT=%SCRIPT% data=fh.ReadAll();fh.close();"
set "SCRIPT=%SCRIPT% x.loadXML(data);"
set "SCRIPT=%SCRIPT% r=x.documentElement;"
set "SCRIPT=%SCRIPT% o.Write('# nodes: ' + r.childNodes.length + '\n');"
set "SCRIPT=%SCRIPT% var n = r.childNodes.item(1);"
set "SCRIPT=%SCRIPT% o.Write(n.xml + '\n');"
set "SCRIPT=%SCRIPT% n = r.childNodes.item(3);"
set "SCRIPT=%SCRIPT% o.Write(n.nodeName + '\n');"
if /i "%DEBUG%" equ "true" set "SCRIPT=%SCRIPT% o.Write(navigator.userAgent + '\n');";
set "SCRIPT=%SCRIPT% close();}""

if /i "%DEBUG%"=="true" echo mshta.exe %SCRIPT% 1>&2
REM uncommenting next leads to annoying and difficult to debug
REM "the handle is invalid" and 
REM "the data necessary to complete this operation is not yet available"
REM errors arising from 
REM o.Write
REM these errors disappear when mshta.exe run is output redirected
REM mshta.exe %SCRIPT%

REM consume the console output from mstha.exe

for /F "delims=" %%_ in ('mshta.exe %SCRIPT% 1 ^| more') do echo %%_
ENDLOCAL
exit /b

