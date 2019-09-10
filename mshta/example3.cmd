@echo off

SETLOCAL

REM set DEBUG to true to print additional innformation to the console
if "%DEBUG%" equ "" set DEBUG=false

REM this example exercises passing batch parameters into geerated jscript inline
set TAG=version

set "SCRIPT=javascript:{"
set "SCRIPT=%SCRIPT% var fso = new ActiveXObject('Scripting.FileSystemObject');"
set "SCRIPT=%SCRIPT% var out = fso.GetStandardStream(1);"
set "SCRIPT=%SCRIPT% var fh = fso.OpenTextFile('pom.xml', 1, true);"
set "SCRIPT=%SCRIPT% var xd = new ActiveXObject('Msxml2.DOMDocument.6.0');"
set "SCRIPT=%SCRIPT%   var tags = ['groupId','artifactId','version'];"
set "SCRIPT=%SCRIPT%   for (var cnt in tags ) {"
if /i "%DEBUG%"=="true"  set "SCRIPT=%SCRIPT%     out.write('cnt = ' + cnt + '\n');"
set "SCRIPT=%SCRIPT%     var tag = tags[cnt];"
set "SCRIPT=%SCRIPT%       if (tag.match(RegExp('%TAG%', 'g'))) {"
set "SCRIPT=%SCRIPT%     out.write(tag + '\n');"
set "SCRIPT=%SCRIPT%   }"
set "SCRIPT=%SCRIPT% }"
set "SCRIPT=%SCRIPT% close();}"

if /i "%DEBUG%"=="true" echo mshta.exe "%SCRIPT%"

for /F "delims=" %%_ in ('mshta.exe "%SCRIPT%" 1 ^| more') do echo %%_
exit /b