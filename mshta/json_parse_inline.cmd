@echo off
REM This exercises JSON processing with mshta.exe doing eval of JSON file contents

if "%DEBUG%" equ "" set DEBUG=false

REM NOTE: mshta.exe fails with a incomprehensive message
REM if the inline script exceeds certain size between 495 and 519 characters
REM making no serious script possible because of script size limit

SETLOCAL
set RESULTS_FILENAME=%1
set RESULTS_DIRECTORY=results
set DEFAULT_RESULTS_FILENAME=result_.json
if "%RESULTS_FILENAME%" equ "" set RESULTS_FILENAME=%DEFAULT_RESULTS_FILENAME%
set RESULTS=%RESULTS_DIRECTORY%\%RESULTS_FILENAME%
if NOT EXIST %RESULTS%  echo Report does not exist %RESULTS% && exit /b 1
pushd %RESULTS_DIRECTORY%
set RESULTS=%RESULTS_FILENAME%
if /i "%DEBUG%" equ "true" 1>&2 echo Parsing %RESULTS%

set "SCRIPT=javascript:{"

set "SCRIPT=%SCRIPT% var f = new ActiveXObject('Scripting.FileSystemObject');"
set "SCRIPT=%SCRIPT% var o = f.GetStandardStream(1);"
set "SCRIPT=%SCRIPT% var h = f.OpenTextFile('%RESULTS%');"
set "SCRIPT=%SCRIPT% t = h.ReadAll();"
set "SCRIPT=%SCRIPT% h.close();"
set "SCRIPT=%SCRIPT% var j=eval('(' + t + ')');"
set "SCRIPT=%SCRIPT% e=j['examples'];"
set "SCRIPT=%SCRIPT% for (i = 0; i != e.length;i ++){"
set "SCRIPT=%SCRIPT% r=e[i];"
set "SCRIPT=%SCRIPT% if (!(r['status'].match(/passed|pending/))) {"
set "SCRIPT=%SCRIPT% o.Write( 'Test: ' + r['full_description'] + '\n' + 'Status: ' +  r['status'] + '\n');
set "SCRIPT=%SCRIPT% }}"
set "SCRIPT=%SCRIPT% o.Write('Summary: '+ '\n' + j['summary_line']+ '\n');"
set "SCRIPT=%SCRIPT% close();}"
REM if /i "%DEBUG%" equ "true" set "SCRIPT=%SCRIPT% out.Write(navigator.userAgent + '\n');";

if /i "%DEBUG%" equ "true" 1>&2 echo mshta.exe "%SCRIPT%"
REM collect the output from mstha.exe
for /F "delims=" %%_ in ('mshta.exe "%SCRIPT%" 1 ^| more') do echo %%_
popd
ENDLOCAL
