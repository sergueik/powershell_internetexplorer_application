@echo off
REM This example exercises JSON processing by mshta.exe doing the JSON.parse method call
REM https://www.w3schools.com/js/js_json_parse.asp
REM https://stackoverflow.com/questions/44547979/batch-parsing-json-file-with-colons-in-value
REM set DEBUG to true to print additional innformation to the console
if "%DEBUG%" equ "" set DEBUG=false

REM NOTE: mshta.exe fails with a incorehensive message once the inline script exceeds certain size:
REM 495 chars is OK
REM 519 chars is not OK


SETLOCAL
set FILEPATH=%1
if "%FILEPATH%" equ "" set FILEPATH=result_.json
REM results\
echo Parsing %FILEPATH%
REM Frequent error
REM The data necessary to  complete this operation is nt yet available
REM JSON is undefined. Does it need Window?
REM https://www.devcurry.com/2010/12/resolve-json-is-undefined-error-in.html
REM https://stackoverflow.com/questions/8332362/script5009-json-is-undefined
REM https://github.com/douglascrockford/JSON-js/blob/master/json2.js
REM which is pretty heavy
REM https://social.msdn.microsoft.com/Forums/ie/en-US/home?forum=iewebdevelopment
REM based on https://deploywindows.com/2010/08/20/force-ie8-mode-in-hta/
REM one has to modify

REM HKLM\SOFTWARE\Wow6432Node\Microsoft\Internet Explorer\MAIN\FeatureControl\FEATURE_BROWSER_EMULATION
REM HKCU\SOFTWARE\Wow6432Node\Microsoft\Internet Explorer\MAIN\FeatureControl\FEATURE_BROWSER_EMULATION
REM or
REM \SOFTWARE\Microsoft\Internet Explorer\MAIN\FeatureControl\FEATURE_BROWSER_EMULATION
REM mshta.exe RegDword 8
REM
set "SCRIPT="javascript:{"

set "SCRIPT=%SCRIPT% var fso = new ActiveXObject('Scripting.FileSystemObject');"
set "SCRIPT=%SCRIPT% var out = fso.GetStandardStream(1);"
set "SCRIPT=%SCRIPT% var fh = fso.OpenTextFile('pom.xml', 1, true);"
REM https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/opentextfile-method
set "SCRIPT=%SCRIPT% var _fh = fso.OpenTextFile('%FILEPATH%', 1, true);"
set "SCRIPT=%SCRIPT% var _text = _fh.ReadAll();"
set "SCRIPT=%SCRIPT% _fh.close();"
REM set "SCRIPT=%SCRIPT% try { "

REM set "SCRIPT=%SCRIPT% var _json = JSON.parse(_text);
REM set "SCRIPT=%SCRIPT% _out.Write(_json.summary_line);"
REM set "SCRIPT=%SCRIPT% } catch (e) {}"
REM  230
if /i "%DEBUG%" equ "true" set "SCRIPT=%SCRIPT% out.Write(navigator.userAgent + '\n');";
set "SCRIPT=%SCRIPT% close();}""

if /i "%DEBUG%" equ "true" echo mshta.exe %SCRIPT%
REM mshta.exe %SCRIPT%
REM the next line demonstrates how to collect the response from mstha.exe
for /F "delims=" %%_ in ('mshta.exe %SCRIPT% 1 ^| more') do echo %%_
ENDLOCAL




