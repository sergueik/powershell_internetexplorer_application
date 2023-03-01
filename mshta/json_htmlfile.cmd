<!-- : batch portion

@echo off & setlocal
REM origin: https://stackoverflow.com/questions/33797681/beautify-curl-download-of-json-before-writing-to-file/33808191#33808191
REM see also: https://stackoverflow.com/questions/19445189/cscript-jscript-json
REM see also: https://github.com/douglascrockford/JSON-js

REM The for /F loop forces mshta to communicate with stdout
REM as a console script host.  Without for /f, attempting
REM to write to stdout results in an invalid handle error.
for /F "delims=" %%_ in ('mshta.exe "%~f0"') do echo.%%_
goto :EOF


REM Usage:
REM   json_htmlfile.cmd < test.json
REM 
REM {
REM         "a": "B",
REM         "status": "passed",
REM         "run_time": 0.013529833,
REM         "pending_message": null,
REM         "array": [
REM                 1,
REM                 2,
REM                 3,
REM                 4,
REM                 5
REM         ]
REM }
REM


REM load htmlfile COM object and declare empty JSON object
REM var htmlfile = WSH.CreateObject('htmlfile'), JSON;

REM force htmlfile to load Chakra engine
REM htmlfile.write('<meta http-equiv="x-ua-compatible" content="IE=9" />');

REM The following statement is an overloaded compound statement, a code golfing trick.
REM The "JSON = htmlfile.parentWindow.JSON" statement is executed first, copying the
REM htmlfile COM object's JSON object and methods into "JSON" declared above; then
REM "htmlfile.close()" ignores its argument and unloads the now unneeded COM object.
REM htmlfile.close(JSON = htmlfile.parentWindow.JSON);

end batch / begin HTA : -->

<meta http-equiv="x-ua-compatible" content="IE=9" />
<script>
var fso = new ActiveXObject('Scripting.FileSystemObject'),
    stdin = fso.GetStandardStream(0),
    stdout = fso.GetStandardStream(1),
    json = stdin.ReadAll(),
    pretty = JSON.stringify(JSON.parse(json), null, '\t');

close(stdout.Write(pretty));
</script>
