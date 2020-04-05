@echo off
REM based on: https://stackoverflow.com/questions/44547979/batch-parsing-json-file-with-colons-in-value

setlocal enableDelayedExpansion

if "%DEBUG%" equ "" set DEBUG=false

set RESULTS_FILENAME=%1
set RESULTS_DIRECTORY=results
set DEFAULT_RESULTS_FILENAME=result_.json
if "%RESULTS_FILENAME%" equ "" set RESULTS_FILENAME=%DEFAULT_RESULTS_FILENAME%
set CURRENT_DIR=%CD%
set RESULTS=%CURRENT_DIR%\%RESULTS_DIRECTORY%\%RESULTS_FILENAME%

REM Extension should be "cmd"

if "%DEBUG%" equ "" set DEBUG=false

REM Special symbol replacement table:
REM ! ^^!
REM < ^<
REM > ^>
REM | ^|
REM % %%

REM When echoing <  or > use syntax
REM echo ^</script^>>>%GENERATED_SCRIPT%
REM Otherwise the syntax is acceptable - not required but appears to be cleaner:
REM echo>>%GENERATED_SCRIPT% window.visible = false;

set GENERATED_SCRIPT=%TEMP%\script%RANDOM%.cmd
if /i "%DEBUG%" equ "true" 1>&2 echo Generating %GENERATED_SCRIPT%

echo. >%GENERATED_SCRIPT%
echo ^<HTA:Application ShowInTaskbar=no WindowsState=Minimize SysMenu=No ShowInTaskbar=No Caption=No Border=Thin^>>>%GENERATED_SCRIPT%
echo ^<^^!-- TODO^: switch IE to standards-mode by adding a valid doctype. --^>>>%GENERATED_SCRIPT%
echo ^<meta http-equiv="x-ua-compatible" content="ie=edge" /^>>>%GENERATED_SCRIPT%
echo ^<script language="javascript" type="text/javascript"^>>>%GENERATED_SCRIPT%
echo>>%GENERATED_SCRIPT% window.visible = false;
echo>>%GENERATED_SCRIPT% var debug = false;
echo>>%GENERATED_SCRIPT% var _out = new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1);
echo>>%GENERATED_SCRIPT% var _in = new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(0).ReadLine();
echo>>%GENERATED_SCRIPT% var _fh = new ActiveXObject("Scripting.FileSystemObject").OpenTextFile(_in, 1);
echo>>%GENERATED_SCRIPT% var _json = JSON.parse(_fh.ReadAll()); _fh.Close();
echo>>%GENERATED_SCRIPT% if (debug){ _out.Write(navigator.userAgent + '\r\n'); }
echo>>%GENERATED_SCRIPT% var _examples = _json['examples'];
echo>>%GENERATED_SCRIPT% statuses = ['passed', 'pending']
echo>>%GENERATED_SCRIPT% /* not counting pending examples */

echo.>>%GENERATED_SCRIPT%
echo>>%GENERATED_SCRIPT% var _stats = { 'passed':0, 'failed':0, 'pending': 0 };
echo>>%GENERATED_SCRIPT% if (debug){
echo>>%GENERATED_SCRIPT%   for (prop in _examples[0]) {
echo>>%GENERATED_SCRIPT%     _out.Write( prop + '=' + _examples[0][prop] + '\n');
echo>>%GENERATED_SCRIPT%   }
echo>>%GENERATED_SCRIPT% }
echo>>%GENERATED_SCRIPT% var statuses_regexp = new RegExp('(' + statuses.join('^|') + ')');
echo>>%GENERATED_SCRIPT% for ( cnt = 0; cnt ^^!= _examples.length;cnt ++){
echo>>%GENERATED_SCRIPT%   var _example = _examples[cnt];
echo>>%GENERATED_SCRIPT%   var _status = _example['status'];
echo>>%GENERATED_SCRIPT%   _stats[_status] = _stats[_status] + 1;
echo>>%GENERATED_SCRIPT%   if ( ^^!(_status.match(statuses_regexp))) {
echo>>%GENERATED_SCRIPT%     var full_description = _example['full_description'];
echo>>%GENERATED_SCRIPT%     short_description = full_description.split(/\n^|\\n/).slice(0,1).join(' ');
echo>>%GENERATED_SCRIPT%    _out.Write( 'Test : ' + short_description + '\n' + 'Status: ' + _status + '\n');
echo>>%GENERATED_SCRIPT%   }
echo>>%GENERATED_SCRIPT% }
echo>>%GENERATED_SCRIPT% _out.Write('Summary:' + '\n' +_json.summary_line + '\n');
echo>>%GENERATED_SCRIPT% _out.Write('Stats: ' + Math.round(100 * _stats['passed'] / (_stats['failed'] + _stats['passed'])) + '%%');
echo>>%GENERATED_SCRIPT% window.close();
echo ^</script^>>>%GENERATED_SCRIPT%
if /i "%DEBUG%" equ "true" 1>&2 type %GENERATED_SCRIPT%

if /i "%DEBUG%" equ "true" 1>&2 echo Parsing %RESULTS%
if NOT EXIST %RESULTS%  echo Report does not exist %RESULTS% && exit /b 1
if /i "%DEBUG%" equ "true" 1>&2 echo echo %RESULTS%^|mshta.exe "%GENERATED_SCRIPT%"

REM NOTE: cannot execute directly:
REM echo %RESULTS% |mshta.exe "%GENERATED_SCRIPT%"
REM Line:  30
REM Char:  4
REM Error: The handle is invalid
REM Code:  0
REM URL:   file:///C:/Users/Serguei/AppData/Local/Temp/script.cmd

REM collect the output from mstha.exe
for /f "tokens=* delims=" %%_ in ('echo %RESULTS%^|mshta.exe "%GENERATED_SCRIPT%"') do echo %%_
if ERRORLEVEL 1 echo Error processing %RESULTS% && exit /b 1
del /q %GENERATED_SCRIPT%
goto :EOF
