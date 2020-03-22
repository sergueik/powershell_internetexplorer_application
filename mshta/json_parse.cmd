@echo off

REM origin: https://stackoverflow.com/questions/44547979/batch-parsing-json-file-with-colons-in-value
setlocal enableDelayedExpansion

if "%DEBUG%" equ "" set DEBUG=false

set REPORT_FILENAME=%1
set REPORT_DIR=results
if "%REPORT_FILENAME%" equ "" set REPORT_FILENAME=result_.json
set REPORT_FILE_PATH=%REPORT_DIR%\%REPORT_FILENAME%

REM echo %REPORT_FILE_PATH%|mshta.exe "%~f0"|more
1>&2 echo Parsing %REPORT_FILE_PATH%
for /f "tokens=* delims=" %%_ in ('echo %REPORT_FILENAME%^|mshta.exe "%~f0"') do (
  echo %%_
)

exit /b %ERRORLEVEL%

<HTA:Application ShowInTaskbar=no WindowsState=Minimize SysMenu=No ShowInTaskbar=No Caption=No Border=Thin>
<!-- TODO: switch IE to standards-mode by adding a valid doctype. -->
<meta http-equiv="x-ua-compatible" content="ie=edge" />
<script language="javascript" type="text/javascript">
window.visible = false;
var _out = new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1);
var _in = new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(0).ReadLine();
var _fh = new ActiveXObject("Scripting.FileSystemObject").OpenTextFile(_in, 1);
var _json = JSON.parse(_fh.ReadAll());
_fh.Close();
_out.Write(navigator.userAgent + '\r\n');
_out.Write(_json.summary_line + '\r\n' + _json['examples'][0]['status'] + '\r\n');
window.close();
</script>

