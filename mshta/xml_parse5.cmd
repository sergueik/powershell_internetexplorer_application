@echo off

SETLOCAL enableDelayedExpansion

REM set DEBUG to true to print additional innformation to the console
if "%DEBUG%" equ "" set DEBUG=false

set XML_FILENAME=%1
set DEFAULT_XML_FILENAME=pom.xml
if "%XML_FILENAME%" equ "" set XML_FILENAME=%DEFAULT_XML_FILENAME%
set CURRENT_DIR=%CD%
set XML_FILEPATH=%CURRENT_DIR%\%XML_FILENAME%
set XML_FILEPATH=%XML_FILEPATH:\=\\%


call :CALL_JAVASCRIPT %XML_FILEPATH% artifactId
set ARTIFACTID=%VALUE%
if /i "%DEBUG%"=="true" goto :REPORT
call :CALL_JAVASCRIPT %XML_FILEPATH% groupId
set GROUPID=%VALUE%

call :CALL_JAVASCRIPT %XML_FILEPATH% version
set VERSION=%VALUE%

:REPORT

echo VERSION="%VERSION%"
echo ARTIFACTID="%ARTIFACTID%"
echo GROUPID="%GROUPID%"
if /i NOT "%DEBUG%"=="true" goto :FINISH

:FINISH

ENDLOCAL
exit /b

:CALL_JAVASCRIPT


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
set XML_FILEPATH=%1
set TAG=%~2
echo. >%GENERATED_SCRIPT%
echo ^<HTA:Application ShowInTaskbar=no WindowsState=Minimize SysMenu=No ShowInTaskbar=No Caption=No Border=Thin^>>>%GENERATED_SCRIPT%
echo ^<^^!-- TODO^: switch IE to standards-mode by adding a valid doctype. --^>>>%GENERATED_SCRIPT%
echo ^<meta http-equiv="x-ua-compatible" content="ie=edge" /^>>>%GENERATED_SCRIPT%
echo ^<script language="javascript" type="text/javascript"^>>>%GENERATED_SCRIPT%
echo>>%GENERATED_SCRIPT% window.visible = false;
echo>>%GENERATED_SCRIPT% var debug = false;
echo>>%GENERATED_SCRIPT% var fso = new ActiveXObject('Scripting.FileSystemObject');
echo>>%GENERATED_SCRIPT% var filepath = '%XML_FILEPATH%';
echo>>%GENERATED_SCRIPT% var out = fso.GetStandardStream(1); var handle = fso.OpenTextFile(filepath,1,1);
echo>>%GENERATED_SCRIPT% var xml = new ActiveXObject('Msxml2.DOMDocument.6.0');
echo>>%GENERATED_SCRIPT% xml.async = false;
echo>>%GENERATED_SCRIPT% xml.loadXML(handle.ReadAll());
echo>>%GENERATED_SCRIPT% root = xml.documentElement;
echo>>%GENERATED_SCRIPT% var tag = '%TAG%';
echo>>%GENERATED_SCRIPT% nodes = root.childNodes;
echo>>%GENERATED_SCRIPT% for(i = 0; i ^^!= nodes .length; i++){
echo>>%GENERATED_SCRIPT%   if (nodes.item(i).nodeName.match(RegExp(tag, 'g'))) {
echo>>%GENERATED_SCRIPT%      out.Write(tag + '=' + nodes.item(i).text + '\n');
echo>>%GENERATED_SCRIPT%   }
echo>>%GENERATED_SCRIPT% }
echo>>%GENERATED_SCRIPT% close();
echo ^</script^>>>%GENERATED_SCRIPT%
if /i "%DEBUG%" equ "true" 1>&2 type %GENERATED_SCRIPT%

if /i "%DEBUG%" equ "true" 1>&2 echo echo %RESULTS%^|mshta.exe "%GENERATED_SCRIPT%"

REM NOTE: cannot execute directly:
REM echo %RESULTS% |mshta.exe "%GENERATED_SCRIPT%"
REM Line:  30
REM Char:  4
REM Error: The handle is invalid
REM Code:  0
REM URL:   file:///C:/Users/Serguei/AppData/Local/Temp/script.cmd

REM collect the output from mstha.exe
for /F "tokens=2 delims==" %%_ in ('echo %RESULTS%^|mshta.exe "%GENERATED_SCRIPT%" 1 ^| more') do set VALUE=%%_
if ERRORLEVEL 1 echo Error processing %RESULTS% && exit /b 1
if /i NOT "%DEBUG%"=="true" del /q %GENERATED_SCRIPT%
goto :EOF
