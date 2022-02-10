@echo off

REM creating two environment parameter SCRIPT for path to the script
REM detecting environment parameter DEBUG
if NOT "%DEBUG%" equ "" echo Running with DEBUG set
set SCRIPT=%~dpnx0

REM NOTE: passing arguments appears tricky when
REM powershell run with command built inline as string


@powershell.exe -executionPolicy bypass -command ^
"$f='%SCRIPT%';$debug=$env:DEBUG;$s=(get-content $f) -join \"`n\"; $s = $s.substring($s.IndexOf(\"goto :\"+\"EOF\")+9);if ($debug -ne $null){write-output (\"Running:`n{0}\" -f $s);} invoke-expression -command $s"
@goto :EOF
# powershell code

write-output 'running powershell code'
write-output ( 'Caller script: {0}' -f $f)
pause

exit 0
