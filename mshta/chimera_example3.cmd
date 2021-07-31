@echo off
REM inspired by https://qna.habr.com/q/1027240
set C=%~nx0
if NOT "%DEBUG%" equ "" echo Running with DEBUG set
REM using two environment parameters: C and DEBUG
REM NOTE:  passing arguments appears tricky when
REM powershell run with command built inline as string
@powershell.exe -ExecutionPolicy Bypass -Command "$debug=$env:DEBUG;$s=(get-content \"%~f0\") -join \"`n\"; $s = $s.substring($s.IndexOf(\"goto :\"+\"EOF\")+9);if ($debug -ne $null){write-output (\"Running:`n{0}\" -f$s);} invoke-expression -command $s"
@goto :EOF
# powershell code

write-output 'running powershell code'
write-output ( 'Caller script: {0}' -f $env:C)
pause

exit 0
<#
# a compact version of this script is:
@echo off
@powershell.exe -ExecutionPolicy Bypass -Command "$_=((Get-Content \"%~f0\") -join \"`n\");iex $_.Substring($_.IndexOf(\"goto :\"+\"EOF\")+9)"
@goto :EOF
# powershell code
#>
