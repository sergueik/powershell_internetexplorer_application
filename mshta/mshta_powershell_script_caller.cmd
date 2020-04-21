@echo OFF
REM based on https://www.cyberforum.ru/powershell/thread2477120-page2.html
set SCRIPT_NAME=%1
set DEMO_SCRIPT_NAME=dummy.ps1
REM vb_input.ps1 is hanging
if "%SCRIPT_NAME%" equ "" set SCRIPT_NAME=%DEMO_SCRIPT_NAME%
REM set DEBUG to true to print additional innformation to the console
if "%DEBUG%" equ "" set DEBUG=false
set SCRIPT_PATH=%~dp0%SCRIPT_NAME%
if /i "%DEBUG%"=="true" echo on
mshta.exe vbscript:Execute("CreateObject(""WScript.Shell"").Run ""C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass & '%SCRIPT_PATH%'"", 0, True:close :")
@echo off

