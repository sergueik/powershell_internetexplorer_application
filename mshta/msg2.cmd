@echo off
REM '∩╗┐@echo' is not recognized as an internal or external command, operable program or batch file. 
set CODEPAGE=%1
REM origin https://www.cyberforum.ru/cmd-bat/thread3224628.html
REM on downlevel platforms fails if the script is saved in UTF-8 and code page is changed by the following command
REM it appears mshta.exe is not UNICODE aware? 
REM choosing 65000, 65001 will make the following command silently fail
if NOT "%CODEPAGE%"=="" c:\WINDOWS\system32\chcp.com %CODEPAGE% 
start C:\Windows\System32\mshta.exe vbscript:Execute("msgbox ""          ОТПРАВЬТЕ ЗАКАЗ!          "", 48 + 65536 + 4096, ""НАПОМИНАНИЕ"":close")
REM the MsgBox Label text is available in Clipboard

exit /b

