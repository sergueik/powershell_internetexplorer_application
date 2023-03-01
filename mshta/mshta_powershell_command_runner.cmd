@echo OFF
REM origin: http://forum.oszone.net/thread-353158.html
REM Warning: notoriously flaky
REM echo X=12>c:\temp\test.txt
REM in Powershell console:
REM $path =( $env:TEMP + '\' + 'test.txt'); (Get-Content -Path $path) -replace 'X=.+$', 'X=0' | Set-Content -Path $path
REM works
REM in CMD console:
REM powershell.exe -Command "&{ $path = ( $env:TEMP + '\' + 'test.txt'); (Get-Content -Path $path) -replace 'X=.+$', 'X=0' | Set-Content -Path $path }"
REM works
mshta vbscript:Execute("command = ""C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe &{ $path = c:\temp\test.txt';  Set-Content -Path $path  -value ''}"":set shell = CreateObject(""WScript.Shell""):shell.Run command, 0:close")
REM does not work ?
REM mshta vbscript:Execute("command = ""powershell.exe -Command """"&{ $path = ( $env:TEMP + '\' + 'test.txt'); (Get-Content -Path $path) -replace 'X=.+$', 'X=0' | Set-Content -Path $path }"""""":set shell = CreateObject(""WScript.Shell""):shell.Run command, 0:close")
REM does not work
type c:\temp\test.txt
REM type %TEMP%\test.txt
