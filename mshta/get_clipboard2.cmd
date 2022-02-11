@echo OFF

REM origin: http://forum.oszone.net/thread-350630.html
REM which is actually a copy of
REM https://stackoverflow.com/questions/128463/use-clipboard-from-vbscript
REM see also 
REM https://gist.github.com/codeartery/fefc96f12dd8789a3621af9ab24dfa1a
REM https://gist.github.com/simply-coded/2a31cbd5a7000cb825907052edbe35fe
REM set DEBUG to true to print additional innformation to the console
if "%DEBUG%" equ "" set DEBUG=false
set outfile=%1
if "%outfile%" equ "" Set "outfile=%CD%\clipboard.txt"

mshta.exe vbscript:execute("A=CreateObject(""HTMLFile"").parentWindow.clipboardData.GetData(""text""): Set B = CreateObject(""Scripting.FileSystemObject"").CreateTextFile(""%OutFILE%"", True): On Error Resume Next: B.Write A: B.Close: close")
exit /B:

