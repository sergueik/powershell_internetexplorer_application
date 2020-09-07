@set @dummy=0 /*
@set SOURCE=C:\TEST1\FILE.ISO
@set DESTINATION=D:\TEST2

@cscript.exe //nologo /e:jscript "%0" "%SOURCE%" "%DESTINATION%"
@exit /b

*/new ActiveXObject("shell.Application").NameSpace(WScript.Arguments(1)).CopyHere(WScript.Arguments(0),16);
// origin: http://forum.oszone.net/post-1486312-4.html
