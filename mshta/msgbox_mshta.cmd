@echo OFF
set MESSAGE=%~1
set TITLE=%~2
1>&2 echo MESSAGE=%MESSAGE%
1>&2 echo TITLE=%TITLE%
REM NOTE that the if removed the colon,
REM the following line will cause cmd show usage of "REM"
REM Records comments (remarks) in a batch file or CONFIG.SYS.
REM REM [comment]
REM see also
: REM http://blog.sevagas.com/?Hacking-around-HTA-files
REM https://www.robvanderwoude.com/usermessages.php#MsHta
REM Microsoft.VisualBasic.MsgBoxStyle
REM https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.msgboxstyle?view=netframework-4.5
REM vbAbortRetryIgnore=2	
REM Abort, Retry, and Ignore buttons
REM
REM vbApplicationModal=0	
REM Application modal message box
REM
REM vbCritical=16	
REM Critical message
REM
REM vbDefaultButton1=0	
REM First button is default
REM
REM vbDefaultButton2=256	
REM Second button is default
REM
REM vbDefaultButton3=512	
REM Third button is default
REM
REM vbExclamation=48	
REM Warning message
REM
REM vbInformation=64	
REM Information message
REM
REM vbMsgBoxHelp=16384	
REM Help text
REM NOTE: icon is not displayed
REM
REM vbMsgBoxRight=524288	
REM Right-aligned text
REM
REM vbMsgBoxRtlReading=1048576	
REM Right-to-left reading text (Hebrew and Arabic systems)
REM
REM vbMsgBoxSetForeground=65536	
REM Foreground message box window
REM
REM vbOKCancel=1	
REM OK and Cancel buttons
REM
REM vbOKOnly=0	
REM OK button only (default)
REM
REM vbQuestion=32	
REM Warning query
REM
REM vbRetryCancel=5	
REM Retry and Cancel buttons
REM
REM vbSystemModal=4096	
REM System modal message box
REM
REM vbYesNo=4	
REM Yes and No buttons
REM
REM vbYesNoCancel=3	
REM Yes, No, and Cancel buttons
REM
REM for reading long message from the file
REM dim fso, file, content: set fso = CreateObject(""Scripting.FileSystemObject""): set file = fso.OpenTextFile("".\message.txt"", 1): content = file.ReadAll: file.close
mshta.exe vbscript:Execute("MsgBox ""%MESSAGE%"",vbQuestion or vbOKOnly ,""%TITLE%"":close()")
