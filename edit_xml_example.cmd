@set @a=0/*&echo off&chcp 1251>nul
echo "%1"|>nul find ":"||goto VP
REM ::::::::::::::::::::::::::::::::::
cscript /nologo /e:jscript %0 %1
REM origin: https://www.cyberforum.ru/cmd-bat/thread2800778-page3.html#post17146613
exit /b
 
:VP
REM ::::::::::::::::::::::::::::::::::
exit /b
 
*/xml = 'C:/Proxifier/Profiles/Default.ppx'
arg = WSH.arguments(0).split(':')
x = WSH.CreateObject('MSXML2.DOMDocument.6.0')
x.async = 0; x.load(xml)
x.selectSingleNode('//Address').text = arg[0]
x.selectSingleNode('//Port').text = arg[1]
x.save(xml); x = null; CollectGarbage()
