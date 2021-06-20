@echo off
REM https://www.cyberforum.ru/vbscript-wsh/thread1223576.html

chcp 866 >NUL
set "string=Всё будет хорошо хорошо хорошо."
SetLocal EnableExtensions
for /F "delims=" %%a in ("%string%") do chcp 1251 >NUL&echo %string%|mshta.exe "%~f0"
@chcp 866 >NUL
 
set "psCmd=powershell -c "add-type -an system.windows.forms;[Console]::outputEncoding=[System.Text.Encoding]::GetEncoding('cp866');$F=[Text.Encoding]::GetEncoding('CP866');$T=[Text.Encoding]::GetEncoding('Windows-1251');$tx=[Windows.Forms.Clipboard]::GetText();$b=$T.GetBytes($tx);$b=[Text.Encoding]::Convert($F,$T,$b);$T.GetString($b)""
 
setlocal enabledelayedexpansion
set counter=0
for /F "usebackq delims=" %%# in (`%psCmd%`) do (
    set name=jsPar!counter!
    for /f %%b in ("!name!") do ( set name=%%b&set "!name!=%%#")
    set /a "counter+=1"
)
:: чтобы не было всё просто перечислим строки через перечисляемую переменную
for /F "tokens=2 delims==" %%s in ('set jsPar') do (
    echo %%s
)
::call echo %%jsPar0%%
::call echo %%jsPar1%%
::echo %jsPar0%
::echo %jsPar1%
pause
endlocal&exit /b %errorlevel%
==========================================================================================
<meta charset=cp-866/>
<script>
    var shell = new ActiveXObject('WScript.Shell');
    var string=new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(0).ReadLine();
    //alert(string+' Всё хорошо?')
    //Это то, что я недопонимал. Кодировка, переданная сюда bat 1251(правильная),
    //но её пришлось переотобразить под нынешнее неправильное отображение символов файла 866 кодировки, как кодировки 1251
    //Перекодировать пришедшую из батника правильную 1251 кодировку не пришлось - только сделал другое отображение.
    //И тут у нас целых 2 простых способа сделать это!
    //var E=shell.Exec('powershell -c [Console]::outputEncoding=[System.Text.Encoding]::GetEncoding(\'windows-1251\');Write-Host \''+string+'\';');
    var E=shell.Exec('cmd /c @echo off&chcp 1251>NUL&echo '+string);
    //Теперь напишем функцию alert, для этого неправильного отображения 
    //(Используем отображение консоли для JS - Windows-1251 и перекодируем текст в родную для этого файла 866 кодировку)
    var alertPS = function(s){
        shell.Run("powershell -c [Console]::outputEncoding = [System.Text.Encoding]::GetEncoding('Windows-1251');$F=[Text.Encoding]::GetEncoding('CP866');$T=[Text.Encoding]::GetEncoding('Windows-1251');$b=$T.GetBytes(\"\"\""+s+"\"\"\");$b=[Text.Encoding]::Convert($F,$T,$b);$f='System.Windows.Forms';Add-Type -AssemblyName $f;($o=New-Object $f'.Form').TopMost=$True;[Windows.Forms.MessageBox]::Show($o,$T.GetString($b),'',0,48)",0,true);
    };
    //Получим результат Exec PS по коррекции отображения символов для этого неправильного javascript
    var str1251 = E.StdOut.ReadAll().replace(new RegExp("\\r?\\n", "g"), ""); //\r\n - не возвращается обратно в Bat
    E.StdIn.Close();
    alertPS(str1251)    
    //alert(str1251)
    //var str866=shell.Exec("powershell -c [Console]::outputEncoding = [System.Text.Encoding]::GetEncoding('CP866');$F=[Text.Encoding]::GetEncoding('CP866');$T=[Text.Encoding]::GetEncoding('Windows-1251');$b=$T.GetBytes(\"\"\""+str1251+"\"\"\");$b=[Text.Encoding]::Convert($F,$T,$b);$T.GetString($b)");
    //new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).Write(str866); //GetStandardStream(1) - отказался работать в 866 кодировке напрочь!
    clipboardData.setData('text',str1251+'\nВсё хорошо?');
    window.close();
</script>
