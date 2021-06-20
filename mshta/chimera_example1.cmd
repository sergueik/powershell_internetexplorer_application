@echo off
REM https://www.cyberforum.ru/vbscript-wsh/thread1223576.html
chcp 1251 >NUL
set "string=Всё будет хорошо хорошо хорошо."
for /f "delims=" %%. in ('echo %string%^|mshta.exe "%~f0"') do set "b64=%%."
echo %b64%
pause
exit /b
==========================================================================================
<script>
    var string=new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(0).ReadLine();
    var E, LineWriteTimerID,txtResults=''
    var E=new ActiveXObject('WScript.Shell').Exec('powershell -c $OutputEncoding = [Console]::outputEncoding =[System.Text.Encoding]::GetEncoding(\'windows-1251\');Write-Host \''+string+'\';');
    var output = E.StdOut.ReadAll().replace(new RegExp("\\r?\\n", "g"), ""); //\r\n - не возвращается обратно в Bat
    E.StdIn.Close();
    alert('resultFromPowershell='+output)   
    txtResults=output+' Хорошо?'
    new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).Write(txtResults);
    window.close();
</script>

