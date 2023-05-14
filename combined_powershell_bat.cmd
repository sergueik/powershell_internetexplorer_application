<# :
@echo off
REM origin: http://forum.oszone.net/thread-353571.html
call :screencode
call :sendmail
echo unsubst Z: drive
subst.exe Z: /d
goto :EOF
:screencode
echo subst Z: drive
echo subst.exe Z: %TEMP%
subst.exe Z: %TEMP%
echo Create screen directory
if NOT EXIST "Z:\screen" mkdir Z:\screen
echo Running %~f0 in Powershell
powershell.exe -noprofile -executionpolicy bypass "&{[ScriptBlock]::Create((get-content -literalpath '%~f0') -join [Char]10).Invoke()}"
exit /b
REM #>
add-type -assemblyname System.Windows.Forms
$s = [Windows.Forms.SystemInformation]::VirtualScreen
$b = new-object System.Drawing.Bitmap $s.Width, $s.Height
$g = [System.Drawing.Graphics]::FromImage($b)
$g.CopyFromScreen($s.Location, [System.Drawing.Point]::Empty, $s.Size)
$g.Dispose()
$f = ('Z:\screen\' + ( get-date  -uformat '%Y-%m-%d-%H-%M-%S' ) + '.png')
write-output $f
$b.Save($f)
$b.Dispose()
<# :
goto :EOF

:sendmail
set "SENDER=Z:\SENDER.exe"
echo sending email
echo "%SENDER%"
REM cmd loop trick
set "$_user1=user1@gmail.com"
set "$_user2=user2@gmail.com"

for /f "tokens=2 delims==" %%. in ('set $_') do call :cmd_env_loop "%%."

goto :EOF

:cmd_env_loop

if "%~1" equ "" goto :EOF
pushd Z:
echo sendng to "%~1"
popd
goto :EOF
REM
REM #>
