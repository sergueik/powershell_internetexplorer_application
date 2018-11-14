# based on: https://4sysops.com/archives/building-a-web-server-with-powershell/
# TODO: run as job
# Start-Job -ArgumentList @('-port', 10001) -scriptblock { ...

param(
  $port = $null
)
try {
  add-type -path 'c:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Web.dll'
} catch [Exception]{
  # add-type : Could not load file or assembly 'file:///C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Web.dll' or one of its dependencies. An attempt was made to load a program with an incorrect format.
  # https://social.technet.microsoft.com/Forums/windowsserver/en-US/44f54abe-8e4c-47fe-a553-568b855a19cb/error-trying-to-load-systemwebdll?forum=winserverpowershell
}

# Unable to find type [System.Web.MimeMapping]
[System.Reflection.Assembly]::LoadWithPartialName('System.Web') | out-null

# Exception calling "Start" with "0" argument(s): "Failed to listen on prefix 'http://localhost:18000/' because it conflicts with an existing registration on the machine."
if ($port -eq $null) {
  $port = Get-Random -minimum 10000 -maximum 20000
}
write-output ('Runnng on port ' + $port)
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:${port}/")
$listener.Start() # blocking call
$drivename =  'MyPowerShellSite'
if ((get-psdrive -name $drivename -errorAction silentlycontinue) -ne $null) {
  remove-psdrive -Name $drivename
}
New-PSDrive -Name $drivename -PSProvider FileSystem -Root $PWD.Path | out-null
$Context = $listener.GetContext()
$URL = $Context.Request.Url.LocalPath
$Content = Get-Content -Encoding Byte -Path "MyPowerShellSite:$URL"

$Context.Response.ContentType = [System.Web.MimeMapping]::GetMimeMapping("MyPowerShellSite:$URL")
write-debug $Context.Response.ContentType
$Context.Response.OutputStream.Write($Content, 0, $Content.Length)
$Context.Response.Close()
$listener.Stop()
