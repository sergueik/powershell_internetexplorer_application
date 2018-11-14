param(
 $html_file = $null,
 [switch]$debug
)
function Get-ScriptDirectory
{
  [string]$scriptDirectory = $null

  if ($host.Version.Major -gt 2) {
    $scriptDirectory = (Get-Variable PSScriptRoot).Value
    Write-Debug ('$PSScriptRoot: {0}' -f $scriptDirectory)
    if ($scriptDirectory -ne $null) {
      return $scriptDirectory;
    }
    $scriptDirectory = [System.IO.Path]::GetDirectoryName($MyInvocation.PSCommandPath)
    Write-Debug ('$MyInvocation.PSCommandPath: {0}' -f $scriptDirectory)
    if ($scriptDirectory -ne $null) {
      return $scriptDirectory;
    }

    $scriptDirectory = Split-Path -Parent $PSCommandPath
    Write-Debug ('$PSCommandPath: {0}' -f $scriptDirectory)
    if ($scriptDirectory -ne $null) {
      return $scriptDirectory;
    }
  } else {
    $scriptDirectory = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)
    if ($scriptDirectory -ne $null) {
      return $scriptDirectory;
    }
    $Invocation = (Get-Variable MyInvocation -Scope 1).Value
    if ($Invocation.PSScriptRoot) {
      $scriptDirectory = $Invocation.PSScriptRoot
    } elseif ($Invocation.MyCommand.Path) {
      $scriptDirectory = Split-Path $Invocation.MyCommand.Path
    } else {
      $scriptDirectory = $Invocation.InvocationName.Substring(0,$Invocation.InvocationName.LastIndexOf('\'))
    }
    return $scriptDirectory
  }
}

# TODO: better  handle output https://stackoverflow.com/questions/11973775/powershell-get-output-from-receive-job
# NOTE: start-job argument passing appears to be leading whitespace-sensitive
Start-Job -scriptblock {param( $port = $null, $webroot = $null, $debug = $false)

  # NOTE: not useful in the job context, stay for one commit only
  function Get-ScriptDirectory
  {
    [string]$scriptDirectory = $null

    if ($host.Version.Major -gt 2) {
      $scriptDirectory = (Get-Variable PSScriptRoot).Value
      Write-Debug ('$PSScriptRoot: {0}' -f $scriptDirectory)
      if ($scriptDirectory -ne $null) {
        return $scriptDirectory;
      }
      $scriptDirectory = [System.IO.Path]::GetDirectoryName($MyInvocation.PSCommandPath)
      Write-Debug ('$MyInvocation.PSCommandPath: {0}' -f $scriptDirectory)
      if ($scriptDirectory -ne $null) {
        return $scriptDirectory;
      }

      $scriptDirectory = Split-Path -Parent $PSCommandPath
      Write-Debug ('$PSCommandPath: {0}' -f $scriptDirectory)
      if ($scriptDirectory -ne $null) {
        return $scriptDirectory;
      }
    } else {
      $scriptDirectory = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)
      if ($scriptDirectory -ne $null) {
        return $scriptDirectory;
      }
      $Invocation = (Get-Variable MyInvocation -Scope 1).Value
      if ($Invocation.PSScriptRoot) {
        $scriptDirectory = $Invocation.PSScriptRoot
      } elseif ($Invocation.MyCommand.Path) {
        $scriptDirectory = Split-Path $Invocation.MyCommand.Path
      } else {
        $scriptDirectory = $Invocation.InvocationName.Substring(0,$Invocation.InvocationName.LastIndexOf('\'))
      }
      return $scriptDirectory
    }
  }

  if ($webroot -ne $null) {
    $logpath =  $webroot
  } else {
    $logpath =  $env:TEMP
  }
  # does not get inherited from caller
  if  ($debugpreference -eq 'continue') {
    $debug = $true
  }
  # based on: https://4sysops.com/archives/building-a-web-server-with-powershell/
  if  ($debug) {
    write-output "Param `$port is ${port}" | out-file "${logpath}\zz.txt" -append
    write-output ("Get-ScriptDirectory output is {0}" -f (Get-ScriptDirectory)) | out-file "${logpath}\zz.txt" -append
    write-output "Param `$webroot is `"${webroot}`"" | out-file "${logpath}\zz.txt" -append
  }

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
  if  ($debug) {
    write-output ('Runnng on port ' + $port) | out-file "${logpath}\zz.txt" -append
  }
  $listener = New-Object System.Net.HttpListener
  $listener.Prefixes.Add("http://localhost:${port}/")
  $listener.Start() # blocking call
  $drivename =  'MyPowerShellSite'
  if ((get-psdrive -name $drivename -errorAction silentlycontinue) -ne $null) {
    remove-psdrive -Name $drivename
  }

  # New-PSDrive -Name $drivename -PSProvider FileSystem -Root $PWD.Path | out-null
  New-PSDrive -Name $drivename -PSProvider FileSystem -Root $webroot | out-null
  $Context = $listener.GetContext()
  $URL = $Context.Request.Url.LocalPath
  $Content = Get-Content -Encoding Byte -Path "MyPowerShellSite:$URL"

  $Context.Response.ContentType = [System.Web.MimeMapping]::GetMimeMapping("MyPowerShellSite:$URL")
  write-debug $Context.Response.ContentType
  $Context.Response.OutputStream.Write($Content, 0, $Content.Length)
  $Context.Response.Close()
  $listener.Stop()
} -ArgumentList @('10001', (Get-ScriptDirectory), 'true')