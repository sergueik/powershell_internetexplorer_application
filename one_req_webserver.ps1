#Copyright (c) 2018 Serguei Kouzmine
#
#Permission is hereby granted, free of charge, to any person obtaining a copy
#of this software and associated documentation files (the "Software"), to deal
#in the Software without restriction, including without limitation the rights
#to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
#copies of the Software, and to permit persons to whom the Software is
#furnished to do so, subject to the following conditions:
#
#The above copyright notice and this permission notice shall be included in
#all copies or substantial portions of the Software.
#
#THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
#IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
#FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
#AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
#LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
#OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
#THE SOFTWARE.

param(
  $datafile = 'data.html',
  $port = '10001',
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
Start-Job  -ArgumentList @($port, (Get-ScriptDirectory), 'true') -scriptblock {param( $port = $null, $webroot = $null, $debug = $false)

  if ($webroot -ne $null) {
    $logpath =  $webroot
  } else {
    $logpath =  $env:TEMP
  }

  $logfile = "${logpath}\zz.txt"
  $drivename = 'MyPowerShellSite'

  # based on: https://4sysops.com/archives/building-a-web-server-with-powershell/
  # See also a similar gist shttps://gist.github.com/19WAS85/5424431
  # See also https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/creating-powershell-web-server
  if  ($debug) {
    write-output "Param `$port = ${port}" | out-file $logfile -append
    write-output "Param `$webroot = `"${webroot}`"" | out-file $logfile -append
  }
  [System.Reflection.Assembly]::LoadWithPartialName('System.Web') | out-null

  if ($port -eq $null) {
    $port = Get-Random -minimum 10000 -maximum 20000
  }
  if  ($debug) {
    write-output ('Runnng on port ' + $port) | out-file $logfile -append
  }
  $listener = New-Object System.Net.HttpListener
  $listener.Prefixes.Add("http://localhost:${port}/")
  $listener.Start() # blocking call

  if ((get-psdrive -name $drivename -errorAction silentlycontinue) -ne $null) {
    remove-psdrive -Name $drivename
  }

  New-PSDrive -Name $drivename -PSProvider FileSystem -Root $webroot | out-null
  $context = $listener.GetContext()
  $localfile_path = "${drivename}:$($context.Request.Url.LocalPath)"
  $content = Get-Content -Encoding Byte -Path $localfile_path

  $context.Response.ContentType = [System.Web.MimeMapping]::GetMimeMapping($localfile_path)
  # write-debug $context.Response.ContentType
  $context.Response.OutputStream.Write($content, 0, $content.Length)
  $context.Response.OutputStream.Flush()
  start-sleep -millisecond 1000
  $context.Response.Close()
  $listener.Stop()
  remove-psdrive -Name $drivename
}


$uri = ('file:///{0}' -f ((resolve-path $datafile).path -replace '\\', '/'))
write-host -foreground 'Green' ('Reading {0}' -f $uri )
[Microsoft.PowerShell.Commands.WebResponseObject]$obj = (Invoke-WebRequest -Uri $uri)
$obj| select-object -property 'BaseResponse','StatusCode','StatusDescription','Headers','RawContentStream' | format-list

$uri = "http://localhost:${port}/${datafile}"
write-host -foreground 'Green' ('Reading {0}' -f $uri )
[Microsoft.PowerShell.Commands.WebResponseObject]$obj = (Invoke-WebRequest -Uri $uri)
$obj| select-object -property  'BaseResponse','StatusCode','StatusDescription','Headers','RawContentStream','ParsedHtml','Links' | format-list