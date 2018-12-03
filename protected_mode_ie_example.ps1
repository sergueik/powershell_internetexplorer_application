# origin: http://forum.oszone.net/thread-337924.html
# (copied from somewhere else)


<#
Set IE = GetObject("new:{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}")
' explanation for the row above:
' https://blogs.msdn.microsoft.com/ieinternals/2011/08/03/default-integrity-level-and-automation/


IE.Visible = False	' set true for debug
IE.ToolBar = 0		' set 1 for debug
IE.StatusBar = 0	' set 1 for debug


IE.Navigate("https://kerio:4081/login")
' login process, must be in intranet zone in Internet Explorer

WScript.Sleep(5000)
' wait for kerio to successfully redirect us on webstatistics page

IE.Quit
' close IE
#>

$clsid = new-object Guid 'D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E'
$type = [Type]::GetTypeFromCLSID($clsid)
$ie = [Activator]::CreateInstance($type)
$object.Drives

$ie.Visible = $true	
# set $true for debug, $false to hide window
$ie.ToolBar = 0		
# may want to set 1 for debug
$ie.StatusBar = 0	
# may want to set 1 for debug

$ie| get-member


# navigate to intranet site
# Internet Explorer’s Protected Mode is a security sandbox that relies upon the integrity level system in Windows

# $ie.Navigate("https://kerio:4081/login")
# NTLM login process, must be in intranet zone in Internet Explorer

$target_url = 'http://store.demoqa.com/products-page/'
$ie.navigate2($target_url)

# NOTE: neither peoperties get response
if (($ie.Busy -ne $null) -and ($ie.ReadyState -ne  $null)){
  while ($ie.Busy -or ($ie.ReadyState -ne 4)) {
    write-debug 'waiting'
    write-debug ('`$ie.Busy = {0}' -f $ie.Busy)
    write-debug ('`$ie.ReadyState = {0}' -f $ie.ReadyState)
    start-sleep -milliseconds 100
  }
} else {
  start-sleep -millisecond 3000
}

# wait for e.g. kerio to successfully redirect the user
try {
  $ie.Quit()
} catch [Exception] {

}


while( ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($ie) <# | out-null #>) ) {}
Remove-Variable ie -ErrorAction SilentlyContinue