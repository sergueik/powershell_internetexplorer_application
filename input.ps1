# http://wsh2.uw.hu/ch08c.html
param($message = 'message')
$ie = new-object -com 'internetexplorer.application'; $ie.visible = $false; $ie.navigate2('about:blank')
while ($ie.Busy -or ($ie.ReadyState -ne 4)) { start-sleep -milliseconds 100}
<#
if (($ie.Busy -ne $null) -and ($ie.ReadyState -ne  $null)){
while ($ie.Busy -or ($ie.ReadyState -ne 4)) { start-sleep -milliseconds 100}
} else {
  start-sleep -milliseconds 100
}
#>
$input = $ie.Document.Script.prompt($message, ''); $ie.quit()
write-output $input

