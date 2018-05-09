### Info

This repository contains projects interacting with Internet Explorer browser from Powershell.
It strated with convering the VB Script snipped discussed
in the forum  http://forum.oszone.net/thread-334713.html (in Russian).


The code pattern thie projet is about to explore is
```powershell
$ie = new-object -com 'internetexplorer.application'
$ie.visible = $true
$url = '...'
$ie.navigate2($url)
while (($s.Busy -eq $true ) -or ($s.ReadyState -ne 4)) {
  start-sleep 100
}
$documentElement = $ie.document.documentElement;
# can navigate the page DOM using $documentElement e.g.
$m = $documentElement.getElementsByClassName('header')
$document = $ie.Document
$document.parentWindow.execScript("alert('Arbitrary javascript code')", "javascript")
# will pop up the Alert on IE
```

### See Also:
* https://stackoverflow.com/questions/3514945/running-a-javascript-function-in-an-instance-of-internet-explorer?utm_medium=organic&utm_source=google_rich_qa&utm_campaign=google_rich_qa
* http://www.vbaexpress.com/forum/showthread.php?9690-Solved-call-a-javascript-function

### License
This project is licensed under the terms of the MIT license.

### Author
[Serguei Kouzmine](kouzmine_serguei@yahoo.com)
