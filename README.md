### Info

This repository contains projects interacting with Internet Explorer browser from Powershell.
It started with converting the legacy VB Script snippet discussed
in the forum  http://forum.oszone.net/thread-334713.html (in Russian) to Powershell. 

It appears that even in 2016 - 2018 a lot of people is still using IE to automate web page processing.

Using IE / COM / Powershell appears useful for automating tasks which do not warrant setting up a full blown 
Selenium / Java application stack, or for restricted environments where installing Java is not an option.

The code pattern thie project is about to explore is
```powershell
$ie = new-object -com 'internetexplorer.application'
$ie.visible = $true
$url = '...'
$ie.navigate2($url)
while (($s.Busy -eq $true ) -or ($s.ReadyState -ne 4)) {
  start-sleep 100
}
$documentElement = $ie.document.documentElement

# can navigate the page DOM using $documentElement e.g.
$m = $documentElement.getElementsByClassName('header')
$e = $documentElement.querySelector('div > div > input')

# some API require refernce to $document:
$document = $ie.Document
$e = $document.getElementById('password_field')

# Running script appears to require a reference to the window object
$window = $document.parentWindow
$window.execScript('alert("this will pop up the Alert on IE")', 'javascript')
```
Note: since COM API offered by `InteretExplorer.Application` appear a fair bridge to the [low-level JavaScript Web API](https://developer.mozilla.org/en-US/docs/Web/API)
one is likely end up with custom implementation of the familar Selenium API like
```powershell

function highlight {
param (
  [System.Management.Automation.PSReference]$window_ref,
  [String]$locator,
  [int]$delay = 100
)
  $window = $window_ref.Value
  $highlightBorderScript = (@"
var selector = '{0}';
var elements = document.querySelectorAll(selector);
elements[0].style.border='3px solid yellow';
"@ -f $locator)
  $window.execScript($highlightBorderScript, 'javascript')
  start-sleep -milliseconds $delay

  $removeBorderScript = ("var selector = '{0}';var elements = document.querySelectorAll(selector);elements[0].style.border='';"  -f $locator)
  $window.execScript($removeBorderScript, 'javascript')

}

```
callable via
```powershell
highlight -locator $locator -window_ref ([ref]$window)
```

or 
```powershell
function sendKeys {
param (
  [System.Management.Automation.PSReference]$window_ref,
  [String]$locator,
  [String]$text = 'this is the text'
)
  $window = $window_ref.Value 
  $textEnterScript = (@"
var selector = '{0}';
var elements = document.querySelectorAll(selector);
elements[0].value  = '{1}';
"@  -f $locator, $text)
  $window.execScript($textEnterScript, 'javascript')
}
```
callable via
```powershell
sendKeys -locator 'form[class = "form-inline"]' -text 'This is the text to input' -window_ref ([ref]$window)
```
### See Also:
* https://stackoverflow.com/questions/3514945/running-a-javascript-function-in-an-instance-of-internet-explorer?utm_medium=organic&utm_source=google_rich_qa&utm_campaign=google_rich_qa
* http://www.vbaexpress.com/forum/showthread.php?9690-Solved-call-a-javascript-function
* [Powershell-With-IE tutorial](http://powershelltutorial.net/technology/Powershell-With-IE)
* [Powershell browser-based tasks](https://westerndevs.com/simple-powershell-automation-browser-based-tasks/)
* [another post indicating some refactoring possible](https://www.gngrninja.com/script-ninja/2016/9/25/powershell-getting-started-controlling-internet-explorer)
* [Accessing Javascript functions e.g. scrollTo](https://geekeefy.wordpress.com/2017/09/07/tip-scrolling-internet-explorer-with-powershell/)


### License
This project is licensed under the terms of the MIT license.

### Author
[Serguei Kouzmine](kouzmine_serguei@yahoo.com)
