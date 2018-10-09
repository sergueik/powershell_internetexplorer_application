### Info

This repository contains projects interacting with Internet Explorer browser from Powershell.
It started with converting the legacy VB Script snippet discussed
in the forum  http://forum.oszone.net/thread-334713.html (in Russian) to Powershell.

It appears that even in 2016 - 2018 a lot of people is still using IE to automate web page processing.

![controlling IE with Powershell](https://github.com/sergueik/powershell_internetexplorer_application/blob/master/screenshots/capture.png)

Using IE / COM / Powershell appears useful for automating tasks which do not warrant setting up a full blown
Selenium / Java application stack, or for restricted environments where installing Java is not an option - Powershell / Internet Explorer has no install dependencies.

The code pattern thie project is about to explore looks like the following:
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

One can also make Powershell invoke some Javascipt code in the browser and receive e.g. the set of specific attributes of a set of elements
using one of the helper functions:

```powershell

$ie = new-object -com 'internetexplorer.application'
$target_url = 'https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/'
$ie.navigate2($target_url)
$document = $ie.document
$document_element = $document.documentElement
$window = $document.parentWindow
$result_tag = 'PSResult'
$element_locator = 'section#downloads ul.driver-downloads li.driver-download > a'

$result_array =  collect_data_array -window_ref ([ref]$window) -document_ref ([ref]$document) `
                -element_locator $element_locator -element_attribute 'href' -result_tag $result_tag -debug

$result_array | format-list

```
would produce array of results (href attributes of the links in the table in this example):
```powershell
"https://download.microsoft.com/download/F/8/A/F8AF50AB-3C3A-4BC4-8773-DC27B32988DD/MicrosoftWebDriver.exe"
"https://download.microsoft.com/download/D/4/1/D417998A-58EE-4EFE-A7CC-39EF9E020768/MicrosoftWebDriver.exe"
...
``` 
while
```powershell
$result_obj =  collect_data_hash -window_ref ([ref]$window) -document_ref ([ref]$document) `
             -element_locator $element_locator -value_attribute 'href' -result_tag $result_tag -debug

format-list -InputObject $result_obj

```
would produce a rowset of hashes:
```powershell
@{key=Release 17134; value=https://download.microsoft.com/download/F/8/A/F8AF50AB-3C3A-4BC4-8773-DC27B32988DD/MicrosoftWebDriver.exe}
@{key=Release 16299; value=https://download.microsoft.com/download/D/4/1/D417998A-58EE-4EFE-A7CC-39EF9E020768/MicrosoftWebDriver.exe}
...
``` 
which can be coverted into a hash with of links with text of the cell serving as the key.

### See Also:

  * [InternetExplorer COM a.k.a. SHDocVw.InternetExplorer object (Windows)](https://msdn.microsoft.com/en-us/ie/aa752084(v=vs.94))
  * [InternetExplorer IHTMLDocument3_Interface interface](https://msdn.microsoft.com/en-us/ie/hh773775(v=vs.94))
  * [InternetExplorer HTML document object](https://msdn.microsoft.com/en-us/ie/ms535862(v=vs.94))
  * [Web API DOM (MDN)](https://developer.mozilla.org/en-US/docs/Web/API/Document_Object_Model)
  * https://stackoverflow.com/questions/3514945/running-a-javascript-function-in-an-instance-of-internet-explorer?utm_medium=organic&utm_source=google_rich_qa&utm_campaign=google_rich_qa
  * http://www.vbaexpress.com/forum/showthread.php?9690-Solved-call-a-javascript-function
  * [Powershell-With-IE tutorial](http://powershelltutorial.net/technology/Powershell-With-IE)
  * [Powershell browser-based tasks](https://westerndevs.com/simple-powershell-automation-browser-based-tasks/)
  * [another post indicating some refactoring possible](https://www.gngrninja.com/script-ninja/2016/9/25/powershell-getting-started-controlling-internet-explorer)
  * [Accessing Javascript functions e.g. scrollTo](https://geekeefy.wordpress.com/2017/09/07/tip-scrolling-internet-explorer-with-powershell/)
  * [using element navigation to interact with ssl alert](https://www.kiloroot.com/powershell-script-to-open-a-web-page-and-bypass-ssl-certificate-errors-2/)
  * [handling alerts via `FindWindow` and `SendMessage`](https://social.technet.microsoft.com/Forums/ie/en-US/d1a556b7-54db-4513-bafd-f16ed000f9ac/vba-to-dismiss-an-ie8-or-ie9-message-from-webpage-popup-window?forum=ieitprocurrentver)
  * [another example](https://www.gngrninja.com/script-ninja/2016/9/25/powershell-getting-started-controlling-internet-explorer)
  * [little known Javascript methods (in Russian)](https://jsonplaceholder.typicode.com/comments?postId=200)

### License
This project is licensed under the terms of the MIT license.

### Author
[Serguei Kouzmine](kouzmine_serguei@yahoo.com)
