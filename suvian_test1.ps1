$ie = new-object -com 'internetexplorer.application'
# see also 'MSXML2.DOMDocument'
$ie.visible = $true
$target_url = 'http://suvian.in/selenium/1.1link.html'
$ie.navigate2($target_url)
# wait for the page to loads
while (($ie.Busy -eq $true ) -or ($ie.ReadyState -ne 4)) { # 4 a.k.a. READYSTATE_COMPLETE
  start-sleep -milliseconds 100
}
$debug =  $false
$documentElement = $ie.document.documentElement
$document = $ie.document
$window = $document.parentWindow
$locator = 'body > div.intro-header > div > div > div > div > h3:nth-child(2) > a'

$m1 = $documentElement.getElementsByClassName('intro-message')
$e1 = $m1[0]
$m2 = $e1.getElementsByTagName('h3')
$e2 = $m2 | where-object { $_.innerText -match '.*Click Here.*' }
$e2.FireEvent('onclick', $null)
# True
# no effect
$e1.click()

# no effect
# $documentElement.FireEvent('onclick', $e2)
# No such interface supported

$highlightBorderScript = ("var selector = '{0}';var elements = document.querySelectorAll(selector);elements[0].style.border='3px solid yellow';" -f $locator)
$window.execScript($highlightBorderScript, 'javascript')
start-sleep -milliseconds 500

$removeBorderScript = ("var selector = '{0}';var elements = document.querySelectorAll(selector);elements[0].style.border='';"  -f $locator)
$window.execScript($removeBorderScript, 'javascript')

$clickScript = ("var selector = '{0}';var elements = document.querySelectorAll(selector);elements[0].click();"  -f $locator)
$window.execScript($clickScript, 'javascript')

start-sleep -milliseconds 10000

# expect the URL to become
if ( -not (write-output $document.url -match '.*/1.1link_validate.html$')) {
  write-output ('Unexpected URL: ' + $document.url )
}
$ie.quit()

