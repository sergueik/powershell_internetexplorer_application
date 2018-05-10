function highlight { 
param (
  [System.Management.Automation.PSReference]$window_ref,
  [String]$locator,
  [int]$delay = 100
)
  $window = $window_ref.Value 
  $highlightBorderScript = ("var selector = '{0}';var elements = document.querySelectorAll(selector);elements[0].style.border='3px solid yellow';" -f $locator)
  $window.execScript($highlightBorderScript, 'javascript')
  start-sleep -milliseconds $delay

  $removeBorderScript = ("var selector = '{0}';var elements = document.querySelectorAll(selector);elements[0].style.border='';"  -f $locator)
  $window.execScript($removeBorderScript, 'javascript')

}

function sendKeys {
param (
  [System.Management.Automation.PSReference]$window_ref,
  [String]$locator,
  [String]$text = 'this is the text'
)
  $window = $window_ref.Value 
  # this will enter text
  $textEnterScript = ("var selector = '{0}';var elements = document.querySelectorAll(selector);elements[0].value  = '{1}';"  -f $locator, $text)
  $window.execScript($textEnterScript, 'javascript')


}
$ie = new-object -com 'internetexplorer.application'
# see also 'MSXML2.DOMDocument'
$ie.visible = $true
$target_url = 'http://suvian.in/selenium/1.2text_field.html'
$ie.navigate2($target_url)
# wait for the page to loads
while (($ie.Busy -eq $true ) -or ($ie.ReadyState -ne 4)) { # 4 a.k.a. READYSTATE_COMPLETE
  start-sleep -milliseconds 100
}
$debug =  $false
$documentElement = $ie.document.documentElement
$document = $ie.document
$window = $document.parentWindow

$m1 = $documentElement.getElementsByClassName('intro-message')
$e1 = $m1[0]
$e2 = $e1.querySelector('form[class = "form-inline"]')

# NOTE: sent to $document, not to $documentElement or $e1 
$e2 = $document.getElementById('namefield')

$locator = '#namefield'

highlight -locator $locator -window_ref ([ref]$window) 

$text = 'this is the text'
sendKeys -text $text -locator $locator -window_ref ([ref]$window) 
start-sleep -milliseconds 10000
write-output ('Document URL: ' + $document.url )

$ie.quit()

