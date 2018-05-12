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

$MODULE_NAME = 'internetexplorer_application_helper.psd1'
Import-Module -Name ('{0}/{1}' -f '.', $MODULE_NAME )

$ie = new-object -com 'internetexplorer.application'
# see also 'MSXML2.DOMDocument'
$ie.visible = $true
$target_url = 'http://suvian.in/selenium/1.1link.html'
$ie.navigate2($target_url)
# wait for the page to loads
# ([ref]$ie) | wait_while_busy
wait_busy -ie_ref ([ref]$ie) 

$debug =  $false
$document_element = $ie.document.documentElement
$document = $ie.document
$window = $document.parentWindow

$m1 = $document_element.getElementsByClassName('intro-message')
$e1 = $m1[0]
$m2 = $e1.getElementsByTagName('h3')
$e2 = $m2 | where-object { $_.innerText -match '.*Click Here.*' }
$e2.FireEvent('onclick', $null) ;
# has no effect
$e1.click()
# has no effect

# $document_element.FireEvent('onclick', $e2)
# No such interface supported

$locator = 'body > div.intro-header > div > div > div > div > h3:nth-child(2) > a'
highlight -locator $locator -window_ref ([ref]$window) -document_element_ref ([ref] $document_element)

# failing test
$locator = 'body > div.intro-header > div > div > div > div > h3:nth-child(2) > b'
highlight -locator $locator -window_ref ([ref]$window) -document_element_ref ([ref] $document_element)

$locator = 'body > div.intro-header > div > div > div > div > h3:nth-child(2) > a'
click -locator $locator -window_ref ([ref]$window)

start-sleep -milliseconds 10000

# expect the URL to become
if ( -not ($document.url -match '.*/1.1link_validate.html$')) {
  write-output ('Unexpected URL: ' + $document.url )
}
$ie.quit()

