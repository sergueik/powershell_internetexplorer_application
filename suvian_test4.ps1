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
$target_url = 'http://suvian.in/selenium/1.5married_radio.html'
$ie.navigate2($target_url)
wait_busy -ie_ref ([ref]$ie) 

$debug =  $false
$document_element = $ie.document.documentElement
$document = $ie.document
$window = $document.parentWindow

$checkbox_value = 0
# $locator = ("//div[@class='intro-header']/div[@class='container']/div[@class='row']/div[@class='col-lg-12']/div[@class='intro-message']/form/input[@name='married' and @value='{0}']" -f $checkbox_value )
# TOO complex, and it was xpath
# hanging IE
$locator = ("form input[name='married'][value='{0}']" -f $checkbox_value )
# https://developer.mozilla.org/en-US/docs/Web/API/Element/querySelector
# does not work:  actully is hanging IE if sent to the COM object
#  $element = $document_element.querySelector($locator, $null)
$locator = "form input[name='married']"
# https://developer.mozilla.org/en-US/docs/Web/API/Element/querySelectorAll
$e2 = $null
$m1 = $document_element.getElementsByClassName('intro-message')
<#
hangong IE , try ...catch deos ot help
try {
  $e2 = $document_element.querySelectorAll($locator)
  $e2
  } catch [Exception] {
  write-output ( 'Exception : ' + $_.Exception.Message)
}
#>
$e1 = $m1[0]
try {
  $e2 = $e1.querySelector($locator)
  $e2
} catch [Exception] {
  write-output ( 'Exception : ' + $_.Exception.Message)
}

# TODO: bug ?
# $document_element.querySelectorAll("form input[name='married']") returns only one element, not two
write-output ('Element 2 {0}' -f $e2.innerHTML)
# $element.click()

highlight -locator $locator -window_ref ([ref]$window) -document_element_ref ([ref] $document_element)
sendEnterKey -locator $locator -window_ref ([ref]$window)  -key 13
# does not work
# $ie.quit()

