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


# This script extracts the download link the edge driver by parsing the page HTML
# without using Selenium itself
# the current version just prints the data in the format:
# href : https://download.microsoft.com/download/C/0/7/C07EBF21-5305-4EC8-83B1-A6
#        FCC8F93F45/MicrosoftWebDriver.exe
# text : Release 10586
#
# href : https://download.microsoft.com/download/8/D/0/8D0D08CF-790D-4586-B726-C6
#        469A9ED49C/MicrosoftWebDriver.exe
# text : Release 10240

$MODULE_NAME = 'internetexplorer_application_helper.psd1'
Import-Module -Name ('{0}/{1}' -f '.', $MODULE_NAME )

$ie = new-object -com 'internetexplorer.application'
# see also 'MSXML2.DOMDocument'
$ie.visible = $true
$target_url = 'https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/'
$ie.navigate2($target_url)
wait_busy -ie_ref ([ref]$ie)
start-sleep -Milliseconds 1000
$debug =  $false
$document = $ie.document
$document_element = $document.documentElement
$window = $document.parentWindow
$result_tag = 'my_data'
# $result_tag = 'PSResult'
$element_locator = 'section#downloads ul.driver-downloads li.driver-download > a'
collect_data_hash -window_ref ([ref]$window) -element_locator $element_locator -value_attribute 'href' -result_tag $result_tag -debug

$result_raw = $document.body.getAttribute($result_tag)
write-output ('Result (raw): ( in "' + $result_tag + '") ' + $result_raw)
# NOTE:
try {
  $result_obj = $result_raw | convertfrom-json
  format-list -InputObject $result_obj
} catch [Exception] {
  write-output ('Exception : ' + $_.Exception.Message)
}

collect_data_array -window_ref ([ref]$window) -element_locator $element_locator -element_attribute 'href' -result_tag $result_tag -debug

$result_raw = $document.body.getAttribute($result_tag)
write-output ('Result (raw): ( in "' + $result_tag + '") ' + $result_raw)
$result_array = ($result_raw -replace '^\[', '' -replace '\]$' ) -split ','
$result_array | format-list
# quit and dispose IE
$ie.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ie) | out-null
Remove-Variable ie