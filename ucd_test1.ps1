#Copyright (c) 2020 Serguei Kouzmine
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

param(
	[String] $username = 'admin',
	[String] $password = 'admin'
)
$MODULE_NAME = 'internetexplorer_application_helper.psd1'
Import-Module -Name ('{0}/{1}' -f '.', $MODULE_NAME )

$ie = new-object -com 'internetexplorer.application'
# see also 'MSXML2.DOMDocument'
$ie.visible = $true
$ucdServerIp = $env:UCD_SERVER_IP
if ($ucdServerIp -eq '' -or $ucdServerIp -eq $null){
  $ucdServerIp = '192.168.0.64'
}

$target_url = ("https://{0}:8443/" -f $ucdServerIp)
<#
  TODO: tweak zones
  Unexpected URL: res://ieframe.dll/invalidcert.htm?SSLError=50331648#https://192.168.0.64:8443/
#>
$ie.navigate2($target_url)
# wait for the page to loads
# ([ref]$ie) | wait_while_busy
wait_busy -ie_ref ([ref]$ie)

$debug =  $false
$document_element = $ie.document.documentElement
$document = $ie.document
$window = $document.parentWindow

$locator = 'form[action = "/tasks/LoginTasks/login" ] input[name = "username"]'
$element = $document_element.querySelector($locator, $null)
highlight -locator $locator -window_ref ([ref]$window) -document_element_ref ([ref] $document_element)
sendKeys -locator $locator -window_ref ([ref]$window) -document_element_ref ([ref] $document_element) -text $username

$locator = 'form input[name = "password"]'
highlight -locator $locator -window_ref ([ref]$window) -document_element_ref ([ref] $document_element)
sendKeys -locator $locator -window_ref ([ref]$window) -document_element_ref ([ref] $document_element) -text $password


$locator = 'form span[widgetid = "submitButton"]'
highlight -locator $locator -window_ref ([ref]$window) -document_element_ref ([ref] $document_element)
click -locator $locator -window_ref ([ref]$window)

start-sleep -milliseconds 10000

# expect the URL to become
if ( -not ($document.url -match '.*welcome.*$')) {
  write-output ('Unexpected URL: ' + $document.url )
}
$ie.quit()

