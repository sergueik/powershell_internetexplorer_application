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
$target_url = 'https://select2.github.io/examples.html'
$ie.navigate2($target_url)
# wait for the page to loads
# ([ref]$ie) | wait_while_busy
wait_busy -ie_ref ([ref]$ie)

$debug =  $false
$document_element = $ie.document.documentElement
$document = $ie.document
$window = $document.parentWindow

$locator =  'select.js-states'
$e1 = $document_element.querySelector($locator )
highlight -locator $locator -window_ref ([ref]$window)
$selectOption = 'FL'
$selectOptionScript = ( @"

var selector = '{0}';
var o_val = '{1}';
var s2_obj = `$(selector).select2();
option = s2_obj.val(o_val);
// debugging
if (debug) {
  alert ('Value or selected option is ' + option.val());
}
// does now work with IE
option.trigger('select');
"@ -f $locator , $selectOption  )
write-output  ('Executing:  {0}' -f $selectOptionScript )
$document.parentWindow.execScript($selectOptionScript, 'javascript')
$querySelectedValueScript = ( @"
var selector = '{0}';
var s2_obj = `$(selector).select2();
return s2_obj.val();
"@ -f $locator )
start-sleep -millisecond 10000
[String]$result = [String] $document.parentWindow.execScript(
$querySelectedValueScript,'javascript')
write-output ("Selected via Javascript: " -f $result)

start-sleep -milliseconds 1000

$ie.quit()

