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
$target_url = 'https://datatables.net/examples/api/highlight.html'
$ie.navigate2($target_url)
# wait for the page to loads
# ([ref]$ie) | wait_while_busy
wait_busy -ie_ref ([ref]$ie) 

$debug =  $false
$document_element = $ie.document.documentElement
$document = $ie.document
$window = $document.parentWindow
# $document | get-member

$table = $document.getElementById('example')
write-output $table.innerHTML

$cell_text = 'Software Engineer'

<#

$script_template = @"

var row_selector = '{0}';
var cell_selector = '{1}';
var cell_text = '{2}';
var rows = document.querySelectorAll(row_selector);
for (var row_cnt = 0 ;row_cnt != rows.length;row_cnt ++ ){
  var row = rows[row_cnt];
  var cells = row.querySelectorAll(cell_selector);
  for (var cell_cnt = 0 ; cell_cnt != cells.length;cell_cnt ++ ){
    var cell = cells[cell_cnt];
      if (cell.innerHTML.indexOf(cell_text) == 0) {
        alert(cell.innerHTML);
    }
  }  
}
"@  -replace "`n", ' '
write-output $script_template 
$script = ( $script_template  -f 'tbody/tr' , 'td', 'Software Engineer' )
write-output $script
Error formatting a string: Input string was not in a correct format..
#>
$row_selector = '#example tbody > tr' 
$cell_selector = 'td'

$script = @"

var row_selector = '#example tbody > tr';
var cell_selector = 'td';
var cell_text = 'Software Engineer';
var rows = document.querySelectorAll(row_selector);
for (var row_cnt = 0 ;row_cnt != rows.length;row_cnt ++ ){
  var row = rows[row_cnt];
  var cells = row.querySelectorAll(cell_selector);
  for (var cell_cnt = 0 ; cell_cnt != cells.length;cell_cnt ++ ){
    var cell = cells[cell_cnt];
      if (cell.innerHTML.indexOf(cell_text) == 0) {
        alert(cell.innerHTML);
    }
  }  
}
"@  -replace "`n", ' '
write-output $script

$window.execScript($script, 'javascript')

$ie.quit()

