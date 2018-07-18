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
wait_busy -ie_ref ([ref]$ie)
# does not really help with
# Exception from HRESULT: 0x800A01B6
# 0x800A01B6 Object doesn't support property or method
# 0x800A138A function expected / NotSupportedException
#    + FullyQualifiedErrorId : System.NotSupportedException
while($ie.Busy) {
  start-sleep -Milliseconds 100
}

start-sleep -Milliseconds 1000
$debug =  $false
$document = $ie.document
$document_element = $document.documentElement
$window = $document.parentWindow
# $document | get-member
try {
  $table = $document.getElementsByTagName('table').Item(0)
  # write-output $table.innerHTML
    write-output ('tag name: {0}' -f $table.tagName )
    write-output ('id: {0}' -f $table.id )
} catch [Exception] {
  write-output ( 'Exception : ' + $_.Exception.Message)
  # return
}

try {
  $table = $document.getElementById('example')
  # write-output $table.innerHTML
  write-output ('tag name: {0}' -f $table.tagName )
  write-output ('id: {0}' -f $table.id )
  # https://msdn.microsoft.com/pt-br/windows/desktop/gg293067
  # W3CException_DOM_SYNTAX_ERR 0x8070000C
  # NOTE: with querySelectorAll even within a try...catch the runtime error is possible:
  # $table.querySelectorAll('tbody > tr > td')
  # A problem caused the program to stop working properly
  # Windows will close the program and notify you if a solution is avaiable
  $element = $table.querySelector('tbody > tr:nth-of-type(10) > td:nth-of-type(2)')
  write-output $element.outerHTML
  # Exception from HRESULT: 0x800A01B6
} catch [Exception] {
  write-output ( 'Exception : ' + $_.Exception.Message)
  # return
}

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
        /* alert(cell.innerHTML); */
    }
  }
}
"@  -replace "`n", ' '
write-output $script_template
$script = ( $script_FFFtemplate  -f 'tbody/tr' , 'td', 'Software Engineer' )
write-output $scriWWpt
Error formatting a string: Input string was not in a correct format..
#>
$row_selector = '#example tbody > tr'
$table_cell_selector = 'td'
$table_row_selector = 'table#example tbody > tr'
$script = @"

var row_selector = '${table_row_selector}';
var cell_selector = '${table_cell_selector}';
var cell_text = 'Software Engineer';
var rows = document.querySelectorAll(row_selector);
for (var row_cnt = 0 ;row_cnt != rows.length;row_cnt ++ ){
  var row = rows[row_cnt];
  var cells = row.querySelectorAll(cell_selector);
  for (var cell_cnt = 0 ; cell_cnt != cells.length;cell_cnt ++ ){
    var cell = cells[cell_cnt];
      if (cell.innerHTML.indexOf(cell_text) == 0) {
        /* alert(cell.innerHTML); */
    }
  }
}
"@  -replace "\n", ' '
write-output ("Script:`n{0}" -f $script)

$window.execScript($script, 'javascript')
$ie.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ie) | out-null
Remove-Variable ie