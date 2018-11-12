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

$datafile = 'data.html'
$uri = ('file:///{0}' -f ((resolve-path $datafile).path -replace '\\', '/'))
# based on http://www.cyberforum.ru/powershell/thread2343207.html
[Microsoft.PowerShell.Commands.WebResponseObject]$obj = (Invoke-WebRequest -Uri $uri)
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/invoke-webrequest?view=powershell-6 

# https://stackoverflow.com/questions/24977233/parse-local-html-file
# Can't use Poweshell invoke-webrequest cmdlet to *parse* HTML from local files, will only have
#   StatusCode
#   StatusDescription
#   Content
#   RawContent
#   Headers
#   RawContentLength
# methods available

$html = new-object -ComObject 'HTMLFile'
$html.IHTMLDocument2_write($obj.RawContent)

# NOTE: can also load like just a file
# $raw_data = get-content -path (resolve-path $datafile)  
# $doc = $html.IHTMLDocument2_write($raw_data)

$document = $html.documentElement

$tables = $document.getElementsByTagName('table')
write-output ('{0} tables' -f $tables.length )
$table = $document.getElementsByTagName('table').item(0)

try {
  # will crash even with try / catch
  # $table.querySelectorAll('tr td')
  $element = $table.querySelector('tr td:first-of-type') # Exception : Invalid argument.
  $element.innerHTML
} catch [Exception]{ 
  write-output ( 'Exception : ' + $_.Exception.Message)
}

<#
$rows = $table.getElementsByTagName('tr')
write-output ('{0} rows' -f $rows.length )
# $row = $rows.item(0)
#  $row | get-member
@(0..$rows.length) | foreach-object { 
  $index = $_
  # write-output $index
  $row = $table.getElementsByTagName('tr').item($index)
  # write-output $row.innerHTML
  try {
    $cols = $row.getElementsByTagName('TD')
    write-output ('{0} columns' -f $cols.length )
    if ($cols.length -eq 9) {
      write-output ('{0} column: {1}' -f '0', $cols.item(0).innerHTML )
    }
  } catch [Exception]{ 
    write-output ( 'Exception : ' + $_.Exception.Message)
  }
}
#>
# $table  | get
# $html.IHTMLDocument2_write($table[0].outerHTML)
# https://msdn.microsoft.com/en-us/windows/cc304115(v=vs.71)
write-output ( 'while rows: {0}' -f $table.querySelectorall('tr[bgcolor]').length)
$length = $table.querySelectorall('tr[bgcolor="#f0f0f0"]').length

@(0..$length) | foreach-object { 
  $index = $_ 
  $row = $table.querySelectorall('tr[bgcolor="#f0f0f0"]').item($index)
  try {
  
    $email = $row.getElementsByTagName('td').item(6)
    <#
    if ($email.querySelectorAll('a.external').length -ne 0 ){
      # Exception : Method invocation failed because [System.__ComObject] does not contain a method named 'querySelectorall'.
      write-output $email.innerHTML
      write-output $row.getElementsByTagName('TD')[0].outerHTML
    }
    #>
    if ($email.innerText -match '[a-z0-9_]+@mail.ru' ){
      # Exception : Method invocation failed because [System.__ComObject] does not contain a method named 'querySelectorall'.
      write-output $email.innerText
      write-output $row.getElementsByTagName('TD').item(0).innerText
    }
  } catch [Exception]{ 
    write-output ( 'Exception : ' + $_.Exception.Message)
  }
}