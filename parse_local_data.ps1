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


# was replied to the question http://www.cyberforum.ru/powershell/thread2343207.html

# load dirty tables file provided by customer
$datafile = 'data.html'
$uri = ('file:///{0}' -f ((resolve-path $datafile).path -replace '\\', '/'))
[Microsoft.PowerShell.Commands.WebResponseObject]$obj = (Invoke-WebRequest -Uri $uri)
# https://stackoverflow.com/questions/24977233/parse-local-html-file
# one can't use Poweshell `invoke-webrequest` cmdlet to DOM parsing of local HTML files
# only methods available: StatusCode, StatusDescription, Content, RawContent, Headers, RawContentLength

$html = new-object -ComObject 'HTMLFile'
$html.IHTMLDocument2_write($obj.rawContent)
$table = $html.documentElement.getElementsByTagName('table').item(0)

write-debug 'select rows by background color'
$length = $table.querySelectorall('tr[bgcolor="#f0f0f0"]').length
write-debug ( 'white rows: {0}' -f  $length)

@(0..$length) | foreach-object {
  $index = $_
  # loading sets into variables appears to be less reliable that re-issuing a COM object call
  $row = $table.querySelectorall('tr[bgcolor="#f0f0f0"]').item($index)
  try {
    write-debug 'select rows which have email'
    $email = $row.getElementsByTagName('td').item(6)
    if ($email.innerText -match '[a-z0-9_]+@mail.ru' ){
      write-output $email.innerText
      write-output $row.getElementsByTagName('td').item(0).innerText
    }
  } catch [Exception]{
    write-debug ('Exception: ' + $_.Exception.Message)
  }
}