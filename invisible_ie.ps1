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


# This script exercises ivisible IE. To prove the success,
# prints some generic element from the page to indicate the success.
# Tested in windows 8.1 (pending windows 10)

$ie = new-object -com 'internetexplorer.application'
$target_url = 'https://www.makemytrip.com/'
$ie.navigate2($target_url)
  while ($ie.Busy -or ($ie.ReadyState -ne 4)) {
    # 4 a.k.a. READYSTATE_COMPLETE
    write-debug 'waiting'
    start-sleep -milliseconds 100
  }
$ie.visible = $false
$debug =  $true
$debugpreference = 'continue'
$document = $ie.document
$document_element = $document.documentElement
$window = $document.parentWindow
if ($window -ne $null) {
  if ($document -ne $null) {

  $elements = $document_element.getElementsByTagName('script')

    write-output ('result: {0}' -f ($elements.Item(0).outerHTML))
  } else {
    write-output 'document is null'
  }
} else {
  write-output 'window is null'
}

# quit and dispose IE
$ie.Quit()
while ( [void]([System.Runtime.Interopservices.Marshal]::ReleaseComObject($ie)) ) {}
Remove-Variable ie -ErrorAction SilentlyContinue
