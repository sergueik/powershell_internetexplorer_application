#Copyright (c) 2021 Serguei Kouzmine
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

# this illustrates that the "or" syntax of css selector is not supported very well by IE
# https://www.w3schools.com/cssref/css_selectors.asp
  # in Powershell Version 4.0 [mshtml.HTMLDivElementClass] does not contain a method named 'getElementsByClassName'.
  # in Powershell Version 4.0 [mshtml.HTMLDivElementClass] does not containa method named 'querySelectorall'.
  # not practical to continue coding against 4.0

param (
  [String]$hub_ip = '',
  [String]$hub_port = '4444'
)
  if ($hub_ip -ne '') {
    $status = test-netconnection -computername $hub_ip -port $hub_port
    if (( -not $status.PingSucceeded )-or (-not $status.TcpTestSucceeded) ) {
      write-output ('http://{0}:4444/grid/console is not responding' -f $hub_ip, $hub_port)
      exit 0
    }
  }
  if ($hub_ip -eq '') {
    $datafile = 'grid_console.html'
    $uri = ('file:///{0}' -f ((resolve-path $datafile).path -replace '\\', '/'))
  } else {
    $uri = ('http://{0}:4444/grid/console' -f $hub_ip, $hub_port)
  }

  [Microsoft.PowerShell.Commands.WebResponseObject]$response_obj = (Invoke-WebRequest -Uri $uri)
  $html = new-object -ComObject 'HTMLFile'
  $html.IHTMLDocument2_write($response_obj.rawContent)

  $css_selector = '#rightColumn,#leftColumn p.proxyid'

  write-output 'attempt 1'
  $content = $html.getElementById('main_content')
  $length = $content.querySelectorall($css_selector).length
  write-output('Found {0} elements' -f $length)

  @(0..($length-1)) | foreach-object {
    $index = $_
    $element = $content.querySelectorall($css_selector).item($index)
    write-output ('processing item # {0}: "{1}"' -f $index, $element.InnerText)
    remove-variable element -ErrorAction SilentlyContinue
  }
  remove-variable content -ErrorAction SilentlyContinue
  remove-variable length -ErrorAction SilentlyContinue


  write-output 'attempt 2'
  # $ids = @('left-column','right-column')
  $ids = @('leftColumn', 'rightColumn')
  $content = $html.getElementById('main_content')

  $ids| foreach-object {

    $column_id = $_
    $css_selector = ('#{0} p[class="proxyid"]' -f $column_id)
    write-output ('css_selector: "{0}"' -f $css_selector ) 
    $content = $html.getElementById('main_content')
    $length = $content.querySelectorall($css_selector).length
    write-output('Found {0} elements' -f $length)

    @(0..($length-1)) | foreach-object {
      $index = $_
      # loading sets into variables appears to be less reliable that re-issuing a COM object call
      $element = $content.querySelectorall($css_selector).item($index)
      write-output ('processing item # {0}: "{1}"' -f $index, $element.InnerHTML)
    }
    remove-variable element -ErrorAction SilentlyContinue
    remove-variable content -ErrorAction SilentlyContinue
  }
  remove-variable length -ErrorAction SilentlyContinue
  write-output 'attempt 3'

  $ids = @('leftColumn', 'rightColumn')

  $ids| foreach-object {
    $column_id = $_
    write-output ('processing column: "{0}"' -f $column_id)
    $column = $html.getElementById($column_id)
    $elements = $column.getElementsByClassName('proxyid')
    if (($elements -ne $null) -and ($elements.length -ne 0 )) {
      $length = $elements.length
      0..($length - 1) | foreach-object {
        $index = $_
        $element = $elements.item($index)
        write-output ('processing item # {0}: "{1}"' -f $index, $element.InnerText)
      }
    }
  }
  # the following two snippets do not work and will hang the IE and calling Powershell
  # with "a problem causes the program to stop working correctly" dialog. It is  helpful to start extra powershell.exe, stacked
  # This Exception cannot be caught. The code is commented out
  <#
  try {
    $length = $html.querySelectorall('#leftColumn,#rightColumn p[class = "proxyid"]').length
    write-output('Found {0} elements' -f $length)
  } catch [Exception] {
    write-output ( 'Exception : ' + $_.Exception.Message)
  }
  try {
    $column = $html.getElementById('leftColumn')
    $element = $context.querySelectorall('p[class= "proxyid"]')|select-object -first 1
    $element.getType().FullName
  } catch [Exception] {
    write-output ( 'Exception : ' + $_.Exception.Message)
  }
  #>
