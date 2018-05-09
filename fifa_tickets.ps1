# based on: http://forum.oszone.net/thread-334713.html
# objective:
# associate the titles of the matches with the presence of the column
# other than the
# "class="categoryBox zeroAvailability" 
# which indicates e.g. lowAvailability

$ie = new-object -com 'internetexplorer.application'
$ie.visible = $true
$target_url = 'https://tickets.fifa.com/Services/ADService.html?lang=ru'
$ie.navigate2($target_url)
# wait for the page to loads
while (($ie.Busy -eq $true ) -or ($ie.ReadyState -ne 4)) { # 4 a.k.a. READYSTATE_COMPLETE
  start-sleep 100
}
$debug =  $false
$documentElement = $ie.document.documentElement

# https://developer.mozilla.org/en-US/docs/Web/API/Document/getElementsByClassName
# $m1 is likely HTMLCollection
$m1 = $documentElement.getElementsByClassName('header')
if ($debug) {
  $m1.item(1)
}

# limit the work to few first nodes
$cnt = 0
$max_items = 100

if ($m1.length -lt $max_items ){
  $max_items = $m1.length - 1
}

1..$m1.length | foreach-object {

 if ($cnt -ge $max_items ) {
   if ($ie -ne  $null) {
     $ie.quit()
     $ie = $null
     return
   }
 } else {
  $cnt = $_
  if ($debug ){
    write-output $cnt
  }
  $e1 = $m1.item($cnt)

  # write-output ( 'Node text: ' + $e1.textContent )
  if ($debug ){
    write-output ( 'Node name: ' + $e1.nodeName ) # DIV
    write-output ( 'Node text: ' + $e1.textContent )
  }
  $e2 = $e1.parentNode
  if ($debug ){
    $e2.innerHTML
    # <div class="header" ng-bind="product.productName">???? 02 -?????? : ??????? - ????????????</div>
  }

  $e3 = $e2.parentNode

  $e4 = $e3.NextSibling.NextSibling

  if ($debug ){
    $e4.textContent
    #     CAT 1
    #     CAT 2
    #     CAT 3
    #     CAT 4
  }

  $m2 = $e4.getElementsByClassName('categoryBox')
  if ($debug ){
    $m2[1].innerHTML
    # CAT 2
    $m2[1].outerHTML
    # <div class="categoryBox zeroAvailability" ng-bind="cat.categoryName" ng-class="cat.availabilityColor">CAT 2</div>
  }
  if ($debug) { 
    write-output $m2[0].outerHTML
  }
  if ($debug) {   
    $m2 | foreach-object {
      write-output ("class: " + $_.className)

      # write-output ("class: " + $_.getAttribute('class'))
      write-output ("ng-class: " + $_.getAttribute('ng-class'))
      write-output ("context: " + $_.textContent)
    }
  }
  <#  $_.className -match '.*(?:zeroAvailability|lowAvailability).*'  #>
  $m2 | where-object { $_.className -match '.*(?:lowAvailability).*' } |
    foreach-object {
      write-output ('Node2 HTML: ' + $_.outerHTML)
      write-output ('Node 2 context: ' + $e1.textContent)
      write-output ('Node1 text: ' + $e1.textContent )
    }
  }
}

<# original script 

#>