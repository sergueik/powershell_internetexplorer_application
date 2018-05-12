
<#
.SYNOPSIS
    Highlights page element
.DESCRIPTION
    Highlights page element by executing Javascript through InternetExplorer.Application
    
.EXAMPLE
    highlight -window_ref ([ref]$window) -locator $locator -delay 1500 -color 'green'
    highlight -window_ref ([ref]$window) -document_element_ref ([ref]$document+element) -locator $locator -delay 1500 -color 'green'
.LINK
    
.NOTES
    VERSION HISTORY
    2018/05/12 Initial Version
#>

function highlight {
  param (
    [System.Management.Automation.PSReference]$window_ref,
    [System.Management.Automation.PSReference]$document_element_ref,
    [String]$locator,
    [string]$color = 'yellow',
    [int]$delay = 100
  )
  $window = $window_ref.Value
  if ($document_element_ref -ne $null) {
    $document_element = $document_element_ref.Value
    $element = $null
    try {
      $element = $document_element.querySelector($locator)
      $element.innerHTML | out-null
    } catch [Exception] {
      write-Debug ( 'Exception : ' + $_.Exception.Message)
      return
    }
    if ($element -eq $null) {
      write-Debug ('unable to find {0}' -f $locator )
      return
    }
  }
  $highlightBorderScript = ( @"
var selector = '{0}';
var elements = document.querySelectorAll(selector);
elements[0].style.border = '3px solid {1}';
"@  -f $locator, $color )
  try {
    $window.execScript($highlightBorderScript, 'javascript')
  } catch [Exception] {
    write-Debug ( 'Exception : ' + $_.Exception.Message)
    return
  }
  start-sleep -milliseconds $delay

  $removeBorderScript = (@"
var selector = '{0}';
var elements = document.querySelectorAll(selector);
elements[0].style.border='';
"@  -f $locator)
  try {
    $window.execScript($removeBorderScript, 'javascript')
  } catch [Exception] {
    write-Debug ( 'Exception : ' + $_.Exception.Message)
    return
  }
}

<#
.SYNOPSIS
    Clicks page element
.DESCRIPTION
    Sends clickk to page element located by Javascript by executing Javascript through InternetExplorer.Application
    
.EXAMPLE
    click -window_ref ([ref]$window) -locator $locator
.LINK
    
.NOTES
    VERSION HISTORY
    2018/05/12 Initial Version
#>

function click {
  param (
    [System.Management.Automation.PSReference]$window_ref,
    [String]$locator
  )
  $window = $window_ref.Value
  $clickScript = (@"
var selector = '{0}';
var elements = document.querySelectorAll(selector);
elements[0].click();
"@  -f $locator)
  $window.execScript($clickScript, 'javascript')
}

<#
.SYNOPSIS
    Sends Text into the page element
.DESCRIPTION
    Sends text into page element located by Javascript by executing Javascript through InternetExplorer.Application
    
.EXAMPLE
    sendKeys -window_ref ([ref]$window) -locator $locator -text 'text'
.LINK
    
.NOTES
    VERSION HISTORY
    2018/05/12 Initial Version
#>
function sendKeys {
  param (
    [System.Management.Automation.PSReference]$window_ref,
    [System.Management.Automation.PSReference]$document_element_ref,
    [String]$locator,
    [String]$text = 'this is the text'
  )
  $window = $window_ref.Value
  # Note: may try the value property
  if ($document_element_ref -ne $null) {
    $document_element = $document_element_ref.Value
    $element = $null
    try {
      $element = $document_element.querySelector($locator)
      $element.innerHTML | out-null
      $element.value = $text
    } catch [Exception] {
      write-Debug ( 'Exception : ' + $_.Exception.Message)
      return
    }
    if ($element -eq $null) {
      write-Debug ('unable to find {0}' -f $locator )
      return
    }
  }
  $textEnterScript = (@"
var selector = '{0}';
var elements = document.querySelectorAll(selector);
elements[0].value  = '{1}';
"@  -f $locator, $text)
  $window.execScript($textEnterScript, 'javascript')
}


<#
.SYNOPSIS
    Sends Text into the page element
.DESCRIPTION
    Sends text into page element located by Javascript by executing Javascript through InternetExplorer.Application
    
.EXAMPLE
    ([ref]$ie) | wait_busy # `valuefrompipeline` does not currently work
    wait_busy -ie_ref ([ref]$ie) 
.LINK
    
.NOTES
    VERSION HISTORY
    2018/05/12 Initial Version
#>
function wait_busy {
  param(
    [System.Management.Automation.PSReference]$ie_ref
  )
  $ie = $ie_ref.Value

  while ($ie.Busy -or ($ie.ReadyState -ne 4)) {
    # 4 a.k.a. READYSTATE_COMPLETE
    start-sleep -milliseconds 100
  }

}

<#
.SYNOPSIS
    Scrolls the browser window (e.g. into the page element offset)
.DESCRIPTION
    Scrolls the browser window (e.g. into the page element offset)
    by executing Javascript through InternetExplorer.Application wrapper
    
.EXAMPLE
    $document = $ie.document
    $window = $document.parentWindow
    scroll_to -winow_ref ([ref]$window) -vertical_scroll 500 
.LINK
    
.NOTES
    VERSION HISTORY
    2018/05/12 Initial Version
#>
function scroll_to { 
  param (
    [System.Management.Automation.PSReference]$window_ref,
    [int]$vertical_scroll = 100
  )

  $window = $window_ref.Value
  $window.scrollTo(0,$vertical_scroll)
}