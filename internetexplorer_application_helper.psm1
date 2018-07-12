
<#
.SYNOPSIS
    Highlights page element
.DESCRIPTION
    Highlights page element by executing Javascript through InternetExplorer.Application
    
.EXAMPLE
    highlight -window_ref ([ref]$window) -locator $locator -delay 1500 -color 'green'
    highlight -window_ref ([ref]$window) -document_element_ref ([ref]$document_element) -locator $locator -delay 1500 -color 'green'
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
      $element = $document_element.querySelector($locator, $null)
      # aterntive: $document_element.querySelectorAll($locator)
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
    Sends Enter Key into the page element (e.g. Select2 element with confirmation behavior)
.DESCRIPTION
    Sends text into page element located by Javascript by executing Javascript through InternetExplorer.Application
    
.EXAMPLE
    sendEnterKey -ie_ref ([ref]$ie) [-key $keycode]
    # ([ref]$ie) | sendEnterKey # `valuefrompipeline` does not currently work
.LINK
    
.NOTES
    VERSION HISTORY
    2018/05/12 Initial Version
#>


function sendEnterKey{ 
  param (
    [System.Management.Automation.PSReference]$window_ref,
    [int]$keycode = 13 
  )
  $window = $window_ref.Value
  # origin: https://stackoverflow.com/questions/596481/is-it-possible-to-simulate-key-press-events-programmatically?utm_medium=organic&utm_source=google_rich_qa&utm_campaign=google_rich_qa

<#
  $sendEnterKeyScript = @'
var keyboardEvent = document.createEvent("KeyboardEvent");
var initMethod = typeof keyboardEvent.initKeyboardEvent !== "undefined" ? "initKeyboardEvent" : "initKeyEvent";
keyboardEvent[initMethod](
                   "keydown", // event type : keydown, keyup, keypress
                    true, // bubbles
                    true, // cancelable
                    window, // viewArg: should be window
                    false, // ctrlKeyArg
                    false, // altKeyArg
                    false, // shiftKeyArg
                    false, // metaKeyArg
                    40, // keyCodeArg : unsigned long the virtual key code, else 0
                    0 // charCodeArgs : unsigned long the Unicode character associated with the depressed key, else 0
);
document.dispatchEvent(keyboardEvent);
'@ -replace '\/\/.*$', '' -replace "`r", ' ' -replace ' +', ' '
  write-debug $sendEnterKeyScript
#>
  <#
$sendEnterKeyScript = @'
  
var eventType = "keydown";
var bubbles = true; 
var cancelable = true; 
var viewArg = window; 
var ctrlKeyArg = false; 
var altKeyArg = false; 
var shiftKeyArg = false; 
var metaKeyArg = false; 
var keyCodeArg = 40; 
var charCodeArgs = 0; 

var keyboardEvent = document.createEvent("KeyboardEvent");
var initMethod = typeof keyboardEvent.initKeyboardEvent !== "undefined" ? "initKeyboardEvent" : "initKeyEvent";
keyboardEvent[initMethod](eventType, bubbles, cancelable, viewArg,ctrlKeyArg, altKeyArg, shiftKeyArg, metaKeyArg, keyCodeArg charCodeArgs);
document.dispatchEvent(keyboardEvent);

'@
    #>
  try {
    $window.execScript(('var keyboardEvent = document.createEvent("KeyboardEvent"); var initMethod = typeof keyboardEvent.initKeyboardEvent !== "undefined" ? "initKeyboardEvent" : "initKeyEvent"; keyboardEvent[initMethod]( "keydown", true, true, window, false, false, false, false, {0}, 0 ); document.dispatchEvent(keyboardEvent); ' -f $keycode), 'javascript')
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
    Locates page element
.DESCRIPTION
    Sends clickk to page element located by Javascript by executing Javascript through InternetExplorer.Application
    
.EXAMPLE
    _locate -window_ref ([ref]$window) -locator $locator
.LINK
    
.NOTES
    VERSION HISTORY
    2018/05/12 Initial Version
#>

function _locate {
  param (
    [String]$locator
  )
   return (@"
var selector = '{0}';
var elements = document.querySelectorAll(selector);
var element = elements[0];
"@  -f $locator)
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
    2018/05/21 Working with selects/options too.
#>
function sendKeys {
  param (
    [System.Management.Automation.PSReference]$window_ref,
    [System.Management.Automation.PSReference]$document_element_ref,
    [String]$locator,
    [String]$text = 'entered text'
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
  $sendKeysScript = (@"
var selector = '{0}';
var elements = document.querySelectorAll(selector);
elements[0].value  = '{1}';
"@  -f $locator, $text)
  $window.execScript($sendKeysScript, 'javascript')
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


<#
.SYNOPSIS
    Scrolls the browser window (e.g. into the page element offset)
.DESCRIPTION
    Scrolls the browser window (e.g. into the page element offset)
    by executing Javascript through InternetExplorer.Application wrapper
    
.EXAMPLE
    $document = $ie.document
    $window = $document.parentWindow
    scroll_element_into_view -winow_ref ([ref]$window) -locator $locator
.LINK
    
.NOTES
    VERSION HISTORY
    2018/07/11 Initial Version
#>
function scroll_element_into_view { 
  param (
    [System.Management.Automation.PSReference]$window_ref,
    [string]$locator = $null
  )

  $window = $window_ref.Value
  [string]$scroll_script = ( @"
var selector = '{0}';
var element = document.querySelector(selector);
element.scrollIntoView();
"@  -f $locator)
  try {
    $window.execScript($scroll_script, 'javascript')
  } catch [Exception] {
    write-Debug ( 'Exception : ' + $_.Exception.Message)
    return
  }
}


