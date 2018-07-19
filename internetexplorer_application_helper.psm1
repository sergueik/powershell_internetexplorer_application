
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

<#
.SYNOPSIS
    Closes the browser window, releases COM reference
.DESCRIPTION
    Closes the browser window, releases COM reference

.EXAMPLE
    finish_test ([ref]$ie)
.LINK

.NOTES
    VERSION HISTORY
    2018/07/11 Initial Version
#>

function finish_test {
  # TODO: valuefrompipeline
  param (
    [System.Management.Automation.PSReference]$ie_ref
  )

  $ie = $ie_ref.Value
  $ie.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ie) | out-null
  Remove-Variable ie
}

<#
.SYNOPSIS
    Runs the javascript in the web page to collect and optionally returns the rowset of attribute or text data paired together from element found via querySelectorAll

.DESCRIPTION
    Runs the javascript in the web page to collect and optionally returns the rowset of attribute pairs or text data from element found via querySelectorAll,
    It  is useful because the querySelectorAll method is not very stable with IE controlled through Powershell
    NOTE: setting debug switch will lead to browser showing the data in the alert dialog.

.EXAMPLE
    $result_tag = 'result'
    collect_data_hash -window_ref ([ref]$window) `
      -element_locator 'a' -key_attribute 'href' -value_attribute 'text' -result_tag $result_tag
    [String]$result_raw = $document.body.getAttribute($result_tag)
    write-output ('Result (raw): ( in "' + $result_tag + '") ' + $result_raw)
    # NOTE: final conversion on the caller side
    try {
      $result_obj = $result_raw   | convertfrom-json
        format-list -InputObject $result_obj
    } catch [Exception] {
        write-output ('Exception : ' + $_.Exception.Message)
    }
    When the document_ref parameter is provided, the function itself collects and formats the result array as in:

.LINK

.NOTES
    VERSION HISTORY
    2018/07/18 Initial Version
#>

function collect_data_hash {
  param (
    [System.Management.Automation.PSReference]$window_ref,
    # NOTE: setting the default value of $null to PSReference paramter is a bad idea:
    # Cannot process argument transformation on parameter'document_ref'. Reference type is expected in argument.
    [System.Management.Automation.PSReference]$document_ref,
    [String]$element_locator,
    [String]$key_attribute = $null,
    [String]$value_attribute = 'class',
    [string]$result_tag = 'PSResult',
    [switch]$debug
  )

  $window = $window_ref.Value

  [bool]$debug_flag = [bool]$PSBoundParameters['debug'].IsPresent
  [string]$debug_str = 'false'
  if ($debug_flag) {
    $debug_str =  'true'
  } else {
    $debug_str = 'false'
  }
  # can not directly return the value, need to place it into the page
  # https://stackoverflow.com/questions/26021813/ie-com-automation-how-to-get-the-return-value-of-window-execscript-in-powersh
  $script = @"
    var element_locator = '${element_locator}';
    var key_attribute = '${key_attribute}';
    var value_attribute = '${value_attribute}';
    var result_tag = '${result_tag}';
    var debug = ${debug_str};

    var elements = document.querySelectorAll(element_locator);
    var result = [];
    for (var cnt =0 ;cnt != elements.length ; cnt ++) {
      var element = elements[cnt];
      var data_key = ''
      if (key_attribute!= ''){
        data_key = element.getAttribute(key_attribute)
      } else {
        data_key = element.innerHTML
      }
      result.push( {
        'key':  data_key,
        'value': element.getAttribute(value_attribute),
      });
    }
    document.body.setAttribute(result_tag , JSON.stringify(result) );
    if (debug) {
      alert('Result: ( in "' + result_tag + '") ' + document.body.getAttribute(result_tag));
    }
"@
# ^^ NOTE: heredoc end mark needs to be placed the beginning or the line.
# TODO: multiline heredoc for $script_template

  if ($debug_flag) {
    write-debug ("Script`n:{0}" -f $script)
  }
  $window.execScript($script, 'javascript')

  [String[]]$result_rowset = @(@{})
  if ($document_ref -ne $null){
    $document = $document_ref.Value
    $result_raw = $document.body.getAttribute($result_tag)
    write-debug ('Result (raw): ( in "' + $result_tag + '") ' + $result_raw)
    try {
      $result_rowset = $result_raw | convertfrom-json
      if (-not ($debugpreference -match 'continue')) {
        format-list -InputObject $result_rowset
      }
    } catch [Exception] {
      if (-not ($debugpreference -match 'continue')) {
        throw $message
      } else {
        write-debug ('Exception : ' + $_.Exception.Message)
      }
    }
  }
  return $result_rowset
}

<#
.SYNOPSIS
    Runs the javascript in the web page to collect and optionally returns the array of attribute or text data paired together from element found via querySelectorAll

.DESCRIPTION
    Runs the javascript in the web page to collect the array of attribute pairs or text data from element found via querySelectorAll,
    It  is useful because the querySelectorAll method is not very stable with IE controlled through Powershell
    NOTE: setting debug switch will lead to browser showing the data in the alert dialog.
.EXAMPLE
    $result_tag = 'result'
    $element_locator = 'section#downloads ul.driver-downloads li.driver-download > a'
    $element_attribute = 'href'
    collect_data_array -window_ref ([ref]$window) `
      -element_locator 'a' -element_attribute 'href' -result_tag $result_tag
    [String]$result_raw = $document.body.getAttribute($result_tag)
    write-output ('Result (raw): ( in "' + $result_tag + '") ' + $result_raw)
    # NOTE: final conversion on the caller side
    $result_array = ($result_raw -replace '^\[', '' -replace '\]$' ) -split ','
    $result_array | format-list
    When the document_ref parameter is provided, the function itself collects and formats the result array as in:
    $result = collect_data_array -window_ref ([ref]$window) `
      -element_locator 'a' -element_attribute 'href' -result_tag $result_tag

.LINK

.NOTES
    VERSION HISTORY
    2018/07/18 Initial Version
#>

# a data collector variant with a different DOM of the response JSON
function collect_data_array {
  param (
    [System.Management.Automation.PSReference]$window_ref,
    # NOTE: setting the default value of $null to PSReference paramter is a bad idea:
    # Cannot process argument transformation on parameter'document_ref'. Reference type is expected in argument.
    [System.Management.Automation.PSReference]$document_ref,
    [String]$element_locator,
    [String]$element_attribute = 'class',
    [string]$result_tag = 'PSResult',
    [switch]$debug
  )

  $window = $window_ref.Value

  [bool]$debug_flag = [bool]$PSBoundParameters['debug'].IsPresent
  [string]$debug_str = 'false'
  if ($debug_flag) {
    $debug_str =  'true'
  } else {
    $debug_str = 'false'
  }
  $script = @"
    var element_locator = '${element_locator}';
    var element_attribute = '${element_attribute}';
    var result_tag = '${result_tag}';
    var debug = ${debug_str};

    var elements = document.querySelectorAll(element_locator);
    var result = [];
    for (var cnt =0 ;cnt != elements.length ; cnt ++) {
      var element = elements[cnt];
      result.push( element.getAttribute(element_attribute));
    }
    document.body.setAttribute(result_tag , JSON.stringify(result) );
    if (debug) {
      alert('Result: ( in "' + result_tag + '") ' + document.body.getAttribute(result_tag));
    }
"@
  if ($debug_flag) {
    write-debug ("Script`n:{0}" -f $script)
  }
  $window.execScript($script, 'javascript')
  [String[]]$result_array = @()
  if ($document_ref -ne $null){
    $document = $document_ref.Value
    [String]$result_raw = $document.body.getAttribute($result_tag)
    write-debug ('Result (raw): ( in "' + $result_tag + '") ' + $result_raw)
    # NOTE: final conversion on the caller side
    $result_array = ($result_raw -replace '^\[', '' -replace '\]$' ) -split ','
    if (-not ($debugpreference -match 'continue')) {
      $result_array | format-list
    }
  }
  return $result_array;
}