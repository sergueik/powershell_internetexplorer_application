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

# based on: https://www.automateexcel.com/vba/automate-internet-explorer-ie-using/

# Poweshell offers somewhat counter-intuitive way of locating elements:
# based  on https://community.spiceworks.com/topic/2114024-powershell-ie-automation-hover-button

$MODULE_NAME = 'internetexplorer_application_helper.psd1' ; 
Import-Module -Name ('{0}/{1}' -f '.', $MODULE_NAME ) ;

$ie = new-object -com 'internetexplorer.application' ;
$ie.visible = $true; 
$target_url = 'https://www.automateexcel.com/excel/vba' ;
$ie.navigate2($target_url) ; 
wait_busy -ie_ref ([ref]$ie)  ;


$document = $ie.document ;


#  https://developer.mozilla.org/en-US/docs/Web/API/Node/nodeType
$ELEMENT_NODE = 1 
$xall = $document.all
# $xall.length ~ 500 
$xinput = $xall | where-object {$_.nodeName -eq 'INPUT'}
# $xinput.length ~> 13
$xinput_interact = $xinput |
where-object { $_.type -ne 'submit' -and $_.type -ne 'hidden' }
# $xinput_interact.legth ~> 3

$xinput_interact.item(1).parentNode.outerHTML

# <div class="ginput_container ginput_container_text"><input name="input_1" class="large" id="input_1_1" aria-invalid="false" aria-required="true" type="text" placeholder="First Name" value=""></div>
$xinput_interact.value  = 'test'
# <input type="text" placeholder="Search...">

# NOTE nodeType method appears somewhat slow and should not be used in filter:
# with 500  elements on this page i takes ~ 5 minutes to filter
# $elements = $document.all | where-object {$_.nodeType -eq 1}

#Release COM Object
$ie.quit
[void][Runtime.Interopservices.Marshal]::ReleaseComObject($ieObject)