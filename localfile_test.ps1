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

# https://stackoverflow.com/questions/33558807/powershell-internet-explorer-com-object-select-class-drop-down-menu-item
# http://www.cyberforum.ru/powershell/thread2281419.html
# (in Russian)

# NOTE: requires elevation
param (
  [switch]$configure
)

$MODULE_NAME = 'internetexplorer_application_helper.psd1'
Import-Module -Name ('{0}/{1}' -f '.', $MODULE_NAME )

# based on: http://blogs.msdn.com/b/virtual_pc_guy/archive/2010/09/23/a-self-elevating-powershell-script.aspx
# Get the ID and security principal of the current user account
$id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object System.Security.Principal.WindowsPrincipal ($id)

# Get the security principal for the Administrator role
$admin = [System.Security.Principal.WindowsBuiltInRole]::Administrator

# Check to see if we are currently running "as Administrator"
if ( -not $principal.IsInRole($admin)) {
  write-error 'Please relaunch "as Administrator"'
  exit 1
}

[bool]$configure_flag = [bool]$PSBoundParameters['configure'].IsPresent

if ($configure_flag) {
  # https://support.microsoft.com/en-us/help/2002093/allow-active-content-to-run-files-on-my-computer-group-policy-setting
  $hive = 'HKLM:'
  $path = '/Software/Microsoft/Internet Explorer/Main/FeatureControl/FEATURE_LOCALMACHINE_LOCKDOWN'
  $name = 'iexplore.exe'
  $value = '0'
  $propertyType = 'Dword'
  change_registry_setting -hive $hive -Name $name -Value $value -PropertyType $propertyType

  $hive = 'HKCU:'
  $path = '/Software/Microsoft/Internet Explorer/Main/FeatureControl'
  $name = 'iexplore.exe'
  $value = '0'
  $propertyType = 'Dword'

  $path_key = 'FEATURE_LOCALMACHINE_LOCKDOWN'

  pushd ('{0}/' -f $hive )
  $registry_path_status = Test-Path -Path ('{0}/{1}' -f $path, $path_key) -ErrorAction 'SilentlyContinue'
  if ($registry_path_status -ne $true) {

  new-item -path $path -name $path_key
  # New-ItemProperty : Attempted to perform an unauthorized operation.
  }
  popd
  change_registry_setting -hive $hive -Name $name -Value $value -PropertyType $propertyType -path ('{0}/{1}' -f $path, $path_key)
}

$ie = new-object -com 'internetexplorer.application'
<#
TODO: detect
new-object : Creating an instance of the COM component with CLSID
{0002DF01-0000-0000-C000-000000000046} from the IClassFactory failed due to the following error: 800704a6 A system shutdown has already been scheduled.
(Exception from HRESULT: 0x800704A6).
#>
# see also 'MSXML2.DOMDocument'
$ie.visible = $true
[string]$url = 'C:\developer\sergueik\powershell_internetexplorer_application\login.html'

$ie.navigate2($url)

wait_busy -ie_ref ([ref]$ie)
start-sleep -Milliseconds 1000
$debug =  $false
$document = $ie.document
$document_element = $document.documentElement
$window = $document.parentWindow
$forms = $document.forms
$form = $forms.namedItem('login')

$login = 'мой_логин'
$password = 'password'

$form.item('j_username').value = $login
$form.item('j_password').value = $password

$select = $document.getElementsByClassName('formStyle') | where-object { $_.name -eq 'domain'}
$options = $select.options
$options.selectedIndex = 3

# quit and dispose IE
start-sleep -second 10
$ie.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ie) | out-null
remove-variable ie
