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

# https://stackoverflow.com/questions/21451412/in-ie-org-openqa-selenium-webdriverexception-this-usually-means-that-a-call-to/41704106
# https://automated-testing.info/t/ne-zapuskayutsya-testy-v-ie-oshibka-navigaczii/21029 (in Russian)

$MODULE_NAME = 'internetexplorer_application_helper.psd1'
Import-Module -Name ('{0}/{1}' -f '.', $MODULE_NAME )

$ie = new-object -com 'internetexplorer.application'
# see also 'MSXML2.DOMDocument'
$ie.visible = $true
# http://a.testaddressbook.com/sign_in

function change_registry_setting {

  param(
    [string]$hive,
    [string]$path,
    [string]$name,
    [string]$value,
    [string]$propertyType,
    # will be converted to 'Microsoft.Win32.RegistryValueKind' enumeration
    # 'String', 'ExpandString', 'Binary', 'DWord', 'MultiString', 'QWord'
    [switch]$debug

  )
  pushd $hive
  cd $path
  $local:setting = Get-ItemProperty -Path ('{0}/{1}' -f $hive,$path) -Name $name -ErrorAction 'SilentlyContinue'
  if ($local:setting -ne $null) {
    if ([bool]$PSBoundParameters['debug'].IsPresent) {
      Select-Object -ExpandProperty $name -InputObject $local:setting
    }
    if ($local:setting -ne $value) {
      Set-ItemProperty -Path ('{0}/{1}' -f $hive,$path) -Name $name -Value $value
    }
  } else {
    New-ItemProperty -Path ('{0}/{1}' -f $hive,$path) -Name $name -Value $value -PropertyType $propertyType
  }
  popd

}
# NOTE: requires elevation
$hive = 'HKLM:'
$path = '/Software/Microsoft/Internet Explorer/Main/FeatureControl/FEATURE_HTTP_USERNAME_PASSWORD_DISABLE'
$name = 'iexplore.exe'
$value = '0'
$propertyType = 'Dword'
change_registry_setting -hive $hive -Name $name -Value $value -PropertyType $propertyType

$hive = 'HKCU:'
$path = '/Software/Microsoft/Internet Explorer/Main/FeatureControl'
$name = 'iexplore.exe'
$value = '0'
$propertyType = 'Dword'

$path_key = 'FEATURE_HTTP_USERNAME_PASSWORD_DISABLE'

pushd $hive
$registry_path_status = Test-Path -Path ('{0}/{1}' -f $path, $path_key) -ErrorAction 'SilentlyContinue'
if ($registry_path_status -ne $true) {

new-item -path $path -name $path_key
# New-ItemProperty : Attempted to perform an unauthorized operation.
}
popd
change_registry_setting -hive $hive -Name $name -Value $value -PropertyType $propertyType -path ('{0}/{1}' -f $path, $path_key)


$target_url = 'http://test%40gmail.com:password@a.testaddressbook.com/sign_in'
# NOTE: this site does not actually recognize authentication. Used only to demonstrate the effect of the registry setting and securiry exception
# https://stackoverflow.com/questions/5725430/http-test-server-that-accepts-get-post-calls
try {
  $ie.navigate2($target_url)
  wait_busy -ie_ref ([ref]$ie)
  start-sleep -Milliseconds 1000
  $debug =  $false
  $document = $ie.document
  $document_element = $document.documentElement
  $window = $document.parentWindow
} catch [Exception] {
  write-Debug ( 'Exception : ' + $_.Exception.Message)
  # Exception : A security problem occurred. (Exception from HRESULT:0x800C000E)
}
finally {
  # quit and dispose IE
  $ie.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ie) | out-null
  Remove-Variable ie
}