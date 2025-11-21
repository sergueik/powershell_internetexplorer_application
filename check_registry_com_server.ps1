function check_registry_com_server { 
param(
  [string]$clssid 
)
if (-not ($clssid -match '{[0-9A-F-]+}')) { 
  $clssid = ('{{{0}}}' -f $clssid )
}
if (test-path -path ('HKLM:\SOFTWARE\Classes\CLSID\{0}' -f $clssid )){
   $x = get-itemproperty -path ('HKLM:\SOFTWARE\Classes\CLSID\{0}' -f $clssid ) -name ''
	 write-host ('found {0}: {1}' -f  $clssid, $x.'(default)')
   return $true
} else {
	 write-host ('not found: {0}' -f  $clssid)
  return $false
}
}


check_registry_com_server -clssid	 '00024500-0000-0000-C000-000000000046'
check_registry_com_server -clssid	 '{00024500-0000-0000-C000-000000000046}'

check_registry_com_server -clssid	 '{25336920-03F9-11cf-8FD0-00AA00686F13}'