# read-host -assecurestring | convertfrom-securestring | out-file C:\temp\pw.txt
$pwcheck = get-content C:\temp\pw.txt | convertto-securestring
$Credential = new-object -typename System.Management.Automation.PSCredential -argumentlist "XXXXXXXXXXXXx", $pwcheck
Start-Sleep -Milliseconds 500
$UserName.value = $Credential.UserName
$Password.value = $credential.GetNetworkCredential().Password
Start-Sleep -Milliseconds 500
