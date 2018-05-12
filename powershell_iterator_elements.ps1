# Poweshell offers somewhat counter-intuitive way of locating elements:
# based  on https://community.spiceworks.com/topic/2114024-powershell-ie-automation-hover-button
$document = $ie.document
$SearchButton = $document.all | Where-Object {$_.tagname -like 'DIV*'} | Where-Object {$_.classname -eq 'ButtonGroup ExecuteSearch'} | foreach {$_.children}

