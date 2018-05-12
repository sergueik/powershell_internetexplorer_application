
$ie=New-object -COM InternetExplorer.Application

$ie.navigate2('about:blank')
$ie.width = 400
$ie.height = 600
$ie.Resizable = $true
$ie.StatusBar = $true
$ie.AddressBar = $false
$ie.MenuBar = $false
$ie.Toolbar = $false

# build the html text to display
# from the running services
foreach ($service_object in  ( get-service | where {$_.status -eq 'running' })) { 
$html = $html + '<font face=Verdana size=2>' + $service_object.Displayname+':  '+ $service_object.status+'</font><br>' 
}

# send the html string to the body innerHTML method
$ie.document.body.innerHTML = $html

# write message in the browser statusline
$ie.StatusText = ( $svc.Count ).ToString() +  ' running services '

$ie.Visible=$true

