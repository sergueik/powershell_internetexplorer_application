param(
  $title,
  $message,
  [string]$logfile = "C:\temp\vb_input.txt",
  [switch]$debug
)
[string] $data = 'vb_input'
[string] $message = ('{0} started' -f $data )
write-host $message
out-file -inputobject $message -filepath $logfile
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
$text = [Microsoft.VisualBasic.Interaction]::InputBox($message, $title)
<#
Add-Type -AssemblyName 'PresentationFramework'
[System.Windows.MessageBox]::Show(('Entered {0}' -f $text))
#>
# https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.interaction.msgbox?view=netframework-4.0
if ($debug){
  [Microsoft.VisualBasic.Interaction]::MsgBox( ('Entered {0}' -f $text))|out-null
}
$message = ('{0} closing' -f $data )
out-file -inputobject $message -filepath $logfile -append
write-output $text


