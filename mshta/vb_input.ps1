param(
  $title,
  $message,
  [switch]$debug
)

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
write-output $text

