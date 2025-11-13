
param(
  [string]$filename = 'ch09s33.html'
)
$elementawcontent = Get-Content -Path $filename -Raw

$htmlfile = New-Object -ComObject 'HTMLFile'
$htmlfile.IHTMLDocument2_write($elementawcontent )
# write-output $htmlfile.title
$table = $htmlfile.documentElement.getElementsByTagName('table').item(0)
# write-output $table.innerHTML
$element = $table.querySelectorall('CODE[class=computeroutput]').item(0)
# write-output $element.innerHTML
$selector = 'span[class="bold"]'
$length = $table.querySelectorall($selector).length
@(0..$length) | foreach-object {
  $index = $_

  $element = $table.querySelectorall($selector).item($index)
  write-output $element.innerText
}
<#
Code
Device
Notes

#>
