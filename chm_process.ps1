function process_file{
  param(
    [string]$filename = 'ch09s33.html'
  )
  $rawcontent = Get-Content -Path $filename -Raw
  $rawcontent = (($rawcontent -join '') - replace '\n','' ) -replace '<head>.*</head>' ,''

  $htmlfile = New-Object -ComObject 'HTMLFile'
  $htmlfile.IHTMLDocument2_write($rawcontent )
  # write-output $htmlfile.title
  $table = $htmlfile.documentElement.getElementsByTagName('table').item(0)
  # write-output $table.innerHTML
  $element = $table.querySelectorall('CODE[class=computeroutput]').item(0)
  # write-output $element.innerHTML
  $selector = 'span[class="bold"]'
  try {
    $length = $table.querySelectorall($selector).length
  } catch [Exception] {
    $length = 0
  }	
  if ($length -eq 0 ){
    return
  }
  @(0..$length) | foreach-object {
    $index = $_
    # pick 1 column from 3 column row
    if (($index % 3 ) -eq 0) {

    $element  = $table.query Selectorall($selector).item($index)
    write-output $element.innerText
  }
}

get-childitem -filter '*html' | foreach-object {
  $filename = $_
  write-output ('filename: {0}' -f $filename )
  process_file -filename $filename
}
<#
Code
Device
Notes

#>
