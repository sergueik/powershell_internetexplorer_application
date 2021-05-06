param (
 [string]$data = 'test',
 [string]$file = "C:\temp\a1.txt"
)
write-host $data
out-file -inputobject $data -filepath $file -append

