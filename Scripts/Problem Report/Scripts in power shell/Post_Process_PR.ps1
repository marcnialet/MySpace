$filePath = ".\EventTraces\Logging\LoggingExport.xml"
$encoding = "UTF8"

# Read the contents of the file
$fileContent = Get-Content -Path $filePath

# Save the file with overwrite and UTF-8 encoding
$fileContent | Out-File -FilePath $filePath -Encoding $encoding -Force

Write-Host "File saved successfully with UTF-8 encoding."

$folderPath = ".\EventTraces\SystemSoftwareLog"  # Replace with the path to your folder
$filePattern = "Roche.C4C.ServiceHosting.BackEnd.log*"  # Replace with the file pattern


$destFolder  = "{0}\Parts\" -f $folderPath
Write-Host "Destination folder: $destFolder"
# delete the destination folder if it exists
if (test-path -path $destfolder) {
 remove-item -path $destfolder -recurse -force
 }
 # create a new destination folder
 new-item -itemtype directory -path $destfolder | out-null

$files = Get-ChildItem -Path $folderPath -Filter $filePattern

$partNumber = 1
$chunkSize = 20000 * 1KB

$index = 0
foreach ($file in $files) {
     Write-Host "Processing file: $($file.Name)"
	 Write-Host "Processing file: $($file.FullName)"
     $content = Get-Content $file.FullName -Raw

     while ($content.Length -gt 0) {
        $chunk = $content.Substring(0, [Math]::Min($chunkSize, $content.Length))
        $content = $content.Substring($chunk.Length)

		$newFileName = "{0}\Parts\Roche.C4C.ServiceHosting.BackEnd.log.{1:D3}.part.{2:D3}" -f $folderPath, $index, $partNumber
        $newFilePath = $newFileName
		
		$newFolder = Split-Path -Path $newFilePath -Parent
		Write-Host "New base folder: $newFolder"
		# if (!(Test-Path -Path $newFilePath)) {
            # New-Item -ItemType Directory -Path $newFolder | Out-Null
        # }
		# Join-Path -Path $folderPath -ChildPath $newFileName

		Write-Host "New filename: $newFileName"
		Write-Host "New filepath: $newFilePath"	
		

        Set-Content -Path $newFilePath -Value $chunk

        Write-Host "Split file created: $newFilePath"

        $partNumber++        
     }
	 $partNumber=0
	 $index++
}
