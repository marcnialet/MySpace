
$filePath = ".\EventTraces\Logging\LoggingExport.xml"
$encoding = "UTF8"

# Read the contents of the file
$fileContent = Get-Content -Path $filePath

# Save the file with overwrite and UTF-8 encoding
$fileContent | Out-File -FilePath $filePath -Encoding $encoding -Force

Write-Host "File saved successfully with UTF-8 encoding."
