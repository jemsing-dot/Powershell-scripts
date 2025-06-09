Start-Transcript

$scriptStartTime = Get-Date


# Define the path to the directory
$directoryPath = "C:\Users\njs361\OneDrive - Newcastle University\DP\DP Format experiments"

# Define the file extensions to delete
$fileExtensions = @(".rf", ".rtf", ".s01", ".sat", ".scd", ".scr", ".spf", ".tbl", ".txt", ".text", ".thm", ".tmp", ".ttf", ".vcf", ".vob", ".wbk", ".wmf", ".xls", ".xlsx", ".zip", ".zpf")  # Add your desired file formats here

# Get all files within the specified path
$files = Get-ChildItem -Path $directoryPath -Recurse -File -Force

# Loop through each file and remove it if it matches the specified extensions
foreach ($file in $files) {
    if ($fileExtensions -contains $file.Extension) {
        Remove-Item -Path $file.FullName -Force
        Write-Host "Deleted file: $($file.FullName)"
    }
}

$scriptEndTime = Get-Date
$scriptDuration = $scriptEndTime - $scriptStartTime

Write-Host "The entire script took $($scriptDuration.TotalSeconds) seconds to complete."

Stop-Transcript