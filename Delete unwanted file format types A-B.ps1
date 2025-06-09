Start-Transcript

$scriptStartTime = Get-Date


# Define the path to the directory
$directoryPath = "insert file path"

# Define the file extensions to delete
$fileExtensions = @(".2", ".150", ".160", ".07", ".07sgm edits", ".19dec03", ".3ds", ".ac$", ".ai", ".avi", ".bak", ".bk1", ".bk2", ".bk3", ".bk4", ".bmp", ".bp2", ".bsw", ".bup")  # Add your desired file formats here

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
