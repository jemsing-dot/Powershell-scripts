Start-Transcript

$scriptStartTime = Get-Date

# Define the path to the directory
$directoryPath = "C:\Users\njs361\OneDrive - Newcastle University\DP\DP Format experiments"

# Get all directories within the specified path, including hidden ones
$directories = Get-ChildItem -Path $directoryPath -Recurse -Directory -Force -ErrorAction SilentlyContinue

# Loop through each directory and remove it if it's empty
foreach ($dir in $directories) {
    if ((Get-ChildItem -Path $dir.FullName -Force -ErrorAction SilentlyContinue).Count -eq 0) {
        Remove-Item -Path $dir.FullName -Force
        Write-Host "Deleted empty folder: $($dir.FullName)"
    }
}

$scriptEndTime = Get-Date
$scriptDuration = $scriptEndTime - $scriptStartTime

Write-Host "The entire script took $($scriptDuration.TotalSeconds) seconds to complete."

Stop-Transcript