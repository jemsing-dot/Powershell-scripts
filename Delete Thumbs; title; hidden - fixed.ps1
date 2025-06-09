Start-Transcript

$scriptStartTime = Get-Date

# Define the path to the directory
$directoryPath = "C:\Users\njs361\OneDrive - Newcastle University\DP\DP Format experiments"

# Initialize a counter for deleted files
$global:deletedFilesCount = 0

# Function to delete files based on filter
function Delete-Files {
    param (
        [string]$filter
    )

    Get-ChildItem -Path $directoryPath -Recurse -Force | Where-Object {
        $_.Name -like $filter -and -not $_.PSIsContainer
    } | ForEach-Object {
        Remove-Item -Path $_.FullName -Force
        Write-Host "Deleted file: $($_.FullName)"
        $global:deletedFilesCount++
    }
}

# Function to delete hidden files
function Delete-HiddenFiles {
    Get-ChildItem -Path $directoryPath -Recurse -Force | Where-Object {
        $_.Attributes -match "Hidden" -and -not $_.PSIsContainer
    } | ForEach-Object {
        Remove-Item -Path $_.FullName -Force
        Write-Host "Deleted hidden file: $($_.FullName)"
        $global:deletedFilesCount++
    }
}

# Function to delete Thumbs.db files
function Delete-ThumbsDbFiles {
    Get-ChildItem -Path $directoryPath -Recurse -Force | Where-Object {
        $_.Name -eq "Thumbs.db" -and -not $_.PSIsContainer
    } | ForEach-Object {
        Remove-Item -Path $_.FullName -Force
        Write-Host "Deleted Thumbs.db file: $($_.FullName)"
        $global:deletedFilesCount++
    }
}

# Delete all ~$ files
Delete-Files -filter "~$*"

# Delete all hidden files
Delete-HiddenFiles

# Delete all Thumbs.db files
Delete-ThumbsDbFiles

Write-Host "Deleted all ~$ files, hidden files, and Thumbs.db files."

$scriptEndTime = Get-Date
$scriptDuration = $scriptEndTime - $scriptStartTime

Write-Host "All specified files have been deleted. Total files deleted: $global:deletedFilesCount. Duration: $($scriptDuration.TotalSeconds) seconds"

Stop-Transcript