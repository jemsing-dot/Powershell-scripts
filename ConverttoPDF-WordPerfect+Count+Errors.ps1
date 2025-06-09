$scriptStartTime = Get-Date

# Initialize failure log
$global:failureLog = @()  # Use global scope to ensure accessibility

# Initialize COM application
$word = $null

# Initialize counter for converted files
$global:convertedCount = 0

function Convert-WordToPDF {
    param (
        [string]$inputFile,
        [string]$outputFile
    )
    try {
        if (-not $word) { $word = New-Object -ComObject Word.Application }
        $doc = $word.Documents.Open($inputFile)
        $doc.SaveAs([ref] $outputFile, [ref] 17) # 17 is the wdFormatPDF constant
        $doc.Close()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
        $global:convertedCount++  # Increment the counter
        Write-Host "Files converted so far: $global:convertedCount"  # Display the current count
    }
    catch {
        Write-Host "Failed to convert Word document: $inputFile to PDF."
        $global:failureLog += [PSCustomObject]@{
            FileName       = $inputFile
            FileType       = "Word Document"
            FailureTime    = (Get-Date)

     }
    }
    finally {
        if ($word) {
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
            $word = $null
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
}

# Main script starts here
$directory = "\\insert your file paths"

# Create a CSV file with headers if it doesn't exist
if (-not (Test-Path "\\insert your file paths.csv")) {
    $header = [PSCustomObject]@{
        FileName    = "File Name"
        FileType    = "File Type"
        FailureTime = "Failure Time"
    }
    $header | Export-Csv -Path "\\insert your file paths.csv" -NoTypeInformation
}

# Get all files in the directory
$files = Get-ChildItem -Path $directory -Recurse -Force

foreach ($file in $files) {
    $pdfPath = [System.IO.Path]::ChangeExtension($file.FullName, ".pdf")

    try {
        Convert-WordToPDF $file.FullName $pdfPath

        # Copy the time and date stamps
        if (Test-Path $pdfPath) {
            $originalFile = Get-Item $file.FullName
            $pdfFile = Get-Item $pdfPath
            $pdfFile.CreationTime = $originalFile.CreationTime
            $pdfFile.LastWriteTime = $originalFile.LastWriteTime
        }
    }
    catch {
        Write-Host "Error during processing: $($file.FullName). Error: $($_.Exception.Message)"
        $global:failureLog += [PSCustomObject]@{
            FileName       = $file.FullName
            FileType       = $file.Extension
            FailureTime    = (Get-Date)
        }
    }
}

# Write failure log to CSV
if ($failureLog.Count -gt 0) {
    Write-Host "Failures detected: $($failureLog.Count)"
    $failureLog | Export-Csv -Path "\\insert your file paths.csv" -NoTypeInformation -Append
    Write-Host "Failure details have been written to error file"
} else { Write-Host "No failures detected."
}

# Output the total number of converted files
Write-Host "Total files converted: $global:convertedCount"

# Clean up COM objects
if ($word) {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    $word = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$scriptEndTime = Get-Date
$scriptDuration = $scriptEndTime - $scriptStartTime

Write-Host "The entire script took $($scriptDuration.TotalSeconds) seconds to complete."
