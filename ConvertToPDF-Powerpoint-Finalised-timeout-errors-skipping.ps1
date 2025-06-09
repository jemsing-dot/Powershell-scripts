$scriptStartTime = Get-Date

# Initialize failure log
$global:failureLog = @()

# Initialize COM applications
$word = $null
$excel = $null
$powerpoint = $null

# Initialize counter for converted files
$global:convertedCount = 0

function Convert-PowerPointToPDF {
    param (
        [string]$inputFile,
        [string]$outputFile
    )
    try {
        if (-not $powerpoint) {
            $powerpoint = New-Object -ComObject PowerPoint.Application
        }

        Write-Host "Opening: $inputFile"
        $presentation = $powerpoint.Presentations.Open($inputFile, [ref] $false, [ref] $false, [ref] $false)

        Write-Host "Saving as PDF to: $outputFile"
        $presentation.SaveAs($outputFile, 32)  # 32 = PDF format
        $presentation.Close()

        $global:convertedCount++
        Write-Host "Files converted so far: $global:convertedCount"
    }
    catch {
        Write-Host "Failed to convert presentation to PDF: $inputFile"
        $global:failureLog += [PSCustomObject]@{
            FileName    = $inputFile
            FileType    = "PowerPoint"
            FailureTime = (Get-Date)
        }

        # Wait for 30 seconds in case of failure
        Write-Host "Waiting 30 seconds in case of failure..."
        Start-Sleep -Seconds 30
    }
    finally {
        if ($powerpoint) {
            $powerpoint.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint) | Out-Null
            $powerpoint = $null
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
}

# Main script starts here
$directory = "H:\Documents"

# Create a CSV file with headers if it doesn't exist
$failureLogPath = "H:\Documents\FailedConversions.csv"
if (-not (Test-Path $failureLogPath)) {
    $header = [PSCustomObject]@{
        FileName    = "File Name"
        FileType    = "File Type"
        FailureTime = "Failure Time"
    }
    $header | Export-Csv -Path $failureLogPath -NoTypeInformation
}

# Get all files with the specified extensions
$files = Get-ChildItem -Path $directory -Include *.ppt, *.pptx -Recurse -Force

foreach ($file in $files) {
    $extension = $file.Extension.ToLower()
    $pdfPath = [System.IO.Path]::ChangeExtension($file.FullName, ".pdf")

    # Skip conversion if PDF already exists
    if (Test-Path $pdfPath) {
        Write-Host "Skipping '$($file.FullName)' – PDF already exists."
        continue
    }

    try {
        switch ($extension) {
            ".pptx" { Convert-PowerPointToPDF $file.FullName $pdfPath }
            ".ppt"  { Convert-PowerPointToPDF $file.FullName $pdfPath }
            default {
                Write-Host "Unsupported file type: $($file.FullName)"
            }
        }

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
            FileName    = $file.FullName
            FileType    = $extension
            FailureTime = (Get-Date)
        }

        # Wait for 30 seconds in case of failure
        Write-Host "Waiting 30 seconds in case of failure..."
        Start-Sleep -Seconds 30
    }
}

# Write failure log to CSV
if ($failureLog.Count -gt 0) {
    Write-Host "Failures detected: $($failureLog.Count)"
    $failureLog | Export-Csv -Path $failureLogPath -NoTypeInformation -Append
    Write-Host "Failure details written to: $failureLogPath"
} else {
    Write-Host "No failures detected."
}

# Output the total number of converted files
Write-Host "Total files converted: $global:convertedCount"

$scriptEndTime = Get-Date
$scriptDuration = $scriptEndTime - $scriptStartTime
Write-Host "The entire script took $($scriptDuration.TotalSeconds) seconds to complete."
