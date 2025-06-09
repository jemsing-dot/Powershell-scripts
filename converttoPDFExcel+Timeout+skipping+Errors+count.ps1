$scriptStartTime = Get-Date

# Initialize failure log
$global:failureLog = @()  # Use global scope to ensure accessibility

# Initialize COM applications
$word = $null
$excel = $null
$powerpoint = $null

# Initialize counter for converted files
$global:convertedCount = 0

function Convert-ExcelToPDF {
    param (
        [string]$inputFile,
        [string]$outputFile
    )
    try {
        if (-not $excel) { 
        $excel = New-Object -ComObject Excel.Application 
        $excel.Visible = $false
        $excel.DisplayAlerts = 0  # Suppress prompts
        }

        $workbook = $excel.Workbooks.Open($inputFile, $false, $true)
        $workbook.ExportAsFixedFormat(0, $outputFile) # 0 is the xlTypePDF constant
        $workbook.Close($false)
        $global:convertedCount++  # Increment the counter
        Write-Host "Files converted so far: $global:convertedCount"  # Display the current count
    }
    catch {
        Write-Host "Failed to convert Excel document: $inputFile to PDF."
        $global:failureLog += [PSCustomObject]@{
            FileName       = $inputFile
            FileType       = "Excel Document"
            FailureTime    = (Get-Date)

      }
    }
    finally {
        if ($excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            $excel = $null
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
}

# Main script starts here
$directory = "H:\Documents"

# Create a CSV file with headers if it doesn't exist
if (-not (Test-Path "H:\Documents\Test failed conversions.csv")) {
    $header = [PSCustomObject]@{
        FileName    = "File Name"
        FileType    = "File Type"
        FailureTime = "Failure Time"
    }
    $header | Export-Csv -Path "H:\Documents\Test failed conversions.csv" -NoTypeInformation 
}


# Get all files with the specified extensions
$files = Get-ChildItem -Path $directory -Include *.xlsx, *.xls -Recurse -Force

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
            ".xlsx" { Convert-ExcelToPDF $file.FullName $pdfPath }
            ".xls" { Convert-ExcelToPDF $file.FullName $pdfPath }
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
            FileName       = $file.FullName
            FileType       = $extension
            FailureTime    = (Get-Date)
        }

        # Start a timer for 30 seconds
        $timeout = 30
        $startTime = Get-Date

        while ((Get-Date).Subtract($startTime).TotalSeconds -lt $timeout) {
            Start-Sleep -Seconds 1
            # Check if the user has taken any action (e.g., pressing a key)
            if ($Host.UI.RawUI.KeyAvailable) {
                $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                Write-Host "User action detected, continuing..."
                break
            }
        }

        Write-Host "No action detected for 30 seconds, moving to next file..."
    }
}
 

# Write failure log to CSV
if ($failureLog.Count -gt 0) {
    Write-Host "Failures detected: $($failureLog.Count)"
    $failureLog | Export-Csv -Path "H:\Documents\Test failed conversions.csv" -NoTypeInformation -Append
    Write-Host "Failure details have been written to Test failed conversions.csv"
} else {Write-Host "No failures detected."
}

# Output the total number of converted files
Write-Host "Total files converted: $global:convertedCount"

# Clean up COM objects
if ($excel) { $excel.Quit() }

$scriptEndTime = Get-Date
$scriptDuration = $scriptEndTime - $scriptStartTime

Write-Host "The entire script took $($scriptDuration.TotalSeconds) seconds to complete."
