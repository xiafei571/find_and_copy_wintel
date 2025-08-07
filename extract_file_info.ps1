# Parameter settings
$targetFolder = "C:\YourTargetFolder"  # Change this to your target folder path

# Get current script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Output file paths
$txtOutputPath = Join-Path $scriptDir "file_names.txt"
$csvOutputPath = Join-Path $scriptDir "file_info.csv"

# Check if target folder exists
if (-not (Test-Path $targetFolder)) {
    Write-Host "Error: Target folder '$targetFolder' does not exist!"
    Write-Host "Please edit the script and set the correct folder path."
    exit 1
}

Write-Host "Scanning folder: $targetFolder"
Write-Host "Please wait..."

# Get all files recursively
try {
    $allFiles = Get-ChildItem -Path $targetFolder -Recurse -File -Force -ErrorAction SilentlyContinue
    Write-Host "Found $($allFiles.Count) files."
} catch {
    Write-Host "Error accessing folder: $_"
    exit 1
}

# Extract file names for TXT file
$fileNames = $allFiles | ForEach-Object { $_.Name }

# Extract detailed info for CSV file
$fileInfo = $allFiles | Select-Object Name, 
    @{Name="DateModified"; Expression={$_.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")}},
    @{Name="Type"; Expression={$_.Extension}}

# Write to TXT file (file names only)
Write-Host "Writing file names to: $txtOutputPath"
$fileNames | Out-File -FilePath $txtOutputPath -Encoding UTF8

# Write to CSV file (Name, Date Modified, Type)
Write-Host "Writing detailed info to: $csvOutputPath"
$fileInfo | Export-Csv -Path $csvOutputPath -NoTypeInformation -Encoding UTF8

# Display summary
Write-Host ""
Write-Host "=== SUMMARY ==="
Write-Host "Total files processed: $($allFiles.Count)"
Write-Host "Files created:"
Write-Host "  - File names (TXT): $txtOutputPath"
Write-Host "  - Detailed info (CSV): $csvOutputPath"
Write-Host "Done!"