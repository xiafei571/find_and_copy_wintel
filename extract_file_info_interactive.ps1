# Set console encoding to UTF-8 for Japanese characters
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 | Out-Null

Write-Host "=== File Information Extractor ==="
Write-Host ""
Write-Host "How to input folder path with Japanese names:"
Write-Host "1. Open File Explorer and navigate to your target folder"
Write-Host "2. Copy the path from the address bar (Ctrl+L, then Ctrl+C)"
Write-Host "3. Paste it below, or drag the folder into this window"
Write-Host ""

# Get target folder from user input
do {
    $targetFolder = Read-Host "Enter target folder path"
    $targetFolder = $targetFolder.Trim('"')  # Remove quotes if present
    
    if ([string]::IsNullOrWhiteSpace($targetFolder)) {
        Write-Host "Please enter a valid path."
        continue
    }
    
    Write-Host "Checking path: $targetFolder"
    
    if (Test-Path -LiteralPath $targetFolder) {
        break
    } else {
        Write-Host "Folder not found. Please try again."
        Write-Host "Make sure to copy the full path from File Explorer."
        Write-Host ""
    }
} while ($true)

# Get current script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Output file paths
$txtOutputPath = Join-Path $scriptDir "file_names.txt"
$csvOutputPath = Join-Path $scriptDir "file_info.csv"

Write-Host ""
Write-Host "Scanning folder: $targetFolder"
Write-Host "Please wait..."

# Get all files recursively
try {
    $allFiles = Get-ChildItem -LiteralPath $targetFolder -Recurse -File -Force -ErrorAction SilentlyContinue
    Write-Host "Found $($allFiles.Count) files."
} catch {
    Write-Host "Error accessing folder: $_"
    Read-Host "Press Enter to exit"
    exit 1
}

if ($allFiles.Count -eq 0) {
    Write-Host "No files found in the specified folder."
    Read-Host "Press Enter to exit"
    exit 0
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
Write-Host ""
Write-Host "Done!"
Read-Host "Press Enter to exit"