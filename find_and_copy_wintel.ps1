# Parameter settings
$csvPath = "C:\your_file_list.csv"       # Input CSV file path
$searchRoot = "C:\"                      # Root directory to search
$destinationDir = "C:\CollectedFiles"    # Destination directory for copying files

# Supported file extensions
$allowedExtensions = @(".csv", ".log", ".xls", ".tsv", ".pdf", ".txt")

# Regular expression: match date yyyymmdd or yyyymmddhhmmss
$datePattern = '\d{8}(\d{6})?'

# Read CSV content
$lines = Get-Content -Path $csvPath -Encoding UTF8
$headers = $lines[0]
$rows = $lines[1..($lines.Length - 1)]

# Create destination directory (if not exists)
if (-not (Test-Path $destinationDir)) {
    New-Item -ItemType Directory -Path $destinationDir | Out-Null
}

# Read CSV first to get the list of files we need to find
Write-Host "Reading CSV file to determine search targets..."

# Initialize output results
$output = @()
$output += "$headers,FullPath,CopyTime"

# Initialize counters
$totalFiles = 0
$foundFiles = 0
$copiedFiles = 0
$notFoundFiles = 0

# Function to search for a specific file
function Find-FileByPattern {
    param(
        [string]$FileName,
        [string]$SearchRoot,
        [array]$AllowedExtensions
    )
    
    if ($FileName -match 'yyyymmddhhmmss|YYYYMMDDHHMMSS') {
        # Search for files with 14-digit datetime pattern
        $searchPattern = $FileName -replace 'yyyymmddhhmmss|YYYYMMDDHHMMSS', '*'
        $regexPattern = $FileName -replace 'yyyymmddhhmmss|YYYYMMDDHHMMSS', '\d{14}'
    } elseif ($FileName -match 'yyyymmdd|YYYYMMDD') {
        # Search for files with 8-digit date pattern
        $searchPattern = $FileName -replace 'yyyymmdd|YYYYMMDD', '*'
        $regexPattern = $FileName -replace 'yyyymmdd|YYYYMMDD', '\d{8}'
    } else {
        # Exact filename search
        return Get-ChildItem -Path $SearchRoot -Filter $FileName -Recurse -File -Force -ErrorAction SilentlyContinue
    }
    
    # Search with wildcard pattern and filter by regex
    return Get-ChildItem -Path $SearchRoot -Filter $searchPattern -Recurse -File -Force -ErrorAction SilentlyContinue |
        Where-Object { 
            $AllowedExtensions -contains $_.Extension.ToLower() -and
            $_.Name -match $regexPattern
        }
}

# Iterate through each CSV row filename
foreach ($line in $rows) {
    $columns = $line -split ","
    $fileName = $columns[0].Trim()
    $fullPath = ""
    $copyTime = ""
    
    $totalFiles++

    # Only process supported extensions
    $ext = [System.IO.Path]::GetExtension($fileName)
    if (-not $allowedExtensions.Contains($ext.ToLower())) {
        $output += "$line,,"
        continue
    }

    Write-Host "Searching for: $fileName"
    
    # Search for this specific file
    $matchedFiles = Find-FileByPattern -FileName $fileName -SearchRoot $searchRoot -AllowedExtensions $allowedExtensions
    
    if ($matchedFiles.Count -gt 0) {
        $foundFiles++
        
        # If multiple matches or date pattern, find the "latest" one
        if ($fileName -match 'yyyymmdd|YYYYMMDD|yyyymmddhhmmss|YYYYMMDDHHMMSS' -or $matchedFiles.Count -gt 1) {
            $targetFile = $matchedFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        } else {
            $targetFile = $matchedFiles[0]
        }

        # Copy to destination path
        $destPath = Join-Path $destinationDir $targetFile.Name
        Copy-Item -Path $targetFile.FullName -Destination $destPath -Force
        $fullPath = $targetFile.FullName
        $copyTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $copiedFiles++
        Write-Host "Copied: $fileName -> $destPath"
    } else {
        $notFoundFiles++
        Write-Host "Not found: $fileName"
    }

    # Record results
    $output += "$fileName,""$fullPath"",""$copyTime"""
}

# Write back to CSV file (overwrite original file)
$output | Set-Content -Path $csvPath -Encoding UTF8

# Display summary
Write-Host ""
Write-Host "=== SUMMARY ==="
Write-Host "Total files in CSV: $totalFiles"
Write-Host "Found files: $foundFiles"
Write-Host "Successfully copied: $copiedFiles"
Write-Host "Not found: $notFoundFiles"
Write-Host "Updated CSV written to: $csvPath"
