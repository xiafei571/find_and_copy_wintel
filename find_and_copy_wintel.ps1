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

# Recursively scan all files (limited to allowed extensions)
Write-Host "Indexing all files under $searchRoot (please wait)..."
$allFiles = Get-ChildItem -Path $searchRoot -Recurse -File -Force -ErrorAction SilentlyContinue |
    Where-Object { $allowedExtensions -contains $_.Extension.ToLower() }

# Build mapping from filename to file object array
$nameToFilesMap = @{}
$patternToFilesMap = @{}

foreach ($file in $allFiles) {
    $fileName = $file.Name.ToLower()
    
    # Exact name mapping
    if (-not $nameToFilesMap.ContainsKey($fileName)) {
        $nameToFilesMap[$fileName] = @()
    }
    $nameToFilesMap[$fileName] += $file
    
    # Pattern mapping for files with dates
    if ($fileName -match '\d{14}') {
        # Replace 14-digit datetime with pattern
        $pattern = $fileName -replace '\d{14}', 'yyyymmddhhmmss'
        if (-not $patternToFilesMap.ContainsKey($pattern)) {
            $patternToFilesMap[$pattern] = @()
        }
        $patternToFilesMap[$pattern] += $file
        
        # Also add uppercase pattern
        $patternUpper = $fileName -replace '\d{14}', 'YYYYMMDDHHMMSS'
        if (-not $patternToFilesMap.ContainsKey($patternUpper)) {
            $patternToFilesMap[$patternUpper] = @()
        }
        $patternToFilesMap[$patternUpper] += $file
    } elseif ($fileName -match '\d{8}') {
        # Replace 8-digit date with pattern
        $pattern = $fileName -replace '\d{8}', 'yyyymmdd'
        if (-not $patternToFilesMap.ContainsKey($pattern)) {
            $patternToFilesMap[$pattern] = @()
        }
        $patternToFilesMap[$pattern] += $file
        
        # Also add uppercase pattern
        $patternUpper = $fileName -replace '\d{8}', 'YYYYMMDD'
        if (-not $patternToFilesMap.ContainsKey($patternUpper)) {
            $patternToFilesMap[$patternUpper] = @()
        }
        $patternToFilesMap[$patternUpper] += $file
    }
}
Write-Host "Indexed $($allFiles.Count) files."

# Initialize output results
$output = @()
$output += "$headers,FullPath,CopyTime"

# Iterate through each CSV row filename
foreach ($line in $rows) {
    $columns = $line -split ","
    $fileName = $columns[0].Trim()
    $key = $fileName.ToLower()
    $fullPath = ""
    $copyTime = ""

    # Only process supported extensions
    $ext = [System.IO.Path]::GetExtension($fileName)
    if (-not $allowedExtensions.Contains($ext.ToLower())) {
        $output += "$line,,"
        continue
    }

    # Check if matching files exist
    $matchedFiles = @()
    
    # Try pattern matching first (much faster)
    if ($patternToFilesMap.ContainsKey($key)) {
        $matchedFiles = $patternToFilesMap[$key]
    } elseif ($nameToFilesMap.ContainsKey($key)) {
        # Exact match lookup
        $matchedFiles = $nameToFilesMap[$key]
    }
    
    if ($matchedFiles.Count -gt 0) {
        # If filename contains date pattern or multiple matches, find the "latest" one
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
        Write-Host "Copied: $fileName -> $destPath"
    } else {
        Write-Host "Not found: $fileName"
    }

    # Record results
    $output += "$fileName,""$fullPath"",""$copyTime"""
}

# Write back to CSV file (overwrite original file)
$output | Set-Content -Path $csvPath -Encoding UTF8
Write-Host "All done. Updated CSV written to $csvPath"
