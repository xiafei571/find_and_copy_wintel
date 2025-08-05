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
foreach ($file in $allFiles) {
    $key = $file.Name.ToLower()
    if (-not $nameToFilesMap.ContainsKey($key)) {
        $nameToFilesMap[$key] = @()
    }
    $nameToFilesMap[$key] += $file
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
    
    # If filename contains date pattern, search for files with actual dates
    if ($fileName -match 'yyyymmddhhmmss|YYYYMMDDHHMMSS') {
        $pattern = $fileName -replace 'yyyymmddhhmmss|YYYYMMDDHHMMSS', '\d{14}'
        $matchedFiles = $allFiles | Where-Object { $_.Name -match $pattern }
    } elseif ($fileName -match 'yyyymmdd|YYYYMMDD') {
        $pattern = $fileName -replace 'yyyymmdd|YYYYMMDD', '\d{8}'
        $matchedFiles = $allFiles | Where-Object { $_.Name -match $pattern }
    } else {
        # Exact match lookup
        if ($nameToFilesMap.ContainsKey($key)) {
            $matchedFiles = $nameToFilesMap[$key]
        }
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
