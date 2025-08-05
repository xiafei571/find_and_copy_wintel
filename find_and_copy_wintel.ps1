# 参数设定
$csvPath = "C:\your_file_list.csv"       # 输入的 CSV 路径
$searchRoot = "C:\"                      # 要搜索的根目录
$destinationDir = "C:\CollectedFiles"    # 要复制到的目标目录

# 支持的文件扩展名
$allowedExtensions = @(".csv", ".log", ".xls", ".tsv", ".pdf", ".txt")

# 正则表达式：匹配日期 yyyymmdd 或 yyyymmddhhmmss
$datePattern = '\d{8}(\d{6})?'

# 读取 CSV 内容
$lines = Get-Content -Path $csvPath -Encoding UTF8
$headers = $lines[0]
$rows = $lines[1..($lines.Length - 1)]

# 创建目标目录（如不存在）
if (-not (Test-Path $destinationDir)) {
    New-Item -ItemType Directory -Path $destinationDir | Out-Null
}

# 递归扫描所有文件（仅限允许后缀）
Write-Host "📂 Indexing all files under $searchRoot (please wait)..."
$allFiles = Get-ChildItem -Path $searchRoot -Recurse -File -Force -ErrorAction SilentlyContinue |
    Where-Object { $allowedExtensions -contains $_.Extension.ToLower() }

# 建立 文件名 → 文件对象数组 的映射
$nameToFilesMap = @{}
foreach ($file in $allFiles) {
    $key = $file.Name.ToLower()
    if (-not $nameToFilesMap.ContainsKey($key)) {
        $nameToFilesMap[$key] = @()
    }
    $nameToFilesMap[$key] += $file
}
Write-Host "✅ Indexed $($allFiles.Count) files."

# 初始化输出结果
$output = @()
$output += "$headers,FullPath,CopyTime"

# 遍历每一行 CSV 文件名
foreach ($line in $rows) {
    $columns = $line -split ","
    $fileName = $columns[0].Trim()
    $key = $fileName.ToLower()
    $fullPath = ""
    $copyTime = ""

    # 仅处理支持后缀
    $ext = [System.IO.Path]::GetExtension($fileName)
    if (-not $allowedExtensions.Contains($ext.ToLower())) {
        $output += "$line,,"
        continue
    }

    # 查找是否存在匹配文件
    if ($nameToFilesMap.ContainsKey($key)) {
        $matchedFiles = $nameToFilesMap[$key]

        # 如果文件名中含日期，则找出“最新”的一个
        if ($fileName -match $datePattern) {
            $targetFile = $matchedFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        } else {
            $targetFile = $matchedFiles[0]
        }

        # 复制到目标路径
        $destPath = Join-Path $destinationDir $targetFile.Name
        Copy-Item -Path $targetFile.FullName -Destination $destPath -Force
        $fullPath = $targetFile.FullName
        $copyTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Host "📁 Copied: $fileName -> $destPath"
    } else {
        Write-Host "❌ Not found: $fileName"
    }

    # 记录结果
    $output += "$fileName,""$fullPath"",""$copyTime"""
}

# 写回 CSV 文件（覆盖原始文件）
$output | Set-Content -Path $csvPath -Encoding UTF8
Write-Host "✅ All done. Updated CSV written to $csvPath"
