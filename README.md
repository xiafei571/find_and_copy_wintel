# find_and_copy_wintel

# 📄 How to Use

## 1. Export Excel Sheets to CSV

- Open each Excel tab (sheet), go to **File > Save As**.
- Save as **CSV UTF-8 (Comma delimited) (*.csv)**.
- Make sure:
  - First row is a header (e.g., `FileName`)
  - First column contains file names.

## 2. Edit the Script

Update these lines in the PowerShell script:

```powershell
$csvPath = "C:\your_file_list.csv"
$searchRoot = "C:\"
$destinationDir = "C:\CollectedFiles"
```

## 3. Run the Script

Open **PowerShell as Administrator**, then:

```powershell
cd C:\path\to\script
.\file_copy.ps1
```

## ✅ What the Script Does

- Recursively searches all files in `C:\` (or your chosen root).
- Matches file names listed in the CSV.
- If the file name includes a date (`yyyymmdd` or `yyyymmddhhmmss`), picks the **latest one**.
- Copies matched files to `C:\CollectedFiles`.
- Updates the original CSV file by filling:
  - **Second column** with the full path of the found file.
  - **Third column** with the copy timestamp.

## ℹ️ Supported File Extensions

- `.csv`, `.log`, `.xls`, `.tsv`, `.pdf`, `.txt`

## 🔁 Can I Run It Multiple Times?

Yes. The script always reprocesses all rows. It will:
- Overwrite previous results in the CSV.
- Copy files again (with `-Force`), replacing older versions in the destination folder.
