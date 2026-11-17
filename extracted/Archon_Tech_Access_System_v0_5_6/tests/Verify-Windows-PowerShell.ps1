[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$scriptFiles = @(Get-ChildItem -LiteralPath $projectRoot -Filter '*.ps1' -File -Recurse -Force | Sort-Object FullName)
if ($scriptFiles.Count -eq 0) { throw 'لم يتم العثور على ملفات PowerShell داخل الحزمة.' }
$failed = $false
foreach ($file in $scriptFiles) {
    $scriptPath = $file.FullName
    $bytes = [IO.File]::ReadAllBytes($scriptPath)
    $hasBom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
    $hasDoubleBom = $bytes.Length -ge 6 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF -and $bytes[3] -eq 0xEF -and $bytes[4] -eq 0xBB -and $bytes[5] -eq 0xBF
    if (-not $hasBom -or $hasDoubleBom) {
        [Console]::Error.WriteLine("ترميز UTF-8 BOM غير صحيح أو مكرر: $scriptPath")
        $failed = $true
    }
    $tokens = $null
    $parseErrors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count -gt 0) {
        foreach ($parseError in $parseErrors) {
            [Console]::Error.WriteLine(("خطأ تحليل في {0}:{1}:{2}: {3}" -f $scriptPath, $parseError.Extent.StartLineNumber, $parseError.Extent.StartColumnNumber, $parseError.Message))
        }
        $failed = $true
    }
    elseif ($hasBom -and -not $hasDoubleBom) {
        Write-Host "نجح تحليل PowerShell: $scriptPath" -ForegroundColor Green
    }
}
if ($failed) { throw 'فشل فحص محلّل Windows PowerShell لبعض الملفات.' }
Write-Host ("نجح فحص محلّل Windows PowerShell لجميع ملفات PS1 ({0})." -f $scriptFiles.Count) -ForegroundColor Green
