[CmdletBinding()]
param([string]$InstallFolder)
$ErrorActionPreference = 'Stop'
$step = 'التحقق من ملفات المصدر'
$access = $null; $project = $null; $dbEngine = $null; $workspace = $null
$createdInstallFolder = $false; $fePath = $null; $bePath = $null; $bootstrapConfigPath = $null

function Select-Folder([string]$Description, [string]$InitialPath) {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Description
    $dialog.ShowNewFolderButton = $true
    if ($InitialPath -and (Test-Path -LiteralPath $InitialPath)) { $dialog.SelectedPath = $InitialPath }
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { throw 'تم إلغاء اختيار المجلد.' }
    return $dialog.SelectedPath
}

function Test-PathWithin([string]$ChildPath, [string]$ParentPath) {
    $child = [IO.Path]::GetFullPath($ChildPath).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    $parent = [IO.Path]::GetFullPath($ParentPath).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    return $child.Equals($parent, [StringComparison]::OrdinalIgnoreCase) -or $child.StartsWith($parent + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)
}

$sourceRoot = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
$vbaRoot = Join-Path $sourceRoot 'src\vba'
if (-not (Test-Path -LiteralPath $vbaRoot -PathType Container)) { throw "لم يتم العثور على مجلد الوحدات: $vbaRoot" }
$modules = @(
    'modArchonCore.bas','modArchonBuild.bas','modArchonFinance.bas','modArchonProperty.bas',
    'modArchonClosing.bas','modArchonOperations.bas','modArchonConfig.bas','modArchonAttachments.bas',
    'modArchonReports.bas','modArchonPrint.bas','modArchonUI.bas','modArchonInstaller.bas'
)
foreach ($module in $modules) {
    $modulePath = Join-Path $vbaRoot $module
    if (-not (Test-Path -LiteralPath $modulePath -PathType Leaf)) { throw "ملف وحدة مفقود: $modulePath" }
}

if ([string]::IsNullOrWhiteSpace($InstallFolder)) {
    $InstallFolder = Select-Folder 'اختر مجلدًا جديدًا وفارغًا لتثبيت نظام أرشون تك' ([Environment]::GetFolderPath('Desktop'))
}
$InstallFolder = [IO.Path]::GetFullPath($InstallFolder)
if (Test-PathWithin $InstallFolder $sourceRoot) { throw 'اختر مجلد تثبيت مستقلًا خارج مجلد ملفات الحزمة.' }
if (Test-Path -LiteralPath $InstallFolder) {
    if (-not (Test-Path -LiteralPath $InstallFolder -PathType Container)) { throw 'مسار التثبيت موجود لكنه ليس مجلدًا.' }
    if (@(Get-ChildItem -LiteralPath $InstallFolder -Force -ErrorAction Stop).Count -gt 0) {
        throw 'مجلد التثبيت ليس فارغًا. اختر مجلدًا جديدًا أو فارغًا حتى لا تتأثر ملفات موجودة.'
    }
}
else { New-Item -ItemType Directory -Path $InstallFolder -ErrorAction Stop | Out-Null; $createdInstallFolder = $true }

$fePath = Join-Path $InstallFolder 'Archon_App_FE.accdb'
$bePath = Join-Path $InstallFolder 'Archon_Data_BE.accdb'
$bootstrapConfigPath = Join-Path $InstallFolder '.archon_backend_path.txt'
$attachmentsPath = Join-Path $InstallFolder 'Attachments'
$backupsPath = Join-Path $InstallFolder 'Backups'
$writeProbe = Join-Path $InstallFolder ('.archon_write_test_' + [Guid]::NewGuid().ToString('N'))
try { [IO.File]::WriteAllText($writeProbe, 'write-test'); Remove-Item -LiteralPath $writeProbe -Force }
catch { throw "لا توجد صلاحية كتابة في مجلد التثبيت '$InstallFolder': $($_.Exception.Message)" }

try {
    $step = 'فحص Microsoft Access وبيئة DAO/COM'
    try { $access = New-Object -ComObject Access.Application }
    catch { throw 'تعذر إنشاء Microsoft Access.Application عبر COM. ثبّت Microsoft Access المكتبي (ليس Runtime فقط). التفاصيل: ' + $_.Exception.Message }
    $accessVersion = [string]$access.Version
    $versionMajor = 0
    if (-not [int]::TryParse(($accessVersion -split '\.')[0], [ref]$versionMajor) -or $versionMajor -lt 16) {
        throw "إصدار Access غير معتمد. الإصدار المكتشف: $accessVersion. يلزم Access 2021 أو 2024 أو إصدار مكتبي أحدث متوافق."
    }
    $access.Visible = $true
    $dbEngine = $access.DBEngine
    $workspace = $dbEngine.Workspaces.Item(0)
    if ($null -eq $workspace) { throw 'لم يتمكن Access من توفير DAO Workspace.' }

    $step = 'إنشاء ملف الواجهة الأمامية'
    $access.NewCurrentDatabase($fePath)
    if (-not (Test-Path -LiteralPath $fePath -PathType Leaf)) { throw 'لم ينشأ ملف الواجهة الأمامية بعد أمر NewCurrentDatabase.' }

    $step = 'التحقق من السماح بالوصول إلى مشروع VBA'
    try {
        $project = $access.VBE.ActiveVBProject
        if ($null -eq $project) { throw 'ActiveVBProject غير متاح.' }
    }
    catch {
        throw 'منع مركز التوثيق الوصول إلى مشروع VBA أو لم يتمكن Access من فتح VBProject. فعّل خيار الوثوق بالوصول إلى نموذج كائن مشروع VBA. التفاصيل: ' + $_.Exception.Message
    }

    $step = 'استيراد وحدات VBA'
    foreach ($module in $modules) {
        $modulePath = Join-Path $vbaRoot $module
        Write-Host "استيراد $module ..."
        [void]$project.VBComponents.Import($modulePath)
    }

    $step = 'بناء قاعدة البيانات الخلفية والجداول والعلاقات والواجهات'
    Write-Host 'تشغيل أداة إنشاء قاعدة البيانات والواجهات ...'
    # Avoid Access.Application.Run argument marshaling entirely: pass the backend path
    # through a temporary UTF-16 config file that the VBA entry point reads itself.
    [IO.File]::WriteAllText($bootstrapConfigPath, $bePath, [Text.Encoding]::Unicode)
    $installResult = $access.Run('AT_InstallSplitProject_Run')
    if (-not [bool]$installResult) { throw 'لم تؤكد نقطة الدخول البرمجية اكتمال البناء.' }
    if (Test-Path -LiteralPath $bootstrapConfigPath -PathType Leaf) { throw 'لم تحذف نقطة دخول VBA ملف إعداد المسار المؤقت؛ تم إيقاف التثبيت للتحقق.' }

    $step = 'التحقق من ملفات البناء الناتجة'
    if (-not (Test-Path -LiteralPath $fePath -PathType Leaf)) { throw 'ملف الواجهة الأمامية غير موجود بعد البناء.' }
    if (-not (Test-Path -LiteralPath $bePath -PathType Leaf)) { throw 'ملف قاعدة البيانات الخلفية غير موجود بعد البناء.' }
    if ((Get-Item -LiteralPath $fePath).Length -le 0 -or (Get-Item -LiteralPath $bePath).Length -le 0) { throw 'أحد ملفي Access الناتجين فارغ.' }

    $access.Quit(1) # acQuitSaveAll: persist Access objects created during the build.
    Write-Host ''
    Write-Host 'اكتمل إنشاء ملفات النظام بنجاح.' -ForegroundColor Green
    Write-Host "إصدار Access المكتشف: $accessVersion"
    Write-Host "ملف الواجهة: $fePath"
    Write-Host "قاعدة البيانات الخلفية: $bePath"
    Write-Host "مجلد المرفقات: $attachmentsPath"
    Write-Host "مجلد النسخ الاحتياطية: $backupsPath"
    Write-Host 'الخطوة التالية: افتح الواجهة ونفّذ Debug > Compile ثم اختبارات القبول. لا تستخدم بيانات إنتاج قبل نجاحها.'
}
catch {
    $failure = $_.Exception.Message
    Write-Host "فشل التثبيت في المرحلة: $step" -ForegroundColor Red
    Write-Host "التفاصيل: $failure" -ForegroundColor Red
    if ($access) { try { $access.CloseCurrentDatabase() } catch {}; try { $access.Quit(2) } catch {} }
    if ($bootstrapConfigPath -and (Test-Path -LiteralPath $bootstrapConfigPath -PathType Leaf)) {
        try { Remove-Item -LiteralPath $bootstrapConfigPath -Force -ErrorAction Stop }
        catch { Write-Warning "تعذر حذف ملف إعداد المثبّت المؤقت '$bootstrapConfigPath': $($_.Exception.Message)" }
    }
    foreach ($partial in @($fePath, $bePath)) {
        if ($partial -and (Test-Path -LiteralPath $partial -PathType Leaf)) {
            try { Remove-Item -LiteralPath $partial -Force -ErrorAction Stop }
            catch { Write-Warning "تعذر حذف الملف الجزئي '$partial': $($_.Exception.Message)" }
        }
    }
    foreach ($partialDirectory in @($attachmentsPath, $backupsPath)) {
        if ($partialDirectory -and (Test-Path -LiteralPath $partialDirectory -PathType Container)) {
            try {
                if (@(Get-ChildItem -LiteralPath $partialDirectory -Force).Count -eq 0) { Remove-Item -LiteralPath $partialDirectory -Force }
                else { Write-Warning "تم الإبقاء على المجلد '$partialDirectory' لأنه يحتوي ملفات؛ راجعه يدويًا." }
            } catch { Write-Warning "تعذر تنظيف '$partialDirectory': $($_.Exception.Message)" }
        }
    }
    if ($createdInstallFolder -and (Test-Path -LiteralPath $InstallFolder -PathType Container)) {
        try { if (@(Get-ChildItem -LiteralPath $InstallFolder -Force).Count -eq 0) { Remove-Item -LiteralPath $InstallFolder -Force } } catch {}
    }
    throw "فشل تثبيت Archon Tech في المرحلة '$step'. لم يُعلن اكتمال التثبيت. السبب: $failure"
}
finally {
    if ($project) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($project) } catch {} }
    if ($workspace) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workspace) } catch {} }
    if ($dbEngine) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($dbEngine) } catch {} }
    if ($access) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($access) } catch {} }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
