[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][ValidateSet('Backup','Restore')][string]$Action,
    [Parameter(Mandatory=$true)][string]$ProjectFolder,
    [string]$BackupFolder
)
$ErrorActionPreference = 'Stop'
$ProjectFolder = [IO.Path]::GetFullPath($ProjectFolder)
$fe = Join-Path $ProjectFolder 'Archon_App_FE.accdb'
$be = Join-Path $ProjectFolder 'Archon_Data_BE.accdb'
$attachments = Join-Path $ProjectFolder 'Attachments'
$auditLog = Join-Path $ProjectFolder 'Archon_Audit_Fallback.log'
$stageRoot = $null
$restoreStarted = $false
$restoreCommitted = $false
$oldFeMoved = $false; $oldBeMoved = $false; $oldAttachmentsMoved = $false; $oldAuditMoved = $false
$newFePlaced = $false; $newBePlaced = $false; $newAttachmentsPlaced = $false; $newAuditPlaced = $false

function Test-PathWithin([string]$ChildPath, [string]$ParentPath) {
    $child = [IO.Path]::GetFullPath($ChildPath).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    $parent = [IO.Path]::GetFullPath($ParentPath).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    return $child.Equals($parent, [StringComparison]::OrdinalIgnoreCase) -or $child.StartsWith($parent + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)
}

function Assert-ProjectClosed {
    $lockFiles = @(Get-ChildItem -LiteralPath $ProjectFolder -Force -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -like '*.laccdb' -or $_.Name -like '*.ldb' })
    foreach ($lockName in @('Archon_App_FE.laccdb','Archon_Data_BE.laccdb')) {
        $lockPath = Join-Path $ProjectFolder $lockName
        if (Test-Path -LiteralPath $lockPath) { $lockFiles += Get-Item -LiteralPath $lockPath -Force }
    }
    if ($lockFiles.Count -gt 0) {
        $names = ($lockFiles | Select-Object -ExpandProperty Name -Unique) -join ', '
        throw "Access ما زال مفتوحًا أو توجد ملفات قفل ($names). أغلق النظام على جميع الأجهزة قبل النسخ أو الاستعادة."
    }
}

function Get-ManifestEntries([string]$Root) {
    $rootFull = [IO.Path]::GetFullPath($Root).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
    $items = @()
    foreach ($file in Get-ChildItem -LiteralPath $Root -File -Recurse -Force) {
        if ($file.Name -eq 'manifest.json' -and $file.DirectoryName -eq $Root) { continue }
        $relative = $file.FullName.Substring($rootFull.Length)
        $items += [pscustomobject]@{ Path=$relative; Length=[long]$file.Length; SHA256=(Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash }
    }
    return @($items | Sort-Object Path)
}

function Write-JsonUtf8([string]$Path, $Object) {
    $json = $Object | ConvertTo-Json -Depth 8
    [IO.File]::WriteAllText($Path, $json, [System.Text.UTF8Encoding]::new($true))
}

function Test-AccessDatabase([string]$Path) {
    $app = $null; $db = $null
    try {
        $app = New-Object -ComObject Access.Application
        $db = $app.DBEngine.Workspaces.Item(0).OpenDatabase($Path, $false, $true)
        if ($db.TableDefs.Count -lt 1) { throw "ملف Access لا يحتوي على تعريفات جداول: $Path" }
    }
    catch {
        throw "تعذر فتح قاعدة Access للتحقق منها. تأكد من تثبيت Access المكتبي وإغلاق النظام. الملف: '$Path'. التفاصيل: $($_.Exception.Message)"
    }
    finally {
        if ($db) { try { $db.Close() } catch {} }
        if ($app) { try { $app.Quit(2) } catch {}; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch {} }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function Test-BackupFolder([string]$Root) {
    $manifestPath = Join-Path $Root 'manifest.json'
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw 'ملف manifest.json غير موجود في النسخة الاحتياطية.' }
    try { $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json }
    catch { throw "ملف manifest.json غير صالح: $($_.Exception.Message)" }
    if ($manifest.Product -ne 'Archon Tech Access System') { throw 'النسخة لا تخص Archon Tech Access System.' }
    if ($null -eq $manifest.Files -or @($manifest.Files).Count -lt 2) { throw 'بيان النسخة لا يحتوي قائمة ملفات كاملة.' }

    $rootFull = [IO.Path]::GetFullPath($Root).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
    $declared = @{}
    foreach ($entry in $manifest.Files) {
        $relative = [string]$entry.Path
        if ([string]::IsNullOrWhiteSpace($relative)) { throw 'يوجد مسار فارغ في بيان النسخة.' }
        if ([IO.Path]::IsPathRooted($relative)) { throw "مسار مطلق غير مسموح في البيان: $relative" }
        $candidate = [IO.Path]::GetFullPath((Join-Path $Root $relative))
        if (-not $candidate.StartsWith($rootFull, [StringComparison]::OrdinalIgnoreCase)) { throw "مسار خارج مجلد النسخة الاحتياطية: $relative" }
        if ($declared.ContainsKey($relative)) { throw "مسار مكرر في البيان: $relative" }
        $declared[$relative] = $true
        if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) { throw "ملف مفقود من النسخة الاحتياطية: $relative" }
        $fileInfo = Get-Item -LiteralPath $candidate
        if ([long]$fileInfo.Length -ne [long]$entry.Length) { throw "حجم الملف لا يطابق البيان: $relative" }
        $hash = (Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash
        if ($hash -ne [string]$entry.SHA256) { throw "فشل تحقق SHA256 للملف: $relative" }
    }
    foreach ($required in @('Database\Archon_App_FE.accdb','Database\Archon_Data_BE.accdb')) {
        if (-not $declared.ContainsKey($required)) { throw "قاعدة بيانات أساسية غير مدرجة في البيان: $required" }
    }
    $actual = @(Get-ChildItem -LiteralPath $Root -File -Recurse -Force | Where-Object { $_.FullName -ne $manifestPath } | ForEach-Object { $_.FullName.Substring($rootFull.Length) } | Sort-Object)
    $listed = @($declared.Keys | Sort-Object)
    if (($actual -join '|') -ne ($listed -join '|')) { throw 'توجد ملفات غير مدرجة أو مفقودة من بيان النسخة الاحتياطية.' }
    if ($manifest.SchemaVersion -ge 2 -and -not (Test-Path -LiteralPath (Join-Path $Root 'Attachments') -PathType Container)) { throw 'مجلد المرفقات المطلوب في النسخة غير موجود.' }
    return $manifest
}

if (-not (Test-Path -LiteralPath $ProjectFolder -PathType Container)) { throw "مجلد المشروع غير موجود: $ProjectFolder" }
if (-not (Test-Path -LiteralPath $fe -PathType Leaf) -or -not (Test-Path -LiteralPath $be -PathType Leaf)) {
    throw 'لم يتم العثور على ملفي الواجهة والبيانات بالمسميات المتوقعة في مجلد المشروع.'
}
Assert-ProjectClosed

if ($Action -eq 'Backup') {
    if ([string]::IsNullOrWhiteSpace($BackupFolder)) {
        Add-Type -AssemblyName System.Windows.Forms
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = 'اختر مكان حفظ النسخة الاحتياطية خارج مجلد تشغيل النظام'
        $dialog.ShowNewFolderButton = $true
        if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { throw 'تم إلغاء النسخ الاحتياطي.' }
        $BackupFolder = $dialog.SelectedPath
    }
    $backupBase = [IO.Path]::GetFullPath($BackupFolder)
    if (Test-PathWithin $backupBase $ProjectFolder) { throw 'يجب حفظ النسخة الاحتياطية خارج مجلد المشروع لتجنب نسخ النسخ داخل بعضها.' }
    New-Item -ItemType Directory -Force -Path $backupBase | Out-Null
    Test-AccessDatabase $fe
    Test-AccessDatabase $be

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $target = Join-Path $backupBase "Archon_Backup_$stamp"
    if (Test-Path -LiteralPath $target) { $target += '_' + [Guid]::NewGuid().ToString('N').Substring(0,8) }
    $partial = $target + '.partial'
    try {
        New-Item -ItemType Directory -Path $partial -ErrorAction Stop | Out-Null
        $databaseFolder = Join-Path $partial 'Database'
        $attachmentFolder = Join-Path $partial 'Attachments'
        $logFolder = Join-Path $partial 'Logs'
        New-Item -ItemType Directory -Path $databaseFolder | Out-Null
        New-Item -ItemType Directory -Path $attachmentFolder | Out-Null
        New-Item -ItemType Directory -Path $logFolder | Out-Null
        Copy-Item -LiteralPath $fe -Destination (Join-Path $databaseFolder 'Archon_App_FE.accdb')
        Copy-Item -LiteralPath $be -Destination (Join-Path $databaseFolder 'Archon_Data_BE.accdb')
        if (Test-Path -LiteralPath $attachments -PathType Container) {
            foreach ($item in Get-ChildItem -LiteralPath $attachments -Force) { Copy-Item -LiteralPath $item.FullName -Destination $attachmentFolder -Recurse -Force }
        }
        if (Test-Path -LiteralPath $auditLog -PathType Leaf) { Copy-Item -LiteralPath $auditLog -Destination $logFolder }
        Test-AccessDatabase (Join-Path $databaseFolder 'Archon_App_FE.accdb')
        Test-AccessDatabase (Join-Path $databaseFolder 'Archon_Data_BE.accdb')
        $manifest = [pscustomobject]@{
            Product='Archon Tech Access System'; SchemaVersion=2
            BackupCreatedAt=(Get-Date).ToString('o'); SourceProjectFolder=$ProjectFolder
            AttachmentsIncluded=$true; AuditFallbackLogIncluded=(Test-Path -LiteralPath (Join-Path $logFolder 'Archon_Audit_Fallback.log') -PathType Leaf); Files=@(Get-ManifestEntries $partial)
        }
        Write-JsonUtf8 (Join-Path $partial 'manifest.json') $manifest
        $null = Test-BackupFolder $partial
        Move-Item -LiteralPath $partial -Destination $target
        Write-Host "اكتملت النسخة الاحتياطية وتحققنا من ملفاتها: $target" -ForegroundColor Green
        Write-Host 'تتضمن النسخة قاعدتي Access والمرفقات وسجل التدقيق الاحتياطي إن وجد وبيان SHA-256.'
    }
    catch {
        if (Test-Path -LiteralPath $partial) { Remove-Item -LiteralPath $partial -Recurse -Force -ErrorAction SilentlyContinue }
        throw "فشل النسخ الاحتياطي ولم يتم اعتماد مجلد جزئي كنسخة سليمة. التفاصيل: $($_.Exception.Message)"
    }
}
else {
    if ([string]::IsNullOrWhiteSpace($BackupFolder)) { throw 'يجب تحديد مجلد النسخة الاحتياطية في وضع Restore.' }
    $BackupFolder = [IO.Path]::GetFullPath($BackupFolder)
    if (Test-PathWithin $BackupFolder $ProjectFolder) { throw 'يجب أن تكون النسخة الاحتياطية خارج مجلد المشروع.' }
    $manifest = Test-BackupFolder $BackupFolder
    if (-not [string]::IsNullOrWhiteSpace([string]$manifest.SourceProjectFolder)) {
        $recordedRoot = [IO.Path]::GetFullPath([string]$manifest.SourceProjectFolder).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
        $currentRoot = [IO.Path]::GetFullPath($ProjectFolder).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
        if (-not $recordedRoot.Equals($currentRoot, [StringComparison]::OrdinalIgnoreCase)) {
            throw "هذه النسخة مرتبطة بمسار المشروع '$recordedRoot' بينما المسار الحالي '$currentRoot'. أوقفنا الاستعادة لمنع FE من الاحتفاظ بروابط إلى مسار خلفية خاطئ. أعد المشروع إلى مساره الأصلي أو نفّذ إعادة ربط مدروسة في Access بعد التحقق من النسخة."
        }
    }
    $savedFE = Join-Path $BackupFolder 'Database\Archon_App_FE.accdb'
    $savedBE = Join-Path $BackupFolder 'Database\Archon_Data_BE.accdb'
    $savedAttachments = Join-Path $BackupFolder 'Attachments'
    $savedAuditLog = Join-Path $BackupFolder 'Logs\Archon_Audit_Fallback.log'
    if ((Test-Path -LiteralPath $auditLog -PathType Leaf) -and -not (Test-Path -LiteralPath $savedAuditLog -PathType Leaf)) { throw 'النسخة الاحتياطية لا تحتوي سجل التدقيق الاحتياطي الحالي. أُوقفت الاستعادة لمنع فقد أحداث التدقيق المحلية.' }
    if (-not (Test-Path -LiteralPath $savedAttachments -PathType Container)) {
        $currentAttachmentFiles = @()
        if (Test-Path -LiteralPath $attachments -PathType Container) { $currentAttachmentFiles = @(Get-ChildItem -LiteralPath $attachments -File -Recurse -Force) }
        if ($currentAttachmentFiles.Count -gt 0) { throw 'هذه النسخة القديمة لا تحتوي مجلد المرفقات، والمجلد الحالي يحتوي مرفقات. أُوقفت الاستعادة لمنع فقد الملفات.' }
    }
    Test-AccessDatabase $savedFE
    Test-AccessDatabase $savedBE
    $confirm = Read-Host 'ستستبدل ملفات النظام الحالية. اكتب RESTORE للموافقة'
    if ($confirm -cne 'RESTORE') { throw 'تم إلغاء الاستعادة.' }

    $guid = [Guid]::NewGuid().ToString('N')
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $stageRoot = Join-Path $ProjectFolder ('.ArchonRestore_' + $guid)
    $incoming = Join-Path $stageRoot 'Incoming'
    $rollback = Join-Path $stageRoot 'Rollback'
    $safety = Join-Path $ProjectFolder ("PreRestore_{0}_{1}" -f $stamp, $guid.Substring(0,8))
    New-Item -ItemType Directory -Path $incoming | Out-Null
    New-Item -ItemType Directory -Path $rollback | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $incoming 'Database') | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $incoming 'Attachments') | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $incoming 'Logs') | Out-Null

    try {
        Copy-Item -LiteralPath $savedFE -Destination (Join-Path $incoming 'Database\Archon_App_FE.accdb')
        Copy-Item -LiteralPath $savedBE -Destination (Join-Path $incoming 'Database\Archon_Data_BE.accdb')
        if (Test-Path -LiteralPath $savedAttachments -PathType Container) { foreach ($item in Get-ChildItem -LiteralPath $savedAttachments -Force) { Copy-Item -LiteralPath $item.FullName -Destination (Join-Path $incoming 'Attachments') -Recurse -Force } }
        if (Test-Path -LiteralPath $savedAuditLog -PathType Leaf) { Copy-Item -LiteralPath $savedAuditLog -Destination (Join-Path $incoming 'Logs\Archon_Audit_Fallback.log') }
        Copy-Item -LiteralPath (Join-Path $BackupFolder 'manifest.json') -Destination (Join-Path $incoming 'manifest.json')
        $null = Test-BackupFolder $incoming
        Test-AccessDatabase (Join-Path $incoming 'Database\Archon_App_FE.accdb')
        Test-AccessDatabase (Join-Path $incoming 'Database\Archon_Data_BE.accdb')

        New-Item -ItemType Directory -Path (Join-Path $safety 'Database') -Force | Out-Null
        Copy-Item -LiteralPath $fe -Destination (Join-Path $safety 'Database\Archon_App_FE.accdb')
        Copy-Item -LiteralPath $be -Destination (Join-Path $safety 'Database\Archon_Data_BE.accdb')
        New-Item -ItemType Directory -Path (Join-Path $safety 'Attachments') -Force | Out-Null
        if (Test-Path -LiteralPath $attachments -PathType Container) {
            foreach ($item in Get-ChildItem -LiteralPath $attachments -Force) { Copy-Item -LiteralPath $item.FullName -Destination (Join-Path $safety 'Attachments') -Recurse -Force }
        }
        New-Item -ItemType Directory -Path (Join-Path $safety 'Logs') -Force | Out-Null
        if (Test-Path -LiteralPath $auditLog -PathType Leaf) { Copy-Item -LiteralPath $auditLog -Destination (Join-Path $safety 'Logs\Archon_Audit_Fallback.log') }
        $safetyManifest = [pscustomobject]@{ Product='Archon Tech Access System'; SchemaVersion=2; BackupCreatedAt=(Get-Date).ToString('o'); Files=@(Get-ManifestEntries $safety) }
        Write-JsonUtf8 (Join-Path $safety 'manifest.json') $safetyManifest
        $null = Test-BackupFolder $safety

        $restoreStarted = $true
        Move-Item -LiteralPath $fe -Destination (Join-Path $rollback 'Archon_App_FE.accdb'); $oldFeMoved = $true
        Move-Item -LiteralPath $be -Destination (Join-Path $rollback 'Archon_Data_BE.accdb'); $oldBeMoved = $true
        if (Test-Path -LiteralPath $attachments -PathType Container) { Move-Item -LiteralPath $attachments -Destination (Join-Path $rollback 'Attachments'); $oldAttachmentsMoved = $true }
        if (Test-Path -LiteralPath $auditLog -PathType Leaf) { Move-Item -LiteralPath $auditLog -Destination (Join-Path $rollback 'Archon_Audit_Fallback.log'); $oldAuditMoved = $true }
        Move-Item -LiteralPath (Join-Path $incoming 'Database\Archon_App_FE.accdb') -Destination $fe; $newFePlaced = $true
        Move-Item -LiteralPath (Join-Path $incoming 'Database\Archon_Data_BE.accdb') -Destination $be; $newBePlaced = $true
        Move-Item -LiteralPath (Join-Path $incoming 'Attachments') -Destination $attachments; $newAttachmentsPlaced = $true
        $incomingAuditLog = Join-Path $incoming 'Logs\Archon_Audit_Fallback.log'
        if (Test-Path -LiteralPath $incomingAuditLog -PathType Leaf) { Move-Item -LiteralPath $incomingAuditLog -Destination $auditLog; $newAuditPlaced = $true }
        Test-AccessDatabase $fe
        Test-AccessDatabase $be
        $restoreCommitted = $true
        try {
            Remove-Item -LiteralPath $stageRoot -Recurse -Force -ErrorAction Stop
            $stageRoot = $null
        }
        catch {
            Write-Warning "اكتملت الاستعادة والتحقق، لكن تعذر تنظيف مجلد الاستعادة المؤقت. احتفظ به وراجعه: $stageRoot. التفاصيل: $($_.Exception.Message)"
        }
        Write-Host "اكتملت الاستعادة والتحقق من قواعد البيانات. نسخة الحالة السابقة محفوظة هنا: $safety" -ForegroundColor Green
        Write-Warning 'أعد فتح Access ونفّذ فحوص البيانات والروابط واختبارات القبول قبل استئناف العمل.'
    }
    catch {
        $restoreError = $_.Exception.Message
        if ($restoreStarted -and -not $restoreCommitted) {
            $rollbackErrors = @()
            $oldFE = Join-Path $rollback 'Archon_App_FE.accdb'
            $oldBE = Join-Path $rollback 'Archon_Data_BE.accdb'
            $oldAttachments = Join-Path $rollback 'Attachments'
            $oldAudit = Join-Path $rollback 'Archon_Audit_Fallback.log'
            if ($newFePlaced -and (Test-Path -LiteralPath $fe -PathType Leaf)) { try { Remove-Item -LiteralPath $fe -Force } catch { $rollbackErrors += $_.Exception.Message } }
            if ($newBePlaced -and (Test-Path -LiteralPath $be -PathType Leaf)) { try { Remove-Item -LiteralPath $be -Force } catch { $rollbackErrors += $_.Exception.Message } }
            if ($newAttachmentsPlaced -and (Test-Path -LiteralPath $attachments -PathType Container)) { try { Remove-Item -LiteralPath $attachments -Recurse -Force } catch { $rollbackErrors += $_.Exception.Message } }
            if ($newAuditPlaced -and (Test-Path -LiteralPath $auditLog -PathType Leaf)) { try { Remove-Item -LiteralPath $auditLog -Force } catch { $rollbackErrors += $_.Exception.Message } }
            if ($oldFeMoved -and (Test-Path -LiteralPath $oldFE -PathType Leaf)) { try { Move-Item -LiteralPath $oldFE -Destination $fe } catch { $rollbackErrors += $_.Exception.Message } }
            if ($oldBeMoved -and (Test-Path -LiteralPath $oldBE -PathType Leaf)) { try { Move-Item -LiteralPath $oldBE -Destination $be } catch { $rollbackErrors += $_.Exception.Message } }
            if ($oldAttachmentsMoved -and (Test-Path -LiteralPath $oldAttachments -PathType Container)) { try { Move-Item -LiteralPath $oldAttachments -Destination $attachments } catch { $rollbackErrors += $_.Exception.Message } }
            if ($oldAuditMoved -and (Test-Path -LiteralPath $oldAudit -PathType Leaf)) { try { Move-Item -LiteralPath $oldAudit -Destination $auditLog } catch { $rollbackErrors += $_.Exception.Message } }
            if ($rollbackErrors.Count -gt 0) {
                throw "فشلت الاستعادة: $restoreError. وتعذر إكمال التراجع تلقائيًا: $($rollbackErrors -join ' | '). لا تحذف مجلد الاستعادة المؤقت '$stageRoot' أو نسخة الأمان '$safety'."
            }
        }
        throw "فشلت الاستعادة وتم إرجاع الملفات الحية إلى حالتها السابقة قدر الإمكان. السبب: $restoreError. نسخة الأمان: $safety"
    }
    finally {
        if ($stageRoot -and (Test-Path -LiteralPath $stageRoot) -and $restoreCommitted) { Remove-Item -LiteralPath $stageRoot -Recurse -Force -ErrorAction SilentlyContinue }
    }
}
