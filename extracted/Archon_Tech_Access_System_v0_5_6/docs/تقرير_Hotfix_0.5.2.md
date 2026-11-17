# تقرير Hotfix والإصلاح المحدد — Archon Tech Access System 0.5.2

**النسخة الأم:** `Archon_Tech_Access_System_Source_v0_5_1.zip`  
**النسخة الناتجة:** `Archon_Tech_Access_System_Source_v0_5_2.zip`  
**النطاق:** إصلاح خطأ PowerShell COM الذي ظهر للمستخدم بعد استيراد الوحدات الاثنتي عشرة، مع تصحيح حفظ كائنات Access عند النجاح. لم يُعاد بناء النظام ولم يتغير مخطط البيانات أو كود VBA.

## 1. التشخيص من الملف والنص الفعلي

فُتح `tools/Install_Archon_Tech.ps1` من أرشيف 0.5.1 نفسه، وظهر في السطر 94:

```powershell
$installResult = $access.Run('AT_InstallSplitProject_Run', $bePath)
```

هذا هو الاستدعاء المطابق للخطأ. `Access.Application.Run` يأخذ اسم الإجراء كوسيط أول ثم وسيطات الإجراء. عند استخدام PowerShell مع Office COM، يجب تمرير وسيط الماكرو بهذه الحالة كمرجع PowerShell (`[ref]`)؛ تمرير المسار كسلسلة عادية ينتج `Argument: '2' should be a System.Management.Automation.PSReference. Use [ref].`، لأن المسار هو الوسيط الثاني في نداء Run. توثيق Access يبيّن أن Arg1 هو وسيط الإجراء الأول، ومثال PowerShell لأتمتة Office يوضح الخطأ نفسه وطريقة `[ref]`.

## 2. التغيير المنفذ

استبدل الاستدعاء بما يلي:

```powershell
$backendPathArgument = [ref]$bePath
$installResult = $access.Run('AT_InstallSplitProject_Run', $backendPathArgument)
```

تم التحقق ساكنًا من وجود دالة VBA المطابقة `AT_InstallSplitProject_Run(ByVal backendPath As String) As Boolean` في `src/vba/modArchonInstaller.bas`. لا يلزم تعديل VBA لأن وسيط الدالة نفسه String ByVal؛ الإصلاح في تمرير الوسيط عبر طبقة Automation.

## 3. خطأ حفظ إضافي عُثر عليه في الملف نفسه

كان مسار النجاح يستخدم `$access.Quit(2)`. حسب تعداد Access `AcQuitOption`، القيمة 2 هي `acQuitSaveNone` (الخروج دون حفظ)، بينما القيمة 1 هي `acQuitSaveAll`. لذلك غُيّر مسار النجاح إلى `$access.Quit(1)`. احتُفظ بـ `Quit(2)` في مسار الفشل فقط، حيث يجري التخلص عمدًا من البناء الجزئي قبل إزالة ملفات FE/BE الجزئية.

## 4. الملفات المعدلة

- `tools/Install_Archon_Tech.ps1`: تمرير Arg1 عبر `[ref]`، وحفظ كائنات Access عند النجاح.
- `tests/verify_package.py`: تعديل اختبار نقطة الدخول وإضافة اختبارين يمنعان رجوع `[ref]` الخاطئ وخيار الحفظ الخاطئ.
- `VERSION.txt`, `README_AR.md`, `installer/تسلسل_الاستيراد.txt`: تحديث إصدار المصدر إلى 0.5.2.
- `docs/دليل_التركيب_والاختبار.md`, `docs/حالة_التنفيذ_والقيود.md`, `docs/سجل_التغييرات.md`, `release/RELEASE_NOTES_AR.md`: توثيق التصحيح وحدود التحقق.
- `docs/تقرير_Hotfix_0.5.2.md`: هذا التقرير داخل الحزمة.

## 5. ما لم يتغير

- جميع وحدات VBA الاثنتي عشرة كما في 0.5.1.
- مخطط 55 جدولًا و444 حقلًا و86 علاقة و43 فهرسًا.
- منطق الخدمات المالية وواجهات النظام والتقارير.
- سكربت النسخ والاستعادة وبقية أدوات البناء.

## 6. نتائج الاختبار وحدودها

تمت إضافة اختبارين ساكنين جديدين للتحقق من صيغة استدعاء COM وخيار Quit. يجب تشغيل `python tests/verify_package.py` على النسخة النهائية، إضافة إلى `tests/Verify-Windows-PowerShell.ps1` على Windows PowerShell 5.1. لم تُشغّل بيئة Windows أو Access COM أو DAO/ACE في بيئة الإنشاء الحالية؛ لذلك لا ندّعي أن التثبيت نجح بعد هذا التصحيح.

## 7. توجيه للمستخدم

لا تعِد تشغيل مجلد 0.5.1. استخدم حزمة 0.5.2 في مجلد استخراج جديد، واختر مجلد تثبيت جديدًا فارغًا؛ لا تحاول الاستمرار في ملفات التثبيت الجزئية التي أنشأتها محاولة 0.5.1. أبقِ هذه الملفات محفوظة حتى التحقق من أنها ليست بيانات مستخدم، ثم يمكن تنظيفها يدويًا بعد نجاح التثبيت الجديد.

**الحكم:** عولج السبب الذي يطابق رسالة الخطأ مباشرةً في المصدر. الإصدار 0.5.2 Hotfix مرشح لإعادة الاختبار على Windows؛ لم يُعتمد بعد كتثبيت ناجح أو كنظام إنتاجي.
