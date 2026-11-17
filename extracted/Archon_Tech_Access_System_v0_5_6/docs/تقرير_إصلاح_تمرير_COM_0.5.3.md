# تقرير إصلاح تمرير COM — Archon Tech Access System 0.5.3

## سبب الإصلاح

أظهر اختبار المستخدم أن تمرير `[ref]$bePath` إلى `Access.Application.Run` تجاوز خطأ PowerShell السابق لكنه فشل داخل COM برسالة `DISP_E_TYPEMISMATCH`. توقيع VBA السابق كان يستقبل `ByVal backendPath As String`، ولم يثبت توافق هذا التمرير عبر COM على جهاز المستخدم.

## التغيير

- أزيل تمرير أي وسيط إلى `Access.Application.Run`.
- يكتب PowerShell مسار `Archon_Data_BE.accdb` في ملف `.archon_backend_path.txt` بترميز UTF-16LE/Unicode.
- تستدعي PowerShell `AT_InstallSplitProject_Run` بلا وسيطات.
- تقرأ دالة VBA المسار باستخدام `Scripting.FileSystemObject.OpenTextFile(..., ForReading, ..., TristateTrue)`، وتتحقق من وجوده، ثم تحذف ملف الإعداد وتنفذ إجراء البناء الأصلي.
- عند فشل المثبّت قبل قراءة ملف الإعداد، يحاول PowerShell حذف الملف المؤقت.
- أضيف ملف `تشغيل_المثبت_وفحصه.cmd` لتشغيل فحص PowerShell ثم المثبّت بالنقر المزدوج، دون الحاجة إلى نسخ أوامر يدويًا.
- لم يتغير مخطط الجداول أو الخدمات أو الواجهات أو التقارير.

## التحقق وحدوده

اجتاز المصدر 45 اختبارًا ساكنًا/حسابيًا في `tests/verify_package.py`. ويشغّل ملف CMD فاحص PowerShell على Windows قبل بدء المثبّت. بيئة الإنشاء الحالية لا تحتوي Microsoft Access أو Windows COM أو DAO/ACE؛ لذلك لا يمكن إثبات نجاح COM/Access runtime هنا. يجب اعتبار 0.5.3 مرشح اختبار ميداني، وليس إصدارًا مثبتًا أو إنتاجيًا، إلى أن ينجح التثبيت الفعلي ثم Debug > Compile واختبارات القبول.
