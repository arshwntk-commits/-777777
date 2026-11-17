Attribute VB_Name = "modArchonUI"
Option Compare Database
Option Explicit

' Generates the first usable desktop navigation and transaction-entry forms.
' This is a form-builder module, not a substitute for a compiled/tested Access front end.

Public Sub AT_BuildInitialUI()
    On Error GoTo EH
    If Not AT_BuildReportQueries() Then Err.Raise vbObjectError + 690, "AT_BuildInitialUI", "فشل إنشاء استعلامات التقارير؛ أُوقف إنشاء الواجهات التابعة لها."
    AT_CreateMainForm
    AT_CreateOwnerEntryForm
    AT_CreateOwnerEditForm
    AT_CreateOwnerSearchForm
    AT_CreateOwnerStatementPicker
    AT_CreateAccountEntryForm
    AT_CreateChargeEntryForm
    AT_CreateReceiptEntryForm
    AT_CreatePaymentEntryForm
    AT_CreateTransferEntryForm
    AT_CreateAdjustmentEntryForm
    AT_CreateCreditApplicationForm
    AT_CreateReceiptCorrectionForm
    AT_CreateReadOnlyGridForm "frmOwnerDirectory", "سجل الملاك", "SELECT OwnerCode, FullName, Phone, HouseCount, ApartmentCount, AdminStatus, DataStatus, ReviewStatus FROM tblOwners ORDER BY FullName;"
    AT_CreateReadOnlyGridForm "frmReceiptList", "سجل الإيصالات", "SELECT N.OfficialNumber, R.ReceiptID, R.PrintedOwnerCode, R.PrintedOwnerName, D.TransactionDate, R.Amount, PM.MethodName, A.AccountName, D.StatusCode FROM (((tblReceipts AS R INNER JOIN tblFinancialDocuments AS D ON R.DocumentID=D.DocumentID) LEFT JOIN tblOfficialDocumentNumbers AS N ON D.DocumentID=N.DocumentID) LEFT JOIN tblPaymentMethods AS PM ON R.PaymentMethodID=PM.PaymentMethodID) LEFT JOIN tblAccounts AS A ON R.AccountID=A.AccountID ORDER BY D.TransactionDate DESC, R.ReceiptID DESC;"
    AT_CreateReadOnlyGridForm "frmPaymentList", "سجل سندات الصرف", "SELECT N.OfficialNumber, P.PaymentID, P.BeneficiaryName, D.TransactionDate, P.Amount, A.AccountName, D.StatusCode FROM ((tblPayments AS P INNER JOIN tblFinancialDocuments AS D ON P.DocumentID=D.DocumentID) LEFT JOIN tblOfficialDocumentNumbers AS N ON D.DocumentID=N.DocumentID) LEFT JOIN tblAccounts AS A ON P.AccountID=A.AccountID ORDER BY D.TransactionDate DESC, P.PaymentID DESC;"
    AT_CreateReadOnlyGridForm "frmTransferList", "سجل التحويلات", "SELECT N.OfficialNumber, T.TransferID, T.Amount, T.FeeAmount, S.AccountName AS SourceAccount, X.AccountName AS DestinationAccount, D.TransactionDate, D.StatusCode, T.ConfirmationRef FROM (((tblAccountTransfers AS T INNER JOIN tblFinancialDocuments AS D ON T.DocumentID=D.DocumentID) LEFT JOIN tblOfficialDocumentNumbers AS N ON D.DocumentID=N.DocumentID) LEFT JOIN tblAccounts AS S ON T.SourceAccountID=S.AccountID) LEFT JOIN tblAccounts AS X ON T.DestinationAccountID=X.AccountID ORDER BY D.TransactionDate DESC, T.TransferID DESC;"
    AT_CreateReadOnlyGridForm "frmOwnerBalances", "كشف أرصدة الملاك", "SELECT * FROM qryAT_OwnerBalances ORDER BY FullName;"
    AT_CreateReadOnlyGridForm "frmAccountBalances", "أرصدة الحسابات", "SELECT * FROM qryAT_AccountBalances ORDER BY AccountName;"
    AT_CreateReadOnlyGridForm "frmAccountMovements", "حركة الحسابات", "SELECT * FROM qryAT_AccountMovements;"
    AT_CreateReadOnlyGridForm "frmPendingAdjustments", "طلبات الخصم والإعفاء المعلقة", "SELECT * FROM qryAT_PendingAdjustments;"
    AT_CreateReadOnlyGridForm "frmCreditLedger", "دفتر أرصدة الملاك الدائنة", "SELECT * FROM qryAT_CreditLedger;"
    AT_CreateReadOnlyGridForm "frmChargeBalances", "كشف الاستحقاقات والمتبقي", "SELECT * FROM qryAT_ChargeBalances;"
    AT_CreateReadOnlyGridForm "frmOwnerStatementList", "كشف حركة حساب المالك", "SELECT * FROM qryAT_OwnerStatement ORDER BY EntryDate, ReferenceText;"
    DoCmd.OpenForm "frmMain", acNormal
    MsgBox "تم إنشاء مجموعة واجهات أولية. افحص كل نموذج داخل Access، ثم نفّذ اختبارات القبول. لم تُنشأ بعد تقارير الطباعة والإقفال والنسخ/الاستعادة.", vbInformation, "Archon Tech"
    Exit Sub
EH:
    Dim eNum As Long, eDesc As String: eNum = Err.Number: eDesc = Err.Description
    MsgBox "تعذر إكمال إنشاء الواجهات. قد تكون بعض النماذج قد أُنشئت. الخطأ " & eNum & ": " & eDesc, vbExclamation, "Archon Tech"
End Sub

Public Function AT_UI_OpenForm(ByVal formName As String) As Boolean
    On Error GoTo EH
    DoCmd.OpenForm formName, acNormal
    AT_UI_OpenForm = True
    Exit Function
EH:
    MsgBox "تعذر فتح الشاشة " & formName & ": " & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_RunIntegrity() As Boolean
    On Error GoTo EH
    AT_RunIntegrityReview
    AT_UI_RunIntegrity = True
    Exit Function
EH:
    MsgBox "تعذر فحص سلامة البيانات: " & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SaveOwnerDraft() As Boolean
    On Error GoTo EH
    Dim f As Form, ownerID As Long
    Set f = Forms("frmOwnerCreate")
    ownerID = AT_CreateOwnerDraft(CStr(Nz(f.Controls("txtFullName").Value, "")), CStr(Nz(f.Controls("txtPhone").Value, "")), CLng(Val(Nz(f.Controls("txtHouseCount").Value, "0"))), CLng(Val(Nz(f.Controls("txtApartmentCount").Value, "0"))), CStr(Nz(f.Controls("txtNotes").Value, "")))
    f.Controls("txtOwnerID").Value = ownerID
    MsgBox "حُفظت مسودة المالك بالمعرف الداخلي " & ownerID & ". لم يُخصص رقم OW بعد. راجع البيانات ثم اعتمدها عند الجاهزية.", vbInformation, "Archon Tech"
    AT_UI_SaveOwnerDraft = True
    Exit Function
EH:
    MsgBox "تعذر حفظ مسودة المالك: " & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ApproveOwner() As Boolean
    On Error GoTo EH
    Dim ownerID As Long
    ownerID = CLng(Val(Nz(Forms("frmOwnerCreate").Controls("txtOwnerID").Value, 0)))
    If ownerID <= 0 Then Err.Raise vbObjectError + 701, , "احفظ مسودة المالك أولًا."
    If MsgBox("هل راجعت بيانات المالك وتريد اعتمادها وتخصيص رقم رسمي لا يعاد استخدامه؟", vbQuestion + vbYesNo, "تأكيد اعتماد المالك") <> vbYes Then Exit Function
    AT_ApproveOwner ownerID, Environ$("Username")
    MsgBox "تم اعتماد المالك وتخصيص الرقم " & Nz(DLookup("OwnerCode", "tblOwners", "OwnerID=" & ownerID), "") & ".", vbInformation, "Archon Tech"
    AT_UI_ApproveOwner = True
    Exit Function
EH:
    MsgBox "تعذر اعتماد المالك: " & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SaveAccount() As Boolean
    On Error GoTo EH
    Dim f As Form, accountID As Long
    Set f = Forms("frmAccountCreate")
    If MsgBox("سيُنشأ الحساب برصيد افتتاحي صفر، ولا تُضاف الأموال إلا عبر حركة مالية معتمدة. هل تريد المتابعة؟", vbQuestion + vbYesNo, "تأكيد إنشاء الحساب") <> vbYes Then Exit Function
    accountID = AT_CreateAccount(CStr(Nz(f.Controls("txtAccountCode").Value, "")), _
        CStr(Nz(f.Controls("txtAccountName").Value, "")), CStr(Nz(f.Controls("cboAccountType").Value, "")))
    f.Controls("txtAccountID").Value = accountID
    MsgBox "تم إنشاء الحساب برقم داخلي " & accountID & " وبرصيد افتتاحي صفر.", vbInformation, "Archon Tech"
    AT_UI_SaveAccount = True
    Exit Function
EH:
    MsgBox "تعذر إنشاء الحساب: " & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SearchOwners() As Boolean
    On Error GoTo EH
    Dim f As Form, rawTerm As String, normalizedTerm As String, whereClause As String
    Set f = Forms("frmOwnerSearch")
    rawTerm = Trim$(CStr(Nz(f.Controls("txtSearchTerm").Value, "")))
    If Len(rawTerm) = 0 Then Err.Raise vbObjectError + 708, , "أدخل الاسم أو رقم المالك أو الهاتف للبحث."
    normalizedTerm = Replace(AT_NormalizeArabicSearch(rawTerm), "'", "''")
    rawTerm = Replace(rawTerm, "'", "''")
    whereClause = "AdminStatus='ACTIVE' AND (FullNameSearch Like '*" & normalizedTerm & "*' OR FullName Like '*" & rawTerm & "*' OR OwnerCode Like '*" & rawTerm & "*' OR Phone Like '*" & rawTerm & "*')"
    If CurrentProject.AllForms("frmOwnerDirectory").IsLoaded Then
        Forms("frmOwnerDirectory").Filter = whereClause
        Forms("frmOwnerDirectory").FilterOn = True
        Forms("frmOwnerDirectory").Requery
        DoCmd.SelectObject acForm, "frmOwnerDirectory", True
    Else
        DoCmd.OpenForm "frmOwnerDirectory", acNormal, , whereClause
    End If
    AT_UI_SearchOwners = True
    Exit Function
EH:
    MsgBox "تعذر البحث عن المالك: " & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_OpenOwnerStatement() As Boolean
    On Error GoTo EH
    Dim f As Form, ownerID As Long
    Set f = Forms("frmOwnerStatementPick")
    ownerID = CLng(Nz(f.Controls("cboOwnerID").Value, 0))
    If ownerID <= 0 Then Err.Raise vbObjectError + 711, , "اختر المالك أولًا."
    If CurrentProject.AllForms("frmOwnerStatementList").IsLoaded Then
        Forms("frmOwnerStatementList").Filter = "OwnerID=" & ownerID
        Forms("frmOwnerStatementList").FilterOn = True
        Forms("frmOwnerStatementList").Requery
        DoCmd.SelectObject acForm, "frmOwnerStatementList", True
    Else
        DoCmd.OpenForm "frmOwnerStatementList", acNormal, , "OwnerID=" & ownerID
    End If
    AT_UI_OpenOwnerStatement = True
    Exit Function
EH:
    MsgBox "تعذر فتح كشف حساب المالك: " & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_LoadOwnerEdit() As Boolean
    On Error GoTo EH
    Dim f As Form, ownerID As Long
    Set f = Forms("frmOwnerEdit")
    ownerID = CLng(Nz(f.Controls("cboOwnerID").Value, 0))
    If ownerID <= 0 Then Err.Raise vbObjectError + 709, , "اختر المالك أولًا."
    f.Controls("txtFullName").Value = Nz(DLookup("FullName", "tblOwners", "OwnerID=" & ownerID), "")
    f.Controls("txtPhone").Value = Nz(DLookup("Phone", "tblOwners", "OwnerID=" & ownerID), "")
    f.Controls("txtHouseCount").Value = Nz(DLookup("HouseCount", "tblOwners", "OwnerID=" & ownerID), 0)
    f.Controls("txtApartmentCount").Value = Nz(DLookup("ApartmentCount", "tblOwners", "OwnerID=" & ownerID), 0)
    f.Controls("txtNotes").Value = Nz(DLookup("Notes", "tblOwners", "OwnerID=" & ownerID), "")
    AT_UI_LoadOwnerEdit = True
    Exit Function
EH:
    MsgBox "تعذر تحميل بيانات المالك: " & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SaveOwnerEdit() As Boolean
    On Error GoTo EH
    Dim f As Form, ownerID As Long
    Set f = Forms("frmOwnerEdit")
    ownerID = CLng(Nz(f.Controls("cboOwnerID").Value, 0))
    If ownerID <= 0 Then Err.Raise vbObjectError + 710, , "اختر المالك أولًا."
    If MsgBox("سيتم تعديل بيانات المالك وتسجيل القيم السابقة والجديدة في سجل التدقيق. هل تريد المتابعة؟", vbQuestion + vbYesNo, "تأكيد التعديل") <> vbYes Then Exit Function
    AT_UpdateOwnerDetails ownerID, CStr(Nz(f.Controls("txtFullName").Value, "")), CStr(Nz(f.Controls("txtPhone").Value, "")), _
        CLng(Val(Nz(f.Controls("txtHouseCount").Value, "0"))), CLng(Val(Nz(f.Controls("txtApartmentCount").Value, "0"))), _
        CStr(Nz(f.Controls("txtNotes").Value, "")), CStr(Nz(f.Controls("txtEditReason").Value, ""))
    MsgBox "تم حفظ التعديل وتسجيل القيم السابقة والجديدة.", vbInformation, "Archon Tech"
    AT_UI_SaveOwnerEdit = True
    Exit Function
EH:
    MsgBox "تعذر تعديل بيانات المالك: " & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SaveChargeDraft() As Boolean
    On Error GoTo EH
    Dim f As Form, chargeID As Long, chargeDate As Date, rateID As Long, ownerID As Long
    Set f = Forms("frmChargeCreate")
    ownerID = CLng(Nz(f.Controls("cboOwnerID").Value, 0)): rateID = CLng(Nz(f.Controls("cboRateID").Value, 0))
    chargeDate = AT_UI_DateValue(f.Controls("txtChargeDate").Value)
    chargeID = AT_CreateChargeForOwner(ownerID, chargeDate, rateID, CStr(Nz(f.Controls("txtReason").Value, "")))
    f.Controls("txtChargeID").Value = chargeID
    MsgBox "أُنشئت مسودة الاستحقاق. الإجمالي المحفوظ: " & FormatCurrency(Nz(DLookup("GrossAmount", "tblOwnerCharges", "ChargeID=" & chargeID), 0), 0) & " ريال. لا يؤثر قبل الاعتماد.", vbInformation, "Archon Tech"
    AT_UI_SaveChargeDraft = True
    Exit Function
EH:
    MsgBox "تعذر إنشاء مسودة الاستحقاق: " & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ApproveCharge() As Boolean
    On Error GoTo EH
    Dim chargeID As Long
    chargeID = CLng(Val(Nz(Forms("frmChargeCreate").Controls("txtChargeID").Value, 0)))
    If chargeID <= 0 Then Err.Raise vbObjectError + 702, , "أنشئ مسودة الاستحقاق أولًا."
    If MsgBox("اعتماد الاستحقاق؟ سيصبح ضمن المستحقات المالية للمالك.", vbQuestion + vbYesNo, "تأكيد الاعتماد") <> vbYes Then Exit Function
    AT_ApproveCharge chargeID, Environ$("Username")
    MsgBox "تم اعتماد الاستحقاق.", vbInformation, "Archon Tech"
    AT_UI_ApproveCharge = True
    Exit Function
EH:
    MsgBox "تعذر اعتماد الاستحقاق: " & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SaveReceiptDraft() As Boolean
    On Error GoTo EH
    Dim f As Form, receiptID As Long, transactionDate As Date
    Set f = Forms("frmReceiptCreate"): transactionDate = AT_UI_DateValue(f.Controls("txtTransactionDate").Value)
    receiptID = AT_CreateReceiptDraft(CLng(Nz(f.Controls("cboOwnerID").Value, 0)), CCur(Nz(f.Controls("txtAmount").Value, 0)), CLng(Nz(f.Controls("cboPaymentMethodID").Value, 0)), CLng(Nz(f.Controls("cboAccountID").Value, 0)), CStr(Nz(f.Controls("txtPayerName").Value, "")), CStr(Nz(f.Controls("txtCollectorName").Value, "")), CStr(Nz(f.Controls("txtNotes").Value, "")), transactionDate)
    f.Controls("txtReceiptID").Value = receiptID
    MsgBox "أُنشئت مسودة الإيصال رقمها الداخلي " & receiptID & ". لم يتغير رصيد الحساب بعد.", vbInformation, "Archon Tech"
    AT_UI_SaveReceiptDraft = True
    Exit Function
EH:
    MsgBox "تعذر إنشاء مسودة الإيصال: " & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ApproveReceipt() As Boolean
    On Error GoTo EH
    Dim receiptID As Long
    receiptID = CLng(Val(Nz(Forms("frmReceiptCreate").Controls("txtReceiptID").Value, 0)))
    If receiptID <= 0 Then Err.Raise vbObjectError + 703, , "أنشئ مسودة الإيصال أولًا."
    If MsgBox("تأكيد: هل استلمت المبلغ فعليًا من الدافع وبالقيمة المسجلة؟", vbQuestion + vbYesNo + vbDefaultButton2, "تأكيد القبض الفعلي") <> vbYes Then Exit Function
    AT_ApproveReceipt receiptID, Environ$("Username"), True
    MsgBox "تم اعتماد الإيصال. الرقم الرسمي: " & Nz(DLookup("OfficialNumber", "tblOfficialDocumentNumbers", "DocumentID=" & Nz(DLookup("DocumentID", "tblReceipts", "ReceiptID=" & receiptID), 0)), "") & ".", vbInformation, "Archon Tech"
    AT_UI_ApproveReceipt = True
    Exit Function
EH:
    MsgBox "تعذر اعتماد الإيصال: " & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SavePaymentDraft() As Boolean
    On Error GoTo EH
    Dim f As Form, paymentID As Long, transactionDate As Date
    Set f = Forms("frmPaymentCreate"): transactionDate = AT_UI_DateValue(f.Controls("txtTransactionDate").Value)
    paymentID = AT_CreatePaymentDraft(CLng(Nz(f.Controls("cboAccountID").Value, 0)), CLng(Nz(f.Controls("cboExpenseCategoryID").Value, 0)), CCur(Nz(f.Controls("txtAmount").Value, 0)), CLng(Nz(f.Controls("cboPaymentMethodID").Value, 0)), CStr(Nz(f.Controls("txtBeneficiary").Value, "")), CStr(Nz(f.Controls("txtRecipient").Value, "")), CStr(Nz(f.Controls("txtNotes").Value, "")), transactionDate, CLng(Val(Nz(f.Controls("txtRefundCreditMovementID").Value, 0))))
    f.Controls("txtPaymentID").Value = paymentID
    MsgBox "أُنشئت مسودة سند الصرف. لم يُخصم شيء من الحساب بعد.", vbInformation, "Archon Tech"
    AT_UI_SavePaymentDraft = True
    Exit Function
EH:
    MsgBox "تعذر إنشاء مسودة سند الصرف: " & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ApprovePayment() As Boolean
    On Error GoTo EH
    Dim paymentID As Long
    paymentID = CLng(Val(Nz(Forms("frmPaymentCreate").Controls("txtPaymentID").Value, 0)))
    If paymentID <= 0 Then Err.Raise vbObjectError + 704, , "أنشئ مسودة سند الصرف أولًا."
    If MsgBox("تأكيد: هل دُفع المبلغ فعليًا للمستفيد؟ سيتم منع الاعتماد إن لم يكن الرصيد المتاح كافيًا.", vbQuestion + vbYesNo + vbDefaultButton2, "تأكيد الصرف الفعلي") <> vbYes Then Exit Function
    AT_ApprovePayment paymentID, Environ$("Username"), True
    MsgBox "تم اعتماد سند الصرف.", vbInformation, "Archon Tech"
    AT_UI_ApprovePayment = True
    Exit Function
EH:
    MsgBox "تعذر اعتماد سند الصرف: " & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SaveTransferDraft() As Boolean
    On Error GoTo EH
    Dim f As Form, transferID As Long, transactionDate As Date
    Set f = Forms("frmTransferCreate"): transactionDate = AT_UI_DateValue(f.Controls("txtTransactionDate").Value)
    transferID = AT_CreateTransferDraft(CLng(Nz(f.Controls("cboSourceAccountID").Value, 0)), CLng(Nz(f.Controls("cboDestinationAccountID").Value, 0)), CCur(Nz(f.Controls("txtAmount").Value, 0)), CCur(Nz(f.Controls("txtFeeAmount").Value, 0)), CStr(Nz(f.Controls("txtNotes").Value, "")), transactionDate)
    f.Controls("txtTransferID").Value = transferID
    MsgBox "أُنشئت مسودة التحويل. لم تُخصم أو تُضف مبالغ بعد.", vbInformation, "Archon Tech"
    AT_UI_SaveTransferDraft = True
    Exit Function
EH:
    MsgBox "تعذر إنشاء مسودة التحويل: " & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ApproveTransfer() As Boolean
    On Error GoTo EH
    Dim transferID As Long
    transferID = CLng(Val(Nz(Forms("frmTransferCreate").Controls("txtTransferID").Value, 0)))
    If transferID <= 0 Then Err.Raise vbObjectError + 705, , "أنشئ مسودة التحويل أولًا."
    If MsgBox("اعتماد التحويل المعلق وحجز المبلغ من الرصيد المتاح؟ لن يُضاف لحساب الوجهة حتى تأكيد الاستلام.", vbQuestion + vbYesNo, "تأكيد التحويل") <> vbYes Then Exit Function
    AT_ApproveTransfer transferID, Environ$("Username")
    MsgBox "تم اعتماد التحويل كمعلق وحُجز المبلغ. أدخل مرجع التأكيد ثم استخدم زر تأكيد الوصول.", vbInformation, "Archon Tech"
    AT_UI_ApproveTransfer = True
    Exit Function
EH:
    MsgBox "تعذر اعتماد التحويل: " & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ConfirmTransfer() As Boolean
    On Error GoTo EH
    Dim f As Form, transferID As Long, refText As String
    Set f = Forms("frmTransferCreate"): transferID = CLng(Val(Nz(f.Controls("txtTransferID").Value, 0))): refText = CStr(Nz(f.Controls("txtTransferRef").Value, ""))
    If transferID <= 0 Then Err.Raise vbObjectError + 706, , "أدخل معرف التحويل الداخلي أو أنشئ المسودة."
    If MsgBox("هل تأكدت فعليًا من وصول التحويل إلى الحساب المستلم؟", vbQuestion + vbYesNo + vbDefaultButton2, "تأكيد الوصول") <> vbYes Then Exit Function
    AT_ConfirmTransfer transferID, Environ$("Username"), refText
    MsgBox "تم ترحيل التحويل إلى الحسابين وتحرير الحجز.", vbInformation, "Archon Tech"
    AT_UI_ConfirmTransfer = True
    Exit Function
EH:
    MsgBox "تعذر تأكيد التحويل: " & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_RequestAdjustment() As Boolean
    On Error GoTo EH
    Dim f As Form, adjustmentID As Long
    Set f = Forms("frmAdjustmentCreate")
    adjustmentID = AT_RequestChargeAdjustment(CLng(Val(Nz(f.Controls("txtChargeID").Value, 0))), -Abs(CCur(Nz(f.Controls("txtDiscountAmount").Value, 0))), CStr(Nz(f.Controls("txtReason").Value, "")), Environ$("Username"), CStr(Nz(f.Controls("cboAdjustmentType").Value, "DISCOUNT")))
    f.Controls("txtAdjustmentID").Value = adjustmentID
    MsgBox "تم حفظ طلب الخصم كمسودة ولا يؤثر في الحساب قبل الاعتماد.", vbInformation, "Archon Tech"
    AT_UI_RequestAdjustment = True
    Exit Function
EH:
    MsgBox "تعذر تسجيل طلب الخصم: " & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ApproveAdjustment() As Boolean
    On Error GoTo EH
    Dim adjustmentID As Long
    adjustmentID = CLng(Val(Nz(Forms("frmAdjustmentCreate").Controls("txtAdjustmentID").Value, 0)))
    If adjustmentID <= 0 Then Err.Raise vbObjectError + 707, , "سجل طلب الخصم أولًا."
    If MsgBox("هل راجعت السبب والمبلغ قبل اعتماد الخصم؟", vbQuestion + vbYesNo + vbDefaultButton2, "اعتماد الخصم") <> vbYes Then Exit Function
    AT_ApproveChargeAdjustment adjustmentID, Environ$("Username")
    MsgBox "تم اعتماد الخصم.", vbInformation, "Archon Tech"
    AT_UI_ApproveAdjustment = True
    Exit Function
EH:
    MsgBox "تعذر اعتماد الخصم: " & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ApplyCredit() As Boolean
    On Error GoTo EH
    Dim f As Form, ownerID As Long, creditMovementID As Long, chargeID As Long, amount As Currency, reason As String
    Set f = Forms("frmCreditApply")
    ownerID = CLng(Nz(f.Controls("cboOwnerID").Value, 0))
    creditMovementID = CLng(Nz(f.Controls("cboCreditMovementID").Value, 0))
    chargeID = CLng(Nz(f.Controls("cboChargeID").Value, 0))
    amount = CCur(Nz(f.Controls("txtAmount").Value, 0)): reason = CStr(Nz(f.Controls("txtReason").Value, ""))
    If MsgBox("سيُعتمد تطبيق الرصيد الدائن مباشرة على الاستحقاق بعد التحقق من المالك والمبلغ. هل تريد المتابعة؟", vbQuestion + vbYesNo + vbDefaultButton2, "تأكيد تطبيق الرصيد الدائن") <> vbYes Then Exit Function
    AT_ApplyOwnerCreditToCharge ownerID, creditMovementID, chargeID, amount, Environ$("Username"), reason
    MsgBox "تم تطبيق الرصيد الدائن واعتماد التسوية.", vbInformation, "Archon Tech"
    AT_UI_ApplyCredit = True
    Exit Function
EH:
    MsgBox "تعذر تطبيق الرصيد الدائن: " & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ReverseAllocation() As Boolean
    On Error GoTo EH
    Dim f As Form, allocationID As Long, amount As Currency, reason As String
    Set f = Forms("frmReceiptCorrection")
    allocationID = CLng(Val(Nz(f.Controls("txtAllocationID").Value, 0)))
    amount = CCur(Nz(f.Controls("txtAmount").Value, 0)): reason = CStr(Nz(f.Controls("txtReason").Value, ""))
    If allocationID <= 0 Then Err.Raise vbObjectError + 708, , "أدخل رقم سطر التخصيص من سجل تخصيصات الإيصال."
    If MsgBox("سيعاد فتح المستحق بقيمة العكس ويصبح المبلغ المقابل رصيدًا دائنًا ما دام النقد لم يُرد. الرد النقدي يحتاج سند صرف منفصلًا. متابعة؟", vbQuestion + vbYesNo + vbDefaultButton2, "تأكيد عكس التخصيص") <> vbYes Then Exit Function
    AT_ReverseReceiptAllocation allocationID, amount, Environ$("Username"), reason
    MsgBox "تم عكس التخصيص وإنشاء الرصيد الدائن المقابل.", vbInformation, "Archon Tech"
    AT_UI_ReverseAllocation = True
    Exit Function
EH:
    MsgBox "تعذر عكس التخصيص: " & Err.Description, vbExclamation, "Archon Tech"
End Function

Private Function AT_UI_DateValue(ByVal v As Variant) As Date
    If IsNull(v) Or Len(Trim$(CStr(v))) = 0 Then AT_UI_DateValue = Date Else AT_UI_DateValue = CDate(v)
End Function

Private Function AT_FormExists(ByVal formName As String) As Boolean
    Dim ao As AccessObject
    For Each ao In CurrentProject.AllForms
        If StrComp(ao.Name, formName, vbTextCompare) = 0 Then AT_FormExists = True: Exit Function
    Next ao
End Function

Private Sub AT_CreateMainForm()
    Dim f As Form, actualName As String, x As Long, y As Long
    If AT_FormExists("frmMain") Then Exit Sub
    Set f = CreateForm: actualName = f.Name
    f.Caption = "Archon Tech | نظام تحصيل مساهمات المجاري": f.Width = 11800
    f.Section(acDetail).Height = 14300: f.Section(acDetail).BackColor = RGB(248, 250, 252)
    On Error Resume Next: f.RightToLeft = True: On Error GoTo 0
    AT_AddLabel actualName, "lblTitle", "نظام إدارة مساهمات مشروع المجاري", 450, 300, 10500, 500, 18, True
    AT_AddLabel actualName, "lblSubtitle", "يبدأ كل ملف حارة مستقلًا من الصفر؛ كل اعتماد مالي يتطلب تأكيدًا صريحًا.", 500, 850, 10500, 350, 10, False
    x = 650: y = 1500
    AT_AddButton actualName, "cmdOwnerCreate", "تسجيل مالك جديد", x, y, 4600, 650, "=AT_UI_OpenForm(""frmOwnerCreate"")"
    AT_AddButton actualName, "cmdOwnerDirectory", "سجل الملاك", 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmOwnerDirectory"")"
    y = y + 850
    AT_AddButton actualName, "cmdChargeCreate", "إنشاء استحقاق", x, y, 4600, 650, "=AT_UI_OpenForm(""frmChargeCreate"")"
    AT_AddButton actualName, "cmdReceiptCreate", "قبض / إيصال جديد", 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmReceiptCreate"")"
    y = y + 850
    AT_AddButton actualName, "cmdReceiptList", "سجل الإيصالات", x, y, 4600, 650, "=AT_UI_OpenForm(""frmReceiptList"")"
    AT_AddButton actualName, "cmdPaymentCreate", "سند صرف جديد", 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmPaymentCreate"")"
    y = y + 850
    AT_AddButton actualName, "cmdPaymentList", "سجل سندات الصرف", x, y, 4600, 650, "=AT_UI_OpenForm(""frmPaymentList"")"
    AT_AddButton actualName, "cmdTransferCreate", "التحويلات بين الحسابات", 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmTransferCreate"")"
    y = y + 850
    AT_AddButton actualName, "cmdAdjustment", "طلبات الخصم والإعفاء", x, y, 4600, 650, "=AT_UI_OpenForm(""frmAdjustmentCreate"")"
    AT_AddButton actualName, "cmdTransferList", "سجل التحويلات", 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmTransferList"")"
    y = y + 850
    AT_AddButton actualName, "cmdIntegrity", "فحص سلامة البيانات", x, y, 4600, 650, "=AT_UI_RunIntegrity()"
    AT_AddButton actualName, "cmdOwnerBalances", "كشف أرصدة الملاك والمتأخرات", 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmOwnerBalances"")"
    y = y + 850
    AT_AddButton actualName, "cmdAccountBalances", "أرصدة الحسابات", x, y, 4600, 650, "=AT_UI_OpenForm(""frmAccountBalances"")"
    AT_AddButton actualName, "cmdAccountMovements", "حركة الحسابات", 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmAccountMovements"")"
    y = y + 850
    AT_AddButton actualName, "cmdPendingAdjustments", "الخصومات المعلقة", x, y, 4600, 650, "=AT_UI_OpenForm(""frmPendingAdjustments"")"
    AT_AddButton actualName, "cmdOwnerEdit", "تعديل بيانات مالك", 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmOwnerEdit"")"
    y = y + 850
    AT_AddButton actualName, "cmdCreditApply", "تطبيق رصيد دائن على استحقاق", x, y, 4600, 650, "=AT_UI_OpenForm(""frmCreditApply"")"
    AT_AddButton actualName, "cmdCreditLedger", "دفتر الأرصدة الدائنة", 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmCreditLedger"")"
    y = y + 850
    AT_AddButton actualName, "cmdReverseAllocation", "عكس تخصيص إيصال", x, y, 4600, 650, "=AT_UI_OpenForm(""frmReceiptCorrection"")"
    AT_AddButton actualName, "cmdChargeBalances", "كشف الاستحقاقات", 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmChargeBalances"")"
    y = y + 850
    AT_AddButton actualName, "cmdOwnerSearch", "بحث عن مالك", x, y, 4600, 650, "=AT_UI_OpenForm(""frmOwnerSearch"")"
    AT_AddButton actualName, "cmdOwnerStatement", "كشف حساب مالك", 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmOwnerStatementPick"")"
    y = y + 850
    AT_AddButton actualName, "cmdAccountCreate", "إضافة صندوق / بنك / محفظة", x, y, 4600, 650, "=AT_UI_OpenForm(""frmAccountCreate"")"
    AT_AddButton actualName, "cmdAccountBalancesSecond", "استعراض أرصدة الحسابات", 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmAccountBalances"")"
    AT_AddLabel actualName, "lblFooter", "تنبيه: الواجهات أولية ولا تغني عن تجميع VBA واختبارات Access والطباعة والنسخ الاحتياطي قبل التحصيل الحقيقي.", 500, 13500, 10500, 350, 9, False
    DoCmd.Save acForm, "frmMain": DoCmd.Close acForm, "frmMain", acSaveYes
End Sub

Private Sub AT_CreateOwnerEntryForm()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmOwnerCreate") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, "إضافة مالك واعتماده", 8300
    y = 500
    AT_AddFieldPair n, "txtFullName", "اسم المالك", y, 6500: y = y + 650
    AT_AddFieldPair n, "txtPhone", "رقم الهاتف", y, 6500: y = y + 650
    AT_AddFieldPair n, "txtHouseCount", "عدد المنازل", y, 6500: y = y + 650
    AT_AddFieldPair n, "txtApartmentCount", "إجمالي الشقق", y, 6500: y = y + 650
    AT_AddFieldPair n, "txtNotes", "ملاحظات", y, 6500: y = y + 650
    AT_AddFieldPair n, "txtOwnerID", "المعرف الداخلي للمسودة", y, 6500: f.Controls("txtOwnerID").Locked = True: y = y + 850
    AT_AddButton n, "cmdSaveOwner", "حفظ كمسودة", 3000, y, 2200, 600, "=AT_UI_SaveOwnerDraft()"
    AT_AddButton n, "cmdApproveOwner", "مراجعة واعتماد", 5500, y, 2500, 600, "=AT_UI_ApproveOwner()"
    DoCmd.Save acForm, "frmOwnerCreate": DoCmd.Close acForm, "frmOwnerCreate", acSaveYes
End Sub

Private Sub AT_CreateAccountEntryForm()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmAccountCreate") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, "إضافة حساب مالي دون رصيد افتتاحي", 5200: y = 650
    AT_AddFieldPair n, "txtAccountCode", "رمز الحساب", y, 6500: y = y + 750
    AT_AddFieldPair n, "txtAccountName", "اسم الحساب", y, 6500: y = y + 750
    Call AT_AddCombo(n, "cboAccountType", "نوع الحساب", y, "SELECT 'CASH' AS TypeCode, 'صندوق نقدي' AS TypeName UNION SELECT 'BANK','حساب بنكي' UNION SELECT 'EWALLET','محفظة إلكترونية' UNION SELECT 'OTHER','أخرى';", 2, "0;4")
    y = y + 750
    AT_AddFieldPair n, "txtAccountID", "المعرف الداخلي", y, 6500: f.Controls("txtAccountID").Locked = True: y = y + 900
    AT_AddButton n, "cmdSaveAccount", "إنشاء الحساب برصيد صفر", 3600, y, 3400, 650, "=AT_UI_SaveAccount()"
    DoCmd.Save acForm, "frmAccountCreate": DoCmd.Close acForm, "frmAccountCreate", acSaveYes
End Sub

Private Sub AT_CreateOwnerSearchForm()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmOwnerSearch") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, "بحث عن مالك", 3600: y = 650
    AT_AddFieldPair n, "txtSearchTerm", "الاسم أو الرقم أو الهاتف", y, 6500: y = y + 1000
    AT_AddButton n, "cmdSearchOwners", "بحث", 4200, y, 2200, 650, "=AT_UI_SearchOwners()"
    DoCmd.Save acForm, "frmOwnerSearch": DoCmd.Close acForm, "frmOwnerSearch", acSaveYes
End Sub

Private Sub AT_CreateOwnerStatementPicker()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmOwnerStatementPick") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, "اختيار مالك لعرض كشف الحساب", 3600: y = 650
    Call AT_AddCombo(n, "cboOwnerID", "المالك", y, "SELECT OwnerID, OwnerCode & ' - ' & FullName AS OwnerLabel FROM tblOwners WHERE AdminStatus='ACTIVE' ORDER BY FullName;", 2, "0;5")
    y = y + 1000
    AT_AddButton n, "cmdOpenOwnerStatement", "عرض كشف الحساب", 3900, y, 2800, 650, "=AT_UI_OpenOwnerStatement()"
    DoCmd.Save acForm, "frmOwnerStatementPick": DoCmd.Close acForm, "frmOwnerStatementPick", acSaveYes
End Sub

Private Sub AT_CreateOwnerEditForm()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmOwnerEdit") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, "تعديل بيانات المالك مع سجل التدقيق", 8500: y = 450
    Call AT_AddCombo(n, "cboOwnerID", "المالك", y, "SELECT OwnerID, OwnerCode & ' - ' & FullName AS OwnerLabel FROM tblOwners WHERE AdminStatus='ACTIVE' ORDER BY FullName;", 2, "0;4"): y = y + 650
    AT_AddFieldPair n, "txtFullName", "اسم المالك", y, 6500: y = y + 650
    AT_AddFieldPair n, "txtPhone", "رقم الهاتف", y, 6500: y = y + 650
    AT_AddFieldPair n, "txtHouseCount", "عدد المنازل", y, 6500: y = y + 650
    AT_AddFieldPair n, "txtApartmentCount", "إجمالي الشقق", y, 6500: y = y + 650
    AT_AddFieldPair n, "txtNotes", "ملاحظات", y, 6500: y = y + 650
    AT_AddFieldPair n, "txtEditReason", "سبب التعديل (إلزامي)", y, 6500: y = y + 850
    AT_AddButton n, "cmdLoadOwner", "تحميل بيانات المالك", 2800, y, 2500, 600, "=AT_UI_LoadOwnerEdit()"
    AT_AddButton n, "cmdSaveOwnerEdit", "حفظ التعديل", 5600, y, 2200, 600, "=AT_UI_SaveOwnerEdit()"
    DoCmd.Save acForm, "frmOwnerEdit": DoCmd.Close acForm, "frmOwnerEdit", acSaveYes
End Sub

Private Sub AT_CreateChargeEntryForm()
    Dim f As Form, n As String, y As Long, c As Control
    If AT_FormExists("frmChargeCreate") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, "إنشاء استحقاق واعتماده", 7600: y = 500
    Set c = AT_AddCombo(n, "cboOwnerID", "المالك", y, "SELECT OwnerID, OwnerCode & ' - ' & FullName AS OwnerLabel FROM tblOwners WHERE AdminStatus='ACTIVE' ORDER BY FullName;", 2, "0;4")
    y = y + 650
    Set c = AT_AddCombo(n, "cboRateID", "سعر المساهمة", y, "SELECT RateID, RateCode, RateAmount FROM tblContributionRates WHERE IsActive=True ORDER BY EffectiveFrom DESC;", 3, "0;2;1.5"): y = y + 650
    AT_AddFieldPair n, "txtChargeDate", "تاريخ الاستحقاق", y, 6500: y = y + 650
    AT_AddFieldPair n, "txtReason", "البيان / السبب", y, 6500: y = y + 650
    AT_AddFieldPair n, "txtChargeID", "معرف الاستحقاق", y, 6500: f.Controls("txtChargeID").Locked = True: y = y + 850
    AT_AddButton n, "cmdSaveCharge", "حفظ كمسودة", 3000, y, 2200, 600, "=AT_UI_SaveChargeDraft()"
    AT_AddButton n, "cmdApproveCharge", "اعتماد الاستحقاق", 5500, y, 2500, 600, "=AT_UI_ApproveCharge()"
    DoCmd.Save acForm, "frmChargeCreate": DoCmd.Close acForm, "frmChargeCreate", acSaveYes
End Sub

Private Sub AT_CreateReceiptEntryForm()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmReceiptCreate") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, "تسجيل قبض وإصدار إيصال", 10200: y = 350
    Call AT_AddCombo(n, "cboOwnerID", "المالك", y, "SELECT OwnerID, OwnerCode & ' - ' & FullName AS OwnerLabel FROM tblOwners WHERE AdminStatus='ACTIVE' ORDER BY FullName;", 2, "0;4"): y = y + 550
    AT_AddFieldPair n, "txtAmount", "المبلغ بالريال", y, 6500: y = y + 550
    Call AT_AddCombo(n, "cboPaymentMethodID", "طريقة الدفع", y, "SELECT PaymentMethodID, MethodName FROM tblPaymentMethods WHERE IsActive=True ORDER BY MethodName;", 2, "0;4"): y = y + 550
    Call AT_AddCombo(n, "cboAccountID", "الحساب الذي استلم المال", y, "SELECT AccountID, AccountCode & ' - ' & AccountName FROM tblAccounts WHERE IsActive=True ORDER BY AccountName;", 2, "0;4"): y = y + 550
    AT_AddFieldPair n, "txtPayerName", "اسم الدافع (اختياري)", y, 6500: y = y + 550
    AT_AddFieldPair n, "txtCollectorName", "اسم المحصل", y, 6500: y = y + 550
    AT_AddFieldPair n, "txtTransactionDate", "تاريخ الحركة", y, 6500: y = y + 550
    AT_AddFieldPair n, "txtNotes", "ملاحظات", y, 6500: y = y + 550
    AT_AddFieldPair n, "txtReceiptID", "المعرف الداخلي للإيصال", y, 6500: f.Controls("txtReceiptID").Locked = True: y = y + 800
    AT_AddButton n, "cmdSaveReceipt", "حفظ كمسودة", 2500, y, 2200, 600, "=AT_UI_SaveReceiptDraft()"
    AT_AddButton n, "cmdApproveReceipt", "تأكيد القبض واعتماد", 5000, y, 3000, 600, "=AT_UI_ApproveReceipt()"
    DoCmd.Save acForm, "frmReceiptCreate": DoCmd.Close acForm, "frmReceiptCreate", acSaveYes
End Sub

Private Sub AT_CreatePaymentEntryForm()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmPaymentCreate") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, "إنشاء سند صرف / رد رصيد دائن", 11200: y = 350
    Call AT_AddCombo(n, "cboAccountID", "حساب الصرف", y, "SELECT AccountID, AccountCode & ' - ' & AccountName FROM tblAccounts WHERE IsActive=True ORDER BY AccountName;", 2, "0;4"): y = y + 550
    Call AT_AddCombo(n, "cboExpenseCategoryID", "فئة المصروف", y, "SELECT ExpenseCategoryID, CategoryName FROM tblExpenseCategories WHERE IsActive=True ORDER BY CategoryName;", 2, "0;4"): y = y + 550
    AT_AddFieldPair n, "txtAmount", "المبلغ بالريال", y, 6500: y = y + 550
    Call AT_AddCombo(n, "cboPaymentMethodID", "طريقة الدفع", y, "SELECT PaymentMethodID, MethodName FROM tblPaymentMethods WHERE IsActive=True ORDER BY MethodName;", 2, "0;4"): y = y + 550
    AT_AddFieldPair n, "txtBeneficiary", "المستفيد", y, 6500: y = y + 550
    AT_AddFieldPair n, "txtRecipient", "اسم المستلم", y, 6500: y = y + 550
    AT_AddFieldPair n, "txtTransactionDate", "تاريخ الحركة", y, 6500: y = y + 550
    AT_AddFieldPair n, "txtNotes", "ملاحظات", y, 6500: y = y + 550
    AT_AddFieldPair n, "txtPaymentID", "المعرف الداخلي للسند", y, 6500: f.Controls("txtPaymentID").Locked = True: y = y + 650
    AT_AddFieldPair n, "txtRefundCreditMovementID", "معرف حركة الدائن للرد (اختياري)", y, 6500: y = y + 800
    AT_AddButton n, "cmdSavePayment", "حفظ كمسودة", 2500, y, 2200, 600, "=AT_UI_SavePaymentDraft()"
    AT_AddButton n, "cmdApprovePayment", "تأكيد الصرف واعتماد", 5000, y, 3000, 600, "=AT_UI_ApprovePayment()"
    DoCmd.Save acForm, "frmPaymentCreate": DoCmd.Close acForm, "frmPaymentCreate", acSaveYes
End Sub

Private Sub AT_CreateTransferEntryForm()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmTransferCreate") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, "تحويل بين الحسابات", 9300: y = 350
    Call AT_AddCombo(n, "cboSourceAccountID", "حساب المصدر", y, "SELECT AccountID, AccountCode & ' - ' & AccountName FROM tblAccounts WHERE IsActive=True ORDER BY AccountName;", 2, "0;4"): y = y + 550
    Call AT_AddCombo(n, "cboDestinationAccountID", "حساب الوجهة", y, "SELECT AccountID, AccountCode & ' - ' & AccountName FROM tblAccounts WHERE IsActive=True ORDER BY AccountName;", 2, "0;4"): y = y + 550
    AT_AddFieldPair n, "txtAmount", "مبلغ التحويل", y, 6500: y = y + 550
    AT_AddFieldPair n, "txtFeeAmount", "رسوم التحويل", y, 6500: y = y + 550
    AT_AddFieldPair n, "txtTransactionDate", "تاريخ الحركة", y, 6500: y = y + 550
    AT_AddFieldPair n, "txtNotes", "ملاحظات", y, 6500: y = y + 550
    AT_AddFieldPair n, "txtTransferID", "المعرف الداخلي للتحويل", y, 6500: f.Controls("txtTransferID").Locked = True: y = y + 550
    AT_AddFieldPair n, "txtTransferRef", "مرجع تأكيد الوصول", y, 6500: y = y + 750
    AT_AddButton n, "cmdSaveTransfer", "حفظ كمسودة", 800, y, 1900, 550, "=AT_UI_SaveTransferDraft()"
    AT_AddButton n, "cmdApproveTransfer", "اعتماد وحجز", 2850, y, 1900, 550, "=AT_UI_ApproveTransfer()"
    AT_AddButton n, "cmdConfirmTransfer", "تأكيد الوصول", 4900, y, 2200, 550, "=AT_UI_ConfirmTransfer()"
    DoCmd.Save acForm, "frmTransferCreate": DoCmd.Close acForm, "frmTransferCreate", acSaveYes
End Sub

Private Sub AT_CreateAdjustmentEntryForm()
    Dim f As Form, n As String, y As Long, c As Control
    If AT_FormExists("frmAdjustmentCreate") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, "طلب خصم أو إعفاء", 6500: y = 500
    AT_AddFieldPair n, "txtChargeID", "معرف الاستحقاق", y, 6500: y = y + 650
    AT_AddFieldPair n, "txtDiscountAmount", "مبلغ الخصم (موجب في الإدخال)", y, 6500: y = y + 650
    Set c = AT_AddCombo(n, "cboAdjustmentType", "نوع التعديل", y, "SELECT 'DISCOUNT' AS TypeCode, 'خصم' AS TypeName UNION SELECT 'EXEMPTION','إعفاء';", 2, "0;4"): y = y + 650
    AT_AddFieldPair n, "txtReason", "سبب الطلب", y, 6500: y = y + 650
    AT_AddFieldPair n, "txtAdjustmentID", "معرف الطلب", y, 6500: f.Controls("txtAdjustmentID").Locked = True: y = y + 850
    AT_AddButton n, "cmdRequestAdjustment", "حفظ الطلب", 2500, y, 2200, 600, "=AT_UI_RequestAdjustment()"
    AT_AddButton n, "cmdApproveAdjustment", "اعتماد الطلب", 5000, y, 2200, 600, "=AT_UI_ApproveAdjustment()"
    DoCmd.Save acForm, "frmAdjustmentCreate": DoCmd.Close acForm, "frmAdjustmentCreate", acSaveYes
End Sub

Private Sub AT_CreateCreditApplicationForm()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmCreditApply") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, "تطبيق رصيد دائن على استحقاق", 6500: y = 450
    Call AT_AddCombo(n, "cboOwnerID", "المالك", y, "SELECT OwnerID, OwnerCode & ' - ' & FullName AS OwnerLabel FROM tblOwners WHERE AdminStatus='ACTIVE' ORDER BY FullName;", 2, "0;4"): y = y + 650
    Call AT_AddCombo(n, "cboCreditMovementID", "مصدر الرصيد الدائن", y, "SELECT CM.CreditMovementID, O.OwnerCode & ' - ' & O.FullName AS CreditOwner, CM.SignedAmount FROM (tblOwnerCreditMovements AS CM INNER JOIN tblOwners AS O ON CM.OwnerID=O.OwnerID) INNER JOIN tblFinancialDocuments AS D ON CM.DocumentID=D.DocumentID WHERE CM.SignedAmount>0 AND CM.SourceCreditMovementID Is Null AND D.StatusCode='APPROVED' ORDER BY CM.CreditMovementID DESC;", 3, "0;3.5;1.2"): y = y + 650
    Call AT_AddCombo(n, "cboChargeID", "الاستحقاق", y, "SELECT C.ChargeID, C.ChargeCode & ' - ' & O.FullName AS ChargeOwner, C.GrossAmount FROM (tblOwnerCharges AS C INNER JOIN tblOwners AS O ON C.OwnerID=O.OwnerID) INNER JOIN tblFinancialDocuments AS D ON C.DocumentID=D.DocumentID WHERE C.ChargeStatus='ACTIVE' AND D.StatusCode='APPROVED' ORDER BY C.ChargeDate, C.ChargeID;", 3, "0;3.5;1.2"): y = y + 650
    AT_AddFieldPair n, "txtAmount", "مبلغ التطبيق", y, 6500: y = y + 650
    AT_AddFieldPair n, "txtReason", "سبب التطبيق", y, 6500: y = y + 850
    AT_AddButton n, "cmdApplyCredit", "تأكيد تطبيق الرصيد", 3800, y, 3000, 600, "=AT_UI_ApplyCredit()"
    DoCmd.Save acForm, "frmCreditApply": DoCmd.Close acForm, "frmCreditApply", acSaveYes
End Sub

Private Sub AT_CreateReceiptCorrectionForm()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmReceiptCorrection") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, "عكس تخصيص إيصال", 5200: y = 550
    AT_AddFieldPair n, "txtAllocationID", "معرف سطر التخصيص", y, 6500: y = y + 700
    AT_AddFieldPair n, "txtAmount", "المبلغ المراد عكسه", y, 6500: y = y + 700
    AT_AddFieldPair n, "txtReason", "سبب العكس (إلزامي)", y, 6500: y = y + 900
    AT_AddButton n, "cmdReverseAllocation", "تأكيد العكس", 3800, y, 2500, 600, "=AT_UI_ReverseAllocation()"
    DoCmd.Save acForm, "frmReceiptCorrection": DoCmd.Close acForm, "frmReceiptCorrection", acSaveYes
End Sub

Private Sub AT_CreateReadOnlyGridForm(ByVal desiredName As String, ByVal captionText As String, ByVal sourceSQL As String)
    Dim f As Form, n As String
    If AT_FormExists(desiredName) Then Exit Sub
    Set f = CreateForm: n = f.Name
    f.Caption = captionText: f.RecordSource = sourceSQL: f.DefaultView = acDefViewDatasheet
    f.AllowEdits = False: f.AllowAdditions = False: f.AllowDeletions = False: f.DataEntry = False
    f.RecordSelectors = True: f.NavigationButtons = True
    DoCmd.Save acForm, desiredName: DoCmd.Close acForm, desiredName, acSaveYes
End Sub

Private Sub AT_SetFormBase(ByVal f As Form, ByVal captionText As String, ByVal heightValue As Long)
    f.Caption = captionText: f.Width = 10500: f.Section(acDetail).Height = heightValue
    f.Section(acDetail).BackColor = RGB(248, 250, 252): f.NavigationButtons = False: f.RecordSelectors = False
    ' Unbound entry controls must remain editable; read-only datasheets set their own lock flags.
    On Error Resume Next: f.RightToLeft = True: On Error GoTo 0
End Sub

Private Sub AT_AddLabel(ByVal formName As String, ByVal controlName As String, ByVal captionText As String, ByVal leftValue As Long, ByVal topValue As Long, ByVal widthValue As Long, ByVal heightValue As Long, Optional ByVal fontSizeValue As Long = 10, Optional ByVal isBold As Boolean = False)
    Dim c As Control
    Set c = CreateControl(formName, acLabel, acDetail, , controlName, leftValue, topValue, widthValue, heightValue)
    c.Caption = captionText: c.FontName = "Tahoma": c.FontSize = fontSizeValue: c.FontBold = isBold: c.TextAlign = 3
End Sub

Private Sub AT_AddFieldPair(ByVal formName As String, ByVal controlName As String, ByVal captionText As String, ByVal topValue As Long, ByVal widthValue As Long)
    Dim c As Control
    AT_AddLabel formName, "lbl_" & controlName, captionText, 550, topValue, 2150, 360, 10, False
    Set c = CreateControl(formName, acTextBox, acDetail, , controlName, 2900, topValue, widthValue, 380)
    c.Name = controlName: c.FontName = "Tahoma": c.FontSize = 10: c.DecimalPlaces = 0: c.TextAlign = 3
End Sub

Private Function AT_AddCombo(ByVal formName As String, ByVal controlName As String, ByVal captionText As String, ByVal topValue As Long, ByVal rowSource As String, ByVal columnCountValue As Long, ByVal columnWidthsValue As String) As Control
    Dim c As Control
    AT_AddLabel formName, "lbl_" & controlName, captionText, 550, topValue, 2150, 360, 10, False
    Set c = CreateControl(formName, acComboBox, acDetail, , controlName, 2900, topValue, 6500, 380)
    c.Name = controlName: c.RowSourceType = "Table/Query": c.RowSource = rowSource
    c.ColumnCount = columnCountValue: c.BoundColumn = 1: c.ColumnWidths = columnWidthsValue: c.LimitToList = True
    c.FontName = "Tahoma": c.FontSize = 10: c.TextAlign = 3
    Set AT_AddCombo = c
End Function

Private Sub AT_AddButton(ByVal formName As String, ByVal controlName As String, ByVal captionText As String, ByVal leftValue As Long, ByVal topValue As Long, ByVal widthValue As Long, ByVal heightValue As Long, ByVal onClickExpression As String)
    Dim c As Control
    Set c = CreateControl(formName, acCommandButton, acDetail, , controlName, leftValue, topValue, widthValue, heightValue)
    c.Name = controlName: c.Caption = captionText: c.OnClick = onClickExpression
    c.FontName = "Tahoma": c.FontSize = 10: c.FontBold = True
    c.BackColor = RGB(16, 185, 129): c.ForeColor = RGB(255, 255, 255)
End Sub
