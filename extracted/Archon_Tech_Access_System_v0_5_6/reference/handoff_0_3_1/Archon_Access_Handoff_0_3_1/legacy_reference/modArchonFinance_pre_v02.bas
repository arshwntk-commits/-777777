Attribute VB_Name = "modArchonFinance"
Option Compare Database
Option Explicit

' Financial services for the Archon Access schema.
' IMPORTANT: do not expose table editing directly to ordinary users.
' All amounts are positive inputs; signed ledger rows represent debits/credits.

Public Function AT_AccountBookBalance(ByVal accountID As Long, Optional ByVal asOfDate As Date = 0) As Currency
    Dim criteria As String
    criteria = "AccountID=" & accountID
    If asOfDate <> 0 Then criteria = criteria & " AND MovementDate<=#" & Format$(asOfDate, "yyyy-mm-dd") & "#"
    AT_AccountBookBalance = Nz(DSum("SignedAmount", "tblAccountMovements", criteria), 0)
End Function

Public Function AT_OwnerAvailableCredit(ByVal ownerID As Long) As Currency
    AT_OwnerAvailableCredit = Nz(DSum("SignedAmount", "tblOwnerCreditMovements", "OwnerID=" & ownerID), 0)
End Function

Public Function AT_ChargeRemaining(ByVal chargeID As Long) As Currency
    Dim gross As Currency, adjustments As Currency, settled As Currency
    gross = Nz(DLookup("GrossAmount", "tblOwnerCharges", "ChargeID=" & chargeID), 0)
    adjustments = Nz(DSum("Amount", "tblOwnerAdjustments", "ChargeID=" & chargeID & " AND ApprovedAt Is Not Null"), 0)
    settled = Nz(DSum("SignedAmount", "tblChargeSettlementLines", "ChargeID=" & chargeID), 0)
    AT_ChargeRemaining = gross - adjustments - settled
End Function

Public Function AT_CreateReceiptDraft(ByVal ownerID As Long, ByVal amount As Currency, ByVal paymentMethod As String, ByVal accountID As Long, ByVal payerName As String, ByVal collectorName As String, Optional ByVal notes As String = "", Optional ByVal transactionDate As Date = 0) As Long
    On Error GoTo EH
    If amount <= 0 Then Err.Raise vbObjectError + 301, , "المبلغ يجب أن يكون أكبر من صفر."
    If DCount("*", "tblOwners", "OwnerID=" & ownerID & " AND AdminStatus='ACTIVE'") = 0 Then Err.Raise vbObjectError + 302, , "المالك غير موجود أو غير نشط."
    If DCount("*", "tblAccounts", "AccountID=" & accountID & " AND IsActive=True") = 0 Then Err.Raise vbObjectError + 303, , "الحساب غير موجود أو غير نشط."
    If Len(Trim$(paymentMethod)) = 0 Then Err.Raise vbObjectError + 304, , "طريقة الدفع مطلوبة."
    If transactionDate = 0 Then transactionDate = Date
    Dim db As DAO.Database, rs As DAO.Recordset, docID As Long, receiptID As Long
    Set db = CurrentDb
    DBEngine(0)(0).BeginTrans
    Set rs = db.OpenRecordset("tblFinancialDocuments", dbOpenDynaset)
    rs.AddNew
    rs!DocumentType = "RECEIPT": rs!DocumentDate = transactionDate: rs!StatusCode = "DRAFT"
    rs!CreatedBy = Environ$("Username"): rs!CreatedAt = Now(): rs.Update
    rs.Bookmark = rs.LastModified: docID = rs!DocumentID: rs.Close
    Set rs = db.OpenRecordset("tblReceipts", dbOpenDynaset)
    rs.AddNew
    rs!DocumentID = docID: rs!OwnerID = ownerID: rs!Amount = amount
    rs!PaymentMethod = paymentMethod: rs!AccountID = accountID
    rs!PayerName = payerName: rs!CollectorName = collectorName: rs!Notes = notes
    rs!PrintedOwnerName = Nz(DLookup("FullName", "tblOwners", "OwnerID=" & ownerID), "")
    rs!PrintedOwnerCode = Nz(DLookup("OwnerCode", "tblOwners", "OwnerID=" & ownerID), "")
    rs!CreatedAt = Now(): rs.Update: rs.Bookmark = rs.LastModified: receiptID = rs!ReceiptID: rs.Close
    DBEngine(0)(0).CommitTrans
    AT_Audit "RECEIPT_DRAFT", "tblReceipts", receiptID, "SUCCESS", "Draft created; no financial effect."
    AT_CreateReceiptDraft = receiptID
    Exit Function
EH:
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "RECEIPT_DRAFT", "tblReceipts", 0, "FAILED", Err.Description
    Err.Raise Err.Number, "AT_CreateReceiptDraft", Err.Description
End Function

Public Sub AT_ApproveReceipt(ByVal receiptID As Long, ByVal approver As String, Optional ByVal confirmedActualReceipt As Boolean = False)
    On Error GoTo EH
    If Not confirmedActualReceipt Then Err.Raise vbObjectError + 321, , "يجب تأكيد استلام المال فعليًا قبل الاعتماد."
    Dim db As DAO.Database, rs As DAO.Recordset, docID As Long, ownerID As Long, amount As Currency, accountID As Long, docDate As Date, remaining As Currency, allocated As Currency, chargeID As Long, available As Currency, credit As Currency
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT R.*, D.StatusCode, D.DocumentDate FROM tblReceipts AS R INNER JOIN tblFinancialDocuments AS D ON R.DocumentID=D.DocumentID WHERE R.ReceiptID=" & receiptID, dbOpenSnapshot)
    If rs.EOF Then Err.Raise vbObjectError + 322, , "الإيصال غير موجود."
    If rs!StatusCode <> "DRAFT" Then Err.Raise vbObjectError + 323, , "لا يمكن اعتماد إيصال ليس في حالة مسودة."
    docID = rs!DocumentID: ownerID = rs!OwnerID: amount = rs!Amount: accountID = rs!AccountID: docDate = rs!DocumentDate: rs.Close
    DBEngine(0)(0).BeginTrans
    available = AT_AccountBookBalance(accountID)
    ' Receipt increases the account balance, so no minimum-balance check is needed.
    ' Allocate oldest open charges first; any excess becomes a single credit movement.
    Dim charges As DAO.Recordset
    Set charges = db.OpenRecordset("SELECT ChargeID, ChargeDate FROM tblOwnerCharges WHERE OwnerID=" & ownerID & " AND ChargeStatus='ACTIVE' ORDER BY ChargeDate, ChargeID", dbOpenSnapshot)
    remaining = amount
    Do While Not charges.EOF And remaining > 0
        chargeID = charges!ChargeID
        Dim due As Currency
        due = AT_ChargeRemaining(chargeID)
        If due > 0 Then
            allocated = IIf(remaining < due, remaining, due)
            db.Execute "INSERT INTO tblReceiptAllocationLines (ReceiptID, ChargeID, AllocatedAmount, LineNo) VALUES (" & receiptID & "," & chargeID & "," & CStr(allocated) & "," & (DCount("*", "tblReceiptAllocationLines", "ReceiptID=" & receiptID) + 1) & ")", dbFailOnError
            db.Execute "INSERT INTO tblChargeSettlementLines (ChargeID, DocumentID, ReceiptID, SignedAmount, SettlementType, CreatedAt) VALUES (" & chargeID & "," & docID & "," & receiptID & "," & CStr(allocated) & ",'RECEIPT',Now())", dbFailOnError
            remaining = remaining - allocated
        End If
        charges.MoveNext
    Loop
    charges.Close
    If remaining > 0 Then
        db.Execute "INSERT INTO tblOwnerCreditMovements (OwnerID, DocumentID, SignedAmount, CreditType, SourceReceiptID, Reason, CreatedAt) VALUES (" & ownerID & "," & docID & "," & CStr(remaining) & ",'RECEIPT_EXCESS'," & receiptID & ",'Excess from approved receipt',Now())", dbFailOnError
    End If
    db.Execute "INSERT INTO tblAccountMovements (AccountID, DocumentID, MovementDate, SignedAmount, MovementType, LineNo, CreatedAt) VALUES (" & accountID & "," & docID & ",#" & Format$(docDate, "yyyy-mm-dd") & "#," & CStr(amount) & ",'RECEIPT',1,Now())", dbFailOnError
    db.Execute "UPDATE tblFinancialDocuments SET StatusCode='APPROVED', OfficialNumber='" & AT_NextOfficialNumber("RECEIPT") & "', ApprovedBy='" & Replace(approver, "'", "''") & "', ApprovedAt=Now() WHERE DocumentID=" & docID, dbFailOnError
    DBEngine(0)(0).CommitTrans
    AT_Audit "RECEIPT_APPROVE", "tblReceipts", receiptID, "SUCCESS", "Approved; excess credit=" & CStr(remaining)
    Exit Sub
EH:
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "RECEIPT_APPROVE", "tblReceipts", receiptID, "FAILED", Err.Description
    Err.Raise Err.Number, "AT_ApproveReceipt", Err.Description
End Sub

Public Sub AT_ApprovePayment(ByVal paymentID As Long, ByVal approver As String, Optional ByVal confirmedPaid As Boolean = False)
    On Error GoTo EH
    If Not confirmedPaid Then Err.Raise vbObjectError + 341, , "يجب تأكيد أن المبلغ دُفع فعليًا قبل الاعتماد."
    Dim rs As DAO.Recordset, db As DAO.Database, docID As Long, accountID As Long, amount As Currency, docDate As Date, balance As Currency
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT P.*, D.StatusCode, D.DocumentDate FROM tblPayments AS P INNER JOIN tblFinancialDocuments AS D ON P.DocumentID=D.DocumentID WHERE P.PaymentID=" & paymentID, dbOpenSnapshot)
    If rs.EOF Then Err.Raise vbObjectError + 342, , "سند الصرف غير موجود."
    If rs!StatusCode <> "DRAFT" Then Err.Raise vbObjectError + 343, , "سند الصرف ليس مسودة."
    docID = rs!DocumentID: accountID = rs!AccountID: amount = rs!Amount: docDate = rs!DocumentDate: rs.Close
    If amount <= 0 Then Err.Raise vbObjectError + 344, , "المبلغ غير صالح."
    DBEngine(0)(0).BeginTrans
    balance = AT_AccountBookBalance(accountID)
    If balance < amount Then Err.Raise vbObjectError + 345, , "الرصيد الدفتري غير كافٍ؛ لا يسمح النظام باستثناء."
    db.Execute "INSERT INTO tblAccountMovements (AccountID, DocumentID, MovementDate, SignedAmount, MovementType, LineNo, CreatedAt) VALUES (" & accountID & "," & docID & ",#" & Format$(docDate, "yyyy-mm-dd") & "#,-" & CStr(amount) & ",'PAYMENT',1,Now())", dbFailOnError
    db.Execute "UPDATE tblFinancialDocuments SET StatusCode='APPROVED', OfficialNumber='" & AT_NextOfficialNumber("PAYMENT") & "', ApprovedBy='" & Replace(approver, "'", "''") & "', ApprovedAt=Now() WHERE DocumentID=" & docID, dbFailOnError
    DBEngine(0)(0).CommitTrans
    AT_Audit "PAYMENT_APPROVE", "tblPayments", paymentID, "SUCCESS", "Payment approved after balance check."
    Exit Sub
EH:
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "PAYMENT_APPROVE", "tblPayments", paymentID, "FAILED", Err.Description
    Err.Raise Err.Number, "AT_ApprovePayment", Err.Description
End Sub

Public Sub AT_RunIntegrityReview()
    Dim runID As Long, rs As DAO.Recordset, db As DAO.Database
    Set db = CurrentDb
    db.Execute "INSERT INTO tblIntegrityRuns (StartedAt, RunBy, ResultSummary) VALUES (Now(),'" & Replace(Environ$("Username"), "'", "''") & "','Running')", dbFailOnError
    Set rs = db.OpenRecordset("SELECT @@IDENTITY AS NewID", dbOpenSnapshot): runID = rs!NewID: rs.Close
    ' Findings are flags only; this routine intentionally does not change financial data.
    db.Execute "INSERT INTO tblIntegrityFindings (IntegrityRunID, Severity, EntityName, EntityID, FindingCode, FindingText, ResolutionStatus) SELECT " & runID & ",'ERROR','tblOwners',OwnerID,'OWNER_CODE_MISSING','Owner has blank visible code','OPEN' FROM tblOwners WHERE OwnerCode Is Null OR OwnerCode=''", dbFailOnError
    db.Execute "INSERT INTO tblIntegrityFindings (IntegrityRunID, Severity, EntityName, EntityID, FindingCode, FindingText, ResolutionStatus) SELECT " & runID & ",'ERROR','tblFinancialDocuments',DocumentID,'APPROVED_NO_NUMBER','Approved document has no official number','OPEN' FROM tblFinancialDocuments WHERE StatusCode='APPROVED' AND (OfficialNumber Is Null OR OfficialNumber='')", dbFailOnError
    db.Execute "INSERT INTO tblIntegrityFindings (IntegrityRunID, Severity, EntityName, EntityID, FindingCode, FindingText, ResolutionStatus) SELECT " & runID & ",'ERROR','tblAccountMovements',AccountMovementID,'ORPHAN_MOVEMENT','Movement references missing document','OPEN' FROM tblAccountMovements WHERE NOT EXISTS (SELECT DocumentID FROM tblFinancialDocuments WHERE tblFinancialDocuments.DocumentID=tblAccountMovements.DocumentID)", dbFailOnError
    db.Execute "UPDATE tblIntegrityRuns SET CompletedAt=Now(), ResultSummary='Static integrity checks completed; findings require human review.' WHERE IntegrityRunID=" & runID, dbFailOnError
    AT_Audit "INTEGRITY_REVIEW", "tblIntegrityRuns", runID, "SUCCESS", "Findings created without auto-correction."
    MsgBox "اكتمل فحص السلامة. راجع tblIntegrityFindings؛ لم تُعدّل أي بيانات مالية تلقائيًا.", vbInformation, "Archon Tech"
End Sub
