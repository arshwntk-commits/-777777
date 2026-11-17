Attribute VB_Name = "modArchonBuild"
Option Compare Database
Option Explicit

' Archon Tech | Sewer Contributions Access Builder
' Run AT_BuildDatabase once in a NEW blank .accdb file.
' Tested here only by static review; final execution must be done in desktop Access for Windows.

Public Sub AT_BuildDatabase()
    On Error GoTo EH
    Dim db As DAO.Database, sqls As Variant, i As Long
    Set db = CurrentDb
    If DCount("*", "MSysObjects", "Type=1 AND Left([Name],3)='tbl'") > 0 Then
        If MsgBox("يوجد بالفعل جدول يبدأ اسمه بـ tbl. هل تريد الإكمال؟ لن تُحذف الجداول الموجودة.", vbQuestion + vbYesNo, "Archon Tech") <> vbYes Then Exit Sub
    End If
    sqls = Array( _
    "CREATE TABLE tblProjectSettings (SettingID COUNTER CONSTRAINT PK_tblProjectSettings PRIMARY KEY, ProjectName TEXT(150), NeighborhoodName TEXT(150), DefaultRate CURRENCY, CurrencyCode TEXT(10), CollectorName TEXT(150), CreatedAt DATETIME)", _
    "CREATE TABLE tblOwners (OwnerID COUNTER CONSTRAINT PK_tblOwners PRIMARY KEY, OwnerCode TEXT(12), FullName TEXT(200), Phone TEXT(40), HouseCount LONG, ApartmentCount LONG, Notes MEMO, AdminStatus TEXT(20), DataStatus TEXT(20), ReviewStatus TEXT(20), CreatedAt DATETIME, UpdatedAt DATETIME)", _
    "CREATE TABLE tblContributionRates (RateID COUNTER CONSTRAINT PK_tblContributionRates PRIMARY KEY, RateCode TEXT(30), RateAmount CURRENCY, EffectiveFrom DATETIME, EffectiveTo DATETIME, IsActive YESNO, CreatedAt DATETIME)", _
    "CREATE TABLE tblOwnerCharges (ChargeID COUNTER CONSTRAINT PK_tblOwnerCharges PRIMARY KEY, ChargeCode TEXT(20), OwnerID LONG, ChargeDate DATETIME, ApartmentCountAtCharge LONG, RateAmount CURRENCY, GrossAmount CURRENCY, ChargeStatus TEXT(20), CreatedAt DATETIME)", _
    "CREATE TABLE tblFinancialDocuments (DocumentID COUNTER CONSTRAINT PK_tblFinancialDocuments PRIMARY KEY, DocumentType TEXT(30), OfficialNumber TEXT(20), DocumentDate DATETIME, StatusCode TEXT(20), CreatedBy TEXT(100), ApprovedBy TEXT(100), CreatedAt DATETIME, ApprovedAt DATETIME, Reason MEMO, ReversalOfDocumentID LONG, CancelReason MEMO)", _
    "CREATE TABLE tblReceipts (ReceiptID COUNTER CONSTRAINT PK_tblReceipts PRIMARY KEY, DocumentID LONG, OwnerID LONG, Amount CURRENCY, PaymentMethod TEXT(40), AccountID LONG, PayerName TEXT(150), CollectorName TEXT(150), Notes MEMO, PrintedOwnerName TEXT(200), PrintedOwnerCode TEXT(12), CreatedAt DATETIME)", _
    "CREATE TABLE tblReceiptAllocationLines (AllocationID COUNTER CONSTRAINT PK_tblReceiptAllocationLines PRIMARY KEY, ReceiptID LONG, ChargeID LONG, AllocatedAmount CURRENCY, LineNo LONG)", _
    "CREATE TABLE tblOwnerAdjustments (AdjustmentID COUNTER CONSTRAINT PK_tblOwnerAdjustments PRIMARY KEY, DocumentID LONG, ChargeID LONG, AdjustmentType TEXT(30), Amount CURRENCY, Reason MEMO, RequestedBy TEXT(100), ApprovedBy TEXT(100), ApprovedAt DATETIME)", _
    "CREATE TABLE tblChargeSettlementLines (SettlementID COUNTER CONSTRAINT PK_tblChargeSettlementLines PRIMARY KEY, ChargeID LONG, DocumentID LONG, ReceiptID LONG, CreditApplicationID LONG, SignedAmount CURRENCY, SettlementType TEXT(30), ReversalOfSettlementID LONG, CreatedAt DATETIME)", _
    "CREATE TABLE tblOwnerCreditMovements (CreditMovementID COUNTER CONSTRAINT PK_tblOwnerCreditMovements PRIMARY KEY, OwnerID LONG, DocumentID LONG, SignedAmount CURRENCY, CreditType TEXT(30), SourceReceiptID LONG, SourceCreditMovementID LONG, Reason MEMO, CreatedAt DATETIME)", _
    "CREATE TABLE tblCreditApplications (CreditApplicationID COUNTER CONSTRAINT PK_tblCreditApplications PRIMARY KEY, DocumentID LONG, OwnerID LONG, Amount CURRENCY, ApprovedBy TEXT(100), ApprovedAt DATETIME, Reason MEMO)", _
    "CREATE TABLE tblCreditApplicationLines (CreditApplicationLineID COUNTER CONSTRAINT PK_tblCreditApplicationLines PRIMARY KEY, CreditApplicationID LONG, CreditMovementID LONG, ChargeID LONG, Amount CURRENCY)", _
    "CREATE TABLE tblAccounts (AccountID COUNTER CONSTRAINT PK_tblAccounts PRIMARY KEY, AccountCode TEXT(20), AccountName TEXT(120), AccountType TEXT(20), IsActive YESNO, CreatedAt DATETIME)", _
    "CREATE TABLE tblPayments (PaymentID COUNTER CONSTRAINT PK_tblPayments PRIMARY KEY, DocumentID LONG, AccountID LONG, ExpenseCategory TEXT(100), BeneficiaryName TEXT(150), Amount CURRENCY, PaymentMethod TEXT(40), RecipientName TEXT(150), Notes MEMO, RefundOfCreditMovementID LONG, RefundOfDocumentID LONG)", _
    "CREATE TABLE tblAccountTransfers (TransferID COUNTER CONSTRAINT PK_tblAccountTransfers PRIMARY KEY, DocumentID LONG, SourceAccountID LONG, DestinationAccountID LONG, Amount CURRENCY, FeeAmount CURRENCY, TransferState TEXT(20), ConfirmedAt DATETIME, ConfirmationRef TEXT(100), Notes MEMO)", _
    "CREATE TABLE tblAccountMovements (AccountMovementID COUNTER CONSTRAINT PK_tblAccountMovements PRIMARY KEY, AccountID LONG, DocumentID LONG, MovementDate DATETIME, SignedAmount CURRENCY, MovementType TEXT(30), LineNo LONG, CreatedAt DATETIME)", _
    "CREATE TABLE tblAuditEvents (AuditEventID COUNTER CONSTRAINT PK_tblAuditEvents PRIMARY KEY, EventAt DATETIME, UserName TEXT(100), ActionCode TEXT(60), EntityName TEXT(60), EntityID LONG, ResultCode TEXT(30), Reason MEMO, Detail MEMO)", _
    "CREATE TABLE tblAuditChanges (AuditChangeID COUNTER CONSTRAINT PK_tblAuditChanges PRIMARY KEY, AuditEventID LONG, FieldName TEXT(100), OldValue MEMO, NewValue MEMO)", _
    "CREATE TABLE tblFinancialClosings (ClosingID COUNTER CONSTRAINT PK_tblFinancialClosings PRIMARY KEY, PeriodStart DATETIME, PeriodEnd DATETIME, ClosedAt DATETIME, ClosedBy TEXT(100), StatusCode TEXT(20), DiscrepancyException YESNO, ExceptionReason MEMO, Summary MEMO)", _
    "CREATE TABLE tblFinancialClosingAccounts (ClosingAccountID COUNTER CONSTRAINT PK_tblFinancialClosingAccounts PRIMARY KEY, ClosingID LONG, AccountID LONG, BookBalance CURRENCY, ActualBalance CURRENCY, DifferenceAmount CURRENCY, Notes MEMO)", _
    "CREATE TABLE tblBackupLog (BackupLogID COUNTER CONSTRAINT PK_tblBackupLog PRIMARY KEY, StartedAt DATETIME, CompletedAt DATETIME, OperationType TEXT(20), SourcePath MEMO, DestinationPath MEMO, ResultCode TEXT(20), FileSizeBytes DOUBLE, Detail MEMO)", _
    "CREATE TABLE tblAttachments (AttachmentID COUNTER CONSTRAINT PK_tblAttachments PRIMARY KEY, DocumentID LONG, RelativePath MEMO, OriginalFileName TEXT(255), ContentType TEXT(100), FileSizeBytes DOUBLE, SHA256 TEXT(64), AddedAt DATETIME, AddedBy TEXT(100))", _
    "CREATE TABLE tblIntegrityRuns (IntegrityRunID COUNTER CONSTRAINT PK_tblIntegrityRuns PRIMARY KEY, StartedAt DATETIME, CompletedAt DATETIME, RunBy TEXT(100), ResultSummary MEMO)", _
    "CREATE TABLE tblIntegrityFindings (FindingID COUNTER CONSTRAINT PK_tblIntegrityFindings PRIMARY KEY, IntegrityRunID LONG, Severity TEXT(20), EntityName TEXT(60), EntityID LONG, FindingCode TEXT(60), FindingText MEMO, ResolutionStatus TEXT(20), ResolutionReason MEMO, ResolvedBy TEXT(100), ResolvedAt DATETIME)", _
    "CREATE TABLE tblNumberSeries (SeriesID COUNTER CONSTRAINT PK_tblNumberSeries PRIMARY KEY, SeriesCode TEXT(30), NextValue LONG, Prefix TEXT(10), PadLength LONG)" _
    )
    For i = LBound(sqls) To UBound(sqls)
        If Not AT_TableExists(CStr(AT_TableNameFromSQL(sqls(i)))) Then db.Execute CStr(sqls(i)), dbFailOnError
    Next i
    AT_CreateIndexes db
    AT_SeedDefaults db
    AT_Audit "BUILD_SCHEMA", "Database", 0, "SUCCESS", "إنشاء مخطط قاعدة البيانات" 
    MsgBox "تم إنشاء الجداول الأساسية. أكمل خطوات إعداد النماذج والاختبار من دليل التشغيل قبل إدخال بيانات حقيقية.", vbInformation, "Archon Tech"
    Exit Sub
EH:
    On Error Resume Next
    AT_Audit "BUILD_SCHEMA", "Database", 0, "FAILED", Err.Number & " - " & Err.Description
    MsgBox "توقف إنشاء المخطط عند الخطأ: " & Err.Number & vbCrLf & Err.Description & vbCrLf & "راجع الجداول الموجودة وسجل التدقيق قبل إعادة المحاولة.", vbCritical, "Archon Tech"
End Sub

Private Function AT_TableNameFromSQL(ByVal sql As String) As String
    Dim p As Long, q As Long
    p = InStr(1, sql, "CREATE TABLE ", vbTextCompare) + Len("CREATE TABLE ")
    q = InStr(p, sql, " ")
    AT_TableNameFromSQL = Mid$(sql, p, q - p)
End Function

Private Function AT_TableExists(ByVal tableName As String) As Boolean
    Dim t As DAO.TableDef
    On Error Resume Next
    Set t = CurrentDb.TableDefs(tableName)
    AT_TableExists = (Err.Number = 0)
    Err.Clear
End Function

Private Sub AT_CreateIndexes(ByVal db As DAO.Database)
    On Error Resume Next
    db.Execute "CREATE UNIQUE INDEX UX_OwnerCode ON tblOwners (OwnerCode)", dbFailOnError
    db.Execute "CREATE UNIQUE INDEX UX_RateCode ON tblContributionRates (RateCode)", dbFailOnError
    db.Execute "CREATE UNIQUE INDEX UX_ChargeCode ON tblOwnerCharges (ChargeCode)", dbFailOnError
    db.Execute "CREATE UNIQUE INDEX UX_OfficialNumber ON tblFinancialDocuments (OfficialNumber)", dbFailOnError
    db.Execute "CREATE UNIQUE INDEX UX_AccountCode ON tblAccounts (AccountCode)", dbFailOnError
    db.Execute "CREATE UNIQUE INDEX UX_SeriesCode ON tblNumberSeries (SeriesCode)", dbFailOnError
    db.Execute "CREATE INDEX IX_OwnerName ON tblOwners (FullName)", dbFailOnError
    db.Execute "CREATE INDEX IX_OwnerPhone ON tblOwners (Phone)", dbFailOnError
    db.Execute "CREATE INDEX IX_DocDateStatus ON tblFinancialDocuments (DocumentDate, StatusCode)", dbFailOnError
    db.Execute "CREATE INDEX IX_ChargeOwner ON tblOwnerCharges (OwnerID, ChargeDate)", dbFailOnError
    On Error GoTo 0
End Sub

Private Sub AT_SeedDefaults(ByVal db As DAO.Database)
    db.Execute "INSERT INTO tblProjectSettings (ProjectName, NeighborhoodName, DefaultRate, CurrencyCode, CollectorName, CreatedAt) VALUES ('مشروع مساهمات المجاري','اسم الحارة',3000,'YER','',Now())", dbFailOnError
    db.Execute "INSERT INTO tblContributionRates (RateCode, RateAmount, EffectiveFrom, IsActive, CreatedAt) VALUES ('RATE-0001',3000,#1/1/2020#,True,Now())", dbFailOnError
    db.Execute "INSERT INTO tblNumberSeries (SeriesCode, NextValue, Prefix, PadLength) VALUES ('OWNER',1,'OW-',4)", dbFailOnError
    db.Execute "INSERT INTO tblNumberSeries (SeriesCode, NextValue, Prefix, PadLength) VALUES ('RECEIPT',1,'RC-',6)", dbFailOnError
    db.Execute "INSERT INTO tblNumberSeries (SeriesCode, NextValue, Prefix, PadLength) VALUES ('PAYMENT',1,'PV-',6)", dbFailOnError
    db.Execute "INSERT INTO tblNumberSeries (SeriesCode, NextValue, Prefix, PadLength) VALUES ('TRANSFER',1,'TR-',6)", dbFailOnError
    db.Execute "INSERT INTO tblAccounts (AccountCode, AccountName, AccountType, IsActive, CreatedAt) VALUES ('CASH-01','الصندوق النقدي','CASH',True,Now())", dbFailOnError
End Sub

Public Function AT_NextOfficialNumber(ByVal seriesCode As String) As String
    On Error GoTo EH
    Dim db As DAO.Database, rs As DAO.Recordset, n As Long, prefix As String, pad As Long
    Set db = CurrentDb
    DBEngine(0)(0).BeginTrans
    Set rs = db.OpenRecordset("SELECT * FROM tblNumberSeries WHERE SeriesCode='" & Replace(seriesCode, "'", "''") & "'", dbOpenDynaset)
    If rs.EOF Then Err.Raise vbObjectError + 210, , "Series not found: " & seriesCode
    n = rs!NextValue: prefix = Nz(rs!Prefix, ""): pad = Nz(rs!PadLength, 6)
    rs.Edit: rs!NextValue = n + 1: rs.Update
    DBEngine(0)(0).CommitTrans
    AT_NextOfficialNumber = prefix & Right$(String$(pad, "0") & CStr(n), pad)
    rs.Close
    Exit Function
EH:
    On Error Resume Next
    DBEngine(0)(0).Rollback
    Err.Raise Err.Number, "AT_NextOfficialNumber", Err.Description
End Function

Public Sub AT_Audit(ByVal actionCode As String, ByVal entityName As String, ByVal entityID As Long, ByVal resultCode As String, ByVal detail As String, Optional ByVal reason As String = "")
    On Error Resume Next
    Dim q As String
    q = "INSERT INTO tblAuditEvents (EventAt, UserName, ActionCode, EntityName, EntityID, ResultCode, Reason, Detail) VALUES (Now(),'" & Replace(Environ$("Username"), "'", "''") & "','" & Replace(actionCode, "'", "''") & "','" & Replace(entityName, "'", "''") & "'," & entityID & ",'" & Replace(resultCode, "'", "''") & "','" & Replace(reason, "'", "''") & "','" & Replace(detail, "'", "''") & "')"
    CurrentDb.Execute q, dbFailOnError
End Sub
