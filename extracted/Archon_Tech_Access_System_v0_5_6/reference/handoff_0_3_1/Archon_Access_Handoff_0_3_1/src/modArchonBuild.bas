Attribute VB_Name = "modArchonBuild"
Option Compare Database
Option Explicit

' Archon Tech | Sewer Contributions Access Builder
' Initial-install builder for a BLANK .accdb. It is not a migration tool.
' Must be compiled and executed inside desktop Microsoft Access on Windows.

Private Const AT_SCHEMA_VERSION As String = "0.3.1"

Public Sub AT_BuildDatabase()
    On Error GoTo EH
    Dim db As DAO.Database
    Dim ddls As Variant
    Dim i As Long
    Dim createdCount As Long

    Set db = CurrentDb
    If AT_HasPriorSchema(db) Then
        MsgBox "توجد قاعدة سبق إعدادها. هذا الباني مخصص للتثبيت الأولي وليس أداة ترقية، ولم يُجر أي تعديل. احتفظ بنسخة احتياطية واستخدم خطة ترحيل مخصصة.", vbExclamation, "Archon Tech"
        Exit Sub
    End If

    ddls = Array( _
        "CREATE TABLE tblProjectSettings (ProjectSettingID COUNTER CONSTRAINT PK_tblProjectSettings PRIMARY KEY, ProjectName TEXT(150), NeighborhoodName TEXT(150), DefaultRate CURRENCY, CurrencyCode TEXT(10), DefaultCollectorName TEXT(150), CreatedAt DATETIME, UpdatedAt DATETIME)", _
        "CREATE TABLE tblContributionRates (RateID COUNTER CONSTRAINT PK_tblContributionRates PRIMARY KEY, RateCode TEXT(30), RateAmount CURRENCY, EffectiveFrom DATETIME, EffectiveTo DATETIME, IsActive YESNO, CreatedAt DATETIME)", _
        "CREATE TABLE tblUsers (UserID COUNTER CONSTRAINT PK_tblUsers PRIMARY KEY, UserName TEXT(100), DisplayName TEXT(150), AuthenticationMode TEXT(30), IsActive YESNO, CreatedAt DATETIME, LastLoginAt DATETIME)", _
        "CREATE TABLE tblRoles (RoleID COUNTER CONSTRAINT PK_tblRoles PRIMARY KEY, RoleCode TEXT(40), RoleName TEXT(100), IsActive YESNO, CreatedAt DATETIME)", _
        "CREATE TABLE tblPermissions (PermissionID COUNTER CONSTRAINT PK_tblPermissions PRIMARY KEY, PermissionCode TEXT(60), PermissionName TEXT(150), IsActive YESNO)", _
        "CREATE TABLE tblUserRoleAssignments (UserRoleAssignmentID COUNTER CONSTRAINT PK_tblUserRoleAssignments PRIMARY KEY, UserID LONG, RoleID LONG, AssignedAt DATETIME, AssignedByUserID LONG, IsActive YESNO)", _
        "CREATE TABLE tblRolePermissions (RolePermissionID COUNTER CONSTRAINT PK_tblRolePermissions PRIMARY KEY, RoleID LONG, PermissionID LONG, GrantedAt DATETIME)", _
        "CREATE TABLE tblOwners (OwnerID COUNTER CONSTRAINT PK_tblOwners PRIMARY KEY, OwnerCode TEXT(12), FullName TEXT(200), FullNameSearch TEXT(200), Phone TEXT(40), HouseCount LONG, ApartmentCount LONG, Notes MEMO, AdminStatus TEXT(20), DataStatus TEXT(20), ReviewStatus TEXT(20), CreatedAt DATETIME, UpdatedAt DATETIME, CreatedByUserID LONG, UpdatedByUserID LONG)", _
        "CREATE TABLE tblChargeBatches (ChargeBatchID COUNTER CONSTRAINT PK_tblChargeBatches PRIMARY KEY, BatchCode TEXT(30), PeriodLabel TEXT(100), ChargeDate DATETIME, RateID LONG, StatusCode TEXT(20), CreatedAt DATETIME, CreatedByUserID LONG, ApprovedAt DATETIME, ApprovedByUserID LONG, Reason MEMO)", _
        "CREATE TABLE tblFinancialDocuments (DocumentID COUNTER CONSTRAINT PK_tblFinancialDocuments PRIMARY KEY, DocumentType TEXT(30), TransactionDate DATETIME, StatusCode TEXT(20), CreatedAt DATETIME, CreatedByUserName TEXT(100), ApprovedAt DATETIME, ApprovedByUserName TEXT(100), Reason MEMO, CancelReason MEMO, ReversalOfDocumentID LONG)", _
        "CREATE TABLE tblOfficialDocumentNumbers (OfficialDocumentNumberID COUNTER CONSTRAINT PK_tblOfficialDocumentNumbers PRIMARY KEY, DocumentID LONG, OfficialNumber TEXT(20), SeriesCode TEXT(30), NumericValue LONG, AssignedAt DATETIME, AssignedByUserName TEXT(100))", _
        "CREATE TABLE tblNumberSeries (SeriesID COUNTER CONSTRAINT PK_tblNumberSeries PRIMARY KEY, SeriesCode TEXT(30), NextValue LONG, Prefix TEXT(10), PadLength LONG)", _
        "CREATE TABLE tblOwnerCharges (ChargeID COUNTER CONSTRAINT PK_tblOwnerCharges PRIMARY KEY, ChargeCode TEXT(20), DocumentID LONG, OwnerID LONG, ChargeBatchID LONG, ChargeDate DATETIME, ApartmentCountAtCharge LONG, RateAmount CURRENCY, GrossAmount CURRENCY, ChargeStatus TEXT(20), CreatedAt DATETIME)", _
        "CREATE TABLE tblReceipts (ReceiptID COUNTER CONSTRAINT PK_tblReceipts PRIMARY KEY, DocumentID LONG, OwnerID LONG, Amount CURRENCY, PaymentMethodID LONG, AccountID LONG, PayerName TEXT(150), CollectorName TEXT(150), Notes MEMO, PrintedOwnerName TEXT(200), PrintedOwnerCode TEXT(12), PrintedNeighborhoodName TEXT(150), CreatedAt DATETIME)", _
        "CREATE TABLE tblReceiptAllocationLines (AllocationID COUNTER CONSTRAINT PK_tblReceiptAllocationLines PRIMARY KEY, ReceiptID LONG, ChargeID LONG, AllocatedAmount CURRENCY, LineNo LONG)", _
        "CREATE TABLE tblOwnerAdjustments (AdjustmentID COUNTER CONSTRAINT PK_tblOwnerAdjustments PRIMARY KEY, DocumentID LONG, ChargeID LONG, AdjustmentType TEXT(30), SignedAmount CURRENCY, Reason MEMO, RequestedByUserName TEXT(100), ApprovedByUserName TEXT(100), ApprovedAt DATETIME)", _
        "CREATE TABLE tblChargeSettlementLines (SettlementID COUNTER CONSTRAINT PK_tblChargeSettlementLines PRIMARY KEY, ChargeID LONG, DocumentID LONG, ReceiptID LONG, CreditApplicationID LONG, ReceiptCorrectionID LONG, SignedAmount CURRENCY, SettlementType TEXT(30), ReversalOfSettlementID LONG, CreatedAt DATETIME)", _
        "CREATE TABLE tblReceiptCorrections (ReceiptCorrectionID COUNTER CONSTRAINT PK_tblReceiptCorrections PRIMARY KEY, DocumentID LONG, OriginalReceiptID LONG, CorrectionType TEXT(30), Reason MEMO)", _
        "CREATE TABLE tblReceiptCorrectionLines (ReceiptCorrectionLineID COUNTER CONSTRAINT PK_tblReceiptCorrectionLines PRIMARY KEY, ReceiptCorrectionID LONG, OriginalAllocationID LONG, Amount CURRENCY, CreatedAt DATETIME)", _
        "CREATE TABLE tblOwnerCreditMovements (CreditMovementID COUNTER CONSTRAINT PK_tblOwnerCreditMovements PRIMARY KEY, OwnerID LONG, DocumentID LONG, SignedAmount CURRENCY, CreditType TEXT(30), SourceReceiptID LONG, SourceCreditMovementID LONG, Reason MEMO, CreatedAt DATETIME)", _
        "CREATE TABLE tblCreditApplications (CreditApplicationID COUNTER CONSTRAINT PK_tblCreditApplications PRIMARY KEY, DocumentID LONG, OwnerID LONG, Amount CURRENCY, ApprovedByUserName TEXT(100), ApprovedAt DATETIME, Reason MEMO)", _
        "CREATE TABLE tblCreditApplicationLines (CreditApplicationLineID COUNTER CONSTRAINT PK_tblCreditApplicationLines PRIMARY KEY, CreditApplicationID LONG, CreditMovementID LONG, ChargeID LONG, Amount CURRENCY)", _
        "CREATE TABLE tblAccounts (AccountID COUNTER CONSTRAINT PK_tblAccounts PRIMARY KEY, AccountCode TEXT(20), AccountName TEXT(120), AccountType TEXT(20), IsActive YESNO, CreatedAt DATETIME)", _
        "CREATE TABLE tblPaymentMethods (PaymentMethodID COUNTER CONSTRAINT PK_tblPaymentMethods PRIMARY KEY, MethodCode TEXT(30), MethodName TEXT(100), IsActive YESNO, CreatedAt DATETIME)", _
        "CREATE TABLE tblExpenseCategories (ExpenseCategoryID COUNTER CONSTRAINT PK_tblExpenseCategories PRIMARY KEY, CategoryCode TEXT(30), CategoryName TEXT(120), IsActive YESNO, CreatedAt DATETIME)", _
        "CREATE TABLE tblPayments (PaymentID COUNTER CONSTRAINT PK_tblPayments PRIMARY KEY, DocumentID LONG, AccountID LONG, ExpenseCategoryID LONG, BeneficiaryName TEXT(150), Amount CURRENCY, PaymentMethodID LONG, RecipientName TEXT(150), Notes MEMO, RefundOfCreditMovementID LONG, RefundOfDocumentID LONG)", _
        "CREATE TABLE tblAccountTransfers (TransferID COUNTER CONSTRAINT PK_tblAccountTransfers PRIMARY KEY, DocumentID LONG, SourceAccountID LONG, DestinationAccountID LONG, Amount CURRENCY, FeeAmount CURRENCY, ConfirmedAt DATETIME, ConfirmedByUserName TEXT(100), ConfirmationRef TEXT(100), Notes MEMO)", _
        "CREATE TABLE tblAccountMovements (AccountMovementID COUNTER CONSTRAINT PK_tblAccountMovements PRIMARY KEY, AccountID LONG, DocumentID LONG, MovementDate DATETIME, SignedAmount CURRENCY, MovementType TEXT(30), LineNo LONG, CreatedAt DATETIME)", _
        "CREATE TABLE tblAccountReservations (ReservationID COUNTER CONSTRAINT PK_tblAccountReservations PRIMARY KEY, DocumentID LONG, AccountID LONG, Amount CURRENCY, StatusCode TEXT(20), CreatedAt DATETIME, ReleasedAt DATETIME, Reason MEMO)", _
        "CREATE TABLE tblFinancialClosings (ClosingID COUNTER CONSTRAINT PK_tblFinancialClosings PRIMARY KEY, PeriodStart DATETIME, PeriodEnd DATETIME, ClosedAt DATETIME, ClosedByUserName TEXT(100), StatusCode TEXT(20), DiscrepancyException YESNO, ExceptionReason MEMO, Summary MEMO, ReopenedAt DATETIME, ReopenedByUserName TEXT(100), ReopenReason MEMO)", _
        "CREATE TABLE tblFinancialClosingAccounts (ClosingAccountID COUNTER CONSTRAINT PK_tblFinancialClosingAccounts PRIMARY KEY, ClosingID LONG, AccountID LONG, BookBalance CURRENCY, ActualBalance CURRENCY, DifferenceAmount CURRENCY, Notes MEMO)", _
        "CREATE TABLE tblAuditEvents (AuditEventID COUNTER CONSTRAINT PK_tblAuditEvents PRIMARY KEY, EventAt DATETIME, UserName TEXT(100), ActionCode TEXT(60), EntityName TEXT(60), EntityID LONG, ResultCode TEXT(30), Reason MEMO, Detail MEMO)", _
        "CREATE TABLE tblAuditChanges (AuditChangeID COUNTER CONSTRAINT PK_tblAuditChanges PRIMARY KEY, AuditEventID LONG, FieldName TEXT(100), OldValue MEMO, NewValue MEMO)", _
        "CREATE TABLE tblAttachments (AttachmentID COUNTER CONSTRAINT PK_tblAttachments PRIMARY KEY, DocumentID LONG, RelativePath MEMO, OriginalFileName TEXT(255), ContentType TEXT(100), FileSizeBytes DOUBLE, SHA256 TEXT(64), AddedAt DATETIME, AddedByUserName TEXT(100))", _
        "CREATE TABLE tblBackupLog (BackupLogID COUNTER CONSTRAINT PK_tblBackupLog PRIMARY KEY, StartedAt DATETIME, CompletedAt DATETIME, OperationType TEXT(20), SourcePath MEMO, DestinationPath MEMO, ResultCode TEXT(20), FileSizeBytes DOUBLE, Detail MEMO)", _
        "CREATE TABLE tblIntegrityRuns (IntegrityRunID COUNTER CONSTRAINT PK_tblIntegrityRuns PRIMARY KEY, StartedAt DATETIME, CompletedAt DATETIME, RunByUserName TEXT(100), ResultSummary MEMO)", _
        "CREATE TABLE tblIntegrityFindings (FindingID COUNTER CONSTRAINT PK_tblIntegrityFindings PRIMARY KEY, IntegrityRunID LONG, Severity TEXT(20), EntityName TEXT(60), EntityID LONG, FindingCode TEXT(60), FindingText MEMO, ResolutionStatus TEXT(20), ResolutionReason MEMO, ResolvedByUserName TEXT(100), ResolvedAt DATETIME)", _
        "CREATE TABLE tblSchemaVersion (SchemaVersionID COUNTER CONSTRAINT PK_tblSchemaVersion PRIMARY KEY, VersionCode TEXT(20), InstalledAt DATETIME, InstalledByUserName TEXT(100), Notes MEMO)" _
    )

    For i = LBound(ddls) To UBound(ddls)
        If Not AT_TableExists(db, AT_TableNameFromSQL(CStr(ddls(i)))) Then
            db.Execute CStr(ddls(i)), dbFailOnError
            createdCount = createdCount + 1
        End If
    Next i

    AT_CreateIndexes db
    AT_CreateRelations db
    AT_SeedDefaults db
    AT_SetSchemaVersion db
    AT_Audit "BUILD_SCHEMA", "Database", 0, "SUCCESS", "Schema version " & AT_SCHEMA_VERSION & "; tables created in this run=" & createdCount
    MsgBox "اكتمل إنشاء مخطط البيانات الأولي وإعداد العلاقات والفهارس. هذا لا ينشئ بعد جميع النماذج والتقارير ولا يثبت سلامة التشغيل؛ نفّذ اختبار Access قبل البيانات الحقيقية.", vbInformation, "Archon Tech"
    Exit Sub
EH:
    Dim eNum As Long, eDesc As String
    eNum = Err.Number: eDesc = Err.Description
    On Error Resume Next
    AT_Audit "BUILD_SCHEMA", "Database", 0, "FAILED", CStr(eNum) & " - " & eDesc
    MsgBox "توقف البناء. قد تكون بعض الجداول قد أُنشئت قبل الخطأ؛ لا تعاود التشغيل على قاعدة مستخدمة ولا تحذف شيئًا. الخطأ " & eNum & ": " & eDesc, vbCritical, "Archon Tech"
End Sub

Private Function AT_HasPriorSchema(ByVal db As DAO.Database) As Boolean
    On Error GoTo CleanExit
    If AT_TableExists(db, "tblSchemaVersion") Then
        AT_HasPriorSchema = (DCount("*", "tblSchemaVersion") > 0)
        Exit Function
    End If
    If AT_TableExists(db, "tblOwners") Or AT_TableExists(db, "tblFinancialDocuments") Then AT_HasPriorSchema = True
CleanExit:
End Function

Private Function AT_TableNameFromSQL(ByVal sql As String) As String
    Dim p As Long, q As Long
    p = InStr(1, sql, "CREATE TABLE ", vbTextCompare) + Len("CREATE TABLE ")
    q = InStr(p, sql, " ")
    AT_TableNameFromSQL = Mid$(sql, p, q - p)
End Function

Private Function AT_TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    Dim tdf As DAO.TableDef
    On Error Resume Next
    Set tdf = db.TableDefs(tableName)
    AT_TableExists = (Err.Number = 0 And Not tdf Is Nothing)
    Err.Clear
End Function

Private Function AT_IndexExists(ByVal db As DAO.Database, ByVal tableName As String, ByVal indexName As String) As Boolean
    Dim idx As DAO.Index
    On Error Resume Next
    AT_IndexExists = False
    For Each idx In db.TableDefs(tableName).Indexes
        If StrComp(idx.Name, indexName, vbTextCompare) = 0 Then AT_IndexExists = True: Exit For
    Next idx
    Err.Clear
End Function

Private Sub AT_CreateIndexes(ByVal db As DAO.Database)
    ' IGNORE NULL is explicit for codes/numbers assigned only at final approval.
    AT_EnsureIndex db, "tblOwners", "UX_OwnerCode", "OwnerCode", True, True
    AT_EnsureIndex db, "tblContributionRates", "UX_RateCode", "RateCode", True, False
    AT_EnsureIndex db, "tblOwnerCharges", "UX_ChargeCode", "ChargeCode", True, True
    AT_EnsureIndex db, "tblOfficialDocumentNumbers", "UX_OfficialNumber", "OfficialNumber", True, False
    AT_EnsureIndex db, "tblOfficialDocumentNumbers", "UX_OfficialDocument", "DocumentID", True, False
    AT_EnsureIndex db, "tblNumberSeries", "UX_SeriesCode", "SeriesCode", True, False
    AT_EnsureIndex db, "tblAccounts", "UX_AccountCode", "AccountCode", True, False
    AT_EnsureIndex db, "tblPaymentMethods", "UX_MethodCode", "MethodCode", True, False
    AT_EnsureIndex db, "tblExpenseCategories", "UX_ExpenseCategoryCode", "CategoryCode", True, False
    AT_EnsureIndex db, "tblUsers", "UX_UserName", "UserName", True, False
    AT_EnsureIndex db, "tblRoles", "UX_RoleCode", "RoleCode", True, False
    AT_EnsureIndex db, "tblPermissions", "UX_PermissionCode", "PermissionCode", True, False
    AT_EnsureIndex db, "tblReceiptAllocationLines", "UX_ReceiptLineNo", "ReceiptID,LineNo", True, False
    AT_EnsureIndex db, "tblFinancialClosingAccounts", "UX_ClosingAccount", "ClosingID,AccountID", True, False
    AT_EnsureIndex db, "tblUserRoleAssignments", "UX_UserRole", "UserID,RoleID", True, False
    AT_EnsureIndex db, "tblRolePermissions", "UX_RolePermission", "RoleID,PermissionID", True, False
    AT_EnsureIndex db, "tblCreditApplicationLines", "UX_CreditApplicationSourceCharge", "CreditApplicationID,CreditMovementID,ChargeID", True, False
    AT_EnsureIndex db, "tblAccountMovements", "UX_DocumentMovementLine", "DocumentID,LineNo", True, False
    AT_EnsureIndex db, "tblAccountMovements", "IX_AccountMovementDate", "AccountID,MovementDate", False, False
    AT_EnsureIndex db, "tblOwners", "IX_OwnerName", "FullName", False, False
    AT_EnsureIndex db, "tblOwners", "IX_OwnerPhone", "Phone", False, False
    AT_EnsureIndex db, "tblOwnerCharges", "IX_ChargeOwnerDate", "OwnerID,ChargeDate", False, False
    AT_EnsureIndex db, "tblFinancialDocuments", "IX_DocumentDateTypeStatus", "TransactionDate,DocumentType,StatusCode", False, False
    AT_EnsureIndex db, "tblAuditEvents", "IX_AuditTimeAction", "EventAt,ActionCode", False, False
End Sub

Private Sub AT_EnsureIndex(ByVal db As DAO.Database, ByVal tableName As String, ByVal indexName As String, ByVal fieldList As String, ByVal isUnique As Boolean, ByVal ignoreNulls As Boolean)
    Dim idx As DAO.Index, f As Variant
    If AT_IndexExists(db, tableName, indexName) Then Exit Sub
    Set idx = db.TableDefs(tableName).CreateIndex(indexName)
    idx.Unique = isUnique
    If isUnique And ignoreNulls Then idx.IgnoreNulls = True
    For Each f In Split(fieldList, ",")
        idx.Fields.Append idx.CreateField(Trim$(CStr(f)))
    Next f
    db.TableDefs(tableName).Indexes.Append idx
End Sub

Private Sub AT_CreateRelations(ByVal db As DAO.Database)
    Dim links As Variant, item As Variant, bits As Variant
    links = Array( _
        "relOwnersCreatedBy|tblOwners|CreatedByUserID|tblUsers|UserID", _
        "relOwnersUpdatedBy|tblOwners|UpdatedByUserID|tblUsers|UserID", _
        "relChargeBatchRate|tblChargeBatches|RateID|tblContributionRates|RateID", _
        "relChargeBatchCreatedBy|tblChargeBatches|CreatedByUserID|tblUsers|UserID", _
        "relChargeBatchApprovedBy|tblChargeBatches|ApprovedByUserID|tblUsers|UserID", _
        "relDocumentReversalOf|tblFinancialDocuments|ReversalOfDocumentID|tblFinancialDocuments|DocumentID", _
        "relChargeOwner|tblOwnerCharges|OwnerID|tblOwners|OwnerID", _
        "relChargeBatch|tblOwnerCharges|ChargeBatchID|tblChargeBatches|ChargeBatchID", _
        "relChargeDocument|tblOwnerCharges|DocumentID|tblFinancialDocuments|DocumentID", _
        "relOfficialNumberDocument|tblOfficialDocumentNumbers|DocumentID|tblFinancialDocuments|DocumentID", _
        "relReceiptDocument|tblReceipts|DocumentID|tblFinancialDocuments|DocumentID", _
        "relReceiptOwner|tblReceipts|OwnerID|tblOwners|OwnerID", _
        "relReceiptAccount|tblReceipts|AccountID|tblAccounts|AccountID", _
        "relReceiptMethod|tblReceipts|PaymentMethodID|tblPaymentMethods|PaymentMethodID", _
        "relAllocationReceipt|tblReceiptAllocationLines|ReceiptID|tblReceipts|ReceiptID", _
        "relAllocationCharge|tblReceiptAllocationLines|ChargeID|tblOwnerCharges|ChargeID", _
        "relAdjustmentDocument|tblOwnerAdjustments|DocumentID|tblFinancialDocuments|DocumentID", _
        "relAdjustmentCharge|tblOwnerAdjustments|ChargeID|tblOwnerCharges|ChargeID", _
        "relSettlementCharge|tblChargeSettlementLines|ChargeID|tblOwnerCharges|ChargeID", _
        "relSettlementDocument|tblChargeSettlementLines|DocumentID|tblFinancialDocuments|DocumentID", _
        "relSettlementReceipt|tblChargeSettlementLines|ReceiptID|tblReceipts|ReceiptID", _
        "relSettlementCreditApplication|tblChargeSettlementLines|CreditApplicationID|tblCreditApplications|CreditApplicationID", _
        "relSettlementCorrection|tblChargeSettlementLines|ReceiptCorrectionID|tblReceiptCorrections|ReceiptCorrectionID", _
        "relCorrectionDocument|tblReceiptCorrections|DocumentID|tblFinancialDocuments|DocumentID", _
        "relCorrectionReceipt|tblReceiptCorrections|OriginalReceiptID|tblReceipts|ReceiptID", _
        "relCorrectionLineHeader|tblReceiptCorrectionLines|ReceiptCorrectionID|tblReceiptCorrections|ReceiptCorrectionID", _
        "relCorrectionLineAllocation|tblReceiptCorrectionLines|OriginalAllocationID|tblReceiptAllocationLines|AllocationID", _
        "relCreditMovementOwner|tblOwnerCreditMovements|OwnerID|tblOwners|OwnerID", _
        "relCreditMovementDocument|tblOwnerCreditMovements|DocumentID|tblFinancialDocuments|DocumentID", _
        "relCreditMovementReceipt|tblOwnerCreditMovements|SourceReceiptID|tblReceipts|ReceiptID", _
        "relCreditMovementSource|tblOwnerCreditMovements|SourceCreditMovementID|tblOwnerCreditMovements|CreditMovementID", _
        "relCreditApplicationDocument|tblCreditApplications|DocumentID|tblFinancialDocuments|DocumentID", _
        "relCreditApplicationOwner|tblCreditApplications|OwnerID|tblOwners|OwnerID", _
        "relCreditLineApplication|tblCreditApplicationLines|CreditApplicationID|tblCreditApplications|CreditApplicationID", _
        "relCreditLineMovement|tblCreditApplicationLines|CreditMovementID|tblOwnerCreditMovements|CreditMovementID", _
        "relCreditLineCharge|tblCreditApplicationLines|ChargeID|tblOwnerCharges|ChargeID", _
        "relPaymentDocument|tblPayments|DocumentID|tblFinancialDocuments|DocumentID", _
        "relPaymentAccount|tblPayments|AccountID|tblAccounts|AccountID", _
        "relPaymentMethod|tblPayments|PaymentMethodID|tblPaymentMethods|PaymentMethodID", _
        "relPaymentCategory|tblPayments|ExpenseCategoryID|tblExpenseCategories|ExpenseCategoryID", _
        "relPaymentRefundCredit|tblPayments|RefundOfCreditMovementID|tblOwnerCreditMovements|CreditMovementID", _
        "relPaymentRefundDocument|tblPayments|RefundOfDocumentID|tblFinancialDocuments|DocumentID", _
        "relTransferDocument|tblAccountTransfers|DocumentID|tblFinancialDocuments|DocumentID", _
        "relTransferSource|tblAccountTransfers|SourceAccountID|tblAccounts|AccountID", _
        "relTransferDestination|tblAccountTransfers|DestinationAccountID|tblAccounts|AccountID", _
        "relMovementAccount|tblAccountMovements|AccountID|tblAccounts|AccountID", _
        "relMovementDocument|tblAccountMovements|DocumentID|tblFinancialDocuments|DocumentID", _
        "relReservationAccount|tblAccountReservations|AccountID|tblAccounts|AccountID", _
        "relReservationDocument|tblAccountReservations|DocumentID|tblFinancialDocuments|DocumentID", _
        "relClosingAccountHeader|tblFinancialClosingAccounts|ClosingID|tblFinancialClosings|ClosingID", _
        "relClosingAccountAccount|tblFinancialClosingAccounts|AccountID|tblAccounts|AccountID", _
        "relAuditChangeEvent|tblAuditChanges|AuditEventID|tblAuditEvents|AuditEventID", _
        "relAttachmentDocument|tblAttachments|DocumentID|tblFinancialDocuments|DocumentID", _
        "relIntegrityFindingRun|tblIntegrityFindings|IntegrityRunID|tblIntegrityRuns|IntegrityRunID", _
        "relUserRoleUser|tblUserRoleAssignments|UserID|tblUsers|UserID", _
        "relUserRoleRole|tblUserRoleAssignments|RoleID|tblRoles|RoleID", _
        "relUserRoleAssignedBy|tblUserRoleAssignments|AssignedByUserID|tblUsers|UserID", _
        "relRolePermissionRole|tblRolePermissions|RoleID|tblRoles|RoleID", _
        "relRolePermissionPermission|tblRolePermissions|PermissionID|tblPermissions|PermissionID" _
    )
    For Each item In links
        bits = Split(CStr(item), "|")
        AT_EnsureRelation db, CStr(bits(0)), CStr(bits(1)), CStr(bits(2)), CStr(bits(3)), CStr(bits(4))
    Next item
End Sub

Private Sub AT_EnsureRelation(ByVal db As DAO.Database, ByVal relName As String, ByVal childTable As String, ByVal childField As String, ByVal parentTable As String, ByVal parentField As String)
    Dim rel As DAO.Relation, fld As DAO.Field
    On Error Resume Next
    Set rel = db.Relations(relName)
    If Not rel Is Nothing Then Exit Sub
    Err.Clear
    On Error GoTo 0
    Set rel = db.CreateRelation(relName, parentTable, childTable, 0)
    Set fld = rel.CreateField(parentField)
    fld.ForeignName = childField
    rel.Fields.Append fld
    db.Relations.Append rel
End Sub

Private Sub AT_SeedDefaults(ByVal db As DAO.Database)
    If DCount("*", "tblProjectSettings") = 0 Then db.Execute "INSERT INTO tblProjectSettings (ProjectName, NeighborhoodName, DefaultRate, CurrencyCode, DefaultCollectorName, CreatedAt) VALUES ('مشروع مساهمات المجاري','اسم الحارة',3000,'YER','',Now())", dbFailOnError
    If DCount("*", "tblContributionRates") = 0 Then db.Execute "INSERT INTO tblContributionRates (RateCode, RateAmount, EffectiveFrom, IsActive, CreatedAt) VALUES ('RATE-0001',3000,#1/1/2020#,True,Now())", dbFailOnError
    If DCount("*", "tblNumberSeries") = 0 Then
        db.Execute "INSERT INTO tblNumberSeries (SeriesCode, NextValue, Prefix, PadLength) VALUES ('OWNER',1,'OW-',4)", dbFailOnError
        db.Execute "INSERT INTO tblNumberSeries (SeriesCode, NextValue, Prefix, PadLength) VALUES ('RECEIPT',1,'RC-',6)", dbFailOnError
        db.Execute "INSERT INTO tblNumberSeries (SeriesCode, NextValue, Prefix, PadLength) VALUES ('PAYMENT',1,'PV-',6)", dbFailOnError
        db.Execute "INSERT INTO tblNumberSeries (SeriesCode, NextValue, Prefix, PadLength) VALUES ('TRANSFER',1,'TR-',6)", dbFailOnError
        db.Execute "INSERT INTO tblNumberSeries (SeriesCode, NextValue, Prefix, PadLength) VALUES ('CHARGE',1,'CH-',8)", dbFailOnError
    End If
    If DCount("*", "tblAccounts") = 0 Then db.Execute "INSERT INTO tblAccounts (AccountCode, AccountName, AccountType, IsActive, CreatedAt) VALUES ('CASH-01','الصندوق النقدي','CASH',True,Now())", dbFailOnError
    If DCount("*", "tblPaymentMethods") = 0 Then
        db.Execute "INSERT INTO tblPaymentMethods (MethodCode, MethodName, IsActive, CreatedAt) VALUES ('CASH','نقدًا',True,Now())", dbFailOnError
        db.Execute "INSERT INTO tblPaymentMethods (MethodCode, MethodName, IsActive, CreatedAt) VALUES ('BANK','تحويل بنكي',True,Now())", dbFailOnError
        db.Execute "INSERT INTO tblPaymentMethods (MethodCode, MethodName, IsActive, CreatedAt) VALUES ('EWALLET','محفظة إلكترونية',True,Now())", dbFailOnError
    End If
    If DCount("*", "tblExpenseCategories") = 0 Then
        db.Execute "INSERT INTO tblExpenseCategories (CategoryCode, CategoryName, IsActive) VALUES ('GENERAL','مصروفات عامة',True)", dbFailOnError
        db.Execute "INSERT INTO tblExpenseCategories (CategoryCode, CategoryName, IsActive) VALUES ('TRANSFER_FEE','رسوم تحويل',True)", dbFailOnError
        db.Execute "INSERT INTO tblExpenseCategories (CategoryCode, CategoryName, IsActive) VALUES ('MAINTENANCE','صيانة',True)", dbFailOnError
    End If
    If DCount("*", "tblRoles") = 0 Then
        db.Execute "INSERT INTO tblRoles (RoleCode, RoleName, IsActive, CreatedAt) VALUES ('ADMIN','مدير النظام',True,Now())", dbFailOnError
        db.Execute "INSERT INTO tblRoles (RoleCode, RoleName, IsActive, CreatedAt) VALUES ('COLLECTOR','محصل',True,Now())", dbFailOnError
        db.Execute "INSERT INTO tblRoles (RoleCode, RoleName, IsActive, CreatedAt) VALUES ('VIEWER','قارئ تقارير',True,Now())", dbFailOnError
    End If
    If DCount("*", "tblPermissions") = 0 Then AT_SeedPermissions db
    AT_SeedRolePermissions db
    AT_SeedLocalWindowsUser db
End Sub

Private Sub AT_SeedPermissions(ByVal db As DAO.Database)
    Dim codes As Variant, names As Variant, i As Long
    codes = Array("OWNER_VIEW", "OWNER_EDIT", "OWNER_APPROVE", "CHARGE_DRAFT", "CHARGE_APPROVE", "RECEIPT_DRAFT", "RECEIPT_APPROVE", "PAYMENT_DRAFT", "PAYMENT_APPROVE", "ADJUSTMENT_REQUEST", "ADJUSTMENT_APPROVE", "CREDIT_APPLY", "TRANSFER_DRAFT", "TRANSFER_APPROVE", "DOCUMENT_CANCEL", "REPORT_VIEW", "FINANCIAL_CLOSE", "BACKUP_RESTORE", "SETTINGS_ADMIN", "AUDIT_VIEW", "INTEGRITY_RUN")
    names = Array("عرض الملاك", "تعديل بيانات الملاك", "اعتماد المالك", "إنشاء مسودة استحقاق", "اعتماد الاستحقاق", "إنشاء مسودة قبض", "اعتماد القبض", "إنشاء مسودة صرف", "اعتماد الصرف", "طلب خصم/إعفاء", "اعتماد الخصومات", "تطبيق الرصيد الدائن", "إنشاء مسودة تحويل", "اعتماد التحويلات", "إلغاء المسودات", "عرض التقارير", "الإقفال المالي", "النسخ والاستعادة", "إدارة الإعدادات", "عرض التدقيق", "فحص سلامة البيانات")
    For i = LBound(codes) To UBound(codes)
        Dim rs As DAO.Recordset
        Set rs = db.OpenRecordset("tblPermissions", dbOpenDynaset)
        rs.AddNew: rs!PermissionCode = CStr(codes(i)): rs!PermissionName = CStr(names(i)): rs!IsActive = True: rs.Update
        rs.Close
    Next i
End Sub

Private Sub AT_SeedRolePermissions(ByVal db As DAO.Database)
    Dim adminID As Long, collectorID As Long, viewerID As Long, permID As Long
    Dim rs As DAO.Recordset, roleCode As Variant, permCode As Variant, allowed As Boolean
    If DCount("*", "tblRolePermissions") > 0 Then Exit Sub
    adminID = Nz(DLookup("RoleID", "tblRoles", "RoleCode='ADMIN'"), 0)
    collectorID = Nz(DLookup("RoleID", "tblRoles", "RoleCode='COLLECTOR'"), 0)
    viewerID = Nz(DLookup("RoleID", "tblRoles", "RoleCode='VIEWER'"), 0)
    For Each permCode In Array("OWNER_VIEW", "OWNER_EDIT", "OWNER_APPROVE", "CHARGE_DRAFT", "CHARGE_APPROVE", "RECEIPT_DRAFT", "RECEIPT_APPROVE", "PAYMENT_DRAFT", "PAYMENT_APPROVE", "ADJUSTMENT_REQUEST", "ADJUSTMENT_APPROVE", "CREDIT_APPLY", "TRANSFER_DRAFT", "TRANSFER_APPROVE", "DOCUMENT_CANCEL", "REPORT_VIEW", "FINANCIAL_CLOSE", "BACKUP_RESTORE", "SETTINGS_ADMIN", "AUDIT_VIEW", "INTEGRITY_RUN")
        permID = Nz(DLookup("PermissionID", "tblPermissions", "PermissionCode='" & CStr(permCode) & "'"), 0)
        If adminID > 0 And permID > 0 Then AT_GrantRolePermission db, adminID, permID
        allowed = (CStr(permCode) = "OWNER_VIEW" Or CStr(permCode) = "RECEIPT_DRAFT" Or CStr(permCode) = "PAYMENT_DRAFT" Or CStr(permCode) = "TRANSFER_DRAFT" Or CStr(permCode) = "DOCUMENT_CANCEL" Or CStr(permCode) = "ADJUSTMENT_REQUEST" Or CStr(permCode) = "REPORT_VIEW")
        If collectorID > 0 And permID > 0 And allowed Then AT_GrantRolePermission db, collectorID, permID
        If viewerID > 0 And permID > 0 And (CStr(permCode) = "OWNER_VIEW" Or CStr(permCode) = "REPORT_VIEW") Then AT_GrantRolePermission db, viewerID, permID
    Next permCode
End Sub

Private Sub AT_GrantRolePermission(ByVal db As DAO.Database, ByVal roleID As Long, ByVal permissionID As Long)
    Dim rs As DAO.Recordset
    If DCount("*", "tblRolePermissions", "RoleID=" & roleID & " AND PermissionID=" & permissionID) > 0 Then Exit Sub
    Set rs = db.OpenRecordset("tblRolePermissions", dbOpenDynaset)
    rs.AddNew: rs!RoleID = roleID: rs!PermissionID = permissionID: rs!GrantedAt = Now(): rs.Update: rs.Close
End Sub

Private Sub AT_SeedLocalWindowsUser(ByVal db As DAO.Database)
    Dim userName As String, userID As Long, roleID As Long, rs As DAO.Recordset
    userName = Environ$("Username")
    If Len(Trim$(userName)) = 0 Then Exit Sub
    userID = Nz(DLookup("UserID", "tblUsers", "UserName='" & Replace(userName, "'", "''") & "'"), 0)
    If userID = 0 Then
        Set rs = db.OpenRecordset("tblUsers", dbOpenDynaset)
        rs.AddNew: rs!UserName = userName: rs!DisplayName = userName: rs!AuthenticationMode = "WINDOWS_LOCAL": rs!IsActive = True: rs!CreatedAt = Now(): rs.Update: rs.Bookmark = rs.LastModified: userID = CLng(rs!UserID): rs.Close
    End If
    roleID = Nz(DLookup("RoleID", "tblRoles", "RoleCode='ADMIN'"), 0)
    If roleID > 0 And DCount("*", "tblUserRoleAssignments", "UserID=" & userID & " AND RoleID=" & roleID) = 0 Then
        Set rs = db.OpenRecordset("tblUserRoleAssignments", dbOpenDynaset)
        rs.AddNew: rs!UserID = userID: rs!RoleID = roleID: rs!AssignedAt = Now(): rs!IsActive = True: rs.Update: rs.Close
    End If
End Sub

Private Sub AT_SetSchemaVersion(ByVal db As DAO.Database)
    If DCount("*", "tblSchemaVersion") = 0 Then
        db.Execute "INSERT INTO tblSchemaVersion (VersionCode, InstalledAt, InstalledByUserName, Notes) VALUES ('" & AT_SCHEMA_VERSION & "',Now(),'" & Replace(Environ$("Username"), "'", "''") & "','Initial schema build; UI and runtime verification still required.')", dbFailOnError
    End If
End Sub

Public Sub AT_Audit(ByVal actionCode As String, ByVal entityName As String, ByVal entityID As Long, ByVal resultCode As String, ByVal detail As String, Optional ByVal reason As String = "")
    On Error GoTo SafeExit
    If Not AT_TableExists(CurrentDb, "tblAuditEvents") Then Exit Sub
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("tblAuditEvents", dbOpenDynaset)
    rs.AddNew
    rs!EventAt = Now(): rs!UserName = Environ$("Username")
    rs!ActionCode = Left$(actionCode, 60): rs!EntityName = Left$(entityName, 60)
    rs!EntityID = entityID: rs!ResultCode = Left$(resultCode, 30)
    rs!Detail = detail: rs!Reason = reason
    rs.Update: rs.Close
SafeExit:
    ' Best-effort audit must never mask the business-operation error.
End Sub
