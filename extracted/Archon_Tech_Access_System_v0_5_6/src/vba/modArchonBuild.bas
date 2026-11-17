Attribute VB_Name = "modArchonBuild"
Option Compare Database
Option Explicit

' Archon Tech | Sewer Contributions Access Builder
' Initial-install builder for a BLANK .accdb. It is not a migration tool.
' Must be compiled and executed inside desktop Microsoft Access on Windows.

Private Const AT_SCHEMA_VERSION As String = "0.5.0"

Public Sub AT_BuildDatabase()
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    Dim db As DAO.Database
    Set db = CurrentDb
    If AT_HasPriorSchema(db) Then
        MsgBox AT_U("062A,0648,062C,062F,0020,0642,0627,0639,062F,0629,0020,0633,0628,0642,0020,0625,0639,062F,0627,062F,0647,0627,002E,0020,0647,0630,0627,0020,0627,0644,0625,062C,0631,0627,0621,0020,0644,0644,062A,062B,0628,064A,062A,0020,0627,0644,0623,0648,0644,064A,0020,0641,0642,0637,060C,0020,0648,0644,0645,0020,064A,064F,062C,0631,0020,0623,064A,0020,062A,0639,062F,064A,0644,002E"), vbExclamation, "Archon Tech"
        Exit Sub
    End If
    AT_CreateSchemaOnDatabase db
    AT_SeedDefaults db
    AT_SetSchemaVersion db
    AT_Audit "BUILD_SCHEMA", "Database", 0, "SUCCESS", "Schema version " & AT_SCHEMA_VERSION
    MsgBox AT_U("0627,0643,062A,0645,0644,0020,0625,0646,0634,0627,0621,0020,0645,062E,0637,0637,0020,0627,0644,0628,064A,0627,0646,0627,062A,0020,0648,0627,0644,0641,0647,0627,0631,0633,0020,0648,0627,0644,0639,0644,0627,0642,0627,062A,002E,0020,064A,0644,0632,0645,0020,062A,0646,0641,064A,0630,0020,0627,062E,062A,0628,0627,0631,0627,062A,0020,0041,0063,0063,0065,0073,0073,0020,0642,0628,0644,0020,0623,064A,0020,0628,064A,0627,0646,0627,062A,0020,062D,0642,064A,0642,064A,0629,002E"), vbInformation, "Archon Tech"
    Exit Sub
EH:

    savedErrNumber = Err.Number: eDesc = Err.Description
    MsgBox AT_U("062A,0648,0642,0641,0020,0627,0644,0628,0646,0627,0621,002E,0020,0627,0644,062E,0637,0623,0020") & savedErrNumber & ": " & eDesc, vbCritical, "Archon Tech"
    Err.Raise savedErrNumber, "AT_BuildDatabase", eDesc
End Sub

Public Sub AT_CreateSchemaOnDatabase(ByVal db As DAO.Database)
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_ExecuteDDLBatch db, Array( _
        "CREATE TABLE tblProjectSettings (ProjectSettingID COUNTER CONSTRAINT PK_tblProjectSettings PRIMARY KEY, ProjectName TEXT(150), NeighborhoodName TEXT(150), DefaultRate CURRENCY, CurrencyCode TEXT(10), DefaultCollectorName TEXT(150), CreatedAt DATETIME, UpdatedAt DATETIME)", _
        "CREATE TABLE tblContributionRates (RateID COUNTER CONSTRAINT PK_tblContributionRates PRIMARY KEY, RateCode TEXT(30), RateAmount CURRENCY, EffectiveFrom DATETIME, EffectiveTo DATETIME, IsActive YESNO, CreatedAt DATETIME)", _
        "CREATE TABLE tblUsers (UserID COUNTER CONSTRAINT PK_tblUsers PRIMARY KEY, UserName TEXT(100), DisplayName TEXT(150), AuthenticationMode TEXT(30), IsActive YESNO, CreatedAt DATETIME, LastLoginAt DATETIME)", _
        "CREATE TABLE tblRoles (RoleID COUNTER CONSTRAINT PK_tblRoles PRIMARY KEY, RoleCode TEXT(40), RoleName TEXT(100), IsActive YESNO, CreatedAt DATETIME)", _
        "CREATE TABLE tblPermissions (PermissionID COUNTER CONSTRAINT PK_tblPermissions PRIMARY KEY, PermissionCode TEXT(60), PermissionName TEXT(150), IsActive YESNO)", _
        "CREATE TABLE tblUserRoleAssignments (UserRoleAssignmentID COUNTER CONSTRAINT PK_tblUserRoleAssignments PRIMARY KEY, UserID LONG, RoleID LONG, AssignedAt DATETIME, AssignedByUserID LONG, IsActive YESNO)", _
        "CREATE TABLE tblRolePermissions (RolePermissionID COUNTER CONSTRAINT PK_tblRolePermissions PRIMARY KEY, RoleID LONG, PermissionID LONG, GrantedAt DATETIME)", _
        "CREATE TABLE tblOwners (OwnerID COUNTER CONSTRAINT PK_tblOwners PRIMARY KEY, OwnerCode TEXT(12), FullName TEXT(200), FullNameSearch TEXT(200), Phone TEXT(40), IdentityNo TEXT(80), HouseCount LONG, ApartmentCount LONG, Notes MEMO, AdminStatus TEXT(20), DataStatus TEXT(20), ReviewStatus TEXT(20), ActivityStatus TEXT(20), CreatedAt DATETIME, UpdatedAt DATETIME, CreatedByUserID LONG, UpdatedByUserID LONG)", _
        "CREATE TABLE tblChargeBatches (ChargeBatchID COUNTER CONSTRAINT PK_tblChargeBatches PRIMARY KEY, BatchCode TEXT(30), PeriodLabel TEXT(100), ChargeDate DATETIME, RateID LONG, StatusCode TEXT(20), CreatedAt DATETIME, CreatedByUserID LONG, ApprovedAt DATETIME, ApprovedByUserID LONG, Reason MEMO)", _
        "CREATE TABLE tblFinancialDocuments (DocumentID COUNTER CONSTRAINT PK_tblFinancialDocuments PRIMARY KEY, DocumentType TEXT(30), TransactionDate DATETIME, StatusCode TEXT(20), CreatedAt DATETIME, CreatedByUserName TEXT(100), ApprovedAt DATETIME, ApprovedByUserName TEXT(100), Reason MEMO, CancelReason MEMO, ReversalOfDocumentID LONG)", _
        "CREATE TABLE tblOfficialDocumentNumbers (OfficialDocumentNumberID COUNTER CONSTRAINT PK_tblOfficialDocumentNumbers PRIMARY KEY, DocumentID LONG, OfficialNumber TEXT(20), SeriesCode TEXT(30), NumericValue LONG, AssignedAt DATETIME, AssignedByUserName TEXT(100))", _
        "CREATE TABLE tblNumberSeries (SeriesID COUNTER CONSTRAINT PK_tblNumberSeries PRIMARY KEY, SeriesCode TEXT(30), NextValue LONG, Prefix TEXT(10), PadLength LONG)", _
        "CREATE TABLE tblOwnerCharges (ChargeID COUNTER CONSTRAINT PK_tblOwnerCharges PRIMARY KEY, ChargeCode TEXT(20), DocumentID LONG, OwnerID LONG, ChargeBatchID LONG, CampaignID LONG, BuildingID LONG, ChargeDate DATETIME, PeriodStart DATETIME, PeriodEnd DATETIME, ApartmentCountAtCharge LONG, RateAmount CURRENCY, GrossAmount CURRENCY, DiscountAmount CURRENCY, NetAmount CURRENCY, OwnerCodeAtCharge TEXT(12), OwnerNameAtCharge TEXT(200), ChargeStatus TEXT(20), CreatedAt DATETIME)", _
        "CREATE TABLE tblReceipts (ReceiptID COUNTER CONSTRAINT PK_tblReceipts PRIMARY KEY, DocumentID LONG, OwnerID LONG, Amount CURRENCY, PaymentMethodID LONG, AccountID LONG, PayerName TEXT(150), CollectorName TEXT(150), Notes MEMO, PrintedOwnerName TEXT(200), PrintedOwnerCode TEXT(12), PrintedNeighborhoodName TEXT(150), CreatedAt DATETIME)", _
        "CREATE TABLE tblReceiptAllocationLines (AllocationID COUNTER CONSTRAINT PK_tblReceiptAllocationLines PRIMARY KEY, ReceiptID LONG, ChargeID LONG, AllocatedAmount CURRENCY, LineNo LONG)", _
        "CREATE TABLE tblOwnerAdjustments (AdjustmentID COUNTER CONSTRAINT PK_tblOwnerAdjustments PRIMARY KEY, DocumentID LONG, ChargeID LONG, AdjustmentType TEXT(30), SignedAmount CURRENCY, Reason MEMO, RequestedByUserName TEXT(100), ApprovedByUserName TEXT(100), ApprovedAt DATETIME)", _
        "CREATE TABLE tblChargeSettlementLines (SettlementID COUNTER CONSTRAINT PK_tblChargeSettlementLines PRIMARY KEY, ChargeID LONG, DocumentID LONG, ReceiptID LONG, CreditApplicationID LONG, ReceiptCorrectionID LONG, SignedAmount CURRENCY, SettlementType TEXT(30), ReversalOfSettlementID LONG, CreatedAt DATETIME)", _
        "CREATE TABLE tblReceiptCorrections (ReceiptCorrectionID COUNTER CONSTRAINT PK_tblReceiptCorrections PRIMARY KEY, DocumentID LONG, OriginalReceiptID LONG, CorrectionType TEXT(30), Reason MEMO)" _
    )
    AT_ExecuteDDLBatch db, Array( _
        "CREATE TABLE tblReceiptCorrectionLines (ReceiptCorrectionLineID COUNTER CONSTRAINT PK_tblReceiptCorrectionLines PRIMARY KEY, ReceiptCorrectionID LONG, OriginalAllocationID LONG, Amount CURRENCY, CreatedAt DATETIME)", _
        "CREATE TABLE tblOwnerCreditMovements (CreditMovementID COUNTER CONSTRAINT PK_tblOwnerCreditMovements PRIMARY KEY, OwnerID LONG, DocumentID LONG, SignedAmount CURRENCY, CreditType TEXT(30), SourceReceiptID LONG, SourceCreditMovementID LONG, Reason MEMO, CreatedAt DATETIME)", _
        "CREATE TABLE tblCreditApplications (CreditApplicationID COUNTER CONSTRAINT PK_tblCreditApplications PRIMARY KEY, DocumentID LONG, OwnerID LONG, Amount CURRENCY, ApprovedByUserName TEXT(100), ApprovedAt DATETIME, Reason MEMO)", _
        "CREATE TABLE tblCreditApplicationLines (CreditApplicationLineID COUNTER CONSTRAINT PK_tblCreditApplicationLines PRIMARY KEY, CreditApplicationID LONG, CreditMovementID LONG, ChargeID LONG, Amount CURRENCY)", _
        "CREATE TABLE tblAccounts (AccountID COUNTER CONSTRAINT PK_tblAccounts PRIMARY KEY, AccountCode TEXT(20), AccountName TEXT(120), AccountType TEXT(20), IsActive YESNO, CreatedAt DATETIME)", _
        "CREATE TABLE tblPaymentMethods (PaymentMethodID COUNTER CONSTRAINT PK_tblPaymentMethods PRIMARY KEY, MethodCode TEXT(30), MethodName TEXT(100), IsActive YESNO, CreatedAt DATETIME)", _
        "CREATE TABLE tblExpenseCategories (ExpenseCategoryID COUNTER CONSTRAINT PK_tblExpenseCategories PRIMARY KEY, CategoryCode TEXT(30), CategoryName TEXT(120), IsActive YESNO, CreatedAt DATETIME)", _
        "CREATE TABLE tblPayments (PaymentID COUNTER CONSTRAINT PK_tblPayments PRIMARY KEY, VoucherNo TEXT(24), DocumentID LONG, ExpenseID LONG, AccountID LONG, ExpenseCategoryID LONG, BeneficiaryName TEXT(150), Amount CURRENCY, PaymentMethodID LONG, RecipientName TEXT(150), Notes MEMO, RefundOfCreditMovementID LONG, RefundOfDocumentID LONG, ReversalOfPaymentID LONG)", _
        "CREATE TABLE tblAccountTransfers (TransferID COUNTER CONSTRAINT PK_tblAccountTransfers PRIMARY KEY, DocumentID LONG, SourceAccountID LONG, DestinationAccountID LONG, Amount CURRENCY, FeeAmount CURRENCY, ConfirmedAt DATETIME, ConfirmedByUserName TEXT(100), ConfirmationRef TEXT(100), Notes MEMO)", _
        "CREATE TABLE tblAccountMovements (AccountMovementID COUNTER CONSTRAINT PK_tblAccountMovements PRIMARY KEY, AccountID LONG, DocumentID LONG, MovementDate DATETIME, SignedAmount CURRENCY, MovementType TEXT(30), LineNo LONG, CreatedAt DATETIME)", _
        "CREATE TABLE tblAccountReservations (ReservationID COUNTER CONSTRAINT PK_tblAccountReservations PRIMARY KEY, DocumentID LONG, AccountID LONG, Amount CURRENCY, StatusCode TEXT(20), CreatedAt DATETIME, ReleasedAt DATETIME, Reason MEMO)", _
        "CREATE TABLE tblFinancialClosings (ClosingID COUNTER CONSTRAINT PK_tblFinancialClosings PRIMARY KEY, PeriodID LONG, PeriodStart DATETIME, PeriodEnd DATETIME, ClosedAt DATETIME, ClosedByUserName TEXT(100), StatusCode TEXT(20), DiscrepancyException YESNO, ExceptionReason MEMO, Summary MEMO, ReopenedAt DATETIME, ReopenedByUserName TEXT(100), ReopenReason MEMO)", _
        "CREATE TABLE tblFinancialClosingAccounts (ClosingAccountID COUNTER CONSTRAINT PK_tblFinancialClosingAccounts PRIMARY KEY, ClosingID LONG, AccountID LONG, BookBalance CURRENCY, ActualBalance CURRENCY, DifferenceAmount CURRENCY, Notes MEMO)", _
        "CREATE TABLE tblAuditEvents (AuditEventID COUNTER CONSTRAINT PK_tblAuditEvents PRIMARY KEY, EventAt DATETIME, UserName TEXT(100), ActionCode TEXT(60), EntityName TEXT(60), EntityID LONG, ResultCode TEXT(30), Reason MEMO, Detail MEMO)", _
        "CREATE TABLE tblAuditChanges (AuditChangeID COUNTER CONSTRAINT PK_tblAuditChanges PRIMARY KEY, AuditEventID LONG, FieldName TEXT(100), OldValue MEMO, NewValue MEMO)", _
        "CREATE TABLE tblAttachments (AttachmentID COUNTER CONSTRAINT PK_tblAttachments PRIMARY KEY, DocumentID LONG, RelativePath MEMO, OriginalFileName TEXT(255), ContentType TEXT(100), FileSizeBytes DOUBLE, SHA256 TEXT(64), AddedAt DATETIME, AddedByUserName TEXT(100))", _
        "CREATE TABLE tblBackupLog (BackupLogID COUNTER CONSTRAINT PK_tblBackupLog PRIMARY KEY, StartedAt DATETIME, CompletedAt DATETIME, OperationType TEXT(20), SourcePath MEMO, DestinationPath MEMO, ResultCode TEXT(20), FileSizeBytes DOUBLE, Detail MEMO)", _
        "CREATE TABLE tblIntegrityRuns (IntegrityRunID COUNTER CONSTRAINT PK_tblIntegrityRuns PRIMARY KEY, StartedAt DATETIME, CompletedAt DATETIME, RunByUserName TEXT(100), ResultSummary MEMO)" _
    )
    AT_ExecuteDDLBatch db, Array( _
        "CREATE TABLE tblIntegrityFindings (FindingID COUNTER CONSTRAINT PK_tblIntegrityFindings PRIMARY KEY, IntegrityRunID LONG, Severity TEXT(20), EntityName TEXT(60), EntityID LONG, FindingCode TEXT(60), FindingText MEMO, ResolutionStatus TEXT(20), ResolutionReason MEMO, ResolvedByUserName TEXT(100), ResolvedAt DATETIME)", _
        "CREATE TABLE tblSchemaVersion (SchemaVersionID COUNTER CONSTRAINT PK_tblSchemaVersion PRIMARY KEY, VersionCode TEXT(20), InstalledAt DATETIME, InstalledByUserName TEXT(100), Notes MEMO)", _
        "CREATE TABLE tblBuildings (BuildingID COUNTER CONSTRAINT PK_tblBuildings PRIMARY KEY, BuildingCode TEXT(30), BuildingName TEXT(150), Address TEXT(255), IsActive YESNO, CreatedAt DATETIME, UpdatedAt DATETIME, Notes MEMO)", _
        "CREATE TABLE tblBuildingOwnership (OwnershipID COUNTER CONSTRAINT PK_tblBuildingOwnership PRIMARY KEY, BuildingID LONG, OwnerID LONG, StartDate DATETIME, EndDate DATETIME, TransferDocumentNo TEXT(100), Reason MEMO, CreatedAt DATETIME, CreatedByUserName TEXT(100))", _
        "CREATE TABLE tblUnits (UnitID COUNTER CONSTRAINT PK_tblUnits PRIMARY KEY, BuildingID LONG, UnitCode TEXT(30), FloorNo TEXT(30), IsActive YESNO, IsContributionEligible YESNO, Notes MEMO, CreatedAt DATETIME, UpdatedAt DATETIME)", _
        "CREATE TABLE tblUnitOccupancy (OccupancyID COUNTER CONSTRAINT PK_tblUnitOccupancy PRIMARY KEY, UnitID LONG, FamilyName TEXT(150), HouseholdCount LONG, StartDate DATETIME, EndDate DATETIME, Notes MEMO, CreatedAt DATETIME)", _
        "CREATE TABLE tblUnitExemptionRequests (ExemptionRequestID COUNTER CONSTRAINT PK_tblUnitExemptionRequests PRIMARY KEY, UnitID LONG, RequestDate DATETIME, EffectiveFrom DATETIME, EffectiveTo DATETIME, Reason MEMO, EvidenceReference TEXT(255), RequestedByUserName TEXT(100), StatusCode TEXT(20), DecidedByUserName TEXT(100), DecidedAt DATETIME, DecisionReason MEMO, CreatedAt DATETIME)", _
        "CREATE TABLE tblContributionTypes (ContributionTypeID COUNTER CONSTRAINT PK_tblContributionTypes PRIMARY KEY, TypeCode TEXT(30), TypeName TEXT(120), CalculationMethod TEXT(30), IsBaseType YESNO, IsActive YESNO, CreatedAt DATETIME)", _
        "CREATE TABLE tblContributionCampaigns (CampaignID COUNTER CONSTRAINT PK_tblContributionCampaigns PRIMARY KEY, CampaignCode TEXT(30), CampaignName TEXT(150), ContributionTypeID LONG, TargetScope TEXT(30), RateID LONG, AmountPerUnit CURRENCY, PeriodStart DATETIME, PeriodEnd DATETIME, DueDate DATETIME, Reason MEMO, StatusCode TEXT(20), CreatedAt DATETIME, CreatedByUserName TEXT(100), ApprovedAt DATETIME, ApprovedByUserName TEXT(100))", _
        "CREATE TABLE tblCampaignBuildingTargets (CampaignBuildingTargetID COUNTER CONSTRAINT PK_tblCampaignBuildingTargets PRIMARY KEY, CampaignID LONG, BuildingID LONG)", _
        "CREATE TABLE tblCampaignUnitTargets (CampaignUnitTargetID COUNTER CONSTRAINT PK_tblCampaignUnitTargets PRIMARY KEY, CampaignID LONG, UnitID LONG)", _
        "CREATE TABLE tblChargeLines (ChargeLineID COUNTER CONSTRAINT PK_tblChargeLines PRIMARY KEY, ChargeID LONG, CampaignID LONG, UnitID LONG, BuildingID LONG, OwnerIDSnapshot LONG, BuildingCodeSnapshot TEXT(30), UnitCodeSnapshot TEXT(30), RateAmount CURRENCY, GrossAmount CURRENCY, DiscountAmount CURRENCY, NetAmount CURRENCY, CreatedAt DATETIME)", _
        "CREATE TABLE tblFinancialPeriods (PeriodID COUNTER CONSTRAINT PK_tblFinancialPeriods PRIMARY KEY, PeriodCode TEXT(30), PeriodName TEXT(100), PeriodStart DATETIME, PeriodEnd DATETIME, StatusCode TEXT(20), CreatedAt DATETIME, CreatedByUserName TEXT(100), ClosedAt DATETIME, ClosedByUserName TEXT(100), ReopenedAt DATETIME, ReopenedByUserName TEXT(100), ReopenReason MEMO)", _
        "CREATE TABLE tblExpenses (ExpenseID COUNTER CONSTRAINT PK_tblExpenses PRIMARY KEY, ExpenseNo TEXT(24), ExpenseDate DATETIME, CategoryID LONG, BeneficiaryName TEXT(150), Description MEMO, Amount CURRENCY, StatusCode TEXT(20), CreatedAt DATETIME, CreatedByUserName TEXT(100), UpdatedAt DATETIME)", _
        "CREATE TABLE tblPaymentVoucherReversals (ReversalID COUNTER CONSTRAINT PK_tblPaymentVoucherReversals PRIMARY KEY, PaymentID LONG, DocumentID LONG, Amount CURRENCY, Reason MEMO, CreatedAt DATETIME, CreatedByUserName TEXT(100))", _
        "CREATE TABLE tblSystemSettings (SystemSettingID COUNTER CONSTRAINT PK_tblSystemSettings PRIMARY KEY, SettingKey TEXT(100), SettingValue MEMO, UpdatedAt DATETIME, UpdatedByUserName TEXT(100))", _
        "CREATE TABLE tblUserPreferences (UserPreferenceID COUNTER CONSTRAINT PK_tblUserPreferences PRIMARY KEY, UserID LONG, PreferenceKey TEXT(100), PreferenceValue MEMO, UpdatedAt DATETIME)", _
        "CREATE TABLE tblReceiptAttachments (ReceiptAttachmentID COUNTER CONSTRAINT PK_tblReceiptAttachments PRIMARY KEY, ReceiptID LONG, AttachmentID LONG)" _
    )
    AT_ExecuteDDLBatch db, Array( _
        "CREATE TABLE tblExpenseAttachments (ExpenseAttachmentID COUNTER CONSTRAINT PK_tblExpenseAttachments PRIMARY KEY, ExpenseID LONG, AttachmentID LONG)" _
    )
    AT_CreateIndexes db
    AT_CreateRelations db
    Exit Sub
EH:

    savedErrNumber = Err.Number: eDesc = Err.Description
    Err.Raise savedErrNumber, "AT_CreateSchemaOnDatabase", eDesc
End Sub

Private Sub AT_ExecuteDDLBatch(ByVal db As DAO.Database, ByVal ddls As Variant)
    Dim i As Long, tableName As String
    For i = LBound(ddls) To UBound(ddls)
        tableName = AT_TableNameFromSQL(CStr(ddls(i)))
        If Not AT_TableExists(db, tableName) Then db.Execute CStr(ddls(i)), dbFailOnError
    Next i
End Sub

Public Sub AT_SeedInitialData(ByVal db As DAO.Database)
    AT_SeedDefaults db
    AT_SetSchemaVersion db
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
    AT_EnsureIndex db, "tblReceiptAllocationLines", "UX_ReceiptCharge", "ReceiptID,ChargeID", True, False
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
    AT_EnsureIndex db, "tblBuildings", "UX_BuildingCode", "BuildingCode", True, False
    AT_EnsureIndex db, "tblUnits", "UX_BuildingUnitCode", "BuildingID,UnitCode", True, False
    AT_EnsureIndex db, "tblBuildingOwnership", "IX_BuildingOwnershipDates", "BuildingID,StartDate,EndDate", False, False
    AT_EnsureIndex db, "tblBuildingOwnership", "IX_OwnerOwnershipDates", "OwnerID,StartDate,EndDate", False, False
    AT_EnsureIndex db, "tblContributionTypes", "UX_ContributionTypeCode", "TypeCode", True, False
    AT_EnsureIndex db, "tblContributionCampaigns", "UX_CampaignCode", "CampaignCode", True, False
    AT_EnsureIndex db, "tblCampaignBuildingTargets", "UX_CampaignBuilding", "CampaignID,BuildingID", True, False
    AT_EnsureIndex db, "tblCampaignUnitTargets", "UX_CampaignUnit", "CampaignID,UnitID", True, False
    AT_EnsureIndex db, "tblChargeLines", "UX_CampaignUnitCharge", "CampaignID,UnitID", True, True
    AT_EnsureIndex db, "tblChargeLines", "IX_ChargeLineHeader", "ChargeID,UnitID", False, False
    AT_EnsureIndex db, "tblUnitExemptionRequests", "IX_UnitExemptionStatus", "UnitID,StatusCode,EffectiveFrom", False, False
    AT_EnsureIndex db, "tblFinancialPeriods", "UX_FinancialPeriodCode", "PeriodCode", True, False
    AT_EnsureIndex db, "tblExpenses", "UX_ExpenseNo", "ExpenseNo", True, True
    AT_EnsureIndex db, "tblPayments", "IX_PaymentExpense", "ExpenseID", False, False
    AT_EnsureIndex db, "tblSystemSettings", "UX_SystemSettingKey", "SettingKey", True, False
    AT_EnsureIndex db, "tblUserPreferences", "UX_UserPreference", "UserID,PreferenceKey", True, False
    AT_EnsureIndex db, "tblReceiptAttachments", "UX_ReceiptAttachment", "ReceiptID,AttachmentID", True, False
    AT_EnsureIndex db, "tblExpenseAttachments", "UX_ExpenseAttachment", "ExpenseID,AttachmentID", True, False
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
    AT_EnsureRelationBatch db, Array( _
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
        "relAdjustmentCharge|tblOwnerAdjustments|ChargeID|tblOwnerCharges|ChargeID" _
    )
    AT_EnsureRelationBatch db, Array( _
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
        "relCreditLineCharge|tblCreditApplicationLines|ChargeID|tblOwnerCharges|ChargeID" _
    )
    AT_EnsureRelationBatch db, Array( _
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
        "relIntegrityFindingRun|tblIntegrityFindings|IntegrityRunID|tblIntegrityRuns|IntegrityRunID" _
    )
    AT_EnsureRelationBatch db, Array( _
        "relUserRoleUser|tblUserRoleAssignments|UserID|tblUsers|UserID", _
        "relUserRoleRole|tblUserRoleAssignments|RoleID|tblRoles|RoleID", _
        "relUserRoleAssignedBy|tblUserRoleAssignments|AssignedByUserID|tblUsers|UserID", _
        "relRolePermissionRole|tblRolePermissions|RoleID|tblRoles|RoleID", _
        "relRolePermissionPermission|tblRolePermissions|PermissionID|tblPermissions|PermissionID", _
        "relBuildingOwnershipBuilding|tblBuildingOwnership|BuildingID|tblBuildings|BuildingID", _
        "relBuildingOwnershipOwner|tblBuildingOwnership|OwnerID|tblOwners|OwnerID", _
        "relUnitBuilding|tblUnits|BuildingID|tblBuildings|BuildingID", _
        "relOccupancyUnit|tblUnitOccupancy|UnitID|tblUnits|UnitID", _
        "relExemptionUnit|tblUnitExemptionRequests|UnitID|tblUnits|UnitID", _
        "relCampaignType|tblContributionCampaigns|ContributionTypeID|tblContributionTypes|ContributionTypeID", _
        "relCampaignRate|tblContributionCampaigns|RateID|tblContributionRates|RateID", _
        "relCampaignBuildingCampaign|tblCampaignBuildingTargets|CampaignID|tblContributionCampaigns|CampaignID", _
        "relCampaignBuildingBuilding|tblCampaignBuildingTargets|BuildingID|tblBuildings|BuildingID", _
        "relCampaignUnitCampaign|tblCampaignUnitTargets|CampaignID|tblContributionCampaigns|CampaignID", _
        "relCampaignUnitUnit|tblCampaignUnitTargets|UnitID|tblUnits|UnitID", _
        "relChargeCampaign|tblOwnerCharges|CampaignID|tblContributionCampaigns|CampaignID", _
        "relChargeBuilding|tblOwnerCharges|BuildingID|tblBuildings|BuildingID" _
    )
    AT_EnsureRelationBatch db, Array( _
        "relChargeLineCharge|tblChargeLines|ChargeID|tblOwnerCharges|ChargeID", _
        "relChargeLineCampaign|tblChargeLines|CampaignID|tblContributionCampaigns|CampaignID", _
        "relChargeLineUnit|tblChargeLines|UnitID|tblUnits|UnitID", _
        "relChargeLineBuilding|tblChargeLines|BuildingID|tblBuildings|BuildingID", _
        "relExpenseCategory|tblExpenses|CategoryID|tblExpenseCategories|ExpenseCategoryID", _
        "relPaymentExpense|tblPayments|ExpenseID|tblExpenses|ExpenseID", _
        "relVoucherReversalPayment|tblPaymentVoucherReversals|PaymentID|tblPayments|PaymentID", _
        "relVoucherReversalDocument|tblPaymentVoucherReversals|DocumentID|tblFinancialDocuments|DocumentID", _
        "relPeriodClosingPeriod|tblFinancialClosings|PeriodID|tblFinancialPeriods|PeriodID", _
        "relUserPreferenceUser|tblUserPreferences|UserID|tblUsers|UserID", _
        "relReceiptAttachmentReceipt|tblReceiptAttachments|ReceiptID|tblReceipts|ReceiptID", _
        "relReceiptAttachmentAttachment|tblReceiptAttachments|AttachmentID|tblAttachments|AttachmentID", _
        "relExpenseAttachmentExpense|tblExpenseAttachments|ExpenseID|tblExpenses|ExpenseID", _
        "relExpenseAttachmentAttachment|tblExpenseAttachments|AttachmentID|tblAttachments|AttachmentID" _
    )
End Sub

Private Sub AT_EnsureRelationBatch(ByVal db As DAO.Database, ByVal links As Variant)
    Dim item As Variant, bits As Variant
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

Public Sub AT_SeedDefaults(ByVal db As DAO.Database)
    If DCount("*", "tblProjectSettings") = 0 Then
        db.Execute "INSERT INTO tblProjectSettings (ProjectName, NeighborhoodName, DefaultRate, CurrencyCode, DefaultCollectorName, CreatedAt) VALUES ('" & AT_U("0645,0634,0631,0648,0639,0020,0645,0633,0627,0647,0645,0627,062A,0020,0627,0644,0645,062C,0627,0631,064A") & "','" & AT_U("0627,0633,0645,0020,0627,0644,062D,0627,0631,0629") & "',3000,'YER','',Now())", dbFailOnError
    End If
    If DCount("*", "tblContributionRates") = 0 Then db.Execute "INSERT INTO tblContributionRates (RateCode, RateAmount, EffectiveFrom, IsActive, CreatedAt) VALUES ('RATE-0001',3000,#1/1/2020#,True,Now())", dbFailOnError
    If DCount("*", "tblNumberSeries") = 0 Then
        db.Execute "INSERT INTO tblNumberSeries (SeriesCode, NextValue, Prefix, PadLength) VALUES ('OWNER',1,'OW-',4)", dbFailOnError
        db.Execute "INSERT INTO tblNumberSeries (SeriesCode, NextValue, Prefix, PadLength) VALUES ('RECEIPT',1,'RC-',6)", dbFailOnError
        db.Execute "INSERT INTO tblNumberSeries (SeriesCode, NextValue, Prefix, PadLength) VALUES ('PAYMENT',1,'PV-',6)", dbFailOnError
        db.Execute "INSERT INTO tblNumberSeries (SeriesCode, NextValue, Prefix, PadLength) VALUES ('TRANSFER',1,'TR-',6)", dbFailOnError
        db.Execute "INSERT INTO tblNumberSeries (SeriesCode, NextValue, Prefix, PadLength) VALUES ('CHARGE',1,'CH-',8)", dbFailOnError
    End If
    If DCount("*", "tblAccounts") = 0 Then db.Execute AT_U("0049,004E,0053,0045,0052,0054,0020,0049,004E,0054,004F,0020,0074,0062,006C,0041,0063,0063,006F,0075,006E,0074,0073,0020,0028,0041,0063,0063,006F,0075,006E,0074,0043,006F,0064,0065,002C,0020,0041,0063,0063,006F,0075,006E,0074,004E,0061,006D,0065,002C,0020,0041,0063,0063,006F,0075,006E,0074,0054,0079,0070,0065,002C,0020,0049,0073,0041,0063,0074,0069,0076,0065,002C,0020,0043,0072,0065,0061,0074,0065,0064,0041,0074,0029,0020,0056,0041,004C,0055,0045,0053,0020,0028,0027,0043,0041,0053,0048,002D,0030,0031,0027,002C,0027,0627,0644,0635,0646,062F,0648,0642,0020,0627,0644,0646,0642,062F,064A,0027,002C,0027,0043,0041,0053,0048,0027,002C,0054,0072,0075,0065,002C,004E,006F,0077,0028,0029,0029"), dbFailOnError
    If DCount("*", "tblPaymentMethods") = 0 Then
        db.Execute AT_U("0049,004E,0053,0045,0052,0054,0020,0049,004E,0054,004F,0020,0074,0062,006C,0050,0061,0079,006D,0065,006E,0074,004D,0065,0074,0068,006F,0064,0073,0020,0028,004D,0065,0074,0068,006F,0064,0043,006F,0064,0065,002C,0020,004D,0065,0074,0068,006F,0064,004E,0061,006D,0065,002C,0020,0049,0073,0041,0063,0074,0069,0076,0065,002C,0020,0043,0072,0065,0061,0074,0065,0064,0041,0074,0029,0020,0056,0041,004C,0055,0045,0053,0020,0028,0027,0043,0041,0053,0048,0027,002C,0027,0646,0642,062F,064B,0627,0027,002C,0054,0072,0075,0065,002C,004E,006F,0077,0028,0029,0029"), dbFailOnError
        db.Execute AT_U("0049,004E,0053,0045,0052,0054,0020,0049,004E,0054,004F,0020,0074,0062,006C,0050,0061,0079,006D,0065,006E,0074,004D,0065,0074,0068,006F,0064,0073,0020,0028,004D,0065,0074,0068,006F,0064,0043,006F,0064,0065,002C,0020,004D,0065,0074,0068,006F,0064,004E,0061,006D,0065,002C,0020,0049,0073,0041,0063,0074,0069,0076,0065,002C,0020,0043,0072,0065,0061,0074,0065,0064,0041,0074,0029,0020,0056,0041,004C,0055,0045,0053,0020,0028,0027,0042,0041,004E,004B,0027,002C,0027,062A,062D,0648,064A,0644,0020,0628,0646,0643,064A,0027,002C,0054,0072,0075,0065,002C,004E,006F,0077,0028,0029,0029"), dbFailOnError
        db.Execute AT_U("0049,004E,0053,0045,0052,0054,0020,0049,004E,0054,004F,0020,0074,0062,006C,0050,0061,0079,006D,0065,006E,0074,004D,0065,0074,0068,006F,0064,0073,0020,0028,004D,0065,0074,0068,006F,0064,0043,006F,0064,0065,002C,0020,004D,0065,0074,0068,006F,0064,004E,0061,006D,0065,002C,0020,0049,0073,0041,0063,0074,0069,0076,0065,002C,0020,0043,0072,0065,0061,0074,0065,0064,0041,0074,0029,0020,0056,0041,004C,0055,0045,0053,0020,0028,0027,0045,0057,0041,004C,004C,0045,0054,0027,002C,0027,0645,062D,0641,0638,0629,0020,0625,0644,0643,062A,0631,0648,0646,064A,0629,0027,002C,0054,0072,0075,0065,002C,004E,006F,0077,0028,0029,0029"), dbFailOnError
    End If
    If DCount("*", "tblContributionTypes") = 0 Then
        db.Execute AT_U("0049,004E,0053,0045,0052,0054,0020,0049,004E,0054,004F,0020,0074,0062,006C,0043,006F,006E,0074,0072,0069,0062,0075,0074,0069,006F,006E,0054,0079,0070,0065,0073,0020,0028,0054,0079,0070,0065,0043,006F,0064,0065,002C,0020,0054,0079,0070,0065,004E,0061,006D,0065,002C,0020,0043,0061,006C,0063,0075,006C,0061,0074,0069,006F,006E,004D,0065,0074,0068,006F,0064,002C,0020,0049,0073,0042,0061,0073,0065,0054,0079,0070,0065,002C,0020,0049,0073,0041,0063,0074,0069,0076,0065,002C,0020,0043,0072,0065,0061,0074,0065,0064,0041,0074,0029,0020,0056,0041,004C,0055,0045,0053,0020,0028,0027,0042,0041,0053,0045,0027,002C,0027,0627,0644,0645,0633,0627,0647,0645,0629,0020,0627,0644,0623,0633,0627,0633,064A,0629,0027,002C,0027,0050,0045,0052,005F,0055,004E,0049,0054,0027,002C,0054,0072,0075,0065,002C,0054,0072,0075,0065,002C,004E,006F,0077,0028,0029,0029"), dbFailOnError
        db.Execute AT_U("0049,004E,0053,0045,0052,0054,0020,0049,004E,0054,004F,0020,0074,0062,006C,0043,006F,006E,0074,0072,0069,0062,0075,0074,0069,006F,006E,0054,0079,0070,0065,0073,0020,0028,0054,0079,0070,0065,0043,006F,0064,0065,002C,0020,0054,0079,0070,0065,004E,0061,006D,0065,002C,0020,0043,0061,006C,0063,0075,006C,0061,0074,0069,006F,006E,004D,0065,0074,0068,006F,0064,002C,0020,0049,0073,0042,0061,0073,0065,0054,0079,0070,0065,002C,0020,0049,0073,0041,0063,0074,0069,0076,0065,002C,0020,0043,0072,0065,0061,0074,0065,0064,0041,0074,0029,0020,0056,0041,004C,0055,0045,0053,0020,0028,0027,0053,0050,0045,0043,0049,0041,004C,0027,002C,0027,0645,0633,0627,0647,0645,0629,0020,0625,0636,0627,0641,064A,0629,0027,002C,0027,0050,0045,0052,005F,0055,004E,0049,0054,0027,002C,0046,0061,006C,0073,0065,002C,0054,0072,0075,0065,002C,004E,006F,0077,0028,0029,0029"), dbFailOnError
    End If
    If DCount("*", "tblExpenseCategories") = 0 Then
        db.Execute AT_U("0049,004E,0053,0045,0052,0054,0020,0049,004E,0054,004F,0020,0074,0062,006C,0045,0078,0070,0065,006E,0073,0065,0043,0061,0074,0065,0067,006F,0072,0069,0065,0073,0020,0028,0043,0061,0074,0065,0067,006F,0072,0079,0043,006F,0064,0065,002C,0020,0043,0061,0074,0065,0067,006F,0072,0079,004E,0061,006D,0065,002C,0020,0049,0073,0041,0063,0074,0069,0076,0065,0029,0020,0056,0041,004C,0055,0045,0053,0020,0028,0027,0047,0045,004E,0045,0052,0041,004C,0027,002C,0027,0645,0635,0631,0648,0641,0627,062A,0020,0639,0627,0645,0629,0027,002C,0054,0072,0075,0065,0029"), dbFailOnError
        db.Execute AT_U("0049,004E,0053,0045,0052,0054,0020,0049,004E,0054,004F,0020,0074,0062,006C,0045,0078,0070,0065,006E,0073,0065,0043,0061,0074,0065,0067,006F,0072,0069,0065,0073,0020,0028,0043,0061,0074,0065,0067,006F,0072,0079,0043,006F,0064,0065,002C,0020,0043,0061,0074,0065,0067,006F,0072,0079,004E,0061,006D,0065,002C,0020,0049,0073,0041,0063,0074,0069,0076,0065,0029,0020,0056,0041,004C,0055,0045,0053,0020,0028,0027,0054,0052,0041,004E,0053,0046,0045,0052,005F,0046,0045,0045,0027,002C,0027,0631,0633,0648,0645,0020,062A,062D,0648,064A,0644,0027,002C,0054,0072,0075,0065,0029"), dbFailOnError
        db.Execute AT_U("0049,004E,0053,0045,0052,0054,0020,0049,004E,0054,004F,0020,0074,0062,006C,0045,0078,0070,0065,006E,0073,0065,0043,0061,0074,0065,0067,006F,0072,0069,0065,0073,0020,0028,0043,0061,0074,0065,0067,006F,0072,0079,0043,006F,0064,0065,002C,0020,0043,0061,0074,0065,0067,006F,0072,0079,004E,0061,006D,0065,002C,0020,0049,0073,0041,0063,0074,0069,0076,0065,0029,0020,0056,0041,004C,0055,0045,0053,0020,0028,0027,004D,0041,0049,004E,0054,0045,004E,0041,004E,0043,0045,0027,002C,0027,0635,064A,0627,0646,0629,0027,002C,0054,0072,0075,0065,0029"), dbFailOnError
    End If
    If DCount("*", "tblSystemSettings") = 0 Then
        db.Execute "INSERT INTO tblSystemSettings (SettingKey, SettingValue, UpdatedAt, UpdatedByUserName) VALUES ('DATABASE_MODE','SPLIT_FE_BE',Now(),'SYSTEM')", dbFailOnError
        db.Execute "INSERT INTO tblSystemSettings (SettingKey, SettingValue, UpdatedAt, UpdatedByUserName) VALUES ('CURRENCY_CODE','YER',Now(),'SYSTEM')", dbFailOnError
    End If
    If DCount("*", "tblRoles") = 0 Then
        db.Execute AT_U("0049,004E,0053,0045,0052,0054,0020,0049,004E,0054,004F,0020,0074,0062,006C,0052,006F,006C,0065,0073,0020,0028,0052,006F,006C,0065,0043,006F,0064,0065,002C,0020,0052,006F,006C,0065,004E,0061,006D,0065,002C,0020,0049,0073,0041,0063,0074,0069,0076,0065,002C,0020,0043,0072,0065,0061,0074,0065,0064,0041,0074,0029,0020,0056,0041,004C,0055,0045,0053,0020,0028,0027,0041,0044,004D,0049,004E,0027,002C,0027,0645,062F,064A,0631,0020,0627,0644,0646,0638,0627,0645,0027,002C,0054,0072,0075,0065,002C,004E,006F,0077,0028,0029,0029"), dbFailOnError
        db.Execute AT_U("0049,004E,0053,0045,0052,0054,0020,0049,004E,0054,004F,0020,0074,0062,006C,0052,006F,006C,0065,0073,0020,0028,0052,006F,006C,0065,0043,006F,0064,0065,002C,0020,0052,006F,006C,0065,004E,0061,006D,0065,002C,0020,0049,0073,0041,0063,0074,0069,0076,0065,002C,0020,0043,0072,0065,0061,0074,0065,0064,0041,0074,0029,0020,0056,0041,004C,0055,0045,0053,0020,0028,0027,0043,004F,004C,004C,0045,0043,0054,004F,0052,0027,002C,0027,0645,062D,0635,0644,0027,002C,0054,0072,0075,0065,002C,004E,006F,0077,0028,0029,0029"), dbFailOnError
        db.Execute AT_U("0049,004E,0053,0045,0052,0054,0020,0049,004E,0054,004F,0020,0074,0062,006C,0052,006F,006C,0065,0073,0020,0028,0052,006F,006C,0065,0043,006F,0064,0065,002C,0020,0052,006F,006C,0065,004E,0061,006D,0065,002C,0020,0049,0073,0041,0063,0074,0069,0076,0065,002C,0020,0043,0072,0065,0061,0074,0065,0064,0041,0074,0029,0020,0056,0041,004C,0055,0045,0053,0020,0028,0027,0056,0049,0045,0057,0045,0052,0027,002C,0027,0642,0627,0631,0626,0020,062A,0642,0627,0631,064A,0631,0027,002C,0054,0072,0075,0065,002C,004E,006F,0077,0028,0029,0029"), dbFailOnError
        db.Execute AT_U("0049,004E,0053,0045,0052,0054,0020,0049,004E,0054,004F,0020,0074,0062,006C,0052,006F,006C,0065,0073,0020,0028,0052,006F,006C,0065,0043,006F,0064,0065,002C,0020,0052,006F,006C,0065,004E,0061,006D,0065,002C,0020,0049,0073,0041,0063,0074,0069,0076,0065,002C,0020,0043,0072,0065,0061,0074,0065,0064,0041,0074,0029,0020,0056,0041,004C,0055,0045,0053,0020,0028,0027,0050,0052,004F,004A,0045,0043,0054,005F,004D,0041,004E,0041,0047,0045,0052,0027,002C,0027,0645,062F,064A,0631,0020,0627,0644,0645,0634,0631,0648,0639,0027,002C,0054,0072,0075,0065,002C,004E,006F,0077,0028,0029,0029"), dbFailOnError
        db.Execute AT_U("0049,004E,0053,0045,0052,0054,0020,0049,004E,0054,004F,0020,0074,0062,006C,0052,006F,006C,0065,0073,0020,0028,0052,006F,006C,0065,0043,006F,0064,0065,002C,0020,0052,006F,006C,0065,004E,0061,006D,0065,002C,0020,0049,0073,0041,0063,0074,0069,0076,0065,002C,0020,0043,0072,0065,0061,0074,0065,0064,0041,0074,0029,0020,0056,0041,004C,0055,0045,0053,0020,0028,0027,0041,0043,0043,004F,0055,004E,0054,0041,004E,0054,0027,002C,0027,0645,062D,0627,0633,0628,0027,002C,0054,0072,0075,0065,002C,004E,006F,0077,0028,0029,0029"), dbFailOnError
        db.Execute AT_U("0049,004E,0053,0045,0052,0054,0020,0049,004E,0054,004F,0020,0074,0062,006C,0052,006F,006C,0065,0073,0020,0028,0052,006F,006C,0065,0043,006F,0064,0065,002C,0020,0052,006F,006C,0065,004E,0061,006D,0065,002C,0020,0049,0073,0041,0063,0074,0069,0076,0065,002C,0020,0043,0072,0065,0061,0074,0065,0064,0041,0074,0029,0020,0056,0041,004C,0055,0045,0053,0020,0028,0027,0043,0041,0053,0048,0049,0045,0052,0027,002C,0027,0623,0645,064A,0646,0020,0635,0646,062F,0648,0642,0027,002C,0054,0072,0075,0065,002C,004E,006F,0077,0028,0029,0029"), dbFailOnError
        db.Execute AT_U("0049,004E,0053,0045,0052,0054,0020,0049,004E,0054,004F,0020,0074,0062,006C,0052,006F,006C,0065,0073,0020,0028,0052,006F,006C,0065,0043,006F,0064,0065,002C,0020,0052,006F,006C,0065,004E,0061,006D,0065,002C,0020,0049,0073,0041,0063,0074,0069,0076,0065,002C,0020,0043,0072,0065,0061,0074,0065,0064,0041,0074,0029,0020,0056,0041,004C,0055,0045,0053,0020,0028,0027,0052,0045,0047,0049,0053,0054,0052,0041,0054,0049,004F,004E,005F,0043,004C,0045,0052,004B,0027,002C,0027,0645,0633,062C,0644,0020,0628,064A,0627,0646,0627,062A,0027,002C,0054,0072,0075,0065,002C,004E,006F,0077,0028,0029,0029"), dbFailOnError
        db.Execute AT_U("0049,004E,0053,0045,0052,0054,0020,0049,004E,0054,004F,0020,0074,0062,006C,0052,006F,006C,0065,0073,0020,0028,0052,006F,006C,0065,0043,006F,0064,0065,002C,0020,0052,006F,006C,0065,004E,0061,006D,0065,002C,0020,0049,0073,0041,0063,0074,0069,0076,0065,002C,0020,0043,0072,0065,0061,0074,0065,0064,0041,0074,0029,0020,0056,0041,004C,0055,0045,0053,0020,0028,0027,0041,0055,0044,0049,0054,004F,0052,0027,002C,0027,0645,0631,0627,062C,0639,0027,002C,0054,0072,0075,0065,002C,004E,006F,0077,0028,0029,0029"), dbFailOnError
        db.Execute AT_U("0049,004E,0053,0045,0052,0054,0020,0049,004E,0054,004F,0020,0074,0062,006C,0052,006F,006C,0065,0073,0020,0028,0052,006F,006C,0065,0043,006F,0064,0065,002C,0020,0052,006F,006C,0065,004E,0061,006D,0065,002C,0020,0049,0073,0041,0063,0074,0069,0076,0065,002C,0020,0043,0072,0065,0061,0074,0065,0064,0041,0074,0029,0020,0056,0041,004C,0055,0045,0053,0020,0028,0027,0052,0045,0050,004F,0052,0054,005F,0052,0045,0041,0044,0045,0052,0027,002C,0027,0642,0627,0631,0626,0020,062A,0642,0627,0631,064A,0631,0027,002C,0054,0072,0075,0065,002C,004E,006F,0077,0028,0029,0029"), dbFailOnError
    End If
    If DCount("*", "tblPermissions") = 0 Then AT_SeedPermissions db
    AT_SeedRolePermissions db
    AT_SeedLocalWindowsUser db
End Sub

Private Sub AT_SeedPermissions(ByVal db As DAO.Database)
    Dim codes As Variant, i As Long
    codes = Array("OWNER_VIEW", "OWNER_EDIT", "OWNER_APPROVE", "BUILDING_EDIT", "UNIT_EDIT", "UNIT_EXEMPTION_REQUEST", "UNIT_EXEMPTION_APPROVE", "CAMPAIGN_DRAFT", "CAMPAIGN_APPROVE", "CHARGE_DRAFT", "CHARGE_APPROVE", "RECEIPT_DRAFT", "RECEIPT_APPROVE", "PAYMENT_DRAFT", "PAYMENT_APPROVE", "EXPENSE_DRAFT", "EXPENSE_APPROVE", "ADJUSTMENT_REQUEST", "ADJUSTMENT_APPROVE", "CREDIT_APPLY", "TRANSFER_DRAFT", "TRANSFER_APPROVE", "DOCUMENT_CANCEL", "REPORT_VIEW", "FINANCIAL_CLOSE", "FINANCIAL_REOPEN", "BACKUP_RESTORE", "ATTACHMENT_ADD", "SETTINGS_ADMIN", "AUDIT_VIEW", "INTEGRITY_RUN")
    For i = LBound(codes) To UBound(codes)
        Dim rs As DAO.Recordset
        Set rs = db.OpenRecordset("tblPermissions", dbOpenDynaset)
        rs.AddNew: rs!PermissionCode = CStr(codes(i)): rs!PermissionName = AT_PermissionName(CStr(codes(i))): rs!IsActive = True: rs.Update
        rs.Close
    Next i
End Sub

Private Function AT_PermissionName(ByVal permissionCode As String) As String
    Select Case permissionCode
        Case "OWNER_VIEW": AT_PermissionName = AT_U("0639,0631,0636,0020,0627,0644,0645,0644,0627,0643")
        Case "OWNER_EDIT": AT_PermissionName = AT_U("062A,0639,062F,064A,0644,0020,0628,064A,0627,0646,0627,062A,0020,0627,0644,0645,0644,0627,0643")
        Case "OWNER_APPROVE": AT_PermissionName = AT_U("0627,0639,062A,0645,0627,062F,0020,0627,0644,0645,0627,0644,0643")
        Case "BUILDING_EDIT": AT_PermissionName = AT_U("0625,062F,0627,0631,0629,0020,0627,0644,0645,0628,0627,0646,064A,0020,0648,0646,0642,0644,0020,0627,0644,0645,0644,0643,064A,0629")
        Case "UNIT_EDIT": AT_PermissionName = AT_U("0625,062F,0627,0631,0629,0020,0627,0644,0634,0642,0642,0020,0648,0627,0644,0623,0633,0631")
        Case "UNIT_EXEMPTION_REQUEST": AT_PermissionName = AT_U("0637,0644,0628,0020,0625,0639,0641,0627,0621,0020,0634,0642,0629")
        Case "UNIT_EXEMPTION_APPROVE": AT_PermissionName = AT_U("0627,0639,062A,0645,0627,062F,0020,0625,0639,0641,0627,0621,0020,0634,0642,0629")
        Case "CAMPAIGN_DRAFT": AT_PermissionName = AT_U("0625,0639,062F,0627,062F,0020,062D,0645,0644,0629,0020,0645,0633,0627,0647,0645,0629")
        Case "CAMPAIGN_APPROVE": AT_PermissionName = AT_U("0627,0639,062A,0645,0627,062F,0020,062D,0645,0644,0627,062A,0020,0627,0644,0645,0633,0627,0647,0645,0629")
        Case "CHARGE_DRAFT": AT_PermissionName = AT_U("0625,0646,0634,0627,0621,0020,0645,0633,0648,062F,0629,0020,0627,0633,062A,062D,0642,0627,0642")
        Case "CHARGE_APPROVE": AT_PermissionName = AT_U("0627,0639,062A,0645,0627,062F,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642")
        Case "RECEIPT_DRAFT": AT_PermissionName = AT_U("0625,0646,0634,0627,0621,0020,0645,0633,0648,062F,0629,0020,0642,0628,0636")
        Case "RECEIPT_APPROVE": AT_PermissionName = AT_U("0627,0639,062A,0645,0627,062F,0020,0627,0644,0642,0628,0636")
        Case "PAYMENT_DRAFT": AT_PermissionName = AT_U("0625,0646,0634,0627,0621,0020,0645,0633,0648,062F,0629,0020,0635,0631,0641")
        Case "PAYMENT_APPROVE": AT_PermissionName = AT_U("0627,0639,062A,0645,0627,062F,0020,0627,0644,0635,0631,0641")
        Case "EXPENSE_DRAFT": AT_PermissionName = AT_U("062A,0633,062C,064A,0644,0020,0645,0635,0631,0648,0641")
        Case "EXPENSE_APPROVE": AT_PermissionName = AT_U("0627,0639,062A,0645,0627,062F,0020,062F,0641,0639,0020,0645,0635,0631,0648,0641")
        Case "ADJUSTMENT_REQUEST": AT_PermissionName = AT_U("0637,0644,0628,0020,062E,0635,0645,002F,0625,0639,0641,0627,0621,0020,0645,0627,0644,064A")
        Case "ADJUSTMENT_APPROVE": AT_PermissionName = AT_U("0627,0639,062A,0645,0627,062F,0020,0627,0644,062E,0635,0648,0645,0627,062A")
        Case "CREDIT_APPLY": AT_PermissionName = AT_U("062A,0637,0628,064A,0642,0020,0627,0644,0631,0635,064A,062F,0020,0627,0644,062F,0627,0626,0646")
        Case "TRANSFER_DRAFT": AT_PermissionName = AT_U("0625,0646,0634,0627,0621,0020,0645,0633,0648,062F,0629,0020,062A,062D,0648,064A,0644")
        Case "TRANSFER_APPROVE": AT_PermissionName = AT_U("0627,0639,062A,0645,0627,062F,0020,0627,0644,062A,062D,0648,064A,0644,0627,062A")
        Case "DOCUMENT_CANCEL": AT_PermissionName = AT_U("0625,0644,063A,0627,0621,0020,0627,0644,0645,0633,0648,062F,0627,062A")
        Case "REPORT_VIEW": AT_PermissionName = AT_U("0639,0631,0636,0020,0627,0644,062A,0642,0627,0631,064A,0631")
        Case "FINANCIAL_CLOSE": AT_PermissionName = AT_U("0627,0644,0625,0642,0641,0627,0644,0020,0627,0644,0645,0627,0644,064A")
        Case "FINANCIAL_REOPEN": AT_PermissionName = AT_U("0625,0639,0627,062F,0629,0020,0641,062A,062D,0020,0641,062A,0631,0629,0020,0645,0627,0644,064A,0629")
        Case "BACKUP_RESTORE": AT_PermissionName = AT_U("0627,0644,0646,0633,062E,0020,0648,0627,0644,0627,0633,062A,0639,0627,062F,0629")
        Case "ATTACHMENT_ADD": AT_PermissionName = AT_U("0625,0636,0627,0641,0629,0020,0645,0631,0641,0642")
        Case "SETTINGS_ADMIN": AT_PermissionName = AT_U("0625,062F,0627,0631,0629,0020,0627,0644,0625,0639,062F,0627,062F,0627,062A")
        Case "AUDIT_VIEW": AT_PermissionName = AT_U("0639,0631,0636,0020,0627,0644,062A,062F,0642,064A,0642")
        Case "INTEGRITY_RUN": AT_PermissionName = AT_U("0641,062D,0635,0020,0633,0644,0627,0645,0629,0020,0627,0644,0628,064A,0627,0646,0627,062A")
    End Select
End Function

Private Sub AT_SeedRolePermissions(ByVal db As DAO.Database)
    Dim adminID As Long, collectorID As Long, viewerID As Long, permID As Long
    Dim rs As DAO.Recordset, roleCode As Variant, permCode As Variant, allowed As Boolean
    If DCount("*", "tblRolePermissions") > 0 Then Exit Sub
    Dim managerID As Long, accountantID As Long, cashierID As Long, registrarID As Long, auditorID As Long, readerID As Long
    adminID = Nz(DLookup("RoleID", "tblRoles", "RoleCode='ADMIN'"), 0)
    collectorID = Nz(DLookup("RoleID", "tblRoles", "RoleCode='COLLECTOR'"), 0)
    viewerID = Nz(DLookup("RoleID", "tblRoles", "RoleCode='VIEWER'"), 0)
    managerID = Nz(DLookup("RoleID", "tblRoles", "RoleCode='PROJECT_MANAGER'"), 0)
    accountantID = Nz(DLookup("RoleID", "tblRoles", "RoleCode='ACCOUNTANT'"), 0)
    cashierID = Nz(DLookup("RoleID", "tblRoles", "RoleCode='CASHIER'"), 0)
    registrarID = Nz(DLookup("RoleID", "tblRoles", "RoleCode='REGISTRATION_CLERK'"), 0)
    auditorID = Nz(DLookup("RoleID", "tblRoles", "RoleCode='AUDITOR'"), 0)
    readerID = Nz(DLookup("RoleID", "tblRoles", "RoleCode='REPORT_READER'"), 0)
    For Each permCode In Array("OWNER_VIEW", "OWNER_EDIT", "OWNER_APPROVE", "BUILDING_EDIT", "UNIT_EDIT", "UNIT_EXEMPTION_REQUEST", "UNIT_EXEMPTION_APPROVE", "CAMPAIGN_DRAFT", "CAMPAIGN_APPROVE", "CHARGE_DRAFT", "CHARGE_APPROVE", "RECEIPT_DRAFT", "RECEIPT_APPROVE", "PAYMENT_DRAFT", "PAYMENT_APPROVE", "EXPENSE_DRAFT", "EXPENSE_APPROVE", "ADJUSTMENT_REQUEST", "ADJUSTMENT_APPROVE", "CREDIT_APPLY", "TRANSFER_DRAFT", "TRANSFER_APPROVE", "DOCUMENT_CANCEL", "REPORT_VIEW", "FINANCIAL_CLOSE", "FINANCIAL_REOPEN", "BACKUP_RESTORE", "ATTACHMENT_ADD", "SETTINGS_ADMIN", "AUDIT_VIEW", "INTEGRITY_RUN")
        permID = Nz(DLookup("PermissionID", "tblPermissions", "PermissionCode='" & CStr(permCode) & "'"), 0)
        If adminID > 0 And permID > 0 Then AT_GrantRolePermission db, adminID, permID
        If permID > 0 Then
            If collectorID > 0 And (CStr(permCode) = "OWNER_VIEW" Or CStr(permCode) = "RECEIPT_DRAFT" Or CStr(permCode) = "ADJUSTMENT_REQUEST" Or CStr(permCode) = "REPORT_VIEW") Then AT_GrantRolePermission db, collectorID, permID
            If registrarID > 0 And (CStr(permCode) = "OWNER_VIEW" Or CStr(permCode) = "OWNER_EDIT" Or CStr(permCode) = "BUILDING_EDIT" Or CStr(permCode) = "UNIT_EDIT" Or CStr(permCode) = "UNIT_EXEMPTION_REQUEST" Or CStr(permCode) = "ATTACHMENT_ADD") Then AT_GrantRolePermission db, registrarID, permID
            If accountantID > 0 And (CStr(permCode) = "OWNER_VIEW" Or CStr(permCode) = "CAMPAIGN_DRAFT" Or CStr(permCode) = "CAMPAIGN_APPROVE" Or CStr(permCode) = "CHARGE_DRAFT" Or CStr(permCode) = "CHARGE_APPROVE" Or CStr(permCode) = "RECEIPT_DRAFT" Or CStr(permCode) = "RECEIPT_APPROVE" Or CStr(permCode) = "PAYMENT_DRAFT" Or CStr(permCode) = "PAYMENT_APPROVE" Or CStr(permCode) = "EXPENSE_DRAFT" Or CStr(permCode) = "EXPENSE_APPROVE" Or CStr(permCode) = "ADJUSTMENT_REQUEST" Or CStr(permCode) = "ADJUSTMENT_APPROVE" Or CStr(permCode) = "CREDIT_APPLY" Or CStr(permCode) = "TRANSFER_DRAFT" Or CStr(permCode) = "TRANSFER_APPROVE" Or CStr(permCode) = "REPORT_VIEW" Or CStr(permCode) = "ATTACHMENT_ADD") Then AT_GrantRolePermission db, accountantID, permID
            If cashierID > 0 And (CStr(permCode) = "RECEIPT_DRAFT" Or CStr(permCode) = "RECEIPT_APPROVE" Or CStr(permCode) = "PAYMENT_DRAFT" Or CStr(permCode) = "PAYMENT_APPROVE" Or CStr(permCode) = "TRANSFER_DRAFT" Or CStr(permCode) = "TRANSFER_APPROVE") Then AT_GrantRolePermission db, cashierID, permID
            If managerID > 0 And (CStr(permCode) = "OWNER_VIEW" Or CStr(permCode) = "OWNER_EDIT" Or CStr(permCode) = "OWNER_APPROVE" Or CStr(permCode) = "BUILDING_EDIT" Or CStr(permCode) = "UNIT_EDIT" Or CStr(permCode) = "UNIT_EXEMPTION_APPROVE" Or CStr(permCode) = "CAMPAIGN_DRAFT" Or CStr(permCode) = "CAMPAIGN_APPROVE" Or CStr(permCode) = "CHARGE_DRAFT" Or CStr(permCode) = "CHARGE_APPROVE" Or CStr(permCode) = "ADJUSTMENT_APPROVE" Or CStr(permCode) = "FINANCIAL_CLOSE" Or CStr(permCode) = "FINANCIAL_REOPEN" Or CStr(permCode) = "REPORT_VIEW" Or CStr(permCode) = "AUDIT_VIEW" Or CStr(permCode) = "INTEGRITY_RUN" Or CStr(permCode) = "ATTACHMENT_ADD") Then AT_GrantRolePermission db, managerID, permID
            If auditorID > 0 And (CStr(permCode) = "OWNER_VIEW" Or CStr(permCode) = "REPORT_VIEW" Or CStr(permCode) = "AUDIT_VIEW" Or CStr(permCode) = "INTEGRITY_RUN") Then AT_GrantRolePermission db, auditorID, permID
            If viewerID > 0 And (CStr(permCode) = "OWNER_VIEW" Or CStr(permCode) = "REPORT_VIEW") Then AT_GrantRolePermission db, viewerID, permID
            If readerID > 0 And (CStr(permCode) = "OWNER_VIEW" Or CStr(permCode) = "REPORT_VIEW") Then AT_GrantRolePermission db, readerID, permID
        End If
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

Public Sub AT_SetSchemaVersion(ByVal db As DAO.Database)
    If DCount("*", "tblSchemaVersion") = 0 Then
        db.Execute "INSERT INTO tblSchemaVersion (VersionCode, InstalledAt, InstalledByUserName, Notes) VALUES ('" & AT_SCHEMA_VERSION & "',Now(),'" & Replace(Environ$("Username"), "'", "''") & "','Initial schema build; UI and runtime verification still required.')", dbFailOnError
    End If
End Sub

Public Sub AT_Audit(ByVal actionCode As String, ByVal entityName As String, ByVal entityID As Long, ByVal resultCode As String, ByVal detail As String, Optional ByVal reason As String = "")
    On Error GoTo FallbackLog
    Dim rs As DAO.Recordset
    Dim logNo As Integer, logPath As String, fallbackText As String
    Dim auditErrNo As Long, auditErrDesc As String, fallbackErrNo As Long, fallbackErrDesc As String

    If Not AT_TableExists(CurrentDb, "tblAuditEvents") Then Err.Raise vbObjectError + 1490, "AT_Audit", "tblAuditEvents is unavailable."
    Set rs = CurrentDb.OpenRecordset("tblAuditEvents", dbOpenDynaset)
    rs.AddNew
    rs!EventAt = Now(): rs!UserName = Environ$("Username")
    rs!ActionCode = Left$(actionCode, 60): rs!EntityName = Left$(entityName, 60)
    rs!EntityID = entityID: rs!ResultCode = Left$(resultCode, 30)
    rs!Detail = detail: rs!Reason = reason
    rs.Update: rs.Close
    Exit Sub

FallbackLog:
    auditErrNo = Err.Number: auditErrDesc = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Err.Clear
    logPath = CurrentProject.Path & "\Archon_Audit_Fallback.log"
    fallbackText = Format$(Now(), "yyyy-mm-dd hh:nn:ss") & "|" & Environ$("Username") & "|" & _
        Replace(actionCode, "|", "/") & "|" & Replace(entityName, "|", "/") & "|" & CStr(entityID) & "|" & _
        Replace(resultCode, "|", "/") & "|" & Replace(Replace(detail, vbCr, " "), vbLf, " ") & "|" & _
        Replace(Replace(reason, vbCr, " "), vbLf, " ") & "|DAO_ERROR=" & CStr(auditErrNo) & ":" & auditErrDesc
    logNo = FreeFile
    Open logPath For Append As #logNo
    Print #logNo, fallbackText
    Close #logNo
    If Err.Number <> 0 Then
        fallbackErrNo = Err.Number: fallbackErrDesc = Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    If fallbackErrNo <> 0 Then
        Err.Raise vbObjectError + 1491, "AT_Audit", _
            "Audit database insert failed (" & auditErrNo & ": " & auditErrDesc & ") and fallback log write failed (" & fallbackErrNo & ": " & fallbackErrDesc & ")."
    End If
    Debug.Print "Archon audit event was written to sidecar fallback log: " & logPath
End Sub
