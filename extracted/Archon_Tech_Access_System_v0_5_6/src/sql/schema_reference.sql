-- Archon Tech Access schema 0.5.0
-- Reference extract from modArchonBuild.bas.
-- The supported installer executes equivalent DDL through DAO in Microsoft Access.
-- Do not run this file in another database engine. It is not the installer.

CREATE TABLE tblProjectSettings (ProjectSettingID COUNTER CONSTRAINT PK_tblProjectSettings PRIMARY KEY, ProjectName TEXT(150), NeighborhoodName TEXT(150), DefaultRate CURRENCY, CurrencyCode TEXT(10), DefaultCollectorName TEXT(150), CreatedAt DATETIME, UpdatedAt DATETIME);

CREATE TABLE tblContributionRates (RateID COUNTER CONSTRAINT PK_tblContributionRates PRIMARY KEY, RateCode TEXT(30), RateAmount CURRENCY, EffectiveFrom DATETIME, EffectiveTo DATETIME, IsActive YESNO, CreatedAt DATETIME);

CREATE TABLE tblUsers (UserID COUNTER CONSTRAINT PK_tblUsers PRIMARY KEY, UserName TEXT(100), DisplayName TEXT(150), AuthenticationMode TEXT(30), IsActive YESNO, CreatedAt DATETIME, LastLoginAt DATETIME);

CREATE TABLE tblRoles (RoleID COUNTER CONSTRAINT PK_tblRoles PRIMARY KEY, RoleCode TEXT(40), RoleName TEXT(100), IsActive YESNO, CreatedAt DATETIME);

CREATE TABLE tblPermissions (PermissionID COUNTER CONSTRAINT PK_tblPermissions PRIMARY KEY, PermissionCode TEXT(60), PermissionName TEXT(150), IsActive YESNO);

CREATE TABLE tblUserRoleAssignments (UserRoleAssignmentID COUNTER CONSTRAINT PK_tblUserRoleAssignments PRIMARY KEY, UserID LONG, RoleID LONG, AssignedAt DATETIME, AssignedByUserID LONG, IsActive YESNO);

CREATE TABLE tblRolePermissions (RolePermissionID COUNTER CONSTRAINT PK_tblRolePermissions PRIMARY KEY, RoleID LONG, PermissionID LONG, GrantedAt DATETIME);

CREATE TABLE tblOwners (OwnerID COUNTER CONSTRAINT PK_tblOwners PRIMARY KEY, OwnerCode TEXT(12), FullName TEXT(200), FullNameSearch TEXT(200), Phone TEXT(40), IdentityNo TEXT(80), HouseCount LONG, ApartmentCount LONG, Notes MEMO, AdminStatus TEXT(20), DataStatus TEXT(20), ReviewStatus TEXT(20), ActivityStatus TEXT(20), CreatedAt DATETIME, UpdatedAt DATETIME, CreatedByUserID LONG, UpdatedByUserID LONG);

CREATE TABLE tblChargeBatches (ChargeBatchID COUNTER CONSTRAINT PK_tblChargeBatches PRIMARY KEY, BatchCode TEXT(30), PeriodLabel TEXT(100), ChargeDate DATETIME, RateID LONG, StatusCode TEXT(20), CreatedAt DATETIME, CreatedByUserID LONG, ApprovedAt DATETIME, ApprovedByUserID LONG, Reason MEMO);

CREATE TABLE tblFinancialDocuments (DocumentID COUNTER CONSTRAINT PK_tblFinancialDocuments PRIMARY KEY, DocumentType TEXT(30), TransactionDate DATETIME, StatusCode TEXT(20), CreatedAt DATETIME, CreatedByUserName TEXT(100), ApprovedAt DATETIME, ApprovedByUserName TEXT(100), Reason MEMO, CancelReason MEMO, ReversalOfDocumentID LONG);

CREATE TABLE tblOfficialDocumentNumbers (OfficialDocumentNumberID COUNTER CONSTRAINT PK_tblOfficialDocumentNumbers PRIMARY KEY, DocumentID LONG, OfficialNumber TEXT(20), SeriesCode TEXT(30), NumericValue LONG, AssignedAt DATETIME, AssignedByUserName TEXT(100));

CREATE TABLE tblNumberSeries (SeriesID COUNTER CONSTRAINT PK_tblNumberSeries PRIMARY KEY, SeriesCode TEXT(30), NextValue LONG, Prefix TEXT(10), PadLength LONG);

CREATE TABLE tblOwnerCharges (ChargeID COUNTER CONSTRAINT PK_tblOwnerCharges PRIMARY KEY, ChargeCode TEXT(20), DocumentID LONG, OwnerID LONG, ChargeBatchID LONG, CampaignID LONG, BuildingID LONG, ChargeDate DATETIME, PeriodStart DATETIME, PeriodEnd DATETIME, ApartmentCountAtCharge LONG, RateAmount CURRENCY, GrossAmount CURRENCY, DiscountAmount CURRENCY, NetAmount CURRENCY, OwnerCodeAtCharge TEXT(12), OwnerNameAtCharge TEXT(200), ChargeStatus TEXT(20), CreatedAt DATETIME);

CREATE TABLE tblReceipts (ReceiptID COUNTER CONSTRAINT PK_tblReceipts PRIMARY KEY, DocumentID LONG, OwnerID LONG, Amount CURRENCY, PaymentMethodID LONG, AccountID LONG, PayerName TEXT(150), CollectorName TEXT(150), Notes MEMO, PrintedOwnerName TEXT(200), PrintedOwnerCode TEXT(12), PrintedNeighborhoodName TEXT(150), CreatedAt DATETIME);

CREATE TABLE tblReceiptAllocationLines (AllocationID COUNTER CONSTRAINT PK_tblReceiptAllocationLines PRIMARY KEY, ReceiptID LONG, ChargeID LONG, AllocatedAmount CURRENCY, LineNo LONG);

CREATE TABLE tblOwnerAdjustments (AdjustmentID COUNTER CONSTRAINT PK_tblOwnerAdjustments PRIMARY KEY, DocumentID LONG, ChargeID LONG, AdjustmentType TEXT(30), SignedAmount CURRENCY, Reason MEMO, RequestedByUserName TEXT(100), ApprovedByUserName TEXT(100), ApprovedAt DATETIME);

CREATE TABLE tblChargeSettlementLines (SettlementID COUNTER CONSTRAINT PK_tblChargeSettlementLines PRIMARY KEY, ChargeID LONG, DocumentID LONG, ReceiptID LONG, CreditApplicationID LONG, ReceiptCorrectionID LONG, SignedAmount CURRENCY, SettlementType TEXT(30), ReversalOfSettlementID LONG, CreatedAt DATETIME);

CREATE TABLE tblReceiptCorrections (ReceiptCorrectionID COUNTER CONSTRAINT PK_tblReceiptCorrections PRIMARY KEY, DocumentID LONG, OriginalReceiptID LONG, CorrectionType TEXT(30), Reason MEMO);

CREATE TABLE tblReceiptCorrectionLines (ReceiptCorrectionLineID COUNTER CONSTRAINT PK_tblReceiptCorrectionLines PRIMARY KEY, ReceiptCorrectionID LONG, OriginalAllocationID LONG, Amount CURRENCY, CreatedAt DATETIME);

CREATE TABLE tblOwnerCreditMovements (CreditMovementID COUNTER CONSTRAINT PK_tblOwnerCreditMovements PRIMARY KEY, OwnerID LONG, DocumentID LONG, SignedAmount CURRENCY, CreditType TEXT(30), SourceReceiptID LONG, SourceCreditMovementID LONG, Reason MEMO, CreatedAt DATETIME);

CREATE TABLE tblCreditApplications (CreditApplicationID COUNTER CONSTRAINT PK_tblCreditApplications PRIMARY KEY, DocumentID LONG, OwnerID LONG, Amount CURRENCY, ApprovedByUserName TEXT(100), ApprovedAt DATETIME, Reason MEMO);

CREATE TABLE tblCreditApplicationLines (CreditApplicationLineID COUNTER CONSTRAINT PK_tblCreditApplicationLines PRIMARY KEY, CreditApplicationID LONG, CreditMovementID LONG, ChargeID LONG, Amount CURRENCY);

CREATE TABLE tblAccounts (AccountID COUNTER CONSTRAINT PK_tblAccounts PRIMARY KEY, AccountCode TEXT(20), AccountName TEXT(120), AccountType TEXT(20), IsActive YESNO, CreatedAt DATETIME);

CREATE TABLE tblPaymentMethods (PaymentMethodID COUNTER CONSTRAINT PK_tblPaymentMethods PRIMARY KEY, MethodCode TEXT(30), MethodName TEXT(100), IsActive YESNO, CreatedAt DATETIME);

CREATE TABLE tblExpenseCategories (ExpenseCategoryID COUNTER CONSTRAINT PK_tblExpenseCategories PRIMARY KEY, CategoryCode TEXT(30), CategoryName TEXT(120), IsActive YESNO, CreatedAt DATETIME);

CREATE TABLE tblPayments (PaymentID COUNTER CONSTRAINT PK_tblPayments PRIMARY KEY, VoucherNo TEXT(24), DocumentID LONG, ExpenseID LONG, AccountID LONG, ExpenseCategoryID LONG, BeneficiaryName TEXT(150), Amount CURRENCY, PaymentMethodID LONG, RecipientName TEXT(150), Notes MEMO, RefundOfCreditMovementID LONG, RefundOfDocumentID LONG, ReversalOfPaymentID LONG);

CREATE TABLE tblAccountTransfers (TransferID COUNTER CONSTRAINT PK_tblAccountTransfers PRIMARY KEY, DocumentID LONG, SourceAccountID LONG, DestinationAccountID LONG, Amount CURRENCY, FeeAmount CURRENCY, ConfirmedAt DATETIME, ConfirmedByUserName TEXT(100), ConfirmationRef TEXT(100), Notes MEMO);

CREATE TABLE tblAccountMovements (AccountMovementID COUNTER CONSTRAINT PK_tblAccountMovements PRIMARY KEY, AccountID LONG, DocumentID LONG, MovementDate DATETIME, SignedAmount CURRENCY, MovementType TEXT(30), LineNo LONG, CreatedAt DATETIME);

CREATE TABLE tblAccountReservations (ReservationID COUNTER CONSTRAINT PK_tblAccountReservations PRIMARY KEY, DocumentID LONG, AccountID LONG, Amount CURRENCY, StatusCode TEXT(20), CreatedAt DATETIME, ReleasedAt DATETIME, Reason MEMO);

CREATE TABLE tblFinancialClosings (ClosingID COUNTER CONSTRAINT PK_tblFinancialClosings PRIMARY KEY, PeriodID LONG, PeriodStart DATETIME, PeriodEnd DATETIME, ClosedAt DATETIME, ClosedByUserName TEXT(100), StatusCode TEXT(20), DiscrepancyException YESNO, ExceptionReason MEMO, Summary MEMO, ReopenedAt DATETIME, ReopenedByUserName TEXT(100), ReopenReason MEMO);

CREATE TABLE tblFinancialClosingAccounts (ClosingAccountID COUNTER CONSTRAINT PK_tblFinancialClosingAccounts PRIMARY KEY, ClosingID LONG, AccountID LONG, BookBalance CURRENCY, ActualBalance CURRENCY, DifferenceAmount CURRENCY, Notes MEMO);

CREATE TABLE tblAuditEvents (AuditEventID COUNTER CONSTRAINT PK_tblAuditEvents PRIMARY KEY, EventAt DATETIME, UserName TEXT(100), ActionCode TEXT(60), EntityName TEXT(60), EntityID LONG, ResultCode TEXT(30), Reason MEMO, Detail MEMO);

CREATE TABLE tblAuditChanges (AuditChangeID COUNTER CONSTRAINT PK_tblAuditChanges PRIMARY KEY, AuditEventID LONG, FieldName TEXT(100), OldValue MEMO, NewValue MEMO);

CREATE TABLE tblAttachments (AttachmentID COUNTER CONSTRAINT PK_tblAttachments PRIMARY KEY, DocumentID LONG, RelativePath MEMO, OriginalFileName TEXT(255), ContentType TEXT(100), FileSizeBytes DOUBLE, SHA256 TEXT(64), AddedAt DATETIME, AddedByUserName TEXT(100));

CREATE TABLE tblBackupLog (BackupLogID COUNTER CONSTRAINT PK_tblBackupLog PRIMARY KEY, StartedAt DATETIME, CompletedAt DATETIME, OperationType TEXT(20), SourcePath MEMO, DestinationPath MEMO, ResultCode TEXT(20), FileSizeBytes DOUBLE, Detail MEMO);

CREATE TABLE tblIntegrityRuns (IntegrityRunID COUNTER CONSTRAINT PK_tblIntegrityRuns PRIMARY KEY, StartedAt DATETIME, CompletedAt DATETIME, RunByUserName TEXT(100), ResultSummary MEMO);

CREATE TABLE tblIntegrityFindings (FindingID COUNTER CONSTRAINT PK_tblIntegrityFindings PRIMARY KEY, IntegrityRunID LONG, Severity TEXT(20), EntityName TEXT(60), EntityID LONG, FindingCode TEXT(60), FindingText MEMO, ResolutionStatus TEXT(20), ResolutionReason MEMO, ResolvedByUserName TEXT(100), ResolvedAt DATETIME);

CREATE TABLE tblSchemaVersion (SchemaVersionID COUNTER CONSTRAINT PK_tblSchemaVersion PRIMARY KEY, VersionCode TEXT(20), InstalledAt DATETIME, InstalledByUserName TEXT(100), Notes MEMO);

CREATE TABLE tblBuildings (BuildingID COUNTER CONSTRAINT PK_tblBuildings PRIMARY KEY, BuildingCode TEXT(30), BuildingName TEXT(150), Address TEXT(255), IsActive YESNO, CreatedAt DATETIME, UpdatedAt DATETIME, Notes MEMO);

CREATE TABLE tblBuildingOwnership (OwnershipID COUNTER CONSTRAINT PK_tblBuildingOwnership PRIMARY KEY, BuildingID LONG, OwnerID LONG, StartDate DATETIME, EndDate DATETIME, TransferDocumentNo TEXT(100), Reason MEMO, CreatedAt DATETIME, CreatedByUserName TEXT(100));

CREATE TABLE tblUnits (UnitID COUNTER CONSTRAINT PK_tblUnits PRIMARY KEY, BuildingID LONG, UnitCode TEXT(30), FloorNo TEXT(30), IsActive YESNO, IsContributionEligible YESNO, Notes MEMO, CreatedAt DATETIME, UpdatedAt DATETIME);

CREATE TABLE tblUnitOccupancy (OccupancyID COUNTER CONSTRAINT PK_tblUnitOccupancy PRIMARY KEY, UnitID LONG, FamilyName TEXT(150), HouseholdCount LONG, StartDate DATETIME, EndDate DATETIME, Notes MEMO, CreatedAt DATETIME);

CREATE TABLE tblUnitExemptionRequests (ExemptionRequestID COUNTER CONSTRAINT PK_tblUnitExemptionRequests PRIMARY KEY, UnitID LONG, RequestDate DATETIME, EffectiveFrom DATETIME, EffectiveTo DATETIME, Reason MEMO, EvidenceReference TEXT(255), RequestedByUserName TEXT(100), StatusCode TEXT(20), DecidedByUserName TEXT(100), DecidedAt DATETIME, DecisionReason MEMO, CreatedAt DATETIME);

CREATE TABLE tblContributionTypes (ContributionTypeID COUNTER CONSTRAINT PK_tblContributionTypes PRIMARY KEY, TypeCode TEXT(30), TypeName TEXT(120), CalculationMethod TEXT(30), IsBaseType YESNO, IsActive YESNO, CreatedAt DATETIME);

CREATE TABLE tblContributionCampaigns (CampaignID COUNTER CONSTRAINT PK_tblContributionCampaigns PRIMARY KEY, CampaignCode TEXT(30), CampaignName TEXT(150), ContributionTypeID LONG, TargetScope TEXT(30), RateID LONG, AmountPerUnit CURRENCY, PeriodStart DATETIME, PeriodEnd DATETIME, DueDate DATETIME, Reason MEMO, StatusCode TEXT(20), CreatedAt DATETIME, CreatedByUserName TEXT(100), ApprovedAt DATETIME, ApprovedByUserName TEXT(100));

CREATE TABLE tblCampaignBuildingTargets (CampaignBuildingTargetID COUNTER CONSTRAINT PK_tblCampaignBuildingTargets PRIMARY KEY, CampaignID LONG, BuildingID LONG);

CREATE TABLE tblCampaignUnitTargets (CampaignUnitTargetID COUNTER CONSTRAINT PK_tblCampaignUnitTargets PRIMARY KEY, CampaignID LONG, UnitID LONG);

CREATE TABLE tblChargeLines (ChargeLineID COUNTER CONSTRAINT PK_tblChargeLines PRIMARY KEY, ChargeID LONG, CampaignID LONG, UnitID LONG, BuildingID LONG, OwnerIDSnapshot LONG, BuildingCodeSnapshot TEXT(30), UnitCodeSnapshot TEXT(30), RateAmount CURRENCY, GrossAmount CURRENCY, DiscountAmount CURRENCY, NetAmount CURRENCY, CreatedAt DATETIME);

CREATE TABLE tblFinancialPeriods (PeriodID COUNTER CONSTRAINT PK_tblFinancialPeriods PRIMARY KEY, PeriodCode TEXT(30), PeriodName TEXT(100), PeriodStart DATETIME, PeriodEnd DATETIME, StatusCode TEXT(20), CreatedAt DATETIME, CreatedByUserName TEXT(100), ClosedAt DATETIME, ClosedByUserName TEXT(100), ReopenedAt DATETIME, ReopenedByUserName TEXT(100), ReopenReason MEMO);

CREATE TABLE tblExpenses (ExpenseID COUNTER CONSTRAINT PK_tblExpenses PRIMARY KEY, ExpenseNo TEXT(24), ExpenseDate DATETIME, CategoryID LONG, BeneficiaryName TEXT(150), Description MEMO, Amount CURRENCY, StatusCode TEXT(20), CreatedAt DATETIME, CreatedByUserName TEXT(100), UpdatedAt DATETIME);

CREATE TABLE tblPaymentVoucherReversals (ReversalID COUNTER CONSTRAINT PK_tblPaymentVoucherReversals PRIMARY KEY, PaymentID LONG, DocumentID LONG, Amount CURRENCY, Reason MEMO, CreatedAt DATETIME, CreatedByUserName TEXT(100));

CREATE TABLE tblSystemSettings (SystemSettingID COUNTER CONSTRAINT PK_tblSystemSettings PRIMARY KEY, SettingKey TEXT(100), SettingValue MEMO, UpdatedAt DATETIME, UpdatedByUserName TEXT(100));

CREATE TABLE tblUserPreferences (UserPreferenceID COUNTER CONSTRAINT PK_tblUserPreferences PRIMARY KEY, UserID LONG, PreferenceKey TEXT(100), PreferenceValue MEMO, UpdatedAt DATETIME);

CREATE TABLE tblReceiptAttachments (ReceiptAttachmentID COUNTER CONSTRAINT PK_tblReceiptAttachments PRIMARY KEY, ReceiptID LONG, AttachmentID LONG);

CREATE TABLE tblExpenseAttachments (ExpenseAttachmentID COUNTER CONSTRAINT PK_tblExpenseAttachments PRIMARY KEY, ExpenseID LONG, AttachmentID LONG);
