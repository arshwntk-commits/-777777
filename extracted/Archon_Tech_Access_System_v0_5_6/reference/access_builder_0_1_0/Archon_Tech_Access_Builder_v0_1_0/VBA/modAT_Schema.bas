Attribute VB_Name = "modAT_Schema"
Option Compare Database
Option Explicit

' Creates the database schema without deleting existing objects or data.
' Existing fields are retained. Missing fields are added; incompatible existing
' field definitions are logged for manual review rather than changed destructively.

Public Sub AT_BuildSchema()
    On Error GoTo BuildFail
    AT_LogBuild "Schema", "Started", "Creating Archon Tech tables"

    AT_EnsureTable "tblAT_AppSettings", Array( _
        Array("SettingID", AT_DB_LONG, 0, True, True, False), Array("SettingKey", AT_DB_TEXT, 100, True, False, False), _
        Array("SettingValue", AT_DB_MEMO, 0, False, False, True), Array("UpdatedAt", AT_DB_DATE, 0, False, False, False))

    AT_EnsureTable "tblAT_Users", Array( _
        Array("UserID", AT_DB_LONG, 0, True, True, False), Array("UserName", AT_DB_TEXT, 80, True, False, False), _
        Array("DisplayName", AT_DB_TEXT, 150, True, False, False), Array("RoleName", AT_DB_TEXT, 50, False, False, True), _
        Array("IsActive", AT_DB_BOOLEAN, 0, True, False, False), Array("CreatedAt", AT_DB_DATE, 0, False, False, False))

    AT_EnsureTable "tblAT_Owners", Array( _
        Array("OwnerID", AT_DB_LONG, 0, True, True, False), Array("OwnerCode", AT_DB_TEXT, 20, True, False, False), _
        Array("FullName", AT_DB_TEXT, 180, True, False, False), Array("Phone", AT_DB_TEXT, 40, False, False, True), _
        Array("AltPhone", AT_DB_TEXT, 40, False, False, True), Array("Address", AT_DB_TEXT, 255, False, False, True), _
        Array("Notes", AT_DB_MEMO, 0, False, False, True), Array("OwnerStatus", AT_DB_TEXT, 20, True, False, False), _
        Array("DataReviewStatus", AT_DB_TEXT, 20, False, False, True), Array("CreatedAt", AT_DB_DATE, 0, False, False, False), _
        Array("UpdatedAt", AT_DB_DATE, 0, False, False, False))

    AT_EnsureTable "tblAT_Units", Array( _
        Array("UnitID", AT_DB_LONG, 0, True, True, False), Array("UnitCode", AT_DB_TEXT, 30, True, False, False), _
        Array("OwnerID", AT_DB_LONG, 0, True, False, False), Array("BuildingNo", AT_DB_TEXT, 40, False, False, True), _
        Array("ApartmentNo", AT_DB_TEXT, 40, False, False, True), Array("FloorNo", AT_DB_TEXT, 30, False, False, True), _
        Array("ContributionAmount", AT_DB_CURRENCY, 0, True, False, False), Array("IsActive", AT_DB_BOOLEAN, 0, True, False, False), _
        Array("Notes", AT_DB_MEMO, 0, False, False, True), Array("CreatedAt", AT_DB_DATE, 0, False, False, False))

    AT_EnsureTable "tblAT_Charges", Array( _
        Array("ChargeID", AT_DB_LONG, 0, True, True, False), Array("ChargeCode", AT_DB_TEXT, 24, True, False, False), _
        Array("OwnerID", AT_DB_LONG, 0, True, False, False), Array("UnitID", AT_DB_LONG, 0, False, False, False), _
        Array("ChargeDate", AT_DB_DATE, 0, True, False, False), Array("DueDate", AT_DB_DATE, 0, False, False, False), _
        Array("Description", AT_DB_TEXT, 255, False, False, True), Array("OriginalAmount", AT_DB_CURRENCY, 0, True, False, False), _
        Array("CurrentAmount", AT_DB_CURRENCY, 0, True, False, False), Array("Status", AT_DB_TEXT, 20, True, False, False), _
        Array("SourceType", AT_DB_TEXT, 30, False, False, True), Array("CreatedAt", AT_DB_DATE, 0, False, False, False), _
        Array("CreatedBy", AT_DB_TEXT, 80, False, False, True))

    AT_EnsureTable "tblAT_PaymentMethods", Array( _
        Array("PaymentMethodID", AT_DB_LONG, 0, True, True, False), Array("MethodName", AT_DB_TEXT, 80, True, False, False), _
        Array("IsActive", AT_DB_BOOLEAN, 0, True, False, False), Array("Notes", AT_DB_TEXT, 255, False, False, True))

    AT_EnsureTable "tblAT_FinancialAccounts", Array( _
        Array("AccountID", AT_DB_LONG, 0, True, True, False), Array("AccountCode", AT_DB_TEXT, 30, True, False, False), _
        Array("AccountName", AT_DB_TEXT, 120, True, False, False), Array("AccountType", AT_DB_TEXT, 30, True, False, False), _
        Array("InstitutionName", AT_DB_TEXT, 120, False, False, True), Array("AccountReference", AT_DB_TEXT, 100, False, False, True), _
        Array("IsActive", AT_DB_BOOLEAN, 0, True, False, False), Array("CreatedAt", AT_DB_DATE, 0, False, False, False))

    AT_EnsureTable "tblAT_Receipts", Array( _
        Array("ReceiptID", AT_DB_LONG, 0, True, True, False), Array("ReceiptNo", AT_DB_TEXT, 24, False, False, True), _
        Array("OwnerID", AT_DB_LONG, 0, True, False, False), Array("ReceiptDate", AT_DB_DATE, 0, True, False, False), _
        Array("Amount", AT_DB_CURRENCY, 0, True, False, False), Array("PaymentMethodID", AT_DB_LONG, 0, True, False, False), _
        Array("AccountID", AT_DB_LONG, 0, True, False, False), Array("PayerName", AT_DB_TEXT, 150, False, False, True), _
        Array("CollectorName", AT_DB_TEXT, 150, False, False, True), Array("Notes", AT_DB_MEMO, 0, False, False, True), _
        Array("Status", AT_DB_TEXT, 20, True, False, False), Array("CreatedAt", AT_DB_DATE, 0, False, False, False), _
        Array("ApprovedAt", AT_DB_DATE, 0, False, False, False), Array("ApprovedBy", AT_DB_TEXT, 80, False, False, True), _
        Array("ReversedAt", AT_DB_DATE, 0, False, False, False), Array("ReversalReason", AT_DB_MEMO, 0, False, False, True))

    AT_EnsureTable "tblAT_ReceiptAllocations", Array( _
        Array("AllocationID", AT_DB_LONG, 0, True, True, False), Array("ReceiptID", AT_DB_LONG, 0, False, False, False), _
        Array("ChargeID", AT_DB_LONG, 0, True, False, False), Array("AmountSigned", AT_DB_CURRENCY, 0, True, False, False), _
        Array("AllocationType", AT_DB_TEXT, 30, True, False, False), Array("OriginalAllocationID", AT_DB_LONG, 0, False, False, False), _
        Array("CreatedAt", AT_DB_DATE, 0, False, False, False), Array("Notes", AT_DB_TEXT, 255, False, False, True))

    AT_EnsureTable "tblAT_OwnerCreditMovements", Array( _
        Array("CreditMovementID", AT_DB_LONG, 0, True, True, False), Array("OwnerID", AT_DB_LONG, 0, True, False, False), _
        Array("ReceiptID", AT_DB_LONG, 0, False, False, False), Array("PaymentVoucherID", AT_DB_LONG, 0, False, False, False), _
        Array("MovementDate", AT_DB_DATE, 0, True, False, False), Array("AmountSigned", AT_DB_CURRENCY, 0, True, False, False), _
        Array("MovementType", AT_DB_TEXT, 30, True, False, False), Array("Reason", AT_DB_TEXT, 255, False, False, True), _
        Array("CreatedAt", AT_DB_DATE, 0, False, False, False), Array("CreatedBy", AT_DB_TEXT, 80, False, False, True))

    AT_EnsureTable "tblAT_ExpenseCategories", Array( _
        Array("CategoryID", AT_DB_LONG, 0, True, True, False), Array("CategoryName", AT_DB_TEXT, 100, True, False, False), _
        Array("IsActive", AT_DB_BOOLEAN, 0, True, False, False))

    AT_EnsureTable "tblAT_Expenses", Array( _
        Array("ExpenseID", AT_DB_LONG, 0, True, True, False), Array("ExpenseNo", AT_DB_TEXT, 24, False, False, True), _
        Array("CategoryID", AT_DB_LONG, 0, False, False, False), Array("ExpenseDate", AT_DB_DATE, 0, True, False, False), _
        Array("Beneficiary", AT_DB_TEXT, 150, True, False, False), Array("Description", AT_DB_MEMO, 0, False, False, True), _
        Array("Amount", AT_DB_CURRENCY, 0, True, False, False), Array("Status", AT_DB_TEXT, 20, True, False, False), _
        Array("CreatedAt", AT_DB_DATE, 0, False, False, False), Array("CreatedBy", AT_DB_TEXT, 80, False, False, True))

    AT_EnsureTable "tblAT_PaymentVouchers", Array( _
        Array("VoucherID", AT_DB_LONG, 0, True, True, False), Array("VoucherNo", AT_DB_TEXT, 24, False, False, True), _
        Array("VoucherDate", AT_DB_DATE, 0, True, False, False), Array("VoucherType", AT_DB_TEXT, 30, True, False, False), _
        Array("ExpenseID", AT_DB_LONG, 0, False, False, False), Array("OwnerID", AT_DB_LONG, 0, False, False, False), _
        Array("PayeeName", AT_DB_TEXT, 150, True, False, False), Array("Amount", AT_DB_CURRENCY, 0, True, False, False), _
        Array("PaymentMethodID", AT_DB_LONG, 0, True, False, False), Array("AccountID", AT_DB_LONG, 0, True, False, False), _
        Array("Notes", AT_DB_MEMO, 0, False, False, True), Array("Status", AT_DB_TEXT, 20, True, False, False), _
        Array("CreatedAt", AT_DB_DATE, 0, False, False, False), Array("ApprovedAt", AT_DB_DATE, 0, False, False, False), _
        Array("ApprovedBy", AT_DB_TEXT, 80, False, False, True))

    AT_EnsureTable "tblAT_Transfers", Array( _
        Array("TransferID", AT_DB_LONG, 0, True, True, False), Array("TransferNo", AT_DB_TEXT, 24, False, False, True), _
        Array("TransferDate", AT_DB_DATE, 0, True, False, False), Array("FromAccountID", AT_DB_LONG, 0, True, False, False), _
        Array("ToAccountID", AT_DB_LONG, 0, True, False, False), Array("Amount", AT_DB_CURRENCY, 0, True, False, False), _
        Array("FeeAmount", AT_DB_CURRENCY, 0, True, False, False), Array("Status", AT_DB_TEXT, 20, True, False, False), _
        Array("Notes", AT_DB_MEMO, 0, False, False, True), Array("CreatedAt", AT_DB_DATE, 0, False, False, False), _
        Array("StartedAt", AT_DB_DATE, 0, False, False, False), Array("CompletedAt", AT_DB_DATE, 0, False, False, False))

    AT_EnsureTable "tblAT_AccountMovements", Array( _
        Array("MovementID", AT_DB_LONG, 0, True, True, False), Array("AccountID", AT_DB_LONG, 0, True, False, False), _
        Array("MovementDate", AT_DB_DATE, 0, True, False, False), Array("SourceType", AT_DB_TEXT, 30, True, False, False), _
        Array("SourceID", AT_DB_LONG, 0, False, False, False), Array("AmountSigned", AT_DB_CURRENCY, 0, True, False, False), _
        Array("Description", AT_DB_TEXT, 255, False, False, True), Array("CreatedAt", AT_DB_DATE, 0, False, False, False), _
        Array("CreatedBy", AT_DB_TEXT, 80, False, False, True), Array("ReversalOfMovementID", AT_DB_LONG, 0, False, False, False))

    AT_EnsureTable "tblAT_DiscountRequests", Array( _
        Array("DiscountRequestID", AT_DB_LONG, 0, True, True, False), Array("DiscountNo", AT_DB_TEXT, 24, False, False, True), _
        Array("ChargeID", AT_DB_LONG, 0, True, False, False), Array("RequestDate", AT_DB_DATE, 0, True, False, False), _
        Array("RequestedAmount", AT_DB_CURRENCY, 0, True, False, False), Array("Reason", AT_DB_MEMO, 0, True, False, False), _
        Array("Status", AT_DB_TEXT, 20, True, False, False), Array("AmountBefore", AT_DB_CURRENCY, 0, False, False, False), _
        Array("AmountAfter", AT_DB_CURRENCY, 0, False, False, False), Array("ApprovedAt", AT_DB_DATE, 0, False, False, False), _
        Array("ApprovedBy", AT_DB_TEXT, 80, False, False, True), Array("CreatedAt", AT_DB_DATE, 0, False, False, False))

    AT_EnsureTable "tblAT_AuditLog", Array( _
        Array("AuditID", AT_DB_LONG, 0, True, True, False), Array("EventDateTime", AT_DB_DATE, 0, True, False, False), _
        Array("UserName", AT_DB_TEXT, 80, True, False, False), Array("ActionName", AT_DB_TEXT, 80, True, False, False), _
        Array("EntityName", AT_DB_TEXT, 80, False, False, True), Array("EntityID", AT_DB_TEXT, 40, False, False, True), _
        Array("Result", AT_DB_TEXT, 30, True, False, False), Array("Details", AT_DB_MEMO, 0, False, False, True), _
        Array("OldValue", AT_DB_MEMO, 0, False, False, True), Array("NewValue", AT_DB_MEMO, 0, False, False, True))

    AT_EnsureTable "tblAT_Reversals", Array( _
        Array("ReversalID", AT_DB_LONG, 0, True, True, False), Array("SourceType", AT_DB_TEXT, 30, True, False, False), _
        Array("SourceID", AT_DB_LONG, 0, True, False, False), Array("ReversalDate", AT_DB_DATE, 0, True, False, False), _
        Array("Reason", AT_DB_MEMO, 0, True, False, False), Array("ApprovedBy", AT_DB_TEXT, 80, False, False, True), _
        Array("CreatedAt", AT_DB_DATE, 0, False, False, False))

    AT_EnsureTable "tblAT_PeriodClosings", Array( _
        Array("ClosingID", AT_DB_LONG, 0, True, True, False), Array("PeriodName", AT_DB_TEXT, 100, True, False, False), _
        Array("StartDate", AT_DB_DATE, 0, True, False, False), Array("EndDate", AT_DB_DATE, 0, True, False, False), _
        Array("CloseDate", AT_DB_DATE, 0, False, False, False), Array("Status", AT_DB_TEXT, 20, True, False, False), _
        Array("Notes", AT_DB_MEMO, 0, False, False, True), Array("CreatedBy", AT_DB_TEXT, 80, False, False, True), _
        Array("ApprovedBy", AT_DB_TEXT, 80, False, False, True), Array("ReopenedAt", AT_DB_DATE, 0, False, False, False), _
        Array("ReopenReason", AT_DB_MEMO, 0, False, False, True))

    AT_EnsureTable "tblAT_ClosingAccountBalances", Array( _
        Array("ClosingBalanceID", AT_DB_LONG, 0, True, True, False), Array("ClosingID", AT_DB_LONG, 0, True, False, False), _
        Array("AccountID", AT_DB_LONG, 0, True, False, False), Array("BookBalance", AT_DB_CURRENCY, 0, True, False, False), _
        Array("ActualBalance", AT_DB_CURRENCY, 0, True, False, False), Array("Difference", AT_DB_CURRENCY, 0, True, False, False), _
        Array("Notes", AT_DB_MEMO, 0, False, False, True))

    AT_EnsureTable "tblAT_DocumentLinks", Array( _
        Array("DocumentLinkID", AT_DB_LONG, 0, True, True, False), Array("EntityType", AT_DB_TEXT, 40, True, False, False), _
        Array("EntityID", AT_DB_LONG, 0, True, False, False), Array("FilePath", AT_DB_TEXT, 255, True, False, False), _
        Array("Description", AT_DB_TEXT, 255, False, False, True), Array("AddedAt", AT_DB_DATE, 0, False, False, False), _
        Array("AddedBy", AT_DB_TEXT, 80, False, False, True))

    AT_EnsureTable "tblAT_BuildLog", Array( _
        Array("BuildLogID", AT_DB_LONG, 0, True, True, False), Array("EventTime", AT_DB_DATE, 0, True, False, False), _
        Array("StepName", AT_DB_TEXT, 100, True, False, False), Array("Result", AT_DB_TEXT, 30, True, False, False), _
        Array("Details", AT_DB_MEMO, 0, False, False, True), Array("WindowsUser", AT_DB_TEXT, 80, False, False, True))

    AT_BuildIndexes
    AT_BuildRelationships
    AT_SeedReferenceData
    AT_LogBuild "Schema", "Success", "Schema and reference data ready"
    Exit Sub
BuildFail:
    AT_LogBuild "Schema", "Failed", Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "AT_BuildSchema", Err.Description
End Sub

Public Sub AT_EnsureTable(ByVal tableName As String, ByVal fieldSpecs As Variant)
    Dim db As Object, tdf As Object, fld As Object
    Dim i As Long, spec As Variant
    Set db = CurrentDb
    If Not AT_TableExists(tableName) Then
        Set tdf = db.CreateTableDef(tableName)
        For i = LBound(fieldSpecs) To UBound(fieldSpecs)
            spec = fieldSpecs(i)
            Set fld = tdf.CreateField(CStr(spec(0)), CLng(spec(1)), CLng(spec(2)))
            If CBool(spec(4)) Then fld.Attributes = AT_DB_AUTOINCR
            On Error Resume Next
            fld.Required = CBool(spec(3))
            fld.AllowZeroLength = CBool(spec(5))
            On Error GoTo 0
            tdf.Fields.Append fld
        Next i
        db.TableDefs.Append tdf
        spec = fieldSpecs(LBound(fieldSpecs))
        AT_AddPrimaryKey tableName, CStr(spec(0))
    Else
        Set tdf = db.TableDefs(tableName)
        For i = LBound(fieldSpecs) To UBound(fieldSpecs)
            spec = fieldSpecs(i)
            If Not AT_FieldExists(tableName, CStr(spec(0))) Then
                If CBool(spec(4)) Then
                    AT_LogBuild "SchemaField:" & tableName & "." & CStr(spec(0)), "Review", "AutoNumber field missing in existing table; create only after backup/review"
                Else
                    Set fld = tdf.CreateField(CStr(spec(0)), CLng(spec(1)), CLng(spec(2)))
                    On Error Resume Next
                    fld.Required = CBool(spec(3))
                    fld.AllowZeroLength = CBool(spec(5))
                    On Error GoTo 0
                    tdf.Fields.Append fld
                End If
            End If
        Next i
        spec = fieldSpecs(LBound(fieldSpecs))
        If AT_FieldExists(tableName, CStr(spec(0))) Then AT_AddPrimaryKey tableName, CStr(spec(0))
    End If
End Sub

Private Sub AT_AddPrimaryKey(ByVal tableName As String, ByVal fieldName As String)
    Dim tdf As Object, idx As Object, existing As Object
    Set tdf = CurrentDb.TableDefs(tableName)
    For Each existing In tdf.Indexes
        If existing.Primary = True Then Exit Sub
    Next
    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Unique = True
    idx.Fields.Append idx.CreateField(fieldName)
    tdf.Indexes.Append idx
End Sub

Private Sub AT_AddIndex(ByVal tableName As String, ByVal indexName As String, ByVal fieldName As String, Optional ByVal isUnique As Boolean = False)
    Dim tdf As Object, idx As Object
    Dim existing As Object
    Set tdf = CurrentDb.TableDefs(tableName)
    For Each existing In tdf.Indexes
        If StrComp(existing.Name, indexName, vbTextCompare) = 0 Then Exit Sub
    Next
    Set idx = tdf.CreateIndex(indexName)
    idx.Unique = isUnique
    If isUnique Then
        On Error Resume Next
        idx.IgnoreNulls = True
        On Error GoTo 0
    End If
    idx.Fields.Append idx.CreateField(fieldName)
    tdf.Indexes.Append idx
End Sub

Private Sub AT_BuildIndexes()
    AT_AddIndex "tblAT_Owners", "UX_AT_OwnerCode", "OwnerCode", True
    AT_AddIndex "tblAT_Owners", "IX_AT_OwnerPhone", "Phone"
    AT_AddIndex "tblAT_Units", "UX_AT_UnitCode", "UnitCode", True
    AT_AddIndex "tblAT_Charges", "UX_AT_ChargeCode", "ChargeCode", True
    AT_AddIndex "tblAT_Receipts", "UX_AT_ReceiptNo", "ReceiptNo", True
    AT_AddIndex "tblAT_Expenses", "UX_AT_ExpenseNo", "ExpenseNo", True
    AT_AddIndex "tblAT_PaymentVouchers", "UX_AT_VoucherNo", "VoucherNo", True
    AT_AddIndex "tblAT_Transfers", "UX_AT_TransferNo", "TransferNo", True
    AT_AddIndex "tblAT_DiscountRequests", "UX_AT_DiscountNo", "DiscountNo", True
    AT_AddIndex "tblAT_AccountMovements", "IX_AT_MovementAccount", "AccountID"
    AT_AddIndex "tblAT_Charges", "IX_AT_ChargeOwner", "OwnerID"
    AT_AddIndex "tblAT_Charges", "IX_AT_ChargeDueDate", "DueDate"
    AT_AddIndex "tblAT_ReceiptAllocations", "IX_AT_AllocationCharge", "ChargeID"
End Sub

Private Sub AT_AddRelation(ByVal relationName As String, ByVal primaryTable As String, ByVal foreignTable As String, ByVal primaryField As String, ByVal foreignField As String)
    Dim db As Object, rel As Object, fld As Object, existing As Object
    Set db = CurrentDb
    For Each existing In db.Relations
        If StrComp(existing.Name, relationName, vbTextCompare) = 0 Then Exit Sub
    Next
    Set rel = db.CreateRelation(relationName, primaryTable, foreignTable, 0)
    Set fld = rel.CreateField(primaryField)
    fld.ForeignName = foreignField
    rel.Fields.Append fld
    db.Relations.Append rel
End Sub

Private Sub AT_BuildRelationships()
    On Error GoTo RelationFail
    AT_AddRelation "REL_AT_Owner_Units", "tblAT_Owners", "tblAT_Units", "OwnerID", "OwnerID"
    AT_AddRelation "REL_AT_Owner_Charges", "tblAT_Owners", "tblAT_Charges", "OwnerID", "OwnerID"
    AT_AddRelation "REL_AT_Unit_Charges", "tblAT_Units", "tblAT_Charges", "UnitID", "UnitID"
    AT_AddRelation "REL_AT_Owner_Receipts", "tblAT_Owners", "tblAT_Receipts", "OwnerID", "OwnerID"
    AT_AddRelation "REL_AT_Owner_Credits", "tblAT_Owners", "tblAT_OwnerCreditMovements", "OwnerID", "OwnerID"
    AT_AddRelation "REL_AT_Owner_Vouchers", "tblAT_Owners", "tblAT_PaymentVouchers", "OwnerID", "OwnerID"
    AT_AddRelation "REL_AT_Receipt_Allocations", "tblAT_Receipts", "tblAT_ReceiptAllocations", "ReceiptID", "ReceiptID"
    AT_AddRelation "REL_AT_Receipt_Credits", "tblAT_Receipts", "tblAT_OwnerCreditMovements", "ReceiptID", "ReceiptID"
    AT_AddRelation "REL_AT_Voucher_Credits", "tblAT_PaymentVouchers", "tblAT_OwnerCreditMovements", "VoucherID", "PaymentVoucherID"
    AT_AddRelation "REL_AT_Charge_Allocations", "tblAT_Charges", "tblAT_ReceiptAllocations", "ChargeID", "ChargeID"
    AT_AddRelation "REL_AT_Account_Movements", "tblAT_FinancialAccounts", "tblAT_AccountMovements", "AccountID", "AccountID"
    AT_AddRelation "REL_AT_Transfer_FromAccount", "tblAT_FinancialAccounts", "tblAT_Transfers", "AccountID", "FromAccountID"
    AT_AddRelation "REL_AT_Transfer_ToAccount", "tblAT_FinancialAccounts", "tblAT_Transfers", "AccountID", "ToAccountID"
    AT_AddRelation "REL_AT_Method_Receipts", "tblAT_PaymentMethods", "tblAT_Receipts", "PaymentMethodID", "PaymentMethodID"
    AT_AddRelation "REL_AT_Method_Vouchers", "tblAT_PaymentMethods", "tblAT_PaymentVouchers", "PaymentMethodID", "PaymentMethodID"
    AT_AddRelation "REL_AT_Account_Receipts", "tblAT_FinancialAccounts", "tblAT_Receipts", "AccountID", "AccountID"
    AT_AddRelation "REL_AT_Account_Vouchers", "tblAT_FinancialAccounts", "tblAT_PaymentVouchers", "AccountID", "AccountID"
    AT_AddRelation "REL_AT_Category_Expenses", "tblAT_ExpenseCategories", "tblAT_Expenses", "CategoryID", "CategoryID"
    AT_AddRelation "REL_AT_Expense_Vouchers", "tblAT_Expenses", "tblAT_PaymentVouchers", "ExpenseID", "ExpenseID"
    AT_AddRelation "REL_AT_Charge_Discounts", "tblAT_Charges", "tblAT_DiscountRequests", "ChargeID", "ChargeID"
    AT_AddRelation "REL_AT_Closing_Balances", "tblAT_PeriodClosings", "tblAT_ClosingAccountBalances", "ClosingID", "ClosingID"
    AT_AddRelation "REL_AT_Account_ClosingBalances", "tblAT_FinancialAccounts", "tblAT_ClosingAccountBalances", "AccountID", "AccountID"
    Exit Sub
RelationFail:
    AT_LogBuild "Relationships", "Warning", Err.Number & " - " & Err.Description
    Err.Clear
End Sub

Private Sub AT_SeedReferenceData()
    If DCount("*", "tblAT_AppSettings", "SettingKey='Currency'") = 0 Then AT_SetSetting "Currency", AT_DEFAULT_CURRENCY
    If DCount("*", "tblAT_AppSettings", "SettingKey='ContributionPerUnit'") = 0 Then AT_SetSetting "ContributionPerUnit", CStr(AT_DEFAULT_CONTRIBUTION)
    If DCount("*", "tblAT_AppSettings", "SettingKey='SystemName'") = 0 Then AT_SetSetting "SystemName", "Archon Tech | Sewer Contributions"
    If DCount("*", "tblAT_AppSettings", "SettingKey='SchemaVersion'") = 0 Then AT_SetSetting "SchemaVersion", AT_VERSION

    If DCount("*", "tblAT_PaymentMethods") = 0 Then
        CurrentDb.Execute "INSERT INTO tblAT_PaymentMethods (MethodName, IsActive, Notes) VALUES ('" & AT_U("0646,0642,062F,064A") & "', True, 'Cash')", AT_DB_FAIL_ON_ERROR
        CurrentDb.Execute "INSERT INTO tblAT_PaymentMethods (MethodName, IsActive, Notes) VALUES ('" & AT_U("062A,062D,0648,064A,0644") & " " & AT_U("0628,0646,0643,064A") & "', True, 'Bank transfer')", AT_DB_FAIL_ON_ERROR
        CurrentDb.Execute "INSERT INTO tblAT_PaymentMethods (MethodName, IsActive, Notes) VALUES ('" & AT_U("0645,062D,0641,0638,0629") & " " & AT_U("0625,0644,0643,062A,0631,0648,0646,064A,0629") & "', True, 'E-wallet')", AT_DB_FAIL_ON_ERROR
    End If
    If DCount("*", "tblAT_FinancialAccounts") = 0 Then
        CurrentDb.Execute "INSERT INTO tblAT_FinancialAccounts (AccountCode, AccountName, AccountType, IsActive, CreatedAt) VALUES ('CASH-001','" & AT_U("0627,0644,0635,0646,062F,0648,0642") & " " & AT_U("0627,0644,0646,0642,062F,064A") & "','Cash',True,Now())", AT_DB_FAIL_ON_ERROR
    End If
    If DCount("*", "tblAT_ExpenseCategories") = 0 Then
        CurrentDb.Execute "INSERT INTO tblAT_ExpenseCategories (CategoryName, IsActive) VALUES ('" & AT_U("0635,064A,0627,0646,0629") & "',True)", AT_DB_FAIL_ON_ERROR
        CurrentDb.Execute "INSERT INTO tblAT_ExpenseCategories (CategoryName, IsActive) VALUES ('" & AT_U("0646,0642,0644") & " " & AT_U("0648,0645,0648,0627,0635,0644,0627,062A") & "',True)", AT_DB_FAIL_ON_ERROR
        CurrentDb.Execute "INSERT INTO tblAT_ExpenseCategories (CategoryName, IsActive) VALUES ('" & AT_U("0645,0633,062A,0644,0632,0645,0627,062A") & "',True)", AT_DB_FAIL_ON_ERROR
        CurrentDb.Execute "INSERT INTO tblAT_ExpenseCategories (CategoryName, IsActive) VALUES ('" & AT_U("0631,0633,0648,0645") & " " & AT_U("062A,062D,0648,064A,0644") & "',True)", AT_DB_FAIL_ON_ERROR
        CurrentDb.Execute "INSERT INTO tblAT_ExpenseCategories (CategoryName, IsActive) VALUES ('" & AT_U("0623,062E,0631,0649") & "',True)", AT_DB_FAIL_ON_ERROR
    End If
End Sub
