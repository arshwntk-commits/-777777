Attribute VB_Name = "modArchonReports"
Option Compare Database
Option Explicit

' Creates saved, read-only query sources for core operational reports.
' These are report datasets; printable report layouts, PDF export, and formatting still require Access-side build and validation.

Public Function AT_BuildReportQueries() As Boolean
    Dim savedErrNumber As Long, savedErrDescription As String, savedErrSource As String
    Dim db As DAO.Database
    On Error GoTo EH
    Set db = CurrentDb

    AT_ReplaceQuery db, "qryAT_ChargesByOwner", _
        "SELECT C.OwnerID, Sum(C.GrossAmount) AS GrossCharges, Sum(Nz(C.NetAmount,C.GrossAmount)) AS NetCharges " & _
        "FROM tblOwnerCharges AS C INNER JOIN tblFinancialDocuments AS D ON C.DocumentID=D.DocumentID " & _
        "WHERE C.ChargeStatus='ACTIVE' AND D.StatusCode='APPROVED' GROUP BY C.OwnerID;"

    AT_ReplaceQuery db, "qryAT_AdjustmentsByOwner", _
        "SELECT C.OwnerID, Sum(A.SignedAmount) AS AdjustmentTotal " & _
        "FROM (tblOwnerAdjustments AS A INNER JOIN tblFinancialDocuments AS D ON A.DocumentID=D.DocumentID) " & _
        "INNER JOIN tblOwnerCharges AS C ON A.ChargeID=C.ChargeID " & _
        "WHERE D.StatusCode='APPROVED' GROUP BY C.OwnerID;"

    AT_ReplaceQuery db, "qryAT_SettlementsByOwner", _
        "SELECT C.OwnerID, Sum(S.SignedAmount) AS SettlementTotal " & _
        "FROM (tblChargeSettlementLines AS S INNER JOIN tblFinancialDocuments AS D ON S.DocumentID=D.DocumentID) " & _
        "INNER JOIN tblOwnerCharges AS C ON S.ChargeID=C.ChargeID " & _
        "WHERE D.StatusCode='APPROVED' GROUP BY C.OwnerID;"

    AT_ReplaceQuery db, "qryAT_CreditsByOwner", _
        "SELECT C.OwnerID, Sum(C.SignedAmount) AS CreditBalance " & _
        "FROM tblOwnerCreditMovements AS C INNER JOIN tblFinancialDocuments AS D ON C.DocumentID=D.DocumentID " & _
        "WHERE D.StatusCode='APPROVED' GROUP BY C.OwnerID;"

    AT_ReplaceQuery db, "qryAT_OwnerBalances", _
        "SELECT O.OwnerID, O.OwnerCode, O.FullName, O.Phone, O.HouseCount, O.ApartmentCount, O.AdminStatus, " & _
        "Nz(C.GrossCharges,0) AS GrossCharges, Nz(C.NetCharges,0) AS NetCharges, Nz(A.AdjustmentTotal,0) AS AdjustmentTotal, " & _
        "(Nz(C.NetCharges,0)+Nz(A.AdjustmentTotal,0)) AS NetDueBase, Nz(S.SettlementTotal,0) AS NetSettlements, " & _
        "IIf((Nz(C.NetCharges,0)+Nz(A.AdjustmentTotal,0)-Nz(S.SettlementTotal,0))<0,0," & _
        "(Nz(C.NetCharges,0)+Nz(A.AdjustmentTotal,0)-Nz(S.SettlementTotal,0))) AS RemainingDue, " & _
        "Nz(CR.CreditBalance,0) AS AvailableCredit, " & _
        "IIf((Nz(C.NetCharges,0)+Nz(A.AdjustmentTotal,0))<=0,Null," & _
        "IIf(Nz(S.SettlementTotal,0)<=0,0,IIf(100*Nz(S.SettlementTotal,0)/IIf((Nz(C.NetCharges,0)+Nz(A.AdjustmentTotal,0))<=0,1,(Nz(C.NetCharges,0)+Nz(A.AdjustmentTotal,0)))>100,100," & _
        "100*Nz(S.SettlementTotal,0)/IIf((Nz(C.NetCharges,0)+Nz(A.AdjustmentTotal,0))<=0,1,(Nz(C.NetCharges,0)+Nz(A.AdjustmentTotal,0)))))) AS CollectionPercentage " & _
        "FROM ((((tblOwners AS O LEFT JOIN qryAT_ChargesByOwner AS C ON O.OwnerID=C.OwnerID) " & _
        "LEFT JOIN qryAT_AdjustmentsByOwner AS A ON O.OwnerID=A.OwnerID) " & _
        "LEFT JOIN qryAT_SettlementsByOwner AS S ON O.OwnerID=S.OwnerID) " & _
        "LEFT JOIN qryAT_CreditsByOwner AS CR ON O.OwnerID=CR.OwnerID);"

    AT_ReplaceQuery db, "qryAT_Arrears", _
        "SELECT * FROM qryAT_OwnerBalances WHERE RemainingDue>0 AND AdminStatus='ACTIVE' ORDER BY RemainingDue DESC, FullName;"

    AT_ReplaceQuery db, "qryAT_BalancesByAccount", _
        "SELECT M.AccountID, Sum(M.SignedAmount) AS BookBalance " & _
        "FROM tblAccountMovements AS M INNER JOIN tblFinancialDocuments AS D ON M.DocumentID=D.DocumentID " & _
        "WHERE D.StatusCode='APPROVED' GROUP BY M.AccountID;"

    AT_ReplaceQuery db, "qryAT_ReservationsByAccount", _
        "SELECT AccountID, Sum(Amount) AS ReservedBalance FROM tblAccountReservations " & _
        "WHERE StatusCode='ACTIVE' GROUP BY AccountID;"

    AT_ReplaceQuery db, "qryAT_AccountBalances", _
        "SELECT A.AccountID, A.AccountCode, A.AccountName, A.AccountType, A.IsActive, " & _
        "Nz(B.BookBalance,0) AS BookBalance, Nz(R.ReservedBalance,0) AS ReservedBalance, " & _
        "(Nz(B.BookBalance,0)-Nz(R.ReservedBalance,0)) AS AvailableBalance " & _
        "FROM (tblAccounts AS A LEFT JOIN qryAT_BalancesByAccount AS B ON A.AccountID=B.AccountID) " & _
        "LEFT JOIN qryAT_ReservationsByAccount AS R ON A.AccountID=R.AccountID ORDER BY A.AccountName;"

    AT_ReplaceQuery db, "qryAT_ReceiptsHistory", _
        "SELECT N.OfficialNumber, R.ReceiptID, R.PrintedOwnerCode AS OwnerCode, R.PrintedOwnerName AS OwnerName, " & _
        "D.TransactionDate, D.StatusCode, R.Amount, PM.MethodName AS PaymentMethod, A.AccountName, R.PayerName, R.CollectorName, R.Notes " & _
        "FROM (((tblReceipts AS R INNER JOIN tblFinancialDocuments AS D ON R.DocumentID=D.DocumentID) " & _
        "LEFT JOIN tblOfficialDocumentNumbers AS N ON D.DocumentID=N.DocumentID) " & _
        "LEFT JOIN tblPaymentMethods AS PM ON R.PaymentMethodID=PM.PaymentMethodID) " & _
        "LEFT JOIN tblAccounts AS A ON R.AccountID=A.AccountID ORDER BY D.TransactionDate DESC, R.ReceiptID DESC;"

    AT_ReplaceQuery db, "qryAT_PaymentsHistory", _
        "SELECT N.OfficialNumber, P.PaymentID, D.TransactionDate, D.StatusCode, P.BeneficiaryName, P.RecipientName, " & _
        "P.Amount, A.AccountName, E.CategoryName, P.PaymentMethodID, P.Notes " & _
        "FROM (((tblPayments AS P INNER JOIN tblFinancialDocuments AS D ON P.DocumentID=D.DocumentID) " & _
        "LEFT JOIN tblOfficialDocumentNumbers AS N ON D.DocumentID=N.DocumentID) " & _
        "LEFT JOIN tblAccounts AS A ON P.AccountID=A.AccountID) " & _
        "LEFT JOIN tblExpenseCategories AS E ON P.ExpenseCategoryID=E.ExpenseCategoryID ORDER BY D.TransactionDate DESC, P.PaymentID DESC;"

    AT_ReplaceQuery db, "qryAT_AccountMovements", _
        "SELECT M.AccountMovementID, A.AccountCode, A.AccountName, D.TransactionDate, D.DocumentType, N.OfficialNumber, " & _
        "M.MovementType, M.SignedAmount, M.LineNo, D.StatusCode " & _
        "FROM ((tblAccountMovements AS M INNER JOIN tblAccounts AS A ON M.AccountID=A.AccountID) " & _
        "INNER JOIN tblFinancialDocuments AS D ON M.DocumentID=D.DocumentID) " & _
        "LEFT JOIN tblOfficialDocumentNumbers AS N ON D.DocumentID=N.DocumentID " & _
        "ORDER BY D.TransactionDate DESC, M.AccountMovementID DESC;"

    AT_ReplaceQuery db, "qryAT_PendingAdjustments", _
        "SELECT A.AdjustmentID, C.OwnerID, O.OwnerCode, O.FullName, C.ChargeCode, A.AdjustmentType, A.SignedAmount, A.Reason, " & _
        "A.RequestedByUserName, D.TransactionDate, D.StatusCode " & _
        "FROM (((tblOwnerAdjustments AS A INNER JOIN tblFinancialDocuments AS D ON A.DocumentID=D.DocumentID) " & _
        "INNER JOIN tblOwnerCharges AS C ON A.ChargeID=C.ChargeID) INNER JOIN tblOwners AS O ON C.OwnerID=O.OwnerID) " & _
        "WHERE D.StatusCode='DRAFT' ORDER BY D.TransactionDate, A.AdjustmentID;"

    AT_ReplaceQuery db, "qryAT_CreditLedger", _
        "SELECT CM.CreditMovementID, O.OwnerCode, O.FullName, D.TransactionDate, N.OfficialNumber, CM.CreditType, " & _
        "CM.SignedAmount, IIf(CM.SignedAmount>0 AND CM.SourceCreditMovementID Is Null,AT_CreditMovementAvailable(CM.CreditMovementID),Null) AS AvailableAmount, " & _
        "CM.SourceReceiptID, CM.SourceCreditMovementID, CM.Reason, D.StatusCode " & _
        "FROM (((tblOwnerCreditMovements AS CM INNER JOIN tblOwners AS O ON CM.OwnerID=O.OwnerID) " & _
        "INNER JOIN tblFinancialDocuments AS D ON CM.DocumentID=D.DocumentID) " & _
        "LEFT JOIN tblOfficialDocumentNumbers AS N ON D.DocumentID=N.DocumentID) " & _
        "ORDER BY D.TransactionDate DESC, CM.CreditMovementID DESC;"

    AT_ReplaceQuery db, "qryAT_ChargeBalances", _
        "SELECT C.ChargeID, C.ChargeCode, O.OwnerCode, O.FullName, C.ChargeDate, C.ApartmentCountAtCharge, C.RateAmount, " & _
        "C.GrossAmount, C.DiscountAmount, C.NetAmount, AT_ChargeRemaining(C.ChargeID) AS RemainingDue, D.StatusCode " & _
        "FROM (tblOwnerCharges AS C INNER JOIN tblOwners AS O ON C.OwnerID=O.OwnerID) " & _
        "INNER JOIN tblFinancialDocuments AS D ON C.DocumentID=D.DocumentID " & _
        "WHERE C.ChargeStatus='ACTIVE' AND D.StatusCode='APPROVED' ORDER BY O.FullName, C.ChargeDate;"

    AT_ReplaceQuery db, "qryAT_OwnerStatement", _
        "SELECT C.OwnerID AS OwnerID, C.ChargeDate AS EntryDate, C.ChargeCode AS ReferenceText, " & _
        AT_U("0027,0043,0048,0041,0052,0047,0045,0027,0020,0041,0053,0020,0045,006E,0074,0072,0079,0054,0079,0070,0065,002C,0020,0027,0627,0633,062A,062D,0642,0627,0642,0020,0627,0644,0631,0633,0648,0645,0027,0020,0041,0053,0020,0044,0065,0074,0061,0069,006C,0054,0065,0078,0074,002C,0020,0043,002E,004E,0065,0074,0041,006D,006F,0075,006E,0074,0020,0041,0053,0020,0044,0065,0062,0069,0074,0041,006D,006F,0075,006E,0074,002C,0020,0043,0043,0075,0072,0028,0030,0029,0020,0041,0053,0020,0043,0072,0065,0064,0069,0074,0041,006D,006F,0075,006E,0074,002C,0020,0043,002E,0043,0068,0061,0072,0067,0065,0049,0044,0020,0041,0053,0020,0053,006F,0075,0072,0063,0065,0049,0044,0020") & _
        "FROM (tblOwnerCharges AS C INNER JOIN tblFinancialDocuments AS D ON C.DocumentID=D.DocumentID) " & _
        "WHERE C.ChargeStatus='ACTIVE' AND D.StatusCode='APPROVED' " & _
        "UNION ALL SELECT C.OwnerID AS OwnerID, D.TransactionDate AS EntryDate, 'ADJ-' & A.AdjustmentID AS ReferenceText, " & _
        "'ADJUSTMENT' AS EntryType, Left(Nz(A.Reason,''),255) AS DetailText, CCur(0) AS DebitAmount, -A.SignedAmount AS CreditAmount, A.AdjustmentID AS SourceID " & _
        "FROM ((tblOwnerAdjustments AS A INNER JOIN tblFinancialDocuments AS D ON A.DocumentID=D.DocumentID) " & _
        "INNER JOIN tblOwnerCharges AS C ON A.ChargeID=C.ChargeID) " & _
        "WHERE D.StatusCode='APPROVED' AND A.SignedAmount<0 " & _
        "UNION ALL SELECT C.OwnerID AS OwnerID, D.TransactionDate AS EntryDate, " & _
        "IIf(N.OfficialNumber Is Null,'DOC-' & D.DocumentID,N.OfficialNumber) AS ReferenceText, " & _
        "S.SettlementType AS EntryType, S.SettlementType AS DetailText, " & _
        "IIf(S.SignedAmount<0,-S.SignedAmount,CCur(0)) AS DebitAmount, " & _
        "IIf(S.SignedAmount>0,S.SignedAmount,CCur(0)) AS CreditAmount, S.SettlementID AS SourceID " & _
        "FROM ((tblChargeSettlementLines AS S INNER JOIN tblFinancialDocuments AS D ON S.DocumentID=D.DocumentID) " & _
        "INNER JOIN tblOwnerCharges AS C ON S.ChargeID=C.ChargeID) " & _
        "LEFT JOIN tblOfficialDocumentNumbers AS N ON D.DocumentID=N.DocumentID " & _
        "WHERE D.StatusCode='APPROVED' ORDER BY EntryDate, ReferenceText;"

    AT_ReplaceQuery db, "qryAT_OwnerCurrentUnitCounts", _
        "SELECT BO.OwnerID, Count(U.UnitID) AS UnitCount FROM (tblBuildingOwnership AS BO INNER JOIN tblBuildings AS B ON BO.BuildingID=B.BuildingID) INNER JOIN tblUnits AS U ON B.BuildingID=U.BuildingID WHERE BO.StartDate<=Date() AND (BO.EndDate Is Null OR BO.EndDate>=Date()) AND B.IsActive=True AND U.IsActive=True GROUP BY BO.OwnerID;"

    ' Filter current ownership in a saved query. Access does not accept the compound
    ' date/OR predicate reliably inside the ON clause of these LEFT JOIN expressions.
    AT_ReplaceQuery db, "qryAT_CurrentBuildingOwnership", _
        "SELECT BO.OwnershipID, BO.BuildingID, BO.OwnerID, BO.StartDate, BO.EndDate " & _
        "FROM tblBuildingOwnership AS BO " & _
        "WHERE BO.StartDate<=Date() AND (BO.EndDate Is Null OR BO.EndDate>=Date());"

    AT_ReplaceQuery db, "qryAT_BuildingCurrentOwners", _
        "SELECT B.BuildingID,B.BuildingCode,B.BuildingName,B.Address,O.OwnerID,O.OwnerCode,O.FullName,BO.StartDate AS OwnershipStart " & _
        "FROM (tblBuildings AS B LEFT JOIN qryAT_CurrentBuildingOwnership AS BO ON B.BuildingID=BO.BuildingID) " & _
        "LEFT JOIN tblOwners AS O ON BO.OwnerID=O.OwnerID ORDER BY B.BuildingCode;"

    AT_ReplaceQuery db, "qryAT_UnitDirectory", _
        "SELECT B.BuildingCode,B.BuildingName,U.UnitID,U.UnitCode,U.FloorNo,U.IsContributionEligible,U.IsActive,BO.OwnerID,O.OwnerCode,O.FullName " & _
        "FROM ((tblUnits AS U INNER JOIN tblBuildings AS B ON U.BuildingID=B.BuildingID) " & _
        "LEFT JOIN qryAT_CurrentBuildingOwnership AS BO ON B.BuildingID=BO.BuildingID) " & _
        "LEFT JOIN tblOwners AS O ON BO.OwnerID=O.OwnerID ORDER BY B.BuildingCode,U.UnitCode;"

    AT_ReplaceQuery db, "qryAT_ChargeLineRegister", _
        "SELECT C.ChargeCode,C.ChargeDate,O.OwnerCode,O.FullName,L.BuildingCodeSnapshot,L.UnitCodeSnapshot,L.RateAmount,L.GrossAmount,L.DiscountAmount,L.NetAmount FROM ((tblChargeLines AS L INNER JOIN tblOwnerCharges AS C ON L.ChargeID=C.ChargeID) INNER JOIN tblOwners AS O ON C.OwnerID=O.OwnerID) ORDER BY C.ChargeDate DESC,C.ChargeID,L.ChargeLineID;"

    AT_ReplaceQuery db, "qryAT_UnitExemptionReview", _
        "SELECT X.ExemptionRequestID,B.BuildingCode,U.UnitCode,X.RequestDate,X.EffectiveFrom,X.EffectiveTo,X.Reason,X.EvidenceReference,X.RequestedByUserName,X.StatusCode,X.DecidedByUserName,X.DecidedAt,X.DecisionReason FROM ((tblUnitExemptionRequests AS X INNER JOIN tblUnits AS U ON X.UnitID=U.UnitID) INNER JOIN tblBuildings AS B ON U.BuildingID=B.BuildingID) ORDER BY X.RequestDate DESC,X.ExemptionRequestID DESC;"

    AT_ReplaceQuery db, "qryAT_CampaignRegister", _
        "SELECT C.CampaignID,C.CampaignCode,C.CampaignName,T.TypeName,C.TargetScope,C.AmountPerUnit,C.PeriodStart,C.PeriodEnd,C.DueDate,C.StatusCode,C.CreatedAt,C.CreatedByUserName,C.ApprovedAt,C.ApprovedByUserName FROM tblContributionCampaigns AS C INNER JOIN tblContributionTypes AS T ON C.ContributionTypeID=T.ContributionTypeID ORDER BY C.CreatedAt DESC;"

    AT_ReplaceQuery db, "qryAT_ExpenseRegister", _
        "SELECT E.ExpenseID,E.ExpenseNo,E.ExpenseDate,E.BeneficiaryName,E.Description,E.Amount,E.StatusCode,C.CategoryName FROM tblExpenses AS E LEFT JOIN tblExpenseCategories AS C ON E.CategoryID=C.ExpenseCategoryID ORDER BY E.ExpenseDate DESC,E.ExpenseID DESC;"

    AT_ReplaceQuery db, "qryAT_ReceiptDraftAllocations", _
        "SELECT A.AllocationID,R.ReceiptID,O.OwnerCode,O.FullName,C.ChargeCode,A.AllocatedAmount,D.StatusCode " & _
        "FROM (((tblReceiptAllocationLines AS A INNER JOIN tblReceipts AS R ON A.ReceiptID=R.ReceiptID) " & _
        "INNER JOIN tblFinancialDocuments AS D ON R.DocumentID=D.DocumentID) " & _
        "INNER JOIN tblOwners AS O ON R.OwnerID=O.OwnerID) INNER JOIN tblOwnerCharges AS C ON A.ChargeID=C.ChargeID " & _
        "WHERE D.StatusCode='DRAFT' ORDER BY R.ReceiptID,A.LineNo;"

    AT_ReplaceQuery db, "qryAT_ReceiptPrint", _
        "SELECT R.ReceiptID,N.OfficialNumber,D.TransactionDate,R.Amount,R.PayerName,R.CollectorName,R.Notes,R.PrintedOwnerCode AS OwnerCode,R.PrintedOwnerName AS OwnerName,R.PrintedNeighborhoodName AS NeighborhoodName,PM.MethodName AS PaymentMethod,A.AccountName FROM ((((tblReceipts AS R INNER JOIN tblFinancialDocuments AS D ON R.DocumentID=D.DocumentID) LEFT JOIN tblOfficialDocumentNumbers AS N ON D.DocumentID=N.DocumentID) LEFT JOIN tblPaymentMethods AS PM ON R.PaymentMethodID=PM.PaymentMethodID) LEFT JOIN tblAccounts AS A ON R.AccountID=A.AccountID) WHERE D.StatusCode='APPROVED';"

    AT_BuildReportQueries = True
    MsgBox AT_U("062A,0645,0020,0625,0646,0634,0627,0621,002F,062A,062D,062F,064A,062B,0020,0627,0633,062A,0639,0644,0627,0645,0627,062A,0020,0627,0644,062A,0642,0627,0631,064A,0631,0020,0627,0644,0623,0633,0627,0633,064A,0629,002E,0020,0647,0630,0647,0020,0645,0635,0627,062F,0631,0020,0628,064A,0627,0646,0627,062A,061B,0020,0644,0627,0020,062A,0632,0627,0644,0020,062A,062E,0637,064A,0637,0627,062A,0020,0627,0644,0637,0628,0627,0639,0629,0020,0648,0050,0044,0046,0020,0628,062D,0627,062C,0629,0020,0625,0644,0649,0020,0628,0646,0627,0621,0020,0648,0627,062E,062A,0628,0627,0631,0020,062F,0627,062E,0644,0020,0041,0063,0063,0065,0073,0073,002E"), vbInformation, "Archon Tech"
    Exit Function
EH:
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    savedErrSource = Err.Source
    If Len(savedErrSource) = 0 Then savedErrSource = "AT_BuildReportQueries"
    Err.Raise savedErrNumber, savedErrSource, savedErrDescription
End Function

Private Sub AT_ReplaceQuery(ByVal db As DAO.Database, ByVal queryName As String, ByVal sqlText As String)
    Dim existingQuery As DAO.QueryDef, validationRecordset As DAO.Recordset
    Dim existingSql As String, hadExisting As Boolean, priorRemoved As Boolean, newDefinitionCreated As Boolean
    Dim lookupErrNumber As Long, lookupErrDescription As String
    Dim savedErrNumber As Long, savedErrDescription As String
    Dim cleanupErrNumber As Long, cleanupErrDescription As String
    Dim restoreErrNumber As Long, restoreErrDescription As String
    On Error GoTo EH

    ' Preserve the previous definition before replacing a builder-managed query.
    On Error Resume Next
    Set existingQuery = db.QueryDefs(queryName)
    lookupErrNumber = Err.Number
    lookupErrDescription = Err.Description
    Err.Clear
    On Error GoTo EH

    If lookupErrNumber = 0 Then
        hadExisting = True
        existingSql = existingQuery.SQL
        Set existingQuery = Nothing
        db.QueryDefs.Delete queryName
        priorRemoved = True
    ElseIf lookupErrNumber <> 3265 Then
        Err.Raise lookupErrNumber, "AT_ReplaceQuery[" & queryName & "]", lookupErrDescription
    End If

    db.CreateQueryDef queryName, sqlText
    newDefinitionCreated = True

    ' Open each saved query as a read-only snapshot. This makes ACE resolve its
    ' tables, fields, expressions, and previously-created query dependencies now,
    ' so an error is attributed to the exact query rather than a later form.
    Set existingQuery = db.QueryDefs(queryName)
    Set validationRecordset = existingQuery.OpenRecordset(dbOpenSnapshot)
    validationRecordset.Close
    Set validationRecordset = Nothing
    Set existingQuery = Nothing
    Exit Sub

EH:
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    On Error Resume Next
    Err.Clear
    If Not validationRecordset Is Nothing Then validationRecordset.Close
    Set validationRecordset = Nothing
    Set existingQuery = Nothing

    If newDefinitionCreated Then
        db.QueryDefs.Delete queryName
        cleanupErrNumber = Err.Number
        cleanupErrDescription = Err.Description
        Err.Clear
    End If

    ' Restore the prior definition only when this routine successfully removed it.
    ' Never delete a pre-existing query just because an earlier lookup/delete failed.
    If hadExisting And priorRemoved And cleanupErrNumber = 0 Then
        db.CreateQueryDef queryName, existingSql
        restoreErrNumber = Err.Number
        restoreErrDescription = Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    If cleanupErrNumber <> 0 Then savedErrDescription = savedErrDescription & " | Failed to remove the invalid replacement: " & cleanupErrDescription
    If restoreErrNumber <> 0 Then savedErrDescription = savedErrDescription & " | Failed to restore the previous query definition: " & restoreErrDescription
    Err.Raise savedErrNumber, "AT_ReplaceQuery[" & queryName & "]", savedErrDescription
End Sub
