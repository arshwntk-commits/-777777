Attribute VB_Name = "modArchonReports"
Option Compare Database
Option Explicit

' Creates saved, read-only query sources for core operational reports.
' These are report datasets; printable report layouts, PDF export, and formatting still require Access-side build and validation.

Public Function AT_BuildReportQueries() As Boolean
    On Error GoTo EH
    Dim db As DAO.Database
    Set db = CurrentDb

    AT_ReplaceQuery db, "qryAT_ChargesByOwner", _
        "SELECT C.OwnerID, Sum(C.GrossAmount) AS GrossCharges " & _
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
        "Nz(C.GrossCharges,0) AS GrossCharges, Nz(A.AdjustmentTotal,0) AS AdjustmentTotal, " & _
        "(Nz(C.GrossCharges,0)+Nz(A.AdjustmentTotal,0)) AS NetDueBase, Nz(S.SettlementTotal,0) AS NetSettlements, " & _
        "IIf((Nz(C.GrossCharges,0)+Nz(A.AdjustmentTotal,0)-Nz(S.SettlementTotal,0))<0,0," & _
        "(Nz(C.GrossCharges,0)+Nz(A.AdjustmentTotal,0)-Nz(S.SettlementTotal,0))) AS RemainingDue, " & _
        "Nz(CR.CreditBalance,0) AS AvailableCredit, " & _
        "IIf((Nz(C.GrossCharges,0)+Nz(A.AdjustmentTotal,0))<=0,Null," & _
        "IIf(Nz(S.SettlementTotal,0)<=0,0,IIf(100*Nz(S.SettlementTotal,0)/IIf((Nz(C.GrossCharges,0)+Nz(A.AdjustmentTotal,0))<=0,1,(Nz(C.GrossCharges,0)+Nz(A.AdjustmentTotal,0)))>100,100," & _
        "100*Nz(S.SettlementTotal,0)/IIf((Nz(C.GrossCharges,0)+Nz(A.AdjustmentTotal,0))<=0,1,(Nz(C.GrossCharges,0)+Nz(A.AdjustmentTotal,0)))))) AS CollectionPercentage " & _
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
        "C.GrossAmount, AT_ChargeRemaining(C.ChargeID) AS RemainingDue, D.StatusCode " & _
        "FROM (tblOwnerCharges AS C INNER JOIN tblOwners AS O ON C.OwnerID=O.OwnerID) " & _
        "INNER JOIN tblFinancialDocuments AS D ON C.DocumentID=D.DocumentID " & _
        "WHERE C.ChargeStatus='ACTIVE' AND D.StatusCode='APPROVED' ORDER BY O.FullName, C.ChargeDate;"

    AT_ReplaceQuery db, "qryAT_OwnerStatement", _
        "SELECT C.OwnerID AS OwnerID, C.ChargeDate AS EntryDate, C.ChargeCode AS ReferenceText, " & _
        "'CHARGE' AS EntryType, 'استحقاق الرسوم' AS DetailText, C.GrossAmount AS DebitAmount, CCur(0) AS CreditAmount, C.ChargeID AS SourceID " & _
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

    AT_BuildReportQueries = True
    MsgBox "تم إنشاء/تحديث استعلامات التقارير الأساسية. هذه مصادر بيانات؛ لا تزال تخطيطات الطباعة وPDF بحاجة إلى بناء واختبار داخل Access.", vbInformation, "Archon Tech"
    Exit Function
EH:
    MsgBox "تعذر بناء استعلامات التقارير. الخطأ " & Err.Number & ": " & Err.Description & vbCrLf & _
        "قد تكون بعض الاستعلامات قد أُنشئت؛ راجعها في جزء التنقل قبل إعادة المحاولة.", vbExclamation, "Archon Tech"
    AT_BuildReportQueries = False
End Function

Private Sub AT_ReplaceQuery(ByVal db As DAO.Database, ByVal queryName As String, ByVal sqlText As String)
    On Error Resume Next
    db.QueryDefs.Delete queryName
    Err.Clear
    On Error GoTo 0
    db.CreateQueryDef queryName, sqlText
End Sub
