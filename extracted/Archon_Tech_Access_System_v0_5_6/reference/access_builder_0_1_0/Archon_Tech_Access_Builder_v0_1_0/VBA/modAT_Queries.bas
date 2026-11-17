Attribute VB_Name = "modAT_Queries"
Option Compare Database
Option Explicit

Public Sub AT_BuildQueries()
    On Error GoTo BuildFail
    AT_LogBuild "Queries", "Started", "Building saved queries"

    AT_EnsureQuery "qAT_ChargeTotals", _
        "SELECT OwnerID, Sum(CurrentAmount) AS TotalCharges FROM tblAT_Charges WHERE Status<>'Cancelled' GROUP BY OwnerID;"

    AT_EnsureQuery "qAT_AllocatedTotals", _
        "SELECT C.OwnerID, Sum(A.AmountSigned) AS TotalAllocated FROM tblAT_ReceiptAllocations AS A INNER JOIN tblAT_Charges AS C ON A.ChargeID=C.ChargeID GROUP BY C.OwnerID;"

    AT_EnsureQuery "qAT_CreditTotals", _
        "SELECT OwnerID, Sum(AmountSigned) AS CreditBalance FROM tblAT_OwnerCreditMovements GROUP BY OwnerID;"

    AT_EnsureQuery "qAT_OwnerBalances", _
        "SELECT O.OwnerID, O.OwnerCode, O.FullName, Nz(Q.TotalCharges,0) AS TotalCharges, Nz(P.TotalAllocated,0) AS TotalAllocated, " & _
        "Nz(Q.TotalCharges,0)-Nz(P.TotalAllocated,0) AS RemainingDue, Nz(C.CreditBalance,0) AS CreditBalance, " & _
        "IIf(Nz(Q.TotalCharges,0)<=0,Null,IIf(Nz(P.TotalAllocated,0)/Nz(Q.TotalCharges,0)>1,1,Nz(P.TotalAllocated,0)/Nz(Q.TotalCharges,0))) AS CollectionRate " & _
        "FROM ((tblAT_Owners AS O LEFT JOIN qAT_ChargeTotals AS Q ON O.OwnerID=Q.OwnerID) " & _
        "LEFT JOIN qAT_AllocatedTotals AS P ON O.OwnerID=P.OwnerID) LEFT JOIN qAT_CreditTotals AS C ON O.OwnerID=C.OwnerID;"

    AT_EnsureQuery "qAT_Arrears", _
        "SELECT OwnerCode, FullName, Phone, TotalCharges, TotalAllocated, RemainingDue FROM qAT_OwnerBalances WHERE RemainingDue>0 ORDER BY RemainingDue DESC, FullName;"

    AT_EnsureQuery "qAT_AccountBalances", _
        "SELECT A.AccountID, A.AccountCode, A.AccountName, A.AccountType, Nz(Sum(M.AmountSigned),0) AS BookBalance " & _
        "FROM tblAT_FinancialAccounts AS A LEFT JOIN tblAT_AccountMovements AS M ON A.AccountID=M.AccountID " & _
        "GROUP BY A.AccountID, A.AccountCode, A.AccountName, A.AccountType;"

    AT_EnsureQuery "qAT_PendingTransferReserves", _
        "SELECT FromAccountID AS AccountID, Sum(Amount+FeeAmount) AS ReservedAmount FROM tblAT_Transfers WHERE Status='Pending' GROUP BY FromAccountID;"

    AT_EnsureQuery "qAT_AccountAvailableBalances", _
        "SELECT B.AccountID, B.AccountCode, B.AccountName, B.AccountType, B.BookBalance, " & _
        "B.BookBalance-Nz(R.ReservedAmount,0) AS AvailableBalance " & _
        "FROM qAT_AccountBalances AS B LEFT JOIN qAT_PendingTransferReserves AS R ON B.AccountID=R.AccountID;"

    AT_EnsureQuery "qAT_ReceiptRegister", _
        "SELECT R.ReceiptID, R.ReceiptNo, R.ReceiptDate, O.OwnerCode, O.FullName, R.Amount, M.MethodName, A.AccountName, R.Status " & _
        "FROM ((tblAT_Receipts AS R INNER JOIN tblAT_Owners AS O ON R.OwnerID=O.OwnerID) " & _
        "INNER JOIN tblAT_PaymentMethods AS M ON R.PaymentMethodID=M.PaymentMethodID) " & _
        "INNER JOIN tblAT_FinancialAccounts AS A ON R.AccountID=A.AccountID ORDER BY R.ReceiptDate DESC, R.ReceiptID DESC;"

    AT_EnsureQuery "qAT_ExpenseRegister", _
        "SELECT E.ExpenseID, E.ExpenseNo, E.ExpenseDate, C.CategoryName, E.Beneficiary, E.Description, E.Amount, E.Status " & _
        "FROM tblAT_Expenses AS E LEFT JOIN tblAT_ExpenseCategories AS C ON E.CategoryID=C.CategoryID ORDER BY E.ExpenseDate DESC;"

    AT_EnsureQuery "qAT_TransferRegister", _
        "SELECT T.TransferID, T.TransferNo, T.TransferDate, F.AccountName AS FromAccount, D.AccountName AS ToAccount, T.Amount, T.FeeAmount, T.Status " & _
        "FROM (tblAT_Transfers AS T INNER JOIN tblAT_FinancialAccounts AS F ON T.FromAccountID=F.AccountID) " & _
        "INNER JOIN tblAT_FinancialAccounts AS D ON T.ToAccountID=D.AccountID ORDER BY T.TransferDate DESC;"

    AT_EnsureQuery "qAT_DiscountReview", _
        "SELECT D.DiscountRequestID, D.DiscountNo, O.OwnerCode, O.FullName, C.ChargeCode, D.RequestDate, D.RequestedAmount, D.Reason, D.Status " & _
        "FROM ((tblAT_DiscountRequests AS D INNER JOIN tblAT_Charges AS C ON D.ChargeID=C.ChargeID) " & _
        "INNER JOIN tblAT_Owners AS O ON C.OwnerID=O.OwnerID) ORDER BY D.RequestDate DESC;"

    AT_EnsureQuery "qAT_OwnerStatement", _
        "SELECT C.OwnerID, O.OwnerCode, O.FullName, C.ChargeDate AS TransactionDate, '" & AT_U("0627,0633,062A,062D,0642,0627,0642") & "' AS TransactionType, C.ChargeCode AS ReferenceNo, C.Description AS Details, C.CurrentAmount AS Amount " & _
        "FROM tblAT_Charges AS C INNER JOIN tblAT_Owners AS O ON C.OwnerID=O.OwnerID " & _
        "UNION ALL SELECT R.OwnerID, O.OwnerCode, O.FullName, R.ReceiptDate, '" & AT_U("0625,064A,0635,0627,0644") & "' AS TransactionType, Nz(R.ReceiptNo,'" & AT_U("0645,0633,0648,062F,0629") & "') AS ReferenceNo, R.Notes AS Details, -R.Amount AS Amount " & _
        "FROM tblAT_Receipts AS R INNER JOIN tblAT_Owners AS O ON R.OwnerID=O.OwnerID WHERE R.Status='Approved' ORDER BY TransactionDate;"

    AT_LogBuild "Queries", "Success", "Saved queries created/updated"
    Exit Sub
BuildFail:
    AT_LogBuild "Queries", "Failed", Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "AT_BuildQueries", Err.Description
End Sub
