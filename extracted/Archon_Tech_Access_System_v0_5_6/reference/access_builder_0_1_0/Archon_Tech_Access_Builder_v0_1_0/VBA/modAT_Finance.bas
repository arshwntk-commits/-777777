Attribute VB_Name = "modAT_Finance"
Option Compare Database
Option Explicit

' Posting services. Draft records never post to account balances.
' Approved financial effects are written to tblAT_AccountMovements.

Private Sub AT_AddAccountMovement(ByVal accountID As Long, ByVal movementDate As Date, ByVal sourceType As String, ByVal sourceID As Long, ByVal amountSigned As Currency, ByVal description As String)
    CurrentDb.Execute "INSERT INTO tblAT_AccountMovements (AccountID, MovementDate, SourceType, SourceID, AmountSigned, Description, CreatedAt, CreatedBy) VALUES (" & _
        accountID & "," & AT_SqlDate(movementDate) & "," & AT_SqlText(sourceType) & "," & sourceID & "," & Replace(CStr(amountSigned), ",", ".") & "," & _
        AT_SqlText(description) & "," & AT_SqlDate(Now()) & "," & AT_SqlText(AT_UserName()) & ")", AT_DB_FAIL_ON_ERROR
End Sub

Private Sub AT_AddCreditMovement(ByVal ownerID As Long, ByVal receiptID As Variant, ByVal voucherID As Variant, ByVal movementDate As Date, ByVal amountSigned As Currency, ByVal movementType As String, ByVal reason As String)
    Dim receiptSql As String, voucherSql As String
    If IsNull(receiptID) Or IsEmpty(receiptID) Then receiptSql = "Null" Else receiptSql = CStr(CLng(receiptID))
    If IsNull(voucherID) Or IsEmpty(voucherID) Then voucherSql = "Null" Else voucherSql = CStr(CLng(voucherID))
    CurrentDb.Execute "INSERT INTO tblAT_OwnerCreditMovements (OwnerID, ReceiptID, PaymentVoucherID, MovementDate, AmountSigned, MovementType, Reason, CreatedAt, CreatedBy) VALUES (" & _
        ownerID & "," & receiptSql & "," & voucherSql & "," & AT_SqlDate(movementDate) & "," & Replace(CStr(amountSigned), ",", ".") & "," & _
        AT_SqlText(movementType) & "," & AT_SqlText(reason) & "," & AT_SqlDate(Now()) & "," & AT_SqlText(AT_UserName()) & ")", AT_DB_FAIL_ON_ERROR
End Sub

Public Function AT_CreateInitialCharges() As Long
    Dim rs As Object
    Dim unitID As Long, ownerID As Long
    Dim amount As Currency, unitAmount As Currency
    Dim chargeCode As String
    Dim createdCount As Long
    On Error GoTo BuildFail

    amount = CCur(Val(AT_GetSetting("ContributionPerUnit", CStr(AT_DEFAULT_CONTRIBUTION))))
    If amount <= 0 Then Err.Raise vbObjectError + 2100, "AT_CreateInitialCharges", AT_U("0642,064A,0645,0629") & " " & AT_U("0627,0644,0645,0633,0627,0647,0645,0629") & " " & AT_U("0627,0644,0627,0641,062A,0631,0627,0636,064A,0629") & " " & AT_U("063A,064A,0631") & " " & AT_U("0635,062D,064A,062D,0629") & "."

    Set rs = CurrentDb.OpenRecordset("SELECT UnitID, OwnerID, ContributionAmount FROM tblAT_Units WHERE IsActive=True ORDER BY UnitID", AT_DB_OPEN_SNAPSHOT)
    Do While Not rs.EOF
        unitID = CLng(rs!UnitID)
        ownerID = CLng(rs!OwnerID)
        If DCount("*", "tblAT_Charges", "UnitID=" & unitID & " AND SourceType='INITIAL'") = 0 Then
            unitAmount = amount
            If Not IsNull(rs!ContributionAmount) Then
                If CCur(rs!ContributionAmount) > 0 Then unitAmount = CCur(rs!ContributionAmount)
            End If
            chargeCode = AT_NextCode("CH-", "tblAT_Charges", "ChargeCode", 6)
            CurrentDb.Execute "INSERT INTO tblAT_Charges (ChargeCode, OwnerID, UnitID, ChargeDate, DueDate, Description, OriginalAmount, CurrentAmount, Status, SourceType, CreatedAt, CreatedBy) VALUES (" & _
                AT_SqlText(chargeCode) & "," & ownerID & "," & unitID & "," & AT_SqlDate(Date()) & "," & AT_SqlDate(Date()) & "," & _
                AT_SqlText(AT_U("0645,0633,0627,0647,0645,0629") & " " & AT_U("0645,0634,0631,0648,0639") & " " & AT_U("0627,0644,0645,062C,0627,0631,064A")) & "," & Replace(CStr(unitAmount), ",", ".") & "," & Replace(CStr(unitAmount), ",", ".") & ",'Open','INITIAL'," & AT_SqlDate(Now()) & "," & AT_SqlText(AT_UserName()) & ")", AT_DB_FAIL_ON_ERROR
            createdCount = createdCount + 1
        End If
        rs.MoveNext
    Loop
    rs.Close
    AT_Audit "CreateInitialCharges", "tblAT_Charges", Null, "Success", "Created=" & createdCount
    AT_CreateInitialCharges = createdCount
    Exit Function
BuildFail:
    AT_LogBuild "CreateInitialCharges", "Failed", Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "AT_CreateInitialCharges", Err.Description
End Function

Public Function AT_ApproveReceipt(ByVal receiptID As Long) As Boolean
    Dim ws As Object, rs As Object
    Dim ownerID As Long, accountID As Long, methodID As Long
    Dim receiptAmount As Currency, dueRemaining As Currency, allocateAmount As Currency, creditAmount As Currency
    Dim receiptDate As Date, receiptNo As String, statusValue As String
    Dim chargeID As Long
    On Error GoTo ApprovalFail

    statusValue = Nz(DLookup("Status", "tblAT_Receipts", "ReceiptID=" & receiptID), "")
    If statusValue <> "Draft" Then Err.Raise vbObjectError + 2110, "AT_ApproveReceipt", AT_U("0644,0627") & " " & AT_U("064A,0645,0643,0646") & " " & AT_U("0627,0639,062A,0645,0627,062F") & " " & AT_U("0627,0644,0625,064A,0635,0627,0644") & ": " & AT_U("062D,0627,0644,062A,0647") & " " & AT_U("0644,064A,0633,062A") & " " & AT_U("0645,0633,0648,062F,0629") & "."
    ownerID = CLng(Nz(DLookup("OwnerID", "tblAT_Receipts", "ReceiptID=" & receiptID), 0))
    accountID = CLng(Nz(DLookup("AccountID", "tblAT_Receipts", "ReceiptID=" & receiptID), 0))
    methodID = CLng(Nz(DLookup("PaymentMethodID", "tblAT_Receipts", "ReceiptID=" & receiptID), 0))
    receiptAmount = CCur(Nz(DLookup("Amount", "tblAT_Receipts", "ReceiptID=" & receiptID), 0))
    receiptDate = Nz(DLookup("ReceiptDate", "tblAT_Receipts", "ReceiptID=" & receiptID), Date())

    If ownerID <= 0 Or DCount("*", "tblAT_Owners", "OwnerID=" & ownerID) = 0 Then Err.Raise vbObjectError + 2111, , AT_U("0627,062E,062A,0631") & " " & AT_U("0645,0627,0644,0643,064B,0627") & " " & AT_U("0635,062D,064A,062D,064B,0627") & "."
    If receiptAmount <= 0 Then Err.Raise vbObjectError + 2112, , AT_U("0645,0628,0644,063A") & " " & AT_U("0627,0644,0625,064A,0635,0627,0644") & " " & AT_U("064A,062C,0628") & " " & AT_U("0623,0646") & " " & AT_U("064A,0643,0648,0646") & " " & AT_U("0623,0643,0628,0631") & " " & AT_U("0645,0646") & " " & AT_U("0635,0641,0631") & "."
    If Not AT_IsAccountActive(accountID) Then Err.Raise vbObjectError + 2113, , AT_U("0627,0644,062D,0633,0627,0628") & " " & AT_U("0627,0644,0645,0627,0644,064A") & " " & AT_U("063A,064A,0631") & " " & AT_U("0646,0634,0637") & " " & AT_U("0623,0648") & " " & AT_U("063A,064A,0631") & " " & AT_U("0645,062D,062F,062F") & "."
    If Not AT_IsMethodActive(methodID) Then Err.Raise vbObjectError + 2114, , AT_U("0637,0631,064A,0642,0629") & " " & AT_U("0627,0644,062F,0641,0639") & " " & AT_U("063A,064A,0631") & " " & AT_U("0646,0634,0637,0629") & " " & AT_U("0623,0648") & " " & AT_U("063A,064A,0631") & " " & AT_U("0645,062D,062F,062F,0629") & "."

    dueRemaining = AT_OwnerDue(ownerID)
    If receiptAmount > dueRemaining Then
        If MsgBox(AT_U("0627,0644,0645,0628,0644,063A") & " " & AT_U("0623,0643,0628,0631") & " " & AT_U("0645,0646") & " " & AT_U("0627,0644,0645,0633,062A,062D,0642") & " " & AT_U("0627,0644,062D,0627,0644,064A") & " " & AT_U("0628,0645,0642,062F,0627,0631") & " " & FormatCurrency(receiptAmount - dueRemaining) & ". " & AT_U("0633,064A,064F,0633,062C,0644") & " " & AT_U("0627,0644,0641,0631,0642") & " " & AT_U("0631,0635,064A,062F,064B,0627") & " " & AT_U("062F,0627,0626,0646,064B,0627") & " " & AT_U("0644,0644,0645,0627,0644,0643") & " " & AT_U("0648,0644,0627") & " " & AT_U("064A,064F,0639,062F") & " " & AT_U("0625,064A,0631,0627,062F,064B,0627") & " " & AT_U("0625,0636,0627,0641,064A,064B,0627") & ". " & AT_U("0647,0644") & " " & AT_U("062A,0631,064A,062F") & " " & AT_U("0627,0644,0645,062A,0627,0628,0639,0629,061F"), vbQuestion + vbYesNo, AT_U("062A,0623,0643,064A,062F") & " " & AT_U("0645,0628,0644,063A") & " " & AT_U("0632,0627,0626,062F")) <> vbYes Then Exit Function
    ElseIf MsgBox(AT_U("0647,0644") & " " & AT_U("062A,0631,064A,062F") & " " & AT_U("0627,0639,062A,0645,0627,062F") & " " & AT_U("0627,0644,0625,064A,0635,0627,0644") & " " & AT_U("0628,0645,0628,0644,063A") & " " & FormatCurrency(receiptAmount) & AT_U("061F") & " " & AT_U("0628,0639,062F") & " " & AT_U("0627,0644,0627,0639,062A,0645,0627,062F") & " " & AT_U("0633,064A,0624,062B,0631") & " " & AT_U("0641,064A") & " " & AT_U("0631,0635,064A,062F") & " " & AT_U("0627,0644,062D,0633,0627,0628") & " " & AT_U("0627,0644,0645,0627,0644,064A") & ".", vbQuestion + vbYesNo, AT_U("062A,0623,0643,064A,062F") & " " & AT_U("0627,0644,0627,0639,062A,0645,0627,062F")) <> vbYes Then
        Exit Function
    End If

    receiptNo = AT_NextCode("RC-", "tblAT_Receipts", "ReceiptNo", 6)
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans

    CurrentDb.Execute "UPDATE tblAT_Receipts SET ReceiptNo=" & AT_SqlText(receiptNo) & ", Status='Approved', ApprovedAt=" & AT_SqlDate(Now()) & ", ApprovedBy=" & AT_SqlText(AT_UserName()) & " WHERE ReceiptID=" & receiptID & " AND Status='Draft'", AT_DB_FAIL_ON_ERROR
    AT_AddAccountMovement accountID, receiptDate, "RECEIPT", receiptID, receiptAmount, AT_U("062A,062D,0635,064A,0644") & " " & AT_U("0625,064A,0635,0627,0644") & " " & receiptNo

    dueRemaining = receiptAmount
    Set rs = CurrentDb.OpenRecordset("SELECT C.ChargeID, C.CurrentAmount-Nz(Sum(A.AmountSigned),0) AS ChargeRemaining " & _
        "FROM tblAT_Charges AS C LEFT JOIN tblAT_ReceiptAllocations AS A ON C.ChargeID=A.ChargeID " & _
        "WHERE C.OwnerID=" & ownerID & " AND C.Status<>'Cancelled' GROUP BY C.ChargeID, C.CurrentAmount, C.DueDate ORDER BY C.DueDate, C.ChargeID", AT_DB_OPEN_SNAPSHOT)
    Do While Not rs.EOF And dueRemaining > 0
        chargeID = CLng(rs!ChargeID)
        allocateAmount = CCur(Nz(rs!ChargeRemaining, 0))
        If allocateAmount > 0 Then
            If allocateAmount > dueRemaining Then allocateAmount = dueRemaining
            CurrentDb.Execute "INSERT INTO tblAT_ReceiptAllocations (ReceiptID, ChargeID, AmountSigned, AllocationType, CreatedAt, Notes) VALUES (" & _
                receiptID & "," & chargeID & "," & Replace(CStr(allocateAmount), ",", ".") & ",'RECEIPT'," & AT_SqlDate(Now()) & "," & AT_SqlText(AT_U("062A,0648,0632,064A,0639") & " " & AT_U("062A,0644,0642,0627,0626,064A") & " " & AT_U("0639,0644,0649") & " " & AT_U("0623,0642,062F,0645") & " " & AT_U("0627,0633,062A,062D,0642,0627,0642")) & ")", AT_DB_FAIL_ON_ERROR
            dueRemaining = dueRemaining - allocateAmount
        End If
        rs.MoveNext
    Loop
    rs.Close

    creditAmount = dueRemaining
    If creditAmount > 0 Then AT_AddCreditMovement ownerID, receiptID, Null, receiptDate, creditAmount, "ReceiptOverpayment", AT_U("0631,0635,064A,062F") & " " & AT_U("0632,0627,0626,062F") & " " & AT_U("0645,0646") & " " & AT_U("0627,0644,0625,064A,0635,0627,0644") & " " & receiptNo
    AT_Audit "ApproveReceipt", "tblAT_Receipts", receiptID, "Success", "ReceiptNo=" & receiptNo & "; Amount=" & CStr(receiptAmount)
    ws.CommitTrans
    AT_ApproveReceipt = True
    Exit Function
ApprovalFail:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Rollback
    AT_Audit "ApproveReceipt", "tblAT_Receipts", receiptID, "Failed", Err.Number & " - " & Err.Description
    MsgBox AT_U("062A,0639,0630,0631") & " " & AT_U("0627,0639,062A,0645,0627,062F") & " " & AT_U("0627,0644,0625,064A,0635,0627,0644") & ": " & Err.Description, vbCritical, "Archon Tech"
    AT_ApproveReceipt = False
End Function

Public Function AT_ApplyOwnerCredit(ByVal ownerID As Long) As Currency
    Dim ws As Object, rs As Object
    Dim creditRemaining As Currency, useAmount As Currency, chargeRemaining As Currency
    Dim chargeID As Long, applied As Currency
    On Error GoTo ApplyFail
    creditRemaining = AT_OwnerCredit(ownerID)
    If creditRemaining <= 0 Then Exit Function
    If AT_OwnerDue(ownerID) <= 0 Then Exit Function
    If MsgBox(AT_U("0633,064A,062A,0645") & " " & AT_U("0627,0633,062A,062E,062F,0627,0645") & " " & AT_U("0627,0644,0631,0635,064A,062F") & " " & AT_U("0627,0644,062F,0627,0626,0646") & " " & AT_U("0627,0644,0645,062A,0627,062D") & " " & AT_U("0644,062A,0633,0648,064A,0629") & " " & AT_U("0623,0642,062F,0645") & " " & AT_U("0627,0644,0627,0633,062A,062D,0642,0627,0642,0627,062A") & ". " & AT_U("0647,0644") & " " & AT_U("062A,0631,064A,062F") & " " & AT_U("0627,0644,0645,062A,0627,0628,0639,0629,061F"), vbQuestion + vbYesNo, AT_U("062A,0623,0643,064A,062F") & " " & AT_U("0627,0633,062A,062E,062F,0627,0645") & " " & AT_U("0627,0644,0631,0635,064A,062F")) <> vbYes Then Exit Function

    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    Set rs = CurrentDb.OpenRecordset("SELECT C.ChargeID, C.CurrentAmount-Nz(Sum(A.AmountSigned),0) AS ChargeRemaining " & _
        "FROM tblAT_Charges AS C LEFT JOIN tblAT_ReceiptAllocations AS A ON C.ChargeID=A.ChargeID " & _
        "WHERE C.OwnerID=" & ownerID & " AND C.Status<>'Cancelled' GROUP BY C.ChargeID, C.CurrentAmount, C.DueDate ORDER BY C.DueDate, C.ChargeID", AT_DB_OPEN_SNAPSHOT)
    Do While Not rs.EOF And creditRemaining > 0
        chargeID = CLng(rs!ChargeID)
        chargeRemaining = CCur(Nz(rs!ChargeRemaining, 0))
        If chargeRemaining > 0 Then
            useAmount = chargeRemaining
            If useAmount > creditRemaining Then useAmount = creditRemaining
            CurrentDb.Execute "INSERT INTO tblAT_ReceiptAllocations (ReceiptID, ChargeID, AmountSigned, AllocationType, CreatedAt, Notes) VALUES (Null," & _
                chargeID & "," & Replace(CStr(useAmount), ",", ".") & ",'OWNER_CREDIT'," & AT_SqlDate(Now()) & "," & AT_SqlText(AT_U("0627,0633,062A,062E,062F,0627,0645") & " " & AT_U("0631,0635,064A,062F") & " " & AT_U("0627,0644,0645,0627,0644,0643")) & ")", AT_DB_FAIL_ON_ERROR
            creditRemaining = creditRemaining - useAmount
            applied = applied + useAmount
        End If
        rs.MoveNext
    Loop
    rs.Close
    If applied > 0 Then AT_AddCreditMovement ownerID, Null, Null, Date(), -applied, "CreditApplied", AT_U("0627,0633,062A,062E,062F,0627,0645") & " " & AT_U("0627,0644,0631,0635,064A,062F") & " " & AT_U("0627,0644,062F,0627,0626,0646") & " " & AT_U("0644,062A,0633,0648,064A,0629") & " " & AT_U("0627,0644,0627,0633,062A,062D,0642,0627,0642,0627,062A")
    AT_Audit "ApplyOwnerCredit", "tblAT_Owners", ownerID, "Success", "Applied=" & CStr(applied)
    ws.CommitTrans
    AT_ApplyOwnerCredit = applied
    Exit Function
ApplyFail:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Rollback
    MsgBox AT_U("062A,0639,0630,0631") & " " & AT_U("0627,0633,062A,062E,062F,0627,0645") & " " & AT_U("0627,0644,0631,0635,064A,062F") & ": " & Err.Description, vbCritical, "Archon Tech"
End Function

Public Function AT_ApproveVoucher(ByVal voucherID As Long) As Boolean
    Dim ws As Object
    Dim accountID As Long, methodID As Long, ownerID As Long, expenseID As Long
    Dim amount As Currency, available As Currency, voucherType As String, statusValue As String
    Dim voucherDate As Date, voucherNo As String, payee As String
    On Error GoTo ApprovalFail

    statusValue = Nz(DLookup("Status", "tblAT_PaymentVouchers", "VoucherID=" & voucherID), "")
    If statusValue <> "Draft" Then Err.Raise vbObjectError + 2130, , AT_U("0627,0644,0633,0646,062F") & " " & AT_U("0644,064A,0633") & " " & AT_U("0641,064A") & " " & AT_U("062D,0627,0644,0629") & " " & AT_U("0645,0633,0648,062F,0629") & "."
    accountID = CLng(Nz(DLookup("AccountID", "tblAT_PaymentVouchers", "VoucherID=" & voucherID), 0))
    methodID = CLng(Nz(DLookup("PaymentMethodID", "tblAT_PaymentVouchers", "VoucherID=" & voucherID), 0))
    ownerID = CLng(Nz(DLookup("OwnerID", "tblAT_PaymentVouchers", "VoucherID=" & voucherID), 0))
    expenseID = CLng(Nz(DLookup("ExpenseID", "tblAT_PaymentVouchers", "VoucherID=" & voucherID), 0))
    amount = CCur(Nz(DLookup("Amount", "tblAT_PaymentVouchers", "VoucherID=" & voucherID), 0))
    voucherType = UCase$(Nz(DLookup("VoucherType", "tblAT_PaymentVouchers", "VoucherID=" & voucherID), "EXPENSE"))
    voucherDate = Nz(DLookup("VoucherDate", "tblAT_PaymentVouchers", "VoucherID=" & voucherID), Date())
    payee = Nz(DLookup("PayeeName", "tblAT_PaymentVouchers", "VoucherID=" & voucherID), "")

    If amount <= 0 Then Err.Raise vbObjectError + 2131, , AT_U("0645,0628,0644,063A") & " " & AT_U("0627,0644,0633,0646,062F") & " " & AT_U("064A,062C,0628") & " " & AT_U("0623,0646") & " " & AT_U("064A,0643,0648,0646") & " " & AT_U("0623,0643,0628,0631") & " " & AT_U("0645,0646") & " " & AT_U("0635,0641,0631") & "."
    If Not AT_IsAccountActive(accountID) Then Err.Raise vbObjectError + 2132, , AT_U("0627,0644,062D,0633,0627,0628") & " " & AT_U("0627,0644,0645,0627,0644,064A") & " " & AT_U("063A,064A,0631") & " " & AT_U("0646,0634,0637") & " " & AT_U("0623,0648") & " " & AT_U("063A,064A,0631") & " " & AT_U("0645,062D,062F,062F") & "."
    If Not AT_IsMethodActive(methodID) Then Err.Raise vbObjectError + 2133, , AT_U("0637,0631,064A,0642,0629") & " " & AT_U("0627,0644,062F,0641,0639") & " " & AT_U("063A,064A,0631") & " " & AT_U("0646,0634,0637,0629") & " " & AT_U("0623,0648") & " " & AT_U("063A,064A,0631") & " " & AT_U("0645,062D,062F,062F,0629") & "."
    If AT_AccountAvailableBalance(accountID) < amount Then Err.Raise vbObjectError + 2134, , AT_U("0627,0644,0631,0635,064A,062F") & " " & AT_U("0627,0644,0645,062A,0627,062D") & " " & AT_U("063A,064A,0631") & " " & AT_U("0643,0627,0641,064D") & ". " & AT_U("0644,0627") & " " & AT_U("064A,0645,0643,0646") & " " & AT_U("062A,062C,0627,0648,0632") & " " & AT_U("0647,0630,0627") & " " & AT_U("0627,0644,0634,0631,0637") & " " & AT_U("062D,062A,0649") & " " & AT_U("0644,0644,0645,0633,0624,0648,0644") & "."
    If voucherType = "OWNER_REFUND" Then
        If ownerID <= 0 Then Err.Raise vbObjectError + 2135, , AT_U("062D,062F,062F") & " " & AT_U("0627,0644,0645,0627,0644,0643") & " " & AT_U("0627,0644,0645,0631,062A,0628,0637") & " " & AT_U("0628,0631,062F") & " " & AT_U("0627,0644,0631,0635,064A,062F") & "."
        If AT_OwnerCredit(ownerID) < amount Then Err.Raise vbObjectError + 2136, , AT_U("0631,0635,064A,062F") & " " & AT_U("0627,0644,0645,0627,0644,0643") & " " & AT_U("0627,0644,062F,0627,0626,0646") & " " & AT_U("063A,064A,0631") & " " & AT_U("0643,0627,0641,064D") & " " & AT_U("0644,0639,0645,0644,064A,0629") & " " & AT_U("0627,0644,0631,062F") & "."
    ElseIf voucherType <> "EXPENSE" Then
        Err.Raise vbObjectError + 2137, , AT_U("0646,0648,0639") & " " & AT_U("0627,0644,0633,0646,062F") & " " & AT_U("063A,064A,0631") & " " & AT_U("0645,062F,0639,0648,0645") & " " & AT_U("0641,064A") & " " & AT_U("0647,0630,0647") & " " & AT_U("0627,0644,0646,0633,062E,0629") & "."
    End If
    If expenseID > 0 Then
        If Nz(DLookup("Status", "tblAT_Expenses", "ExpenseID=" & expenseID), "") = "Paid" Then Err.Raise vbObjectError + 2138, , AT_U("0627,0644,0645,0635,0631,0648,0641") & " " & AT_U("0645,0631,062A,0628,0637") & " " & AT_U("0628,0633,0646,062F") & " " & AT_U("0645,0639,062A,0645,062F") & " " & AT_U("0645,0633,0628,0642,064B,0627") & "."
        If CCur(Nz(DLookup("Amount", "tblAT_Expenses", "ExpenseID=" & expenseID), 0)) <> amount Then Err.Raise vbObjectError + 2139, , AT_U("0645,0628,0644,063A") & " " & AT_U("0627,0644,0633,0646,062F") & " " & AT_U("064A,062C,0628") & " " & AT_U("0623,0646") & " " & AT_U("064A,0637,0627,0628,0642") & " " & AT_U("0645,0628,0644,063A") & " " & AT_U("0627,0644,0645,0635,0631,0648,0641") & " " & AT_U("0627,0644,0645,0631,062A,0628,0637") & "."
    End If

    If MsgBox(AT_U("0647,0644") & " " & AT_U("062A,0624,0643,062F") & " " & AT_U("0635,0631,0641") & " " & AT_U("0645,0628,0644,063A") & " " & FormatCurrency(amount) & " " & AT_U("0625,0644,0649") & " " & payee & AT_U("061F"), vbQuestion + vbYesNo, AT_U("062A,0623,0643,064A,062F") & " " & AT_U("0627,0644,0635,0631,0641")) <> vbYes Then Exit Function
    voucherNo = AT_NextCode("PV-", "tblAT_PaymentVouchers", "VoucherNo", 6)
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    CurrentDb.Execute "UPDATE tblAT_PaymentVouchers SET VoucherNo=" & AT_SqlText(voucherNo) & ", Status='Approved', ApprovedAt=" & AT_SqlDate(Now()) & ", ApprovedBy=" & AT_SqlText(AT_UserName()) & " WHERE VoucherID=" & voucherID & " AND Status='Draft'", AT_DB_FAIL_ON_ERROR
    AT_AddAccountMovement accountID, voucherDate, "PAYMENT_VOUCHER", voucherID, -amount, AT_U("0633,0646,062F") & " " & AT_U("0635,0631,0641") & " " & voucherNo & " - " & payee
    If voucherType = "OWNER_REFUND" Then AT_AddCreditMovement ownerID, Null, voucherID, voucherDate, -amount, "CreditRefund", AT_U("0631,062F") & " " & AT_U("0631,0635,064A,062F") & " " & AT_U("062F,0627,0626,0646") & " " & AT_U("0628,0645,0648,062C,0628") & " " & AT_U("0633,0646,062F") & " " & voucherNo
    If expenseID > 0 Then CurrentDb.Execute "UPDATE tblAT_Expenses SET Status='Paid', ExpenseNo=IIf(ExpenseNo Is Null," & AT_SqlText(AT_NextCode("EX-", "tblAT_Expenses", "ExpenseNo", 6)) & ",ExpenseNo) WHERE ExpenseID=" & expenseID, AT_DB_FAIL_ON_ERROR
    AT_Audit "ApproveVoucher", "tblAT_PaymentVouchers", voucherID, "Success", "VoucherNo=" & voucherNo & "; Amount=" & CStr(amount)
    ws.CommitTrans
    AT_ApproveVoucher = True
    Exit Function
ApprovalFail:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Rollback
    AT_Audit "ApproveVoucher", "tblAT_PaymentVouchers", voucherID, "Failed", Err.Number & " - " & Err.Description
    MsgBox AT_U("062A,0639,0630,0631") & " " & AT_U("0627,0639,062A,0645,0627,062F") & " " & AT_U("0633,0646,062F") & " " & AT_U("0627,0644,0635,0631,0641") & ": " & Err.Description, vbCritical, "Archon Tech"
    AT_ApproveVoucher = False
End Function

Public Function AT_StartTransfer(ByVal transferID As Long) As Boolean
    Dim fromID As Long, toID As Long
    Dim amount As Currency, fee As Currency, available As Currency, transferNo As String
    Dim statusValue As String
    On Error GoTo StartFail
    statusValue = Nz(DLookup("Status", "tblAT_Transfers", "TransferID=" & transferID), "")
    If statusValue <> "Draft" Then Err.Raise vbObjectError + 2150, , AT_U("0627,0644,062A,062D,0648,064A,0644") & " " & AT_U("0644,064A,0633") & " " & AT_U("0645,0633,0648,062F,0629") & "."
    fromID = CLng(Nz(DLookup("FromAccountID", "tblAT_Transfers", "TransferID=" & transferID), 0))
    toID = CLng(Nz(DLookup("ToAccountID", "tblAT_Transfers", "TransferID=" & transferID), 0))
    amount = CCur(Nz(DLookup("Amount", "tblAT_Transfers", "TransferID=" & transferID), 0))
    fee = CCur(Nz(DLookup("FeeAmount", "tblAT_Transfers", "TransferID=" & transferID), 0))
    If fromID = toID Or fromID <= 0 Or toID <= 0 Then Err.Raise vbObjectError + 2151, , AT_U("064A,062C,0628") & " " & AT_U("0627,062E,062A,064A,0627,0631") & " " & AT_U("062D,0633,0627,0628,064A,0646") & " " & AT_U("0645,062E,062A,0644,0641,064A,0646") & "."
    If Not AT_IsAccountActive(fromID) Or Not AT_IsAccountActive(toID) Then Err.Raise vbObjectError + 2152, , AT_U("0623,062D,062F") & " " & AT_U("0627,0644,062D,0633,0627,0628,064A,0646") & " " & AT_U("063A,064A,0631") & " " & AT_U("0646,0634,0637") & "."
    If amount <= 0 Or fee < 0 Then Err.Raise vbObjectError + 2153, , AT_U("0645,0628,0644,063A") & " " & AT_U("0627,0644,062A,062D,0648,064A,0644") & " " & AT_U("0623,0648") & " " & AT_U("0627,0644,0631,0633,0648,0645") & " " & AT_U("063A,064A,0631") & " " & AT_U("0635,062D,064A,062D") & "."
    available = AT_AccountAvailableBalance(fromID)
    If available < amount + fee Then Err.Raise vbObjectError + 2154, , AT_U("0627,0644,0631,0635,064A,062F") & " " & AT_U("0627,0644,0645,062A,0627,062D") & " " & AT_U("0644,0627") & " " & AT_U("064A,0643,0641,064A") & " " & AT_U("0644,0644,062A,062D,0648,064A,0644") & " " & AT_U("0648,0627,0644,0631,0633,0648,0645") & "."
    transferNo = AT_NextCode("TR-", "tblAT_Transfers", "TransferNo", 6)
    CurrentDb.Execute "UPDATE tblAT_Transfers SET TransferNo=" & AT_SqlText(transferNo) & ", Status='Pending', StartedAt=" & AT_SqlDate(Now()) & " WHERE TransferID=" & transferID & " AND Status='Draft'", AT_DB_FAIL_ON_ERROR
    AT_Audit "StartTransfer", "tblAT_Transfers", transferID, "Success", "Reserved=" & CStr(amount + fee)
    AT_StartTransfer = True
    Exit Function
StartFail:
    MsgBox AT_U("062A,0639,0630,0631") & " " & AT_U("0628,062F,0621") & " " & AT_U("0627,0644,062A,062D,0648,064A,0644") & ": " & Err.Description, vbCritical, "Archon Tech"
End Function

Public Function AT_CompleteTransfer(ByVal transferID As Long) As Boolean
    Dim ws As Object
    Dim fromID As Long, toID As Long
    Dim amount As Currency, fee As Currency, available As Currency
    Dim transferNo As String, transferDate As Date
    On Error GoTo CompleteFail
    If Nz(DLookup("Status", "tblAT_Transfers", "TransferID=" & transferID), "") <> "Pending" Then Err.Raise vbObjectError + 2160, , AT_U("0627,0644,062A,062D,0648,064A,0644") & " " & AT_U("0644,064A,0633") & " " & AT_U("0645,0639,0644,0642,064B,0627") & "."
    fromID = CLng(DLookup("FromAccountID", "tblAT_Transfers", "TransferID=" & transferID))
    toID = CLng(DLookup("ToAccountID", "tblAT_Transfers", "TransferID=" & transferID))
    amount = CCur(DLookup("Amount", "tblAT_Transfers", "TransferID=" & transferID))
    fee = CCur(Nz(DLookup("FeeAmount", "tblAT_Transfers", "TransferID=" & transferID), 0))
    transferNo = Nz(DLookup("TransferNo", "tblAT_Transfers", "TransferID=" & transferID), "")
    transferDate = Nz(DLookup("TransferDate", "tblAT_Transfers", "TransferID=" & transferID), Date())
    available = AT_AccountAvailableBalance(fromID, transferID)
    If available < amount + fee Then Err.Raise vbObjectError + 2161, , AT_U("0627,0644,0631,0635,064A,062F") & " " & AT_U("0627,0644,0645,062A,0627,062D") & " " & AT_U("0628,0639,062F") & " " & AT_U("0627,0633,062A,0628,0639,0627,062F") & " " & AT_U("0647,0630,0627") & " " & AT_U("0627,0644,062D,062C,0632") & " " & AT_U("063A,064A,0631") & " " & AT_U("0643,0627,0641,064D") & "."
    If MsgBox(AT_U("0647,0644") & " " & AT_U("062A,0645") & " " & AT_U("062A,0646,0641,064A,0630") & " " & AT_U("0627,0644,062A,062D,0648,064A,0644") & " " & AT_U("0641,0639,0644,064A,064B,0627") & " " & AT_U("0648,0648,0635,0644") & " " & AT_U("0627,0644,0645,0628,0644,063A") & " " & AT_U("0625,0644,0649") & " " & AT_U("0627,0644,062D,0633,0627,0628") & " " & AT_U("0627,0644,0648,062C,0647,0629,061F") & " " & AT_U("0644,0646") & " " & AT_U("064A,064F,0636,0627,0641") & " " & AT_U("0644,0644,062D,0633,0627,0628") & " " & AT_U("0627,0644,0648,062C,0647,0629") & " " & AT_U("0642,0628,0644") & " " & AT_U("0647,0630,0627") & " " & AT_U("0627,0644,062A,0623,0643,064A,062F") & ".", vbQuestion + vbYesNo, AT_U("062A,0623,0643,064A,062F") & " " & AT_U("0648,0635,0648,0644") & " " & AT_U("0627,0644,062A,062D,0648,064A,0644")) <> vbYes Then Exit Function
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    CurrentDb.Execute "UPDATE tblAT_Transfers SET Status='Completed', CompletedAt=" & AT_SqlDate(Now()) & " WHERE TransferID=" & transferID & " AND Status='Pending'", AT_DB_FAIL_ON_ERROR
    AT_AddAccountMovement fromID, transferDate, "TRANSFER_OUT", transferID, -amount, AT_U("062A,062D,0648,064A,0644") & " " & AT_U("0635,0627,062F,0631") & " " & transferNo
    AT_AddAccountMovement toID, transferDate, "TRANSFER_IN", transferID, amount, AT_U("062A,062D,0648,064A,0644") & " " & AT_U("0648,0627,0631,062F") & " " & transferNo
    If fee > 0 Then AT_AddAccountMovement fromID, transferDate, "TRANSFER_FEE", transferID, -fee, AT_U("0631,0633,0648,0645") & " " & AT_U("062A,062D,0648,064A,0644") & " " & transferNo
    AT_Audit "CompleteTransfer", "tblAT_Transfers", transferID, "Success", "Completed=" & transferNo
    ws.CommitTrans
    AT_CompleteTransfer = True
    Exit Function
CompleteFail:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Rollback
    AT_Audit "CompleteTransfer", "tblAT_Transfers", transferID, "Failed", Err.Number & " - " & Err.Description
    MsgBox AT_U("062A,0639,0630,0631") & " " & AT_U("0625,0643,0645,0627,0644") & " " & AT_U("0627,0644,062A,062D,0648,064A,0644") & ": " & Err.Description, vbCritical, "Archon Tech"
End Function

Public Function AT_ApproveDiscount(ByVal requestID As Long) As Boolean
    Dim ws As Object
    Dim chargeID As Long
    Dim requested As Currency, beforeAmount As Currency, afterAmount As Currency, paidAmount As Currency
    On Error GoTo DiscountFail
    If Nz(DLookup("Status", "tblAT_DiscountRequests", "DiscountRequestID=" & requestID), "") <> "Pending" Then Err.Raise vbObjectError + 2170, , AT_U("0637,0644,0628") & " " & AT_U("0627,0644,062E,0635,0645") & " " & AT_U("0644,064A,0633") & " " & AT_U("0645,0639,0644,0642,064B,0627") & "."
    chargeID = CLng(DLookup("ChargeID", "tblAT_DiscountRequests", "DiscountRequestID=" & requestID))
    requested = CCur(DLookup("RequestedAmount", "tblAT_DiscountRequests", "DiscountRequestID=" & requestID))
    beforeAmount = CCur(DLookup("CurrentAmount", "tblAT_Charges", "ChargeID=" & chargeID))
    paidAmount = CCur(Nz(DSum("AmountSigned", "tblAT_ReceiptAllocations", "ChargeID=" & chargeID), 0))
    afterAmount = beforeAmount - requested
    If requested <= 0 Or afterAmount < 0 Then Err.Raise vbObjectError + 2171, , AT_U("0642,064A,0645,0629") & " " & AT_U("0627,0644,062E,0635,0645") & " " & AT_U("063A,064A,0631") & " " & AT_U("0635,062D,064A,062D,0629") & "."
    If afterAmount < paidAmount Then Err.Raise vbObjectError + 2172, , AT_U("0644,0627") & " " & AT_U("064A,0645,0643,0646") & " " & AT_U("062A,062E,0641,064A,0636") & " " & AT_U("0627,0644,0627,0633,062A,062D,0642,0627,0642") & " " & AT_U("0625,0644,0649") & " " & AT_U("0623,0642,0644") & " " & AT_U("0645,0646") & " " & AT_U("0627,0644,0645,0628,0644,063A") & " " & AT_U("0627,0644,0645,0633,062F,062F") & " " & AT_U("0639,0644,064A,0647") & "."
    If MsgBox(AT_U("0633,064A,062A,0645") & " " & AT_U("062A,062E,0641,064A,0636") & " " & AT_U("0642,064A,0645,0629") & " " & AT_U("0627,0644,0627,0633,062A,062D,0642,0627,0642") & " " & AT_U("0645,0646") & " " & FormatCurrency(beforeAmount) & " " & AT_U("0625,0644,0649") & " " & FormatCurrency(afterAmount) & ". " & AT_U("0647,0644") & " " & AT_U("062A,0648,0627,0641,0642,061F"), vbQuestion + vbYesNo, AT_U("0627,0639,062A,0645,0627,062F") & " " & AT_U("0627,0644,062E,0635,0645")) <> vbYes Then Exit Function
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    CurrentDb.Execute "UPDATE tblAT_Charges SET CurrentAmount=" & Replace(CStr(afterAmount), ",", ".") & " WHERE ChargeID=" & chargeID, AT_DB_FAIL_ON_ERROR
    CurrentDb.Execute "UPDATE tblAT_DiscountRequests SET Status='Approved', AmountBefore=" & Replace(CStr(beforeAmount), ",", ".") & ", AmountAfter=" & Replace(CStr(afterAmount), ",", ".") & ", ApprovedAt=" & AT_SqlDate(Now()) & ", ApprovedBy=" & AT_SqlText(AT_UserName()) & " WHERE DiscountRequestID=" & requestID, AT_DB_FAIL_ON_ERROR
    AT_Audit "ApproveDiscount", "tblAT_DiscountRequests", requestID, "Success", "Charge=" & chargeID, CStr(beforeAmount), CStr(afterAmount)
    ws.CommitTrans
    AT_ApproveDiscount = True
    Exit Function
DiscountFail:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Rollback
    MsgBox AT_U("062A,0639,0630,0631") & " " & AT_U("0627,0639,062A,0645,0627,062F") & " " & AT_U("0627,0644,062E,0635,0645") & ": " & Err.Description, vbCritical, "Archon Tech"
End Function

Public Function AT_ReverseReceipt(ByVal receiptID As Long, ByVal reason As String) As Boolean
    Dim ws As Object
    Dim ownerID As Long, accountID As Long
    Dim amount As Currency, linkedCredit As Currency
    Dim receiptNo As String
    On Error GoTo ReverseFail
    If Nz(DLookup("Status", "tblAT_Receipts", "ReceiptID=" & receiptID), "") <> "Approved" Then Err.Raise vbObjectError + 2180, , AT_U("064A,0645,0643,0646") & " " & AT_U("0639,0643,0633") & " " & AT_U("0627,0644,0625,064A,0635,0627,0644") & " " & AT_U("0627,0644,0645,0639,062A,0645,062F") & " " & AT_U("0641,0642,0637") & "."
    If Len(Trim$(reason)) = 0 Then Err.Raise vbObjectError + 2181, , AT_U("0633,0628,0628") & " " & AT_U("0627,0644,0639,0643,0633") & " " & AT_U("0645,0637,0644,0648,0628") & "."
    ownerID = CLng(DLookup("OwnerID", "tblAT_Receipts", "ReceiptID=" & receiptID))
    accountID = CLng(DLookup("AccountID", "tblAT_Receipts", "ReceiptID=" & receiptID))
    amount = CCur(DLookup("Amount", "tblAT_Receipts", "ReceiptID=" & receiptID))
    receiptNo = Nz(DLookup("ReceiptNo", "tblAT_Receipts", "ReceiptID=" & receiptID), "")
    linkedCredit = CCur(Nz(DSum("AmountSigned", "tblAT_OwnerCreditMovements", "ReceiptID=" & receiptID & " AND AmountSigned>0"), 0))
    If AT_OwnerCredit(ownerID) < linkedCredit Then Err.Raise vbObjectError + 2182, , AT_U("0644,0627") & " " & AT_U("064A,0645,0643,0646") & " " & AT_U("0639,0643,0633") & " " & AT_U("0627,0644,0625,064A,0635,0627,0644") & " " & AT_U("0644,0623,0646") & " " & AT_U("062C,0632,0621,064B,0627") & " " & AT_U("0645,0646") & " " & AT_U("0631,0635,064A,062F,0647") & " " & AT_U("0627,0644,062F,0627,0626,0646") & " " & AT_U("0627,0633,062A,064F,062E,062F,0645") & " " & AT_U("0628,0627,0644,0641,0639,0644") & ". " & AT_U("0639,0627,0644,062C") & " " & AT_U("0627,0633,062A,062E,062F,0627,0645") & " " & AT_U("0627,0644,0631,0635,064A,062F") & " " & AT_U("0623,0648,0644,064B,0627") & "."
    If MsgBox(AT_U("0633,064A,062A,0645") & " " & AT_U("0639,0643,0633") & " " & AT_U("0627,0644,0623,062B,0631") & " " & AT_U("0627,0644,0645,062D,0627,0633,0628,064A") & " " & AT_U("0644,0644,0625,064A,0635,0627,0644") & " " & receiptNo & AT_U("060C") & " " & AT_U("0648,0644,0646") & " " & AT_U("064A,0639,0646,064A") & " " & AT_U("0630,0644,0643") & " " & AT_U("0631,062F") & " " & AT_U("0627,0644,0646,0642,062F") & " " & AT_U("0641,0639,0644,064A,064B,0627") & ". " & AT_U("0647,0644") & " " & AT_U("062A,0631,064A,062F") & " " & AT_U("0627,0644,0645,062A,0627,0628,0639,0629,061F"), vbQuestion + vbYesNo, AT_U("0639,0643,0633") & " " & AT_U("0625,064A,0635,0627,0644")) <> vbYes Then Exit Function
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans
    CurrentDb.Execute "INSERT INTO tblAT_ReceiptAllocations (ReceiptID, ChargeID, AmountSigned, AllocationType, CreatedAt, Notes) " & _
        "SELECT ReceiptID, ChargeID, -AmountSigned, 'REVERSAL', " & AT_SqlDate(Now()) & ", " & AT_SqlText(reason) & " FROM tblAT_ReceiptAllocations WHERE ReceiptID=" & receiptID & " AND AmountSigned>0", AT_DB_FAIL_ON_ERROR
    If linkedCredit > 0 Then AT_AddCreditMovement ownerID, receiptID, Null, Date(), -linkedCredit, "ReceiptReversal", reason
    AT_AddAccountMovement accountID, Date(), "RECEIPT_REVERSAL", receiptID, -amount, AT_U("0639,0643,0633") & " " & AT_U("0627,0644,0625,064A,0635,0627,0644") & " " & receiptNo
    CurrentDb.Execute "UPDATE tblAT_Receipts SET Status='Reversed', ReversedAt=" & AT_SqlDate(Now()) & ", ReversalReason=" & AT_SqlText(reason) & " WHERE ReceiptID=" & receiptID, AT_DB_FAIL_ON_ERROR
    CurrentDb.Execute "INSERT INTO tblAT_Reversals (SourceType, SourceID, ReversalDate, Reason, ApprovedBy, CreatedAt) VALUES ('RECEIPT'," & receiptID & "," & AT_SqlDate(Now()) & "," & AT_SqlText(reason) & "," & AT_SqlText(AT_UserName()) & "," & AT_SqlDate(Now()) & ")", AT_DB_FAIL_ON_ERROR
    AT_Audit "ReverseReceipt", "tblAT_Receipts", receiptID, "Success", reason
    ws.CommitTrans
    AT_ReverseReceipt = True
    Exit Function
ReverseFail:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Rollback
    MsgBox AT_U("062A,0639,0630,0631") & " " & AT_U("0639,0643,0633") & " " & AT_U("0627,0644,0625,064A,0635,0627,0644") & ": " & Err.Description, vbCritical, "Archon Tech"
End Function
