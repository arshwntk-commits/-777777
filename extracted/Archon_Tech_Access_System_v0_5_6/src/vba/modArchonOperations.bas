Attribute VB_Name = "modArchonOperations"
Option Compare Database
Option Explicit

' Expense entry is kept separate from the payment voucher. Only payment approval posts cash movement.
Public Function AT_CreateExpenseDraft(ByVal expenseDate As Date, ByVal categoryID As Long, ByVal beneficiary As String, _
                                      ByVal descriptionText As String, ByVal amount As Currency) As Long
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "EXPENSE_DRAFT"
    If expenseDate = 0 Then expenseDate = Date
    AT_EnsureDateOpen expenseDate
    If amount <= 0 Or amount <> Fix(amount) Then Err.Raise vbObjectError + 1200, , AT_U("064A,062C,0628,0020,0623,0646,0020,064A,0643,0648,0646,0020,0645,0628,0644,063A,0020,0627,0644,0645,0635,0631,0648,0641,0020,0635,062D,064A,062D,0627,064B,0020,0648,0628,0631,064A,0627,0644,0020,0645,0648,062C,0628,002E")
    If Len(Trim$(beneficiary)) = 0 Or Len(Trim$(descriptionText)) = 0 Then Err.Raise vbObjectError + 1201, , AT_U("0627,0633,0645,0020,0627,0644,0645,0633,062A,0641,064A,062F,0020,0648,0648,0635,0641,0020,0627,0644,0645,0635,0631,0648,0641,0020,0645,0637,0644,0648,0628,0627,0646,002E")
    If DCount("*", "tblExpenseCategories", "ExpenseCategoryID=" & categoryID & " AND IsActive=True") = 0 Then Err.Raise vbObjectError + 1202, , AT_U("062A,0635,0646,064A,0641,0020,0627,0644,0645,0635,0631,0648,0641,0020,063A,064A,0631,0020,0635,0627,0644,062D,002E")
    Dim db As DAO.Database, rs As DAO.Recordset, expenseID As Long
    Set db = CurrentDb: DBEngine(0)(0).BeginTrans
    Set rs = db.OpenRecordset("tblExpenses", dbOpenDynaset)
    rs.AddNew: rs!ExpenseDate = expenseDate: rs!CategoryID = categoryID: rs!BeneficiaryName = Trim$(beneficiary)
    rs!Description = descriptionText: rs!Amount = amount: rs!StatusCode = "DRAFT": rs!CreatedAt = Now()
    rs!CreatedByUserName = Environ$("Username"): rs!UpdatedAt = Now(): rs.Update: rs.Bookmark = rs.LastModified
    expenseID = CLng(rs!ExpenseID): rs.Close
    db.Execute "UPDATE tblExpenses SET ExpenseNo='EX-' & Format([ExpenseID],'000000') WHERE ExpenseID=" & expenseID, dbFailOnError
    DBEngine(0)(0).CommitTrans
    AT_Audit "EXPENSE_DRAFT", "tblExpenses", expenseID, "SUCCESS", "No payment has been made or posted."
    AT_CreateExpenseDraft = expenseID: Exit Function
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "EXPENSE_DRAFT", "tblExpenses", 0, "FAILED", eDesc
    Err.Raise savedErrNumber, "AT_CreateExpenseDraft", eDesc
End Function

Public Function AT_CreateExpensePaymentDraft(ByVal expenseID As Long, ByVal accountID As Long, ByVal paymentMethodID As Long, _
                                             ByVal recipientName As String, Optional ByVal notes As String = "") As Long
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "PAYMENT_DRAFT"
    Dim db As DAO.Database, rs As DAO.Recordset, docID As Long, paymentID As Long
    Dim amount As Currency, expenseDate As Date, categoryID As Long, beneficiary As String, descriptionText As String
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tblExpenses WHERE ExpenseID=" & expenseID, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 1210, , AT_U("0627,0644,0645,0635,0631,0648,0641,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,002E")
    If rs!StatusCode <> "DRAFT" Then rs.Close: Err.Raise vbObjectError + 1211, , AT_U("0647,0630,0627,0020,0627,0644,0645,0635,0631,0648,0641,0020,0644,064A,0633,0020,0645,0633,0648,062F,0629,0020,0645,062A,0627,062D,0629,0020,0644,0644,0633,062F,0627,062F,002E")
    amount = CCur(rs!Amount): expenseDate = CDate(rs!ExpenseDate): categoryID = CLng(rs!CategoryID)
    beneficiary = CStr(rs!BeneficiaryName): descriptionText = CStr(rs!Description): rs.Close
    AT_EnsureDateOpen expenseDate
    If DCount("*", "tblPayments", "ExpenseID=" & expenseID) > 0 Then Err.Raise vbObjectError + 1212, , AT_U("064A,0648,062C,062F,0020,0633,0646,062F,0020,0635,0631,0641,0020,0645,0631,062A,0628,0637,0020,0628,0647,0630,0627,0020,0627,0644,0645,0635,0631,0648,0641,0020,0645,0633,0628,0642,064B,0627,061B,0020,0644,0627,0020,062A,0633,0645,062D,0020,0627,0644,0646,0633,062E,0629,0020,0627,0644,062D,0627,0644,064A,0629,0020,0628,062A,062C,0632,0626,0629,0020,0645,0635,0631,0648,0641,0020,0648,0627,062D,062F,0020,0625,0644,0649,0020,0639,062F,0629,0020,0633,0646,062F,0627,062A,002E")
    If DCount("*", "tblAccounts", "AccountID=" & accountID & " AND IsActive=True") = 0 Then Err.Raise vbObjectError + 1213, , AT_U("062D,0633,0627,0628,0020,0627,0644,0635,0631,0641,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0020,0623,0648,0020,063A,064A,0631,0020,0646,0634,0637,002E")
    If DCount("*", "tblPaymentMethods", "PaymentMethodID=" & paymentMethodID & " AND IsActive=True") = 0 Then Err.Raise vbObjectError + 1214, , AT_U("0637,0631,064A,0642,0629,0020,0627,0644,062F,0641,0639,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0629,0020,0623,0648,0020,063A,064A,0631,0020,0646,0634,0637,0629,002E")
    DBEngine(0)(0).BeginTrans
    docID = AT_CreateDocument("PAYMENT", expenseDate, descriptionText)
    Set rs = db.OpenRecordset("tblPayments", dbOpenDynaset)
    rs.AddNew: rs!DocumentID = docID: rs!ExpenseID = expenseID: rs!AccountID = accountID
    rs!ExpenseCategoryID = categoryID: rs!BeneficiaryName = beneficiary: rs!Amount = amount
    rs!PaymentMethodID = paymentMethodID: rs!RecipientName = recipientName: rs!Notes = notes
    rs.Update: rs.Bookmark = rs.LastModified: paymentID = CLng(rs!PaymentID): rs.Close
    db.Execute "UPDATE tblExpenses SET StatusCode='PAYMENT_DRAFT', UpdatedAt=Now() WHERE ExpenseID=" & expenseID & " AND StatusCode='DRAFT'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 1215, , AT_U("062A,063A,064A,0631,062A,0020,062D,0627,0644,0629,0020,0627,0644,0645,0635,0631,0648,0641,0020,0642,0628,0644,0020,0625,0646,0634,0627,0621,0020,0633,0646,062F,0020,0627,0644,0635,0631,0641,002E")
    DBEngine(0)(0).CommitTrans
    AT_Audit "EXPENSE_PAYMENT_DRAFT", "tblPayments", paymentID, "SUCCESS", "Linked expense=" & expenseID & "; no ledger effect."
    AT_CreateExpensePaymentDraft = paymentID: Exit Function
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "EXPENSE_PAYMENT_DRAFT", "tblPayments", 0, "FAILED", eDesc
    Err.Raise savedErrNumber, "AT_CreateExpensePaymentDraft", eDesc
End Function

Public Sub AT_CancelExpenseDraft(ByVal expenseID As Long, ByVal reason As String)
    On Error GoTo EH
    AT_RequirePermission "DOCUMENT_CANCEL"
    If Len(Trim$(reason)) = 0 Then Err.Raise vbObjectError + 1220, , AT_U("0633,0628,0628,0020,0627,0644,0625,0644,063A,0627,0621,0020,0625,0644,0632,0627,0645,064A,002E")
    If DCount("*", "tblPayments", "ExpenseID=" & expenseID) > 0 Then Err.Raise vbObjectError + 1221, , AT_U("0644,0627,0020,064A,0645,0643,0646,0020,0625,0644,063A,0627,0621,0020,0627,0644,0645,0635,0631,0648,0641,0020,0628,0639,062F,0020,0625,0646,0634,0627,0621,0020,0633,0646,062F,0020,0635,0631,0641,0020,0645,0631,062A,0628,0637,0020,0628,0647,061B,0020,0623,0644,063A,0650,0020,0645,0633,0648,062F,0629,0020,0627,0644,0633,0646,062F,0020,0623,0648,0644,064B,0627,0020,0645,0646,0020,0645,0633,0627,0631,0020,0627,0644,0645,0633,062A,0646,062F,0627,062A,002E")
    CurrentDb.Execute "UPDATE tblExpenses SET StatusCode='CANCELLED', UpdatedAt=Now() WHERE ExpenseID=" & expenseID & " AND StatusCode='DRAFT'", dbFailOnError
    If CurrentDb.RecordsAffected <> 1 Then Err.Raise vbObjectError + 1222, , AT_U("0644,0627,0020,064A,0645,0643,0646,0020,0625,0644,063A,0627,0621,0020,0645,0635,0631,0648,0641,0020,063A,064A,0631,0020,0645,0633,0648,062F,0629,002E")
    AT_Audit "EXPENSE_CANCEL", "tblExpenses", expenseID, "SUCCESS", "Expense draft cancelled.", reason
    Exit Sub
EH:
    AT_Audit "EXPENSE_CANCEL", "tblExpenses", expenseID, "FAILED", Err.Description, reason
    Err.Raise Err.Number, "AT_CancelExpenseDraft", Err.Description
End Sub
