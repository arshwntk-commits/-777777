Attribute VB_Name = "modAT_Tests"
Option Compare Database
Option Explicit

Public Function AT_RunSelfTests(Optional ByVal showSummary As Boolean = True) As Long
    Dim tables As Variant, queries As Variant, forms As Variant, reports As Variant
    Dim i As Long, passed As Long, failed As Long, missing As String

    tables = Array("tblAT_AppSettings", "tblAT_Users", "tblAT_Owners", "tblAT_Units", "tblAT_Charges", _
        "tblAT_PaymentMethods", "tblAT_FinancialAccounts", "tblAT_Receipts", "tblAT_ReceiptAllocations", _
        "tblAT_OwnerCreditMovements", "tblAT_ExpenseCategories", "tblAT_Expenses", "tblAT_PaymentVouchers", _
        "tblAT_Transfers", "tblAT_AccountMovements", "tblAT_DiscountRequests", "tblAT_AuditLog", _
        "tblAT_Reversals", "tblAT_PeriodClosings", "tblAT_ClosingAccountBalances", "tblAT_DocumentLinks", "tblAT_BuildLog")
    queries = Array("qAT_ChargeTotals", "qAT_AllocatedTotals", "qAT_CreditTotals", "qAT_OwnerBalances", _
        "qAT_Arrears", "qAT_AccountBalances", "qAT_PendingTransferReserves", "qAT_AccountAvailableBalances", _
        "qAT_ReceiptRegister", "qAT_ExpenseRegister", "qAT_TransferRegister", "qAT_DiscountReview", "qAT_OwnerStatement")
    forms = Array("frmAT_Home", "frmAT_Owners", "frmAT_Units", "frmAT_Charges", "frmAT_Receipts", _
        "frmAT_Expenses", "frmAT_Vouchers", "frmAT_Transfers", "frmAT_Accounts", "frmAT_Discounts", "frmAT_PaymentMethods", "frmAT_ExpenseCategories")
    reports = Array("rptAT_Arrears", "rptAT_AccountBalances", "rptAT_ReceiptRegister", "rptAT_ExpenseRegister")

    For i = LBound(tables) To UBound(tables)
        If AT_TableExists(CStr(tables(i))) Then
            passed = passed + 1: AT_LogBuild "Test:Table:" & tables(i), "PASS"
        Else
            failed = failed + 1: missing = missing & vbCrLf & AT_U("062C,062F,0648,0644") & " " & AT_U("0645,0641,0642,0648,062F") & ": " & tables(i): AT_LogBuild "Test:Table:" & tables(i), "FAIL", "Missing"
        End If
    Next i
    For i = LBound(queries) To UBound(queries)
        If AT_QueryExists(CStr(queries(i))) Then
            passed = passed + 1: AT_LogBuild "Test:Query:" & queries(i), "PASS"
        Else
            failed = failed + 1: missing = missing & vbCrLf & AT_U("0627,0633,062A,0639,0644,0627,0645") & " " & AT_U("0645,0641,0642,0648,062F") & ": " & queries(i): AT_LogBuild "Test:Query:" & queries(i), "FAIL", "Missing"
        End If
    Next i
    For i = LBound(forms) To UBound(forms)
        If AT_FormExists(CStr(forms(i))) Then
            passed = passed + 1: AT_LogBuild "Test:Form:" & forms(i), "PASS"
        Else
            failed = failed + 1: missing = missing & vbCrLf & AT_U("0646,0645,0648,0630,062C") & " " & AT_U("0645,0641,0642,0648,062F") & ": " & forms(i): AT_LogBuild "Test:Form:" & forms(i), "FAIL", "Missing"
        End If
    Next i
    For i = LBound(reports) To UBound(reports)
        If AT_ReportAvailable(CStr(reports(i))) Then
            passed = passed + 1: AT_LogBuild "Test:Report:" & reports(i), "PASS"
        Else
            failed = failed + 1: missing = missing & vbCrLf & AT_U("062A,0642,0631,064A,0631") & " " & AT_U("0645,0641,0642,0648,062F") & ": " & reports(i): AT_LogBuild "Test:Report:" & reports(i), "FAIL", "Missing"
        End If
    Next i

    AT_RunSelfTests = failed
    If showSummary Then
        If failed = 0 Then
            MsgBox AT_U("0627,062C,062A,0627,0632") & " " & AT_U("0627,0644,0641,062D,0635") & " " & AT_U("0627,0644,0628,0646,064A,0648,064A") & " " & passed & " " & AT_U("062A,062D,0642,0642,064B,0627") & ". " & AT_U("0647,0630,0627") & " " & AT_U("0627,0644,0641,062D,0635") & " " & AT_U("0644,0627") & " " & AT_U("064A,064F,0639,062F") & " " & AT_U("0627,062E,062A,0628,0627,0631,064B,0627") & " " & AT_U("062A,0634,063A,064A,0644,064A,064B,0627") & " " & AT_U("0644,0644,0639,0645,0644,064A,0627,062A") & " " & AT_U("0627,0644,0645,0627,0644,064A,0629") & ".", vbInformation, "Archon Tech - " & AT_U("0641,062D,0635") & " " & AT_U("0627,0644,0628,0646,0627,0621")
        Else
            MsgBox AT_U("0646,062C,062D") & " " & passed & " " & AT_U("0648,0641,0634,0644") & " " & failed & "." & vbCrLf & missing, vbExclamation, "Archon Tech - " & AT_U("0641,062D,0635") & " " & AT_U("0627,0644,0628,0646,0627,0621")
        End If
    End If
End Function

Private Function AT_ReportAvailable(ByVal reportName As String) As Boolean
    On Error GoTo NotFound
    AT_ReportAvailable = CurrentProject.AllReports(reportName).Name = reportName
    Exit Function
NotFound:
    AT_ReportAvailable = False
End Function
