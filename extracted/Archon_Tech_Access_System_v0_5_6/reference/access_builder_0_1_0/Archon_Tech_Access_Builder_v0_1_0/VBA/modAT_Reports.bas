Attribute VB_Name = "modAT_Reports"
Option Compare Database
Option Explicit

Private Const AT_AC_REPORT As Long = 3
Private Const AT_AC_LABEL As Long = 100
Private Const AT_AC_TEXTBOX As Long = 109
Private Const AT_AC_REPORT_HEADER As Long = 3
Private Const AT_AC_PAGE_HEADER As Long = 2
Private Const AT_AC_DETAIL As Long = 0
Private Const AT_AC_SAVE_YES As Long = 1
Private Const AT_AC_VIEW_PREVIEW As Long = 2

Public Sub AT_BuildReports()
    On Error GoTo BuildFail
    AT_CreateTabularReport "rptAT_Arrears", "qAT_Arrears", AT_U("062A,0642,0631,064A,0631") & " " & AT_U("0627,0644,0645,062A,0623,062E,0631,0627,062A"), Array( _
        Array("OwnerCode", AT_U("0631,0642,0645") & " " & AT_U("0627,0644,0645,0627,0644,0643"), 1300), Array("FullName", AT_U("0627,0633,0645") & " " & AT_U("0627,0644,0645,0627,0644,0643"), 2500), Array("Phone", AT_U("0627,0644,0647,0627,062A,0641"), 1700), _
        Array("TotalCharges", AT_U("0625,062C,0645,0627,0644,064A") & " " & AT_U("0627,0644,0627,0633,062A,062D,0642,0627,0642"), 1700), Array("TotalAllocated", AT_U("0627,0644,0645,0633,062F,062F"), 1500), Array("RemainingDue", AT_U("0627,0644,0645,062A,0628,0642,064A"), 1700))
    AT_CreateTabularReport "rptAT_AccountBalances", "qAT_AccountAvailableBalances", AT_U("062A,0642,0631,064A,0631") & " " & AT_U("0623,0631,0635,062F,0629") & " " & AT_U("0627,0644,062D,0633,0627,0628,0627,062A"), Array( _
        Array("AccountCode", AT_U("0631,0645,0632") & " " & AT_U("0627,0644,062D,0633,0627,0628"), 1700), Array("AccountName", AT_U("0627,0644,062D,0633,0627,0628"), 2600), Array("AccountType", AT_U("0627,0644,0646,0648,0639"), 1600), _
        Array("BookBalance", AT_U("0627,0644,0631,0635,064A,062F") & " " & AT_U("0627,0644,062F,0641,062A,0631,064A"), 1900), Array("AvailableBalance", AT_U("0627,0644,0645,062A,0627,062D"), 1900))
    AT_CreateTabularReport "rptAT_ReceiptRegister", "qAT_ReceiptRegister", AT_U("0633,062C,0644") & " " & AT_U("0625,064A,0635,0627,0644,0627,062A") & " " & AT_U("0627,0644,062A,062D,0635,064A,0644"), Array( _
        Array("ReceiptNo", AT_U("0631,0642,0645") & " " & AT_U("0627,0644,0625,064A,0635,0627,0644"), 1700), Array("ReceiptDate", AT_U("0627,0644,062A,0627,0631,064A,062E"), 1600), Array("OwnerCode", AT_U("0631,0642,0645") & " " & AT_U("0627,0644,0645,0627,0644,0643"), 1300), _
        Array("FullName", AT_U("0627,0633,0645") & " " & AT_U("0627,0644,0645,0627,0644,0643"), 2200), Array("Amount", AT_U("0627,0644,0645,0628,0644,063A"), 1700), Array("MethodName", AT_U("0637,0631,064A,0642,0629") & " " & AT_U("0627,0644,062F,0641,0639"), 1700), Array("Status", AT_U("0627,0644,062D,0627,0644,0629"), 1300))
    AT_CreateTabularReport "rptAT_ExpenseRegister", "qAT_ExpenseRegister", AT_U("0633,062C,0644") & " " & AT_U("0627,0644,0645,0635,0631,0648,0641,0627,062A"), Array( _
        Array("ExpenseNo", AT_U("0631,0642,0645") & " " & AT_U("0627,0644,0645,0635,0631,0648,0641"), 1500), Array("ExpenseDate", AT_U("0627,0644,062A,0627,0631,064A,062E"), 1500), Array("CategoryName", AT_U("0627,0644,062A,0635,0646,064A,0641"), 1700), _
        Array("Beneficiary", AT_U("0627,0644,0645,0633,062A,0641,064A,062F"), 2200), Array("Amount", AT_U("0627,0644,0645,0628,0644,063A"), 1700), Array("Status", AT_U("0627,0644,062D,0627,0644,0629"), 1300))
    AT_LogBuild "Reports", "Success", "Printable report templates created where missing"
    Exit Sub
BuildFail:
    AT_LogBuild "Reports", "Failed", Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "AT_BuildReports", Err.Description
End Sub

Private Sub AT_CreateTabularReport(ByVal reportName As String, ByVal queryName As String, ByVal titleText As String, ByVal columns As Variant)
    Dim rpt As Object, tmpName As String
    Dim i As Long, spec As Variant, x As Long
    Dim ctl As Object
    If AT_ReportExists(reportName) Then Exit Sub
    Set rpt = CreateReport
    tmpName = rpt.Name
    rpt.RecordSource = queryName
    rpt.Caption = titleText
    rpt.HasReportHeader = True
    rpt.HasPageHeader = True
    On Error Resume Next
    rpt.RightToLeft = True
    On Error GoTo BuildFail
    rpt.Width = 14500
    rpt.Section(AT_AC_REPORT_HEADER).Height = 750
    rpt.Section(AT_AC_PAGE_HEADER).Height = 500
    rpt.Section(AT_AC_DETAIL).Height = 420

    Set ctl = CreateReportControl(tmpName, AT_AC_LABEL, AT_AC_REPORT_HEADER, , titleText, 300, 100, 12500, 500)
    ctl.FontName = "Tahoma"
    ctl.FontSize = 18
    ctl.FontBold = True
    x = 250
    For i = LBound(columns) To UBound(columns)
        spec = columns(i)
        Set ctl = CreateReportControl(tmpName, AT_AC_LABEL, AT_AC_PAGE_HEADER, , CStr(spec(1)), x, 50, CLng(spec(2)), 350)
        ctl.FontName = "Tahoma"
        ctl.FontBold = True
        ctl.TextAlign = 3
        Set ctl = CreateReportControl(tmpName, AT_AC_TEXTBOX, AT_AC_DETAIL, , CStr(spec(0)), x, 20, CLng(spec(2)), 330)
        ctl.Name = "txt" & CStr(spec(0))
        ctl.FontName = "Tahoma"
        ctl.FontSize = 9
        ctl.TextAlign = 3
        x = x + CLng(spec(2)) + 80
    Next i
    DoCmd.Save AT_AC_REPORT, tmpName
    DoCmd.Close AT_AC_REPORT, tmpName, AT_AC_SAVE_YES
    DoCmd.Rename reportName, AT_AC_REPORT, tmpName
End Sub

Private Function AT_ReportExists(ByVal reportName As String) As Boolean
    On Error GoTo NoReport
    AT_ReportExists = CurrentProject.AllReports(reportName).Name = reportName
    Exit Function
NoReport:
    AT_ReportExists = False
End Function

Public Function AT_OpenReport(ByVal reportName As String) As Boolean
    On Error GoTo OpenFail
    DoCmd.OpenReport reportName, AT_AC_VIEW_PREVIEW
    AT_OpenReport = True
    Exit Function
OpenFail:
    MsgBox AT_U("062A,0639,0630,0631") & " " & AT_U("0641,062A,062D") & " " & AT_U("0627,0644,062A,0642,0631,064A,0631") & " " & reportName & ": " & Err.Description, vbExclamation, "Archon Tech"
End Function
