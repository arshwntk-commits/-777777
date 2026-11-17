Attribute VB_Name = "modArchonPrint"
Option Compare Database
Option Explicit

' Creates printable Access report objects from the saved report queries.
Private Const AT_AC_REPORT_HEADER As Long = 3
Private Const AT_AC_PAGE_HEADER As Long = 2
Private Const AT_AC_DETAIL As Long = 0
Private Const AT_AC_LABEL As Long = 100
Private Const AT_AC_TEXTBOX As Long = 109
Private Const AT_AC_SAVE_YES As Long = 1
Private Const AT_AC_VIEW_PREVIEW As Long = 2

Public Sub AT_BuildPrintableReports()
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_CreateReceiptReport
    AT_CreateTabularReport "rptAT_OwnerBalances", "qryAT_OwnerBalances", AT_U("0643,0634,0641,0020,0623,0631,0635,062F,0629,0020,0627,0644,0645,0644,0627,0643"), Array( _
        Array("OwnerCode", AT_U("0631,0642,0645,0020,0627,0644,0645,0627,0644,0643"), 1500), Array("FullName", AT_U("0627,0633,0645,0020,0627,0644,0645,0627,0644,0643"), 2800), Array("NetDue", AT_U("0635,0627,0641,064A,0020,0627,0644,0645,0633,062A,062D,0642"), 1800), Array("PaidAmount", AT_U("0627,0644,0645,062F,0641,0648,0639"), 1600), Array("RemainingDue", AT_U("0627,0644,0645,062A,0628,0642,064A"), 1600))
    AT_CreateTabularReport "rptAT_AccountBalances", "qryAT_AccountBalances", AT_U("0623,0631,0635,062F,0629,0020,0627,0644,062D,0633,0627,0628,0627,062A"), Array( _
        Array("AccountCode", AT_U("0631,0645,0632,0020,0627,0644,062D,0633,0627,0628"), 1800), Array("AccountName", AT_U("0627,0644,062D,0633,0627,0628"), 3000), Array("BookBalance", AT_U("0627,0644,0631,0635,064A,062F,0020,0627,0644,062F,0641,062A,0631,064A"), 2200), Array("ReservedAmount", AT_U("0627,0644,0645,062D,062C,0648,0632"), 1800), Array("AvailableBalance", AT_U("0627,0644,0645,062A,0627,062D"), 1800))
    AT_CreateTabularReport "rptAT_BuildingDirectory", "qryAT_BuildingCurrentOwners", AT_U("062F,0644,064A,0644,0020,0627,0644,0645,0628,0627,0646,064A"), Array( _
        Array("BuildingCode", AT_U("0631,0642,0645,0020,0627,0644,0645,0628,0646,0649"), 1800), Array("BuildingName", AT_U("0627,0644,0645,0628,0646,0649"), 3000), Array("OwnerCode", AT_U("0631,0642,0645,0020,0627,0644,0645,0627,0644,0643"), 1700), Array("FullName", AT_U("0627,0644,0645,0627,0644,0643"), 3000))
    AT_CreateTabularReport "rptAT_UnitDirectory", "qryAT_UnitDirectory", AT_U("062F,0644,064A,0644,0020,0627,0644,0634,0642,0642"), Array( _
        Array("BuildingCode", AT_U("0631,0642,0645,0020,0627,0644,0645,0628,0646,0649"), 1800), Array("UnitCode", AT_U("0631,0642,0645,0020,0627,0644,0634,0642,0629"), 1800), Array("FloorNo", AT_U("0627,0644,0637,0627,0628,0642"), 1400), Array("FullName", AT_U("0627,0644,0645,0627,0644,0643"), 3000))
    AT_CreateTabularReport "rptAT_ExpenseRegister", "qryAT_ExpenseRegister", AT_U("0633,062C,0644,0020,0627,0644,0645,0635,0631,0648,0641,0627,062A"), Array( _
        Array("ExpenseNo", AT_U("0631,0642,0645,0020,0627,0644,0645,0635,0631,0648,0641"), 1700), Array("ExpenseDate", AT_U("0627,0644,062A,0627,0631,064A,062E"), 1600), Array("BeneficiaryName", AT_U("0627,0644,0645,0633,062A,0641,064A,062F"), 2300), Array("CategoryName", AT_U("0627,0644,062A,0635,0646,064A,0641"), 2000), Array("Amount", AT_U("0627,0644,0645,0628,0644,063A"), 1600), Array("StatusCode", AT_U("0627,0644,062D,0627,0644,0629"), 1400))
    MsgBox AT_U("062A,0645,0020,0625,0646,0634,0627,0621,0020,0642,0648,0627,0644,0628,0020,0627,0644,062A,0642,0627,0631,064A,0631,0020,0627,0644,0642,0627,0628,0644,0629,0020,0644,0644,0645,0639,0627,064A,0646,0629,0020,0648,0627,0644,0637,0628,0627,0639,0629,002E,0020,064A,062C,0628,0020,0641,062D,0635,0020,062A,0646,0633,064A,0642,0020,0648,0627,062A,062C,0627,0647,0020,0627,0644,0637,0628,0627,0639,0629,0020,0645,0646,0020,062F,0627,062E,0644,0020,0041,0063,0063,0065,0073,0073,002E"), vbInformation, "Archon Tech"
    Exit Sub
EH:

    savedErrNumber = Err.Number: eDesc = Err.Description
    MsgBox AT_U("062A,0639,0630,0631,0020,0625,0646,0634,0627,0621,0020,0627,0644,062A,0642,0631,064A,0631,002E") & " " & savedErrNumber & " - " & eDesc, vbExclamation, "Archon Tech"
    Err.Raise savedErrNumber, "AT_BuildPrintableReports", eDesc
End Sub

Private Sub AT_CreateReceiptReport()
    Dim rpt As Report, tempName As String, ctl As Control, x As Long
    If AT_ReportExists("rptAT_Receipt") Then DoCmd.DeleteObject acReport, "rptAT_Receipt"
    Set rpt = CreateReport: tempName = rpt.Name
    rpt.RecordSource = "qryAT_ReceiptPrint": rpt.Caption = AT_U("0625,064A,0635,0627,0644,0020,0642,0628,0636,0020,0645,0633,0627,0647,0645,0629,0020,0645,0634,0631,0648,0639,0020,0627,0644,0645,062C,0627,0631,064A")
    rpt.HasReportHeader = True: rpt.HasPageHeader = False
    On Error Resume Next: rpt.RightToLeft = True: On Error GoTo 0
    rpt.Width = 10500: rpt.Section(AT_AC_REPORT_HEADER).Height = 7800: rpt.Section(AT_AC_DETAIL).Height = 300
    Set ctl = CreateReportControl(tempName, AT_AC_LABEL, AT_AC_REPORT_HEADER, , AT_U("0625,064A,0635,0627,0644,0020,0642,0628,0636,0020,0645,0633,0627,0647,0645,0629,0020,0645,0634,0631,0648,0639,0020,0627,0644,0645,062C,0627,0631,064A"), 450, 300, 9500, 600)
    ctl.FontName = "Tahoma": ctl.FontSize = 18: ctl.FontBold = True: ctl.TextAlign = 3
    x = 450
    AT_ReceiptField tempName, "OfficialNumber", AT_U("0631,0642,0645,0020,0627,0644,0625,064A,0635,0627,0644"), x, 1150, 500, 550: x = x + 4600
    AT_ReceiptField tempName, "TransactionDate", AT_U("062A,0627,0631,064A,062E,0020,0627,0644,0642,0628,0636"), x, 1150, 500, 550
    x = 450: AT_ReceiptField tempName, "OwnerCode", AT_U("0631,0642,0645,0020,0627,0644,0645,0627,0644,0643"), x, 1150, 1400, 550: x = x + 4600
    AT_ReceiptField tempName, "OwnerName", AT_U("0627,0633,0645,0020,0627,0644,0645,0627,0644,0643"), x, 1150, 1400, 550
    x = 450: AT_ReceiptField tempName, "NeighborhoodName", AT_U("0627,0644,062D,0627,0631,0629"), x, 1150, 2300, 550: x = x + 4600
    AT_ReceiptField tempName, "Amount", AT_U("0627,0644,0645,0628,0644,063A,0020,0628,0627,0644,0631,064A,0627,0644,0020,0627,0644,064A,0645,0646,064A"), x, 1150, 2300, 550
    x = 450: AT_ReceiptField tempName, "PaymentMethod", AT_U("0637,0631,064A,0642,0629,0020,0627,0644,062F,0641,0639"), x, 1150, 3200, 550: x = x + 4600
    AT_ReceiptField tempName, "AccountName", AT_U("0627,0644,062D,0633,0627,0628"), x, 1150, 3200, 550
    x = 450: AT_ReceiptField tempName, "PayerName", AT_U("0627,0633,0645,0020,0627,0644,062F,0627,0641,0639"), x, 1150, 4100, 550: x = x + 4600
    AT_ReceiptField tempName, "CollectorName", AT_U("0627,0633,0645,0020,0627,0644,0645,062D,0635,0644"), x, 1150, 4100, 550
    Set ctl = CreateReportControl(tempName, AT_AC_LABEL, AT_AC_REPORT_HEADER, , AT_U("0634,0643,0631,064B,0627,0020,0644,0645,0633,0627,0647,0645,062A,0643,0645,002E,0020,064A,062D,062A,0641,0638,0020,0647,0630,0627,0020,0627,0644,0625,064A,0635,0627,0644,0020,0628,0627,0644,0628,064A,0627,0646,0627,062A,0020,0627,0644,062A,0627,0631,064A,062E,064A,0629,0020,0648,0642,062A,0020,0627,0644,0627,0639,062A,0645,0627,062F,002E"), 500, 5900, 9500, 500)
    ctl.FontName = "Tahoma": ctl.FontSize = 10: ctl.TextAlign = 3
    DoCmd.Save acReport, tempName: DoCmd.Close acReport, tempName, acSaveYes: DoCmd.Rename "rptAT_Receipt", acReport, tempName
End Sub

Private Sub AT_ReceiptField(ByVal reportName As String, ByVal fieldName As String, ByVal labelText As String, ByVal leftValue As Long, ByVal labelTop As Long, ByVal controlTop As Long, ByVal widthValue As Long)
    Dim ctl As Control
    Set ctl = CreateReportControl(reportName, AT_AC_LABEL, AT_AC_REPORT_HEADER, , labelText, leftValue, labelTop, widthValue, 300)
    ctl.FontName = "Tahoma": ctl.FontBold = True: ctl.TextAlign = 3
    Set ctl = CreateReportControl(reportName, AT_AC_TEXTBOX, AT_AC_REPORT_HEADER, , , leftValue, controlTop, widthValue, 400)
    ctl.Name = "txt" & fieldName: ctl.ControlSource = fieldName: ctl.FontName = "Tahoma": ctl.FontSize = 11: ctl.TextAlign = 3
End Sub

Private Sub AT_CreateTabularReport(ByVal reportName As String, ByVal queryName As String, ByVal titleText As String, ByVal columns As Variant)
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    Dim rpt As Report, tempName As String, ctl As Control, i As Long, spec As Variant, x As Long
    If AT_ReportExists(reportName) Then Exit Sub
    Set rpt = CreateReport: tempName = rpt.Name
    rpt.RecordSource = queryName: rpt.Caption = titleText: rpt.HasReportHeader = True: rpt.HasPageHeader = True
    On Error Resume Next: rpt.RightToLeft = True: On Error GoTo EH
    rpt.Width = 15000: rpt.Section(AT_AC_REPORT_HEADER).Height = 700: rpt.Section(AT_AC_PAGE_HEADER).Height = 500: rpt.Section(AT_AC_DETAIL).Height = 400
    Set ctl = CreateReportControl(tempName, AT_AC_LABEL, AT_AC_REPORT_HEADER, , titleText, 300, 100, 14200, 500)
    ctl.FontName = "Tahoma": ctl.FontSize = 16: ctl.FontBold = True: ctl.TextAlign = 3
    x = 250
    For i = LBound(columns) To UBound(columns)
        spec = columns(i)
        Set ctl = CreateReportControl(tempName, AT_AC_LABEL, AT_AC_PAGE_HEADER, , CStr(spec(1)), x, 50, CLng(spec(2)), 350)
        ctl.FontName = "Tahoma": ctl.FontSize = 9: ctl.FontBold = True: ctl.TextAlign = 3
        Set ctl = CreateReportControl(tempName, AT_AC_TEXTBOX, AT_AC_DETAIL, , , x, 20, CLng(spec(2)), 330)
        ctl.Name = "txt" & CStr(spec(0)): ctl.ControlSource = CStr(spec(0)): ctl.FontName = "Tahoma": ctl.FontSize = 8: ctl.TextAlign = 3
        x = x + CLng(spec(2)) + 80
    Next i
    DoCmd.Save acReport, tempName: DoCmd.Close acReport, tempName, acSaveYes: DoCmd.Rename reportName, acReport, tempName
    Exit Sub
EH:

    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: If Not rpt Is Nothing Then DoCmd.Close acReport, tempName, acSaveNo
    Err.Raise savedErrNumber, "AT_CreateTabularReport", eDesc
End Sub

Private Function AT_ReportExists(ByVal reportName As String) As Boolean
    On Error Resume Next
    AT_ReportExists = (CurrentProject.AllReports(reportName).Name = reportName)
    Err.Clear
End Function

Public Sub AT_PrintReceipt(ByVal receiptID As Long)
    On Error GoTo EH
    Dim docID As Long
    docID = CLng(Nz(DLookup("DocumentID", "tblReceipts", "ReceiptID=" & receiptID), 0))
    If docID = 0 Or Nz(DLookup("StatusCode", "tblFinancialDocuments", "DocumentID=" & docID), "") <> "APPROVED" Then Err.Raise vbObjectError + 1400, , AT_U("0644,0627,0020,064A,0645,0643,0646,0020,0637,0628,0627,0639,0629,0020,0625,064A,0635,0627,0644,0020,063A,064A,0631,0020,0645,0639,062A,0645,062F,0020,0623,0648,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,002E")
    DoCmd.OpenReport "rptAT_Receipt", AT_AC_VIEW_PREVIEW, , "ReceiptID=" & receiptID
    Exit Sub
EH:
    MsgBox AT_U("062A,0639,0630,0631,062A,0020,0645,0639,0627,064A,0646,0629,0020,0627,0644,0625,064A,0635,0627,0644,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Sub
