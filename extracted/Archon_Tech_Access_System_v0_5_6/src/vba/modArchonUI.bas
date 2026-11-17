Attribute VB_Name = "modArchonUI"
Option Compare Database
Option Explicit

' Generates the first usable desktop navigation and transaction-entry forms.
' This is a form-builder module, not a substitute for a compiled/tested Access front end.

Public Sub AT_BuildInitialUI()
    Dim savedErrNumber As Long, eDesc As String, savedErrSource As String
    Dim formsOpenBeforeBuild As Collection, cleanupDetails As String
    On Error GoTo EH
    Set formsOpenBeforeBuild = AT_CaptureOpenForms()
    Call AT_BuildReportQueries
    AT_CreateMainForm
    AT_CreateOwnerEntryForm
    AT_CreateOwnerEditForm
    AT_CreateOwnerSearchForm
    AT_CreateOwnerStatementPicker
    AT_CreateAccountEntryForm
    AT_CreateChargeEntryForm
    AT_CreateReceiptEntryForm
    AT_CreatePaymentEntryForm
    AT_CreateTransferEntryForm
    AT_CreateAdjustmentEntryForm
    AT_CreateCreditApplicationForm
    AT_CreateReceiptCorrectionForm
    AT_CreatePropertyForms
    AT_CreateCampaignForms
    AT_CreateReceiptAllocationForm
    AT_CreateCampaignTargetForm
    AT_CreateOccupancyForm
    AT_CreateExpenseForms
    AT_CreateConfigurationForms
    AT_CreateFinancialClosingForms
    AT_CreateReadOnlyGridForm "frmOwnerDirectory", AT_U("0633,062C,0644,0020,0627,0644,0645,0644,0627,0643"), "SELECT OwnerCode, FullName, Phone, HouseCount, ApartmentCount, AdminStatus, DataStatus, ReviewStatus FROM tblOwners ORDER BY FullName;"
    AT_CreateReadOnlyGridForm "frmReceiptList", AT_U("0633,062C,0644,0020,0627,0644,0625,064A,0635,0627,0644,0627,062A"), "SELECT N.OfficialNumber, R.ReceiptID, R.PrintedOwnerCode, R.PrintedOwnerName, D.TransactionDate, R.Amount, PM.MethodName, A.AccountName, D.StatusCode FROM (((tblReceipts AS R INNER JOIN tblFinancialDocuments AS D ON R.DocumentID=D.DocumentID) LEFT JOIN tblOfficialDocumentNumbers AS N ON D.DocumentID=N.DocumentID) LEFT JOIN tblPaymentMethods AS PM ON R.PaymentMethodID=PM.PaymentMethodID) LEFT JOIN tblAccounts AS A ON R.AccountID=A.AccountID ORDER BY D.TransactionDate DESC, R.ReceiptID DESC;"
    AT_CreateReadOnlyGridForm "frmPaymentList", AT_U("0633,062C,0644,0020,0633,0646,062F,0627,062A,0020,0627,0644,0635,0631,0641"), "SELECT N.OfficialNumber, P.PaymentID, P.BeneficiaryName, D.TransactionDate, P.Amount, A.AccountName, D.StatusCode FROM ((tblPayments AS P INNER JOIN tblFinancialDocuments AS D ON P.DocumentID=D.DocumentID) LEFT JOIN tblOfficialDocumentNumbers AS N ON D.DocumentID=N.DocumentID) LEFT JOIN tblAccounts AS A ON P.AccountID=A.AccountID ORDER BY D.TransactionDate DESC, P.PaymentID DESC;"
    AT_CreateReadOnlyGridForm "frmTransferList", AT_U("0633,062C,0644,0020,0627,0644,062A,062D,0648,064A,0644,0627,062A"), "SELECT N.OfficialNumber, T.TransferID, T.Amount, T.FeeAmount, S.AccountName AS SourceAccount, X.AccountName AS DestinationAccount, D.TransactionDate, D.StatusCode, T.ConfirmationRef FROM (((tblAccountTransfers AS T INNER JOIN tblFinancialDocuments AS D ON T.DocumentID=D.DocumentID) LEFT JOIN tblOfficialDocumentNumbers AS N ON D.DocumentID=N.DocumentID) LEFT JOIN tblAccounts AS S ON T.SourceAccountID=S.AccountID) LEFT JOIN tblAccounts AS X ON T.DestinationAccountID=X.AccountID ORDER BY D.TransactionDate DESC, T.TransferID DESC;"
    AT_CreateReadOnlyGridForm "frmOwnerBalances", AT_U("0643,0634,0641,0020,0623,0631,0635,062F,0629,0020,0627,0644,0645,0644,0627,0643"), "SELECT * FROM qryAT_OwnerBalances ORDER BY FullName;"
    AT_CreateReadOnlyGridForm "frmAccountBalances", AT_U("0623,0631,0635,062F,0629,0020,0627,0644,062D,0633,0627,0628,0627,062A"), "SELECT * FROM qryAT_AccountBalances ORDER BY AccountName;"
    AT_CreateReadOnlyGridForm "frmAccountMovements", AT_U("062D,0631,0643,0629,0020,0627,0644,062D,0633,0627,0628,0627,062A"), "SELECT * FROM qryAT_AccountMovements;"
    AT_CreateReadOnlyGridForm "frmPendingAdjustments", AT_U("0637,0644,0628,0627,062A,0020,0627,0644,062E,0635,0645,0020,0648,0627,0644,0625,0639,0641,0627,0621,0020,0627,0644,0645,0639,0644,0642,0629"), "SELECT * FROM qryAT_PendingAdjustments;"
    AT_CreateReadOnlyGridForm "frmCreditLedger", AT_U("062F,0641,062A,0631,0020,0623,0631,0635,062F,0629,0020,0627,0644,0645,0644,0627,0643,0020,0627,0644,062F,0627,0626,0646,0629"), "SELECT * FROM qryAT_CreditLedger;"
    AT_CreateReadOnlyGridForm "frmChargeBalances", AT_U("0643,0634,0641,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,0627,062A,0020,0648,0627,0644,0645,062A,0628,0642,064A"), "SELECT * FROM qryAT_ChargeBalances;"
    AT_CreateReadOnlyGridForm "frmOwnerStatementList", AT_U("0643,0634,0641,0020,062D,0631,0643,0629,0020,062D,0633,0627,0628,0020,0627,0644,0645,0627,0644,0643"), "SELECT * FROM qryAT_OwnerStatement ORDER BY EntryDate, ReferenceText;"
    AT_CreateReadOnlyGridForm "frmBuildingDirectory", AT_U("062F,0644,064A,0644,0020,0627,0644,0645,0628,0627,0646,064A"), "SELECT * FROM qryAT_BuildingCurrentOwners;"
    AT_CreateReadOnlyGridForm "frmUnitDirectory", AT_U("062F,0644,064A,0644,0020,0627,0644,0634,0642,0642"), "SELECT * FROM qryAT_UnitDirectory;"
    AT_CreateReadOnlyGridForm "frmCampaignDirectory", AT_U("0633,062C,0644,0020,0627,0644,062D,0645,0644,0627,062A"), "SELECT * FROM qryAT_CampaignRegister;"
    AT_CreateReadOnlyGridForm "frmExpenseDirectory", AT_U("0633,062C,0644,0020,0627,0644,0645,0635,0631,0648,0641,0627,062A"), "SELECT * FROM qryAT_ExpenseRegister;"
    AT_CreateReadOnlyGridForm "frmDraftReceiptAllocations", AT_U("062A,062E,0635,064A,0635,0627,062A,0020,0627,0644,0625,064A,0635,0627,0644,0627,062A,0020,0627,0644,0645,0639,0644,0642,0629"), "SELECT * FROM qryAT_ReceiptDraftAllocations;"
    AT_CreateReadOnlyGridForm "frmExemptionReview", AT_U("0637,0644,0628,0627,062A,0020,0625,0639,0641,0627,0621,0020,0627,0644,0634,0642,0642"), "SELECT * FROM qryAT_UnitExemptionReview;"
    AT_VerifyRequiredGeneratedForms
    If Not AT_FormIsOpen("frmMain") Then DoCmd.OpenForm "frmMain", acNormal
    MsgBox AT_U("062A,0645,0020,0625,0646,0634,0627,0621,0020,0627,0644,0648,0627,062C,0647,0627,062A,0020,0627,0644,0623,0648,0644,064A,0629,002E,0020,0631,0627,062C,0639,0020,0627,0644,0646,0645,0627,0630,062C,0020,0648,0627,0644,0648,0638,0627,0626,0641,0020,062F,0627,062E,0644,0020,0041,0063,0063,0065,0073,0073,060C,0020,0648,0646,0641,0651,0630,0020,0627,0644,062A,0631,062C,0645,0629,0020,0648,0627,0644,0627,062E,062A,0628,0627,0631,0627,062A,0020,0642,0628,0644,0020,0627,0644,0627,0633,062A,062E,062F,0627,0645,0020,0627,0644,0641,0639,0644,064A,002E"), vbInformation, "Archon Tech"
    Exit Sub
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description: savedErrSource = Err.Source
    If Not formsOpenBeforeBuild Is Nothing Then AT_CloseFormsOpenedDuringBuild formsOpenBeforeBuild, cleanupDetails
    If Len(cleanupDetails) > 0 Then eDesc = eDesc & " | UI cleanup: " & cleanupDetails
    MsgBox AT_U("062A,0639,0630,0631,0020,0625,0643,0645,0627,0644,0020,0625,0646,0634,0627,0621,0020,0627,0644,0648,0627,062C,0647,0627,062A,002E,0020,0627,0644,062E,0637,0623,0020") & savedErrNumber & " [" & savedErrSource & "]: " & eDesc, vbExclamation, "Archon Tech"
    Err.Raise savedErrNumber, savedErrSource, eDesc
End Sub

Public Function AT_UI_AttachReceiptFile() As Boolean
    On Error GoTo EH
    Dim receiptID As Long, documentID As Long, attachmentID As Long
    receiptID = CLng(Nz(Forms("frmReceiptCreate").Controls("txtReceiptID").Value, 0))
    If receiptID <= 0 Then Err.Raise vbObjectError + 711, "AT_UI_AttachReceiptFile", AT_U("0627,062D,0641,0638,0020,0645,0633,0648,062F,0629,0020,0627,0644,0625,064A,0635,0627,0644,0020,0623,0648,0644,0627,064B,0020,0642,0628,0644,0020,0625,0631,0641,0627,0642,0020,0645,0644,0641,002E")
    documentID = CLng(Nz(DLookup("DocumentID", "tblReceipts", "ReceiptID=" & receiptID), 0))
    If documentID <= 0 Then Err.Raise vbObjectError + 712, "AT_UI_AttachReceiptFile", "Receipt document link was not found."
    attachmentID = AT_SelectAndAttachFile(documentID)
    If attachmentID > 0 Then
        MsgBox AT_U("062A,0645,0020,0625,0631,0641,0627,0642,0020,0627,0644,0645,0644,0641,0020,0628,0646,062C,0627,062D,002E,0020,0645,0639,0631,0641,0020,0627,0644,0645,0631,0641,0642,003A,0020") & attachmentID, vbInformation, "Archon Tech"
        AT_UI_AttachReceiptFile = True
    End If
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0625,0631,0641,0627,0642,0020,0645,0631,0641,0642,0020,0628,0627,0644,0625,064A,0635,0627,0644,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_AttachPaymentFile() As Boolean
    On Error GoTo EH
    Dim paymentID As Long, documentID As Long, attachmentID As Long
    paymentID = CLng(Nz(Forms("frmPaymentCreate").Controls("txtPaymentID").Value, 0))
    If paymentID <= 0 Then Err.Raise vbObjectError + 713, "AT_UI_AttachPaymentFile", AT_U("0627,062D,0641,0638,0020,0645,0633,0648,062F,0629,0020,0633,0646,062F,0020,0627,0644,0635,0631,0641,0020,0623,0648,0644,0627,064B,0020,0642,0628,0644,0020,0625,0631,0641,0627,0642,0020,0645,0644,0641,002E")
    documentID = CLng(Nz(DLookup("DocumentID", "tblPayments", "PaymentID=" & paymentID), 0))
    If documentID <= 0 Then Err.Raise vbObjectError + 714, "AT_UI_AttachPaymentFile", "Payment document link was not found."
    attachmentID = AT_SelectAndAttachFile(documentID)
    If attachmentID > 0 Then
        MsgBox AT_U("062A,0645,0020,0625,0631,0641,0627,0642,0020,0627,0644,0645,0644,0641,0020,0628,0646,062C,0627,062D,002E,0020,0645,0639,0631,0641,0020,0627,0644,0645,0631,0641,0642,003A,0020") & attachmentID, vbInformation, "Archon Tech"
        AT_UI_AttachPaymentFile = True
    End If
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0625,0631,0641,0627,0642,0020,0645,0631,0641,0642,0020,0628,0633,0646,062F,0020,0627,0644,0635,0631,0641,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_AttachExpensePaymentFile() As Boolean
    On Error GoTo EH
    Dim paymentID As Long, documentID As Long, attachmentID As Long
    paymentID = CLng(Nz(Forms("frmExpensePay").Controls("txtPaymentID").Value, 0))
    If paymentID <= 0 Then Err.Raise vbObjectError + 715, "AT_UI_AttachExpensePaymentFile", AT_U("0627,0646,0634,0626,0020,0645,0633,0648,062F,0629,0020,0633,0646,062F,0020,0627,0644,0635,0631,0641,0020,0642,0628,0644,0020,0625,0631,0641,0627,0642,0020,0645,0644,0641,002E")
    documentID = CLng(Nz(DLookup("DocumentID", "tblPayments", "PaymentID=" & paymentID), 0))
    If documentID <= 0 Then Err.Raise vbObjectError + 716, "AT_UI_AttachExpensePaymentFile", "Expense payment document link was not found."
    attachmentID = AT_SelectAndAttachFile(documentID)
    If attachmentID > 0 Then
        MsgBox AT_U("062A,0645,0020,0625,0631,0641,0627,0642,0020,0627,0644,0645,0644,0641,0020,0628,0646,062C,0627,062D,002E,0020,0645,0639,0631,0641,0020,0627,0644,0645,0631,0641,0642,003A,0020") & attachmentID, vbInformation, "Archon Tech"
        AT_UI_AttachExpensePaymentFile = True
    End If
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0625,0631,0641,0627,0642,0020,0645,0631,0641,0642,0020,0628,0633,0646,062F,0020,0627,0644,0635,0631,0641,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_OpenForm(ByVal formName As String) As Boolean
    On Error GoTo EH
    DoCmd.OpenForm formName, acNormal
    AT_UI_OpenForm = True
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0641,062A,062D,0020,0627,0644,0634,0627,0634,0629,0020") & formName & ": " & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_RunIntegrity() As Boolean
    On Error GoTo EH
    AT_RunIntegrityReview
    AT_UI_RunIntegrity = True
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0641,062D,0635,0020,0633,0644,0627,0645,0629,0020,0627,0644,0628,064A,0627,0646,0627,062A,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SaveOwnerDraft() As Boolean
    On Error GoTo EH
    Dim f As Form, ownerID As Long
    Set f = Forms("frmOwnerCreate")
    ownerID = AT_CreateOwnerDraft(CStr(Nz(f.Controls("txtFullName").Value, "")), CStr(Nz(f.Controls("txtPhone").Value, "")), 0, 0, CStr(Nz(f.Controls("txtNotes").Value, "")), CStr(Nz(f.Controls("txtIdentityNo").Value, "")))
    f.Controls("txtOwnerID").Value = ownerID
    MsgBox AT_U("062D,064F,0641,0638,062A,0020,0645,0633,0648,062F,0629,0020,0627,0644,0645,0627,0644,0643,0020,0628,0627,0644,0645,0639,0631,0641,0020,0627,0644,062F,0627,062E,0644,064A,0020") & ownerID & AT_U("002E,0020,0644,0645,0020,064A,064F,062E,0635,0635,0020,0631,0642,0645,0020,004F,0057,0020,0628,0639,062F,002E,0020,0631,0627,062C,0639,0020,0627,0644,0628,064A,0627,0646,0627,062A,0020,062B,0645,0020,0627,0639,062A,0645,062F,0647,0627,0020,0639,0646,062F,0020,0627,0644,062C,0627,0647,0632,064A,0629,002E"), vbInformation, "Archon Tech"
    AT_UI_SaveOwnerDraft = True
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,062D,0641,0638,0020,0645,0633,0648,062F,0629,0020,0627,0644,0645,0627,0644,0643,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ApproveOwner() As Boolean
    On Error GoTo EH
    Dim ownerID As Long
    ownerID = CLng(Val(Nz(Forms("frmOwnerCreate").Controls("txtOwnerID").Value, 0)))
    If ownerID <= 0 Then Err.Raise vbObjectError + 701, , AT_U("0627,062D,0641,0638,0020,0645,0633,0648,062F,0629,0020,0627,0644,0645,0627,0644,0643,0020,0623,0648,0644,064B,0627,002E")
    If MsgBox(AT_U("0647,0644,0020,0631,0627,062C,0639,062A,0020,0628,064A,0627,0646,0627,062A,0020,0627,0644,0645,0627,0644,0643,0020,0648,062A,0631,064A,062F,0020,0627,0639,062A,0645,0627,062F,0647,0627,0020,0648,062A,062E,0635,064A,0635,0020,0631,0642,0645,0020,0631,0633,0645,064A,0020,0644,0627,0020,064A,0639,0627,062F,0020,0627,0633,062A,062E,062F,0627,0645,0647,061F"), vbQuestion + vbYesNo, AT_U("062A,0623,0643,064A,062F,0020,0627,0639,062A,0645,0627,062F,0020,0627,0644,0645,0627,0644,0643")) <> vbYes Then Exit Function
    AT_ApproveOwner ownerID, Environ$("Username")
    MsgBox AT_U("062A,0645,0020,0627,0639,062A,0645,0627,062F,0020,0627,0644,0645,0627,0644,0643,0020,0648,062A,062E,0635,064A,0635,0020,0627,0644,0631,0642,0645,0020") & Nz(DLookup("OwnerCode", "tblOwners", "OwnerID=" & ownerID), "") & ".", vbInformation, "Archon Tech"
    AT_UI_ApproveOwner = True
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0627,0639,062A,0645,0627,062F,0020,0627,0644,0645,0627,0644,0643,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SaveAccount() As Boolean
    On Error GoTo EH
    Dim f As Form, accountID As Long
    Set f = Forms("frmAccountCreate")
    If MsgBox(AT_U("0633,064A,064F,0646,0634,0623,0020,0627,0644,062D,0633,0627,0628,0020,0628,0631,0635,064A,062F,0020,0627,0641,062A,062A,0627,062D,064A,0020,0635,0641,0631,060C,0020,0648,0644,0627,0020,062A,064F,0636,0627,0641,0020,0627,0644,0623,0645,0648,0627,0644,0020,0625,0644,0627,0020,0639,0628,0631,0020,062D,0631,0643,0629,0020,0645,0627,0644,064A,0629,0020,0645,0639,062A,0645,062F,0629,002E,0020,0647,0644,0020,062A,0631,064A,062F,0020,0627,0644,0645,062A,0627,0628,0639,0629,061F"), vbQuestion + vbYesNo, AT_U("062A,0623,0643,064A,062F,0020,0625,0646,0634,0627,0621,0020,0627,0644,062D,0633,0627,0628")) <> vbYes Then Exit Function
    accountID = AT_CreateAccount(CStr(Nz(f.Controls("txtAccountCode").Value, "")), _
        CStr(Nz(f.Controls("txtAccountName").Value, "")), CStr(Nz(f.Controls("cboAccountType").Value, "")))
    f.Controls("txtAccountID").Value = accountID
    MsgBox AT_U("062A,0645,0020,0625,0646,0634,0627,0621,0020,0627,0644,062D,0633,0627,0628,0020,0628,0631,0642,0645,0020,062F,0627,062E,0644,064A,0020") & accountID & AT_U("0020,0648,0628,0631,0635,064A,062F,0020,0627,0641,062A,062A,0627,062D,064A,0020,0635,0641,0631,002E"), vbInformation, "Archon Tech"
    AT_UI_SaveAccount = True
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0625,0646,0634,0627,0621,0020,0627,0644,062D,0633,0627,0628,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SearchOwners() As Boolean
    On Error GoTo EH
    Dim f As Form, rawTerm As String, normalizedTerm As String, whereClause As String
    Set f = Forms("frmOwnerSearch")
    rawTerm = Trim$(CStr(Nz(f.Controls("txtSearchTerm").Value, "")))
    If Len(rawTerm) = 0 Then Err.Raise vbObjectError + 708, , AT_U("0623,062F,062E,0644,0020,0627,0644,0627,0633,0645,0020,0623,0648,0020,0631,0642,0645,0020,0627,0644,0645,0627,0644,0643,0020,0623,0648,0020,0627,0644,0647,0627,062A,0641,0020,0644,0644,0628,062D,062B,002E")
    normalizedTerm = Replace(AT_NormalizeArabicSearch(rawTerm), "'", "''")
    rawTerm = Replace(rawTerm, "'", "''")
    whereClause = "AdminStatus='ACTIVE' AND (FullNameSearch Like '*" & normalizedTerm & "*' OR FullName Like '*" & rawTerm & "*' OR OwnerCode Like '*" & rawTerm & "*' OR Phone Like '*" & rawTerm & "*')"
    If CurrentProject.AllForms("frmOwnerDirectory").IsLoaded Then
        Forms("frmOwnerDirectory").Filter = whereClause
        Forms("frmOwnerDirectory").FilterOn = True
        Forms("frmOwnerDirectory").Requery
        DoCmd.SelectObject acForm, "frmOwnerDirectory", True
    Else
        DoCmd.OpenForm "frmOwnerDirectory", acNormal, , whereClause
    End If
    AT_UI_SearchOwners = True
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0627,0644,0628,062D,062B,0020,0639,0646,0020,0627,0644,0645,0627,0644,0643,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_OpenOwnerStatement() As Boolean
    On Error GoTo EH
    Dim f As Form, ownerID As Long
    Set f = Forms("frmOwnerStatementPick")
    ownerID = CLng(Nz(f.Controls("cboOwnerID").Value, 0))
    If ownerID <= 0 Then Err.Raise vbObjectError + 711, , AT_U("0627,062E,062A,0631,0020,0627,0644,0645,0627,0644,0643,0020,0623,0648,0644,064B,0627,002E")
    If CurrentProject.AllForms("frmOwnerStatementList").IsLoaded Then
        Forms("frmOwnerStatementList").Filter = "OwnerID=" & ownerID
        Forms("frmOwnerStatementList").FilterOn = True
        Forms("frmOwnerStatementList").Requery
        DoCmd.SelectObject acForm, "frmOwnerStatementList", True
    Else
        DoCmd.OpenForm "frmOwnerStatementList", acNormal, , "OwnerID=" & ownerID
    End If
    AT_UI_OpenOwnerStatement = True
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0641,062A,062D,0020,0643,0634,0641,0020,062D,0633,0627,0628,0020,0627,0644,0645,0627,0644,0643,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_LoadOwnerEdit() As Boolean
    On Error GoTo EH
    Dim f As Form, ownerID As Long
    Set f = Forms("frmOwnerEdit")
    ownerID = CLng(Nz(f.Controls("cboOwnerID").Value, 0))
    If ownerID <= 0 Then Err.Raise vbObjectError + 709, , AT_U("0627,062E,062A,0631,0020,0627,0644,0645,0627,0644,0643,0020,0623,0648,0644,064B,0627,002E")
    f.Controls("txtFullName").Value = Nz(DLookup("FullName", "tblOwners", "OwnerID=" & ownerID), "")
    f.Controls("txtPhone").Value = Nz(DLookup("Phone", "tblOwners", "OwnerID=" & ownerID), "")
    f.Controls("txtIdentityNo").Value = Nz(DLookup("IdentityNo", "tblOwners", "OwnerID=" & ownerID), "")
    f.Controls("txtNotes").Value = Nz(DLookup("Notes", "tblOwners", "OwnerID=" & ownerID), "")
    AT_UI_LoadOwnerEdit = True
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,062A,062D,0645,064A,0644,0020,0628,064A,0627,0646,0627,062A,0020,0627,0644,0645,0627,0644,0643,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SaveOwnerEdit() As Boolean
    On Error GoTo EH
    Dim f As Form, ownerID As Long
    Set f = Forms("frmOwnerEdit")
    ownerID = CLng(Nz(f.Controls("cboOwnerID").Value, 0))
    If ownerID <= 0 Then Err.Raise vbObjectError + 710, , AT_U("0627,062E,062A,0631,0020,0627,0644,0645,0627,0644,0643,0020,0623,0648,0644,064B,0627,002E")
    If MsgBox(AT_U("0633,064A,062A,0645,0020,062A,0639,062F,064A,0644,0020,0628,064A,0627,0646,0627,062A,0020,0627,0644,0645,0627,0644,0643,0020,0648,062A,0633,062C,064A,0644,0020,0627,0644,0642,064A,0645,0020,0627,0644,0633,0627,0628,0642,0629,0020,0648,0627,0644,062C,062F,064A,062F,0629,0020,0641,064A,0020,0633,062C,0644,0020,0627,0644,062A,062F,0642,064A,0642,002E,0020,0647,0644,0020,062A,0631,064A,062F,0020,0627,0644,0645,062A,0627,0628,0639,0629,061F"), vbQuestion + vbYesNo, AT_U("062A,0623,0643,064A,062F,0020,0627,0644,062A,0639,062F,064A,0644")) <> vbYes Then Exit Function
    AT_UpdateOwnerDetails ownerID, CStr(Nz(f.Controls("txtFullName").Value, "")), CStr(Nz(f.Controls("txtPhone").Value, "")), _
        0, 0, _
        CStr(Nz(f.Controls("txtNotes").Value, "")), CStr(Nz(f.Controls("txtEditReason").Value, "")), CStr(Nz(f.Controls("txtIdentityNo").Value, ""))
    MsgBox AT_U("062A,0645,0020,062D,0641,0638,0020,0627,0644,062A,0639,062F,064A,0644,0020,0648,062A,0633,062C,064A,0644,0020,0627,0644,0642,064A,0645,0020,0627,0644,0633,0627,0628,0642,0629,0020,0648,0627,0644,062C,062F,064A,062F,0629,002E"), vbInformation, "Archon Tech"
    AT_UI_SaveOwnerEdit = True
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,062A,0639,062F,064A,0644,0020,0628,064A,0627,0646,0627,062A,0020,0627,0644,0645,0627,0644,0643,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SaveChargeDraft() As Boolean
    On Error GoTo EH
    Dim f As Form, chargeID As Long, chargeDate As Date, rateID As Long, ownerID As Long
    Set f = Forms("frmChargeCreate")
    ownerID = CLng(Nz(f.Controls("cboOwnerID").Value, 0)): rateID = CLng(Nz(f.Controls("cboRateID").Value, 0))
    chargeDate = AT_UI_DateValue(f.Controls("txtChargeDate").Value)
    chargeID = AT_CreateChargeForOwner(ownerID, chargeDate, rateID, CStr(Nz(f.Controls("txtReason").Value, "")))
    f.Controls("txtChargeID").Value = chargeID
    MsgBox AT_U("0623,064F,0646,0634,0626,062A,0020,0645,0633,0648,062F,0629,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,002E,0020,0627,0644,0625,062C,0645,0627,0644,064A,0020,0627,0644,0645,062D,0641,0648,0638,003A,0020") & FormatCurrency(Nz(DLookup("GrossAmount", "tblOwnerCharges", "ChargeID=" & chargeID), 0), 0) & AT_U("0020,0631,064A,0627,0644,002E,0020,0644,0627,0020,064A,0624,062B,0631,0020,0642,0628,0644,0020,0627,0644,0627,0639,062A,0645,0627,062F,002E"), vbInformation, "Archon Tech"
    AT_UI_SaveChargeDraft = True
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0625,0646,0634,0627,0621,0020,0645,0633,0648,062F,0629,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ApproveCharge() As Boolean
    On Error GoTo EH
    Dim chargeID As Long
    chargeID = CLng(Val(Nz(Forms("frmChargeCreate").Controls("txtChargeID").Value, 0)))
    If chargeID <= 0 Then Err.Raise vbObjectError + 702, , AT_U("0623,0646,0634,0626,0020,0645,0633,0648,062F,0629,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,0020,0623,0648,0644,064B,0627,002E")
    If MsgBox(AT_U("0627,0639,062A,0645,0627,062F,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,061F,0020,0633,064A,0635,0628,062D,0020,0636,0645,0646,0020,0627,0644,0645,0633,062A,062D,0642,0627,062A,0020,0627,0644,0645,0627,0644,064A,0629,0020,0644,0644,0645,0627,0644,0643,002E"), vbQuestion + vbYesNo, AT_U("062A,0623,0643,064A,062F,0020,0627,0644,0627,0639,062A,0645,0627,062F")) <> vbYes Then Exit Function
    AT_ApproveCharge chargeID, Environ$("Username")
    MsgBox AT_U("062A,0645,0020,0627,0639,062A,0645,0627,062F,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,002E"), vbInformation, "Archon Tech"
    AT_UI_ApproveCharge = True
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0627,0639,062A,0645,0627,062F,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SaveReceiptDraft() As Boolean
    On Error GoTo EH
    Dim f As Form, receiptID As Long, transactionDate As Date
    Set f = Forms("frmReceiptCreate"): transactionDate = AT_UI_DateValue(f.Controls("txtTransactionDate").Value)
    receiptID = AT_CreateReceiptDraft(CLng(Nz(f.Controls("cboOwnerID").Value, 0)), CCur(Nz(f.Controls("txtAmount").Value, 0)), CLng(Nz(f.Controls("cboPaymentMethodID").Value, 0)), CLng(Nz(f.Controls("cboAccountID").Value, 0)), CStr(Nz(f.Controls("txtPayerName").Value, "")), CStr(Nz(f.Controls("txtCollectorName").Value, "")), CStr(Nz(f.Controls("txtNotes").Value, "")), transactionDate)
    f.Controls("txtReceiptID").Value = receiptID
    MsgBox AT_U("0623,064F,0646,0634,0626,062A,0020,0645,0633,0648,062F,0629,0020,0627,0644,0625,064A,0635,0627,0644,0020,0631,0642,0645,0647,0627,0020,0627,0644,062F,0627,062E,0644,064A,0020") & receiptID & AT_U("002E,0020,0644,0645,0020,064A,062A,063A,064A,0631,0020,0631,0635,064A,062F,0020,0627,0644,062D,0633,0627,0628,0020,0628,0639,062F,002E"), vbInformation, "Archon Tech"
    AT_UI_SaveReceiptDraft = True
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0625,0646,0634,0627,0621,0020,0645,0633,0648,062F,0629,0020,0627,0644,0625,064A,0635,0627,0644,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ApproveReceipt() As Boolean
    On Error GoTo EH
    Dim receiptID As Long
    receiptID = CLng(Val(Nz(Forms("frmReceiptCreate").Controls("txtReceiptID").Value, 0)))
    If receiptID <= 0 Then Err.Raise vbObjectError + 703, , AT_U("0623,0646,0634,0626,0020,0645,0633,0648,062F,0629,0020,0627,0644,0625,064A,0635,0627,0644,0020,0623,0648,0644,064B,0627,002E")
    If MsgBox(AT_U("062A,0623,0643,064A,062F,003A,0020,0647,0644,0020,0627,0633,062A,0644,0645,062A,0020,0627,0644,0645,0628,0644,063A,0020,0641,0639,0644,064A,064B,0627,0020,0645,0646,0020,0627,0644,062F,0627,0641,0639,0020,0648,0628,0627,0644,0642,064A,0645,0629,0020,0627,0644,0645,0633,062C,0644,0629,061F"), vbQuestion + vbYesNo + vbDefaultButton2, AT_U("062A,0623,0643,064A,062F,0020,0627,0644,0642,0628,0636,0020,0627,0644,0641,0639,0644,064A")) <> vbYes Then Exit Function
    AT_ApproveReceipt receiptID, Environ$("Username"), True
    MsgBox AT_U("062A,0645,0020,0627,0639,062A,0645,0627,062F,0020,0627,0644,0625,064A,0635,0627,0644,002E,0020,0627,0644,0631,0642,0645,0020,0627,0644,0631,0633,0645,064A,003A,0020") & Nz(DLookup("OfficialNumber", "tblOfficialDocumentNumbers", "DocumentID=" & Nz(DLookup("DocumentID", "tblReceipts", "ReceiptID=" & receiptID), 0)), "") & ".", vbInformation, "Archon Tech"
    AT_UI_ApproveReceipt = True
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0627,0639,062A,0645,0627,062F,0020,0627,0644,0625,064A,0635,0627,0644,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SavePaymentDraft() As Boolean
    On Error GoTo EH
    Dim f As Form, paymentID As Long, transactionDate As Date
    Set f = Forms("frmPaymentCreate"): transactionDate = AT_UI_DateValue(f.Controls("txtTransactionDate").Value)
    paymentID = AT_CreatePaymentDraft(CLng(Nz(f.Controls("cboAccountID").Value, 0)), CLng(Nz(f.Controls("cboExpenseCategoryID").Value, 0)), CCur(Nz(f.Controls("txtAmount").Value, 0)), CLng(Nz(f.Controls("cboPaymentMethodID").Value, 0)), CStr(Nz(f.Controls("txtBeneficiary").Value, "")), CStr(Nz(f.Controls("txtRecipient").Value, "")), CStr(Nz(f.Controls("txtNotes").Value, "")), transactionDate, CLng(Val(Nz(f.Controls("txtRefundCreditMovementID").Value, 0))))
    f.Controls("txtPaymentID").Value = paymentID
    MsgBox AT_U("0623,064F,0646,0634,0626,062A,0020,0645,0633,0648,062F,0629,0020,0633,0646,062F,0020,0627,0644,0635,0631,0641,002E,0020,0644,0645,0020,064A,064F,062E,0635,0645,0020,0634,064A,0621,0020,0645,0646,0020,0627,0644,062D,0633,0627,0628,0020,0628,0639,062F,002E"), vbInformation, "Archon Tech"
    AT_UI_SavePaymentDraft = True
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0625,0646,0634,0627,0621,0020,0645,0633,0648,062F,0629,0020,0633,0646,062F,0020,0627,0644,0635,0631,0641,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ApprovePayment() As Boolean
    On Error GoTo EH
    Dim paymentID As Long
    paymentID = CLng(Val(Nz(Forms("frmPaymentCreate").Controls("txtPaymentID").Value, 0)))
    If paymentID <= 0 Then Err.Raise vbObjectError + 704, , AT_U("0623,0646,0634,0626,0020,0645,0633,0648,062F,0629,0020,0633,0646,062F,0020,0627,0644,0635,0631,0641,0020,0623,0648,0644,064B,0627,002E")
    If MsgBox(AT_U("062A,0623,0643,064A,062F,003A,0020,0647,0644,0020,062F,064F,0641,0639,0020,0627,0644,0645,0628,0644,063A,0020,0641,0639,0644,064A,064B,0627,0020,0644,0644,0645,0633,062A,0641,064A,062F,061F,0020,0633,064A,062A,0645,0020,0645,0646,0639,0020,0627,0644,0627,0639,062A,0645,0627,062F,0020,0625,0646,0020,0644,0645,0020,064A,0643,0646,0020,0627,0644,0631,0635,064A,062F,0020,0627,0644,0645,062A,0627,062D,0020,0643,0627,0641,064A,064B,0627,002E"), vbQuestion + vbYesNo + vbDefaultButton2, AT_U("062A,0623,0643,064A,062F,0020,0627,0644,0635,0631,0641,0020,0627,0644,0641,0639,0644,064A")) <> vbYes Then Exit Function
    AT_ApprovePayment paymentID, Environ$("Username"), True
    MsgBox AT_U("062A,0645,0020,0627,0639,062A,0645,0627,062F,0020,0633,0646,062F,0020,0627,0644,0635,0631,0641,002E"), vbInformation, "Archon Tech"
    AT_UI_ApprovePayment = True
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0627,0639,062A,0645,0627,062F,0020,0633,0646,062F,0020,0627,0644,0635,0631,0641,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SaveTransferDraft() As Boolean
    On Error GoTo EH
    Dim f As Form, transferID As Long, transactionDate As Date
    Set f = Forms("frmTransferCreate"): transactionDate = AT_UI_DateValue(f.Controls("txtTransactionDate").Value)
    transferID = AT_CreateTransferDraft(CLng(Nz(f.Controls("cboSourceAccountID").Value, 0)), CLng(Nz(f.Controls("cboDestinationAccountID").Value, 0)), CCur(Nz(f.Controls("txtAmount").Value, 0)), CCur(Nz(f.Controls("txtFeeAmount").Value, 0)), CStr(Nz(f.Controls("txtNotes").Value, "")), transactionDate)
    f.Controls("txtTransferID").Value = transferID
    MsgBox AT_U("0623,064F,0646,0634,0626,062A,0020,0645,0633,0648,062F,0629,0020,0627,0644,062A,062D,0648,064A,0644,002E,0020,0644,0645,0020,062A,064F,062E,0635,0645,0020,0623,0648,0020,062A,064F,0636,0641,0020,0645,0628,0627,0644,063A,0020,0628,0639,062F,002E"), vbInformation, "Archon Tech"
    AT_UI_SaveTransferDraft = True
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0625,0646,0634,0627,0621,0020,0645,0633,0648,062F,0629,0020,0627,0644,062A,062D,0648,064A,0644,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ApproveTransfer() As Boolean
    On Error GoTo EH
    Dim transferID As Long
    transferID = CLng(Val(Nz(Forms("frmTransferCreate").Controls("txtTransferID").Value, 0)))
    If transferID <= 0 Then Err.Raise vbObjectError + 705, , AT_U("0623,0646,0634,0626,0020,0645,0633,0648,062F,0629,0020,0627,0644,062A,062D,0648,064A,0644,0020,0623,0648,0644,064B,0627,002E")
    If MsgBox(AT_U("0627,0639,062A,0645,0627,062F,0020,0627,0644,062A,062D,0648,064A,0644,0020,0627,0644,0645,0639,0644,0642,0020,0648,062D,062C,0632,0020,0627,0644,0645,0628,0644,063A,0020,0645,0646,0020,0627,0644,0631,0635,064A,062F,0020,0627,0644,0645,062A,0627,062D,061F,0020,0644,0646,0020,064A,064F,0636,0627,0641,0020,0644,062D,0633,0627,0628,0020,0627,0644,0648,062C,0647,0629,0020,062D,062A,0649,0020,062A,0623,0643,064A,062F,0020,0627,0644,0627,0633,062A,0644,0627,0645,002E"), vbQuestion + vbYesNo, AT_U("062A,0623,0643,064A,062F,0020,0627,0644,062A,062D,0648,064A,0644")) <> vbYes Then Exit Function
    AT_ApproveTransfer transferID, Environ$("Username")
    MsgBox AT_U("062A,0645,0020,0627,0639,062A,0645,0627,062F,0020,0627,0644,062A,062D,0648,064A,0644,0020,0643,0645,0639,0644,0642,0020,0648,062D,064F,062C,0632,0020,0627,0644,0645,0628,0644,063A,002E,0020,0623,062F,062E,0644,0020,0645,0631,062C,0639,0020,0627,0644,062A,0623,0643,064A,062F,0020,062B,0645,0020,0627,0633,062A,062E,062F,0645,0020,0632,0631,0020,062A,0623,0643,064A,062F,0020,0627,0644,0648,0635,0648,0644,002E"), vbInformation, "Archon Tech"
    AT_UI_ApproveTransfer = True
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0627,0639,062A,0645,0627,062F,0020,0627,0644,062A,062D,0648,064A,0644,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ConfirmTransfer() As Boolean
    On Error GoTo EH
    Dim f As Form, transferID As Long, refText As String
    Set f = Forms("frmTransferCreate"): transferID = CLng(Val(Nz(f.Controls("txtTransferID").Value, 0))): refText = CStr(Nz(f.Controls("txtTransferRef").Value, ""))
    If transferID <= 0 Then Err.Raise vbObjectError + 706, , AT_U("0623,062F,062E,0644,0020,0645,0639,0631,0641,0020,0627,0644,062A,062D,0648,064A,0644,0020,0627,0644,062F,0627,062E,0644,064A,0020,0623,0648,0020,0623,0646,0634,0626,0020,0627,0644,0645,0633,0648,062F,0629,002E")
    If MsgBox(AT_U("0647,0644,0020,062A,0623,0643,062F,062A,0020,0641,0639,0644,064A,064B,0627,0020,0645,0646,0020,0648,0635,0648,0644,0020,0627,0644,062A,062D,0648,064A,0644,0020,0625,0644,0649,0020,0627,0644,062D,0633,0627,0628,0020,0627,0644,0645,0633,062A,0644,0645,061F"), vbQuestion + vbYesNo + vbDefaultButton2, AT_U("062A,0623,0643,064A,062F,0020,0627,0644,0648,0635,0648,0644")) <> vbYes Then Exit Function
    AT_ConfirmTransfer transferID, Environ$("Username"), refText
    MsgBox AT_U("062A,0645,0020,062A,0631,062D,064A,0644,0020,0627,0644,062A,062D,0648,064A,0644,0020,0625,0644,0649,0020,0627,0644,062D,0633,0627,0628,064A,0646,0020,0648,062A,062D,0631,064A,0631,0020,0627,0644,062D,062C,0632,002E"), vbInformation, "Archon Tech"
    AT_UI_ConfirmTransfer = True
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,062A,0623,0643,064A,062F,0020,0627,0644,062A,062D,0648,064A,0644,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_RequestAdjustment() As Boolean
    On Error GoTo EH
    Dim f As Form, adjustmentID As Long
    Set f = Forms("frmAdjustmentCreate")
    adjustmentID = AT_RequestChargeAdjustment(CLng(Val(Nz(f.Controls("txtChargeID").Value, 0))), -Abs(CCur(Nz(f.Controls("txtDiscountAmount").Value, 0))), CStr(Nz(f.Controls("txtReason").Value, "")), Environ$("Username"), CStr(Nz(f.Controls("cboAdjustmentType").Value, "DISCOUNT")))
    f.Controls("txtAdjustmentID").Value = adjustmentID
    MsgBox AT_U("062A,0645,0020,062D,0641,0638,0020,0637,0644,0628,0020,0627,0644,062E,0635,0645,0020,0643,0645,0633,0648,062F,0629,0020,0648,0644,0627,0020,064A,0624,062B,0631,0020,0641,064A,0020,0627,0644,062D,0633,0627,0628,0020,0642,0628,0644,0020,0627,0644,0627,0639,062A,0645,0627,062F,002E"), vbInformation, "Archon Tech"
    AT_UI_RequestAdjustment = True
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,062A,0633,062C,064A,0644,0020,0637,0644,0628,0020,0627,0644,062E,0635,0645,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ApproveAdjustment() As Boolean
    On Error GoTo EH
    Dim adjustmentID As Long
    adjustmentID = CLng(Val(Nz(Forms("frmAdjustmentCreate").Controls("txtAdjustmentID").Value, 0)))
    If adjustmentID <= 0 Then Err.Raise vbObjectError + 707, , AT_U("0633,062C,0644,0020,0637,0644,0628,0020,0627,0644,062E,0635,0645,0020,0623,0648,0644,064B,0627,002E")
    If MsgBox(AT_U("0647,0644,0020,0631,0627,062C,0639,062A,0020,0627,0644,0633,0628,0628,0020,0648,0627,0644,0645,0628,0644,063A,0020,0642,0628,0644,0020,0627,0639,062A,0645,0627,062F,0020,0627,0644,062E,0635,0645,061F"), vbQuestion + vbYesNo + vbDefaultButton2, AT_U("0627,0639,062A,0645,0627,062F,0020,0627,0644,062E,0635,0645")) <> vbYes Then Exit Function
    AT_ApproveChargeAdjustment adjustmentID, Environ$("Username")
    MsgBox AT_U("062A,0645,0020,0627,0639,062A,0645,0627,062F,0020,0627,0644,062E,0635,0645,002E"), vbInformation, "Archon Tech"
    AT_UI_ApproveAdjustment = True
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0627,0639,062A,0645,0627,062F,0020,0627,0644,062E,0635,0645,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ApplyCredit() As Boolean
    On Error GoTo EH
    Dim f As Form, ownerID As Long, creditMovementID As Long, chargeID As Long, amount As Currency, reason As String
    Set f = Forms("frmCreditApply")
    ownerID = CLng(Nz(f.Controls("cboOwnerID").Value, 0))
    creditMovementID = CLng(Nz(f.Controls("cboCreditMovementID").Value, 0))
    chargeID = CLng(Nz(f.Controls("cboChargeID").Value, 0))
    amount = CCur(Nz(f.Controls("txtAmount").Value, 0)): reason = CStr(Nz(f.Controls("txtReason").Value, ""))
    If MsgBox(AT_U("0633,064A,064F,0639,062A,0645,062F,0020,062A,0637,0628,064A,0642,0020,0627,0644,0631,0635,064A,062F,0020,0627,0644,062F,0627,0626,0646,0020,0645,0628,0627,0634,0631,0629,0020,0639,0644,0649,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,0020,0628,0639,062F,0020,0627,0644,062A,062D,0642,0642,0020,0645,0646,0020,0627,0644,0645,0627,0644,0643,0020,0648,0627,0644,0645,0628,0644,063A,002E,0020,0647,0644,0020,062A,0631,064A,062F,0020,0627,0644,0645,062A,0627,0628,0639,0629,061F"), vbQuestion + vbYesNo + vbDefaultButton2, AT_U("062A,0623,0643,064A,062F,0020,062A,0637,0628,064A,0642,0020,0627,0644,0631,0635,064A,062F,0020,0627,0644,062F,0627,0626,0646")) <> vbYes Then Exit Function
    AT_ApplyOwnerCreditToCharge ownerID, creditMovementID, chargeID, amount, Environ$("Username"), reason
    MsgBox AT_U("062A,0645,0020,062A,0637,0628,064A,0642,0020,0627,0644,0631,0635,064A,062F,0020,0627,0644,062F,0627,0626,0646,0020,0648,0627,0639,062A,0645,0627,062F,0020,0627,0644,062A,0633,0648,064A,0629,002E"), vbInformation, "Archon Tech"
    AT_UI_ApplyCredit = True
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,062A,0637,0628,064A,0642,0020,0627,0644,0631,0635,064A,062F,0020,0627,0644,062F,0627,0626,0646,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ReverseAllocation() As Boolean
    On Error GoTo EH
    Dim f As Form, allocationID As Long, amount As Currency, reason As String
    Set f = Forms("frmReceiptCorrection")
    allocationID = CLng(Val(Nz(f.Controls("txtAllocationID").Value, 0)))
    amount = CCur(Nz(f.Controls("txtAmount").Value, 0)): reason = CStr(Nz(f.Controls("txtReason").Value, ""))
    If allocationID <= 0 Then Err.Raise vbObjectError + 708, , AT_U("0623,062F,062E,0644,0020,0631,0642,0645,0020,0633,0637,0631,0020,0627,0644,062A,062E,0635,064A,0635,0020,0645,0646,0020,0633,062C,0644,0020,062A,062E,0635,064A,0635,0627,062A,0020,0627,0644,0625,064A,0635,0627,0644,002E")
    If MsgBox(AT_U("0633,064A,0639,0627,062F,0020,0641,062A,062D,0020,0627,0644,0645,0633,062A,062D,0642,0020,0628,0642,064A,0645,0629,0020,0627,0644,0639,0643,0633,0020,0648,064A,0635,0628,062D,0020,0627,0644,0645,0628,0644,063A,0020,0627,0644,0645,0642,0627,0628,0644,0020,0631,0635,064A,062F,064B,0627,0020,062F,0627,0626,0646,064B,0627,0020,0645,0627,0020,062F,0627,0645,0020,0627,0644,0646,0642,062F,0020,0644,0645,0020,064A,064F,0631,062F,002E,0020,0627,0644,0631,062F,0020,0627,0644,0646,0642,062F,064A,0020,064A,062D,062A,0627,062C,0020,0633,0646,062F,0020,0635,0631,0641,0020,0645,0646,0641,0635,0644,064B,0627,002E,0020,0645,062A,0627,0628,0639,0629,061F"), vbQuestion + vbYesNo + vbDefaultButton2, AT_U("062A,0623,0643,064A,062F,0020,0639,0643,0633,0020,0627,0644,062A,062E,0635,064A,0635")) <> vbYes Then Exit Function
    AT_ReverseReceiptAllocation allocationID, amount, Environ$("Username"), reason
    MsgBox AT_U("062A,0645,0020,0639,0643,0633,0020,0627,0644,062A,062E,0635,064A,0635,0020,0648,0625,0646,0634,0627,0621,0020,0627,0644,0631,0635,064A,062F,0020,0627,0644,062F,0627,0626,0646,0020,0627,0644,0645,0642,0627,0628,0644,002E"), vbInformation, "Archon Tech"
    AT_UI_ReverseAllocation = True
    Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0639,0643,0633,0020,0627,0644,062A,062E,0635,064A,0635,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function


Public Function AT_UI_SaveReceiptAllocation() As Boolean
    On Error GoTo EH
    Dim f As Form, allocationID As Long
    Set f = Forms("frmReceiptAllocations")
    allocationID = AT_SetReceiptAllocation(CLng(Nz(f.Controls("cboReceiptID").Value, 0)), CLng(Nz(f.Controls("cboChargeID").Value, 0)), CCur(Nz(f.Controls("txtAllocatedAmount").Value, 0)))
    MsgBox AT_U("062A,0645,0020,062D,0641,0638,0020,0627,0644,062A,062E,0635,064A,0635,0020,0627,0644,064A,062F,0648,064A,0020,0643,0645,0633,0648,062F,0629,0020,0631,0642,0645,0020") & allocationID & AT_U("002E,0020,0644,0646,0020,064A,0624,062B,0631,0020,0641,064A,0020,0631,0635,064A,062F,0020,0627,0644,0645,0627,0644,0643,0020,062D,062A,0649,0020,0627,0639,062A,0645,0627,062F,0020,0627,0644,0625,064A,0635,0627,0644,002E"), vbInformation, "Archon Tech"
    AT_UI_SaveReceiptAllocation = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,062D,0641,0638,0020,0627,0644,062A,062E,0635,064A,0635,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ClearReceiptAllocations() As Boolean
    On Error GoTo EH
    Dim receiptID As Long
    receiptID = CLng(Nz(Forms("frmReceiptAllocations").Controls("cboReceiptID").Value, 0))
    If receiptID <= 0 Then Err.Raise vbObjectError + 923, , AT_U("0627,062E,062A,0631,0020,0627,0644,0625,064A,0635,0627,0644,0020,0623,0648,0644,064B,0627,002E")
    If MsgBox(AT_U("0633,064A,062A,0645,0020,062D,0630,0641,0020,062A,062E,0635,064A,0635,0627,062A,0020,0627,0644,0645,0633,0648,062F,0629,0020,0644,0647,0630,0627,0020,0627,0644,0625,064A,0635,0627,0644,0020,0641,0642,0637,002E,0020,0645,062A,0627,0628,0639,0629,061F"), vbQuestion + vbYesNo + vbDefaultButton2, AT_U("062A,0623,0643,064A,062F,0020,0645,0633,062D,0020,0627,0644,062A,062E,0635,064A,0635,0627,062A")) <> vbYes Then Exit Function
    AT_ClearReceiptDraftAllocations receiptID
    MsgBox AT_U("062A,0645,0020,0645,0633,062D,0020,062A,062E,0635,064A,0635,0627,062A,0020,0627,0644,0645,0633,0648,062F,0629,002E"), vbInformation, "Archon Tech"
    AT_UI_ClearReceiptAllocations = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0645,0633,062D,0020,0627,0644,062A,062E,0635,064A,0635,0627,062A,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SaveCampaignTarget() As Boolean
    On Error GoTo EH
    Dim f As Form, campaignID As Long, targetKind As String
    Set f = Forms("frmCampaignTargets")
    campaignID = CLng(Nz(f.Controls("cboCampaignID").Value, 0)): targetKind = UCase$(CStr(Nz(f.Controls("cboTargetKind").Value, "")))
    If campaignID <= 0 Then Err.Raise vbObjectError + 924, , AT_U("0627,062E,062A,0631,0020,0627,0644,062D,0645,0644,0629,0020,0623,0648,0644,064B,0627,002E")
    Select Case targetKind
        Case "BUILDING": AT_AddCampaignBuildingTarget campaignID, CLng(Nz(f.Controls("cboBuildingID").Value, 0))
        Case "UNIT": AT_AddCampaignUnitTarget campaignID, CLng(Nz(f.Controls("cboUnitID").Value, 0))
        Case Else: Err.Raise vbObjectError + 925, , AT_U("0627,062E,062A,0631,0020,0646,0648,0639,0020,0627,0644,0647,062F,0641,003A,0020,0645,0628,0646,0649,0020,0623,0648,0020,0634,0642,0629,002E")
    End Select
    MsgBox AT_U("062A,0645,062A,0020,0625,0636,0627,0641,0629,0020,0627,0644,0647,062F,0641,0020,0625,0644,0649,0020,0646,0637,0627,0642,0020,0627,0644,062D,0645,0644,0629,002E"), vbInformation, "Archon Tech"
    AT_UI_SaveCampaignTarget = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0625,0636,0627,0641,0629,0020,0647,062F,0641,0020,0627,0644,062D,0645,0644,0629,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SaveOccupancy() As Boolean
    On Error GoTo EH
    Dim f As Form, occupancyID As Long
    Set f = Forms("frmUnitOccupancy")
    occupancyID = AT_AddUnitOccupancy(CLng(Nz(f.Controls("cboUnitID").Value, 0)), CStr(Nz(f.Controls("txtFamilyName").Value, "")), CLng(Nz(f.Controls("txtHouseholdCount").Value, 0)), AT_UI_DateValue(f.Controls("txtStartDate").Value), CStr(Nz(f.Controls("txtNotes").Value, "")))
    f.Controls("txtOccupancyID").Value = occupancyID
    MsgBox AT_U("062A,0645,0020,062A,0633,062C,064A,0644,0020,0627,0644,0625,0634,063A,0627,0644,0020,0648,062D,0641,0638,0020,062A,0627,0631,064A,062E,0647,0020,0628,0627,0644,0645,0639,0631,0641,0020") & occupancyID & ".", vbInformation, "Archon Tech"
    AT_UI_SaveOccupancy = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,062A,0633,062C,064A,0644,0020,0627,0644,0625,0634,063A,0627,0644,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function


Public Function AT_UI_SaveProjectSettings() As Boolean
    On Error GoTo EH
    Dim f As Form
    Set f = Forms("frmProjectSettings")
    AT_UpdateProjectSettings CStr(Nz(f.Controls("txtProjectName").Value, "")), CStr(Nz(f.Controls("txtNeighborhoodName").Value, "")), CStr(Nz(f.Controls("txtDefaultCollector").Value, "")), CStr(Nz(f.Controls("txtContactInfo").Value, ""))
    MsgBox AT_U("062A,0645,0020,062D,0641,0638,0020,0625,0639,062F,0627,062F,0627,062A,0020,0627,0633,0645,0020,0627,0644,0645,0634,0631,0648,0639,0020,0648,0627,0644,062D,0627,0631,0629,0020,0648,0627,0644,0645,062D,0635,0651,0644,002E"), vbInformation, "Archon Tech"
    AT_UI_SaveProjectSettings = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,062D,0641,0638,0020,0625,0639,062F,0627,062F,0627,062A,0020,0627,0644,0645,0634,0631,0648,0639,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_LoadProjectSettings() As Boolean
    On Error GoTo EH
    Dim f As Form
    Set f = Forms("frmProjectSettings")
    f.Controls("txtProjectName").Value = Nz(DLookup("ProjectName", "tblProjectSettings"), "")
    f.Controls("txtNeighborhoodName").Value = Nz(DLookup("NeighborhoodName", "tblProjectSettings"), "")
    f.Controls("txtDefaultCollector").Value = Nz(DLookup("DefaultCollectorName", "tblProjectSettings"), "")
    f.Controls("txtContactInfo").Value = Nz(DLookup("SettingValue", "tblSystemSettings", "SettingKey='PROJECT_CONTACT'"), "")
    AT_UI_LoadProjectSettings = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,062A,062D,0645,064A,0644,0020,0625,0639,062F,0627,062F,0627,062A,0020,0627,0644,0645,0634,0631,0648,0639,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_CreateRate() As Boolean
    On Error GoTo EH
    Dim f As Form, rateID As Long
    Set f = Forms("frmRateAdmin")
    rateID = AT_CreateContributionRate(CStr(Nz(f.Controls("txtRateCode").Value, "")), CCur(Nz(f.Controls("txtRateAmount").Value, 0)), AT_UI_DateValue(f.Controls("txtEffectiveFrom").Value))
    f.Controls("txtRateID").Value = rateID
    MsgBox AT_U("062A,0645,0020,062A,0633,062C,064A,0644,0020,0627,0644,0633,0639,0631,0020,0627,0644,062C,062F,064A,062F,002E,0020,0628,0642,064A,062A,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,0627,062A,0020,0627,0644,0642,062F,064A,0645,0629,0020,0639,0644,0649,0020,0623,0633,0639,0627,0631,0647,0627,0020,0627,0644,062A,0627,0631,064A,062E,064A,0629,002E"), vbInformation, "Archon Tech"
    AT_UI_CreateRate = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,062A,0633,062C,064A,0644,0020,0627,0644,0633,0639,0631,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_CreatePeriod() As Boolean
    On Error GoTo EH
    Dim f As Form, periodID As Long
    Set f = Forms("frmFinancialPeriod")
    periodID = AT_CreateFinancialPeriod(CStr(Nz(f.Controls("txtPeriodCode").Value, "")), CStr(Nz(f.Controls("txtPeriodName").Value, "")), AT_UI_DateValue(f.Controls("txtPeriodStart").Value), AT_UI_DateValue(f.Controls("txtPeriodEnd").Value))
    f.Controls("txtPeriodID").Value = periodID
    MsgBox AT_U("062A,0645,0020,0625,0646,0634,0627,0621,0020,0627,0644,0641,062A,0631,0629,0020,0627,0644,0645,0627,0644,064A,0629,0020,0628,0627,0644,0645,0639,0631,0641,0020") & periodID & ".", vbInformation, "Archon Tech"
    AT_UI_CreatePeriod = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0625,0646,0634,0627,0621,0020,0627,0644,0641,062A,0631,0629,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_BeginClosing() As Boolean
    On Error GoTo EH
    Dim f As Form, closingID As Long
    Set f = Forms("frmFinancialClosing")
    closingID = AT_BeginFinancialClosing(CLng(Nz(f.Controls("cboPeriodID").Value, 0)))
    f.Controls("txtClosingID").Value = closingID
    MsgBox AT_U("0628,062F,0623,0020,0627,0644,0625,0642,0641,0627,0644,002E,0020,0623,062F,062E,0644,0020,0627,0644,0631,0635,064A,062F,0020,0627,0644,0641,0639,0644,064A,0020,0644,0643,0644,0020,062D,0633,0627,0628,0020,0645,0646,0020,0634,0627,0634,0629,0020,0625,062F,062E,0627,0644,0020,0627,0644,0623,0631,0635,062F,0629,002E"), vbInformation, "Archon Tech"
    AT_UI_BeginClosing = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0628,062F,0621,0020,0627,0644,0625,0642,0641,0627,0644,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SaveClosingActual() As Boolean
    On Error GoTo EH
    Dim f As Form
    Set f = Forms("frmFinancialClosing")
    AT_SaveClosingActual CLng(Nz(f.Controls("txtClosingID").Value, 0)), CLng(Nz(f.Controls("cboAccountID").Value, 0)), CCur(Nz(f.Controls("txtActualBalance").Value, 0)), CStr(Nz(f.Controls("txtClosingNotes").Value, ""))
    MsgBox AT_U("062A,0645,0020,062D,0641,0638,0020,0627,0644,0631,0635,064A,062F,0020,0627,0644,0641,0639,0644,064A,0020,0644,0644,062D,0633,0627,0628,0020,0641,064A,0020,062A,0641,0627,0635,064A,0644,0020,0627,0644,0625,0642,0641,0627,0644,002E"), vbInformation, "Archon Tech"
    AT_UI_SaveClosingActual = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,062D,0641,0638,0020,0627,0644,0631,0635,064A,062F,0020,0627,0644,0641,0639,0644,064A,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_FinalizeClosing() As Boolean
    On Error GoTo EH
    Dim f As Form
    Set f = Forms("frmFinancialClosing")
    If MsgBox(AT_U("0633,064A,062A,0645,0020,0627,0644,062A,062D,0642,0642,0020,0645,0646,0020,0627,0644,0641,0631,0648,0642,0627,062A,0020,0648,0627,0644,062D,062C,0648,0632,0627,062A,0020,0648,0625,063A,0644,0627,0642,0020,0627,0644,0641,062A,0631,0629,002E,0020,0647,0644,0020,0623,062F,062E,0644,062A,0020,062C,0645,064A,0639,0020,0627,0644,0623,0631,0635,062F,0629,0020,0627,0644,0641,0639,0644,064A,0629,0020,0648,0631,0627,062C,0639,062A,0647,0627,061F"), vbQuestion + vbYesNo + vbDefaultButton2, AT_U("062A,0623,0643,064A,062F,0020,0627,0644,0625,0642,0641,0627,0644,0020,0627,0644,0645,0627,0644,064A")) <> vbYes Then Exit Function
    AT_FinalizeFinancialClosing CLng(Nz(f.Controls("txtClosingID").Value, 0)), CStr(Nz(f.Controls("txtExceptionReason").Value, ""))
    MsgBox AT_U("062A,0645,0020,0625,063A,0644,0627,0642,0020,0627,0644,0641,062A,0631,0629,0020,0628,0639,062F,0020,0627,062C,062A,064A,0627,0632,0020,0636,0648,0627,0628,0637,0020,0627,0644,0625,0642,0641,0627,0644,002E"), vbInformation, "Archon Tech"
    AT_UI_FinalizeClosing = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0625,063A,0644,0627,0642,0020,0627,0644,0641,062A,0631,0629,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ReopenPeriod() As Boolean
    On Error GoTo EH
    Dim f As Form
    Set f = Forms("frmFinancialPeriod")
    If MsgBox(AT_U("0625,0639,0627,062F,0629,0020,0627,0644,0641,062A,062D,0020,062A,0633,0645,062D,0020,0628,062D,0631,0643,0627,062A,0020,062A,0627,0631,064A,062E,064A,0629,0020,0636,0645,0646,0020,0627,0644,0641,062A,0631,0629,002E,0020,064A,062C,0628,0020,0643,062A,0627,0628,0629,0020,0633,0628,0628,0020,0645,0648,062B,0642,002E,0020,0645,062A,0627,0628,0639,0629,061F"), vbQuestion + vbYesNo + vbDefaultButton2, AT_U("062A,0623,0643,064A,062F,0020,0625,0639,0627,062F,0629,0020,0641,062A,062D,0020,0627,0644,0641,062A,0631,0629")) <> vbYes Then Exit Function
    AT_ReopenFinancialPeriod CLng(Nz(f.Controls("cboClosedPeriodID").Value, 0)), CStr(Nz(f.Controls("txtReopenReason").Value, ""))
    MsgBox AT_U("062A,0645,062A,0020,0625,0639,0627,062F,0629,0020,0641,062A,062D,0020,0627,0644,0641,062A,0631,0629,0020,0628,0639,062F,0020,0627,0644,062A,062D,0642,0642,0020,0645,0646,0020,0627,0644,0635,0644,0627,062D,064A,0629,0020,0648,0627,0644,0633,0628,0628,002E"), vbInformation, "Archon Tech"
    AT_UI_ReopenPeriod = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0625,0639,0627,062F,0629,0020,0641,062A,062D,0020,0627,0644,0641,062A,0631,0629,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Private Function AT_UI_DateValue(ByVal v As Variant) As Date
    If IsNull(v) Or Len(Trim$(CStr(v))) = 0 Then AT_UI_DateValue = Date Else AT_UI_DateValue = CDate(v)
End Function

Private Function AT_UI_OptionalDateValue(ByVal v As Variant) As Date
    If IsNull(v) Or Len(Trim$(CStr(v))) = 0 Then AT_UI_OptionalDateValue = 0 Else AT_UI_OptionalDateValue = CDate(v)
End Function

Private Function AT_FormExists(ByVal formName As String) As Boolean
    Dim ao As AccessObject
    For Each ao In CurrentProject.AllForms
        If StrComp(ao.Name, formName, vbTextCompare) = 0 Then AT_FormExists = True: Exit Function
    Next ao
End Function


Private Function AT_FormIsOpen(ByVal formName As String) As Boolean
    Dim f As Form
    For Each f In Forms
        If StrComp(f.Name, formName, vbTextCompare) = 0 Then AT_FormIsOpen = True: Exit Function
    Next f
End Function

Private Function AT_CaptureOpenForms() As Collection
    Dim names As Collection, f As Form
    Set names = New Collection
    For Each f In Forms
        names.Add f.Name
    Next f
    Set AT_CaptureOpenForms = names
End Function

Private Function AT_FormNameWasOpen(ByVal formNames As Collection, ByVal formName As String) As Boolean
    Dim item As Variant
    For Each item In formNames
        If StrComp(CStr(item), formName, vbTextCompare) = 0 Then AT_FormNameWasOpen = True: Exit Function
    Next item
End Function

Private Sub AT_CloseFormsOpenedDuringBuild(ByVal formsOpenBeforeBuild As Collection, ByRef cleanupDetails As String)
    Dim i As Long, formName As String, savedNumber As Long, savedDescription As String
    For i = Forms.Count - 1 To 0 Step -1
        formName = Forms(i).Name
        If Not AT_FormNameWasOpen(formsOpenBeforeBuild, formName) Then
            On Error Resume Next
            Err.Clear
            DoCmd.Close acForm, formName, acSaveNo
            savedNumber = Err.Number: savedDescription = Err.Description
            On Error GoTo 0
            If savedNumber <> 0 Then
                If Len(cleanupDetails) > 0 Then cleanupDetails = cleanupDetails & "; "
                cleanupDetails = cleanupDetails & formName & " (" & savedNumber & "): " & savedDescription
            End If
        End If
    Next i
End Sub

Private Sub AT_SaveAndCloseForm(ByVal temporaryName As String, ByVal desiredName As String)
    Dim savedErrNumber As Long, savedErrSource As String, eDesc As String
    Dim temporaryExistedBefore As Boolean
    Dim finalObjectWasCreated As Boolean
    Dim cleanupNumber As Long, cleanupDescription As String
    On Error GoTo EH

    If Len(Trim$(temporaryName)) = 0 Or Len(Trim$(desiredName)) = 0 Then
        Err.Raise vbObjectError + 748, "AT_SaveAndCloseForm", "A temporary and final form name are required."
    End If
    If StrComp(temporaryName, desiredName, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 749, "AT_SaveAndCloseForm", "The temporary name must differ from the requested final name."
    End If
    If AT_FormExists(desiredName) Then
        Err.Raise vbObjectError + 750, "AT_SaveAndCloseForm", "The destination form already exists; it was not replaced: " & desiredName
    End If
    If Not AT_FormIsOpen(temporaryName) Then
        Err.Raise vbObjectError + 751, "AT_SaveAndCloseForm", "The newly created temporary form is not open: " & temporaryName
    End If
    temporaryExistedBefore = AT_FormExists(temporaryName)
    If temporaryExistedBefore Then
        Err.Raise vbObjectError + 763, "AT_SaveAndCloseForm", "The temporary form name is already persisted; it was not modified: " & temporaryName
    End If

    ' Save the open Design-view form under the name Access itself assigned.
    DoCmd.Save acForm, temporaryName
    If Not AT_FormExists(temporaryName) Then
        Err.Raise vbObjectError + 752, "AT_SaveAndCloseForm", "Access did not persist the temporary form: " & temporaryName
    End If

    ' Rename only after the temporary object has been saved and closed.
    DoCmd.Close acForm, temporaryName, acSaveNo
    If AT_FormIsOpen(temporaryName) Then
        Err.Raise vbObjectError + 753, "AT_SaveAndCloseForm", "The temporary form remained open after close: " & temporaryName
    End If
    DoCmd.Rename desiredName, acForm, temporaryName
    finalObjectWasCreated = True

    If Not AT_FormExists(desiredName) Then
        Err.Raise vbObjectError + 754, "AT_SaveAndCloseForm", "The renamed form is not present in CurrentProject.AllForms: " & desiredName
    End If
    If AT_FormExists(temporaryName) Then
        Err.Raise vbObjectError + 755, "AT_SaveAndCloseForm", "The temporary form name still exists after rename: " & temporaryName
    End If

    ' Reopen the persisted object in hidden Design view to verify later openability,
    ' then close it without saving additional changes.
    DoCmd.OpenForm desiredName, acDesign, , , , acHidden
    If Not AT_FormIsOpen(desiredName) Then
        Err.Raise vbObjectError + 756, "AT_SaveAndCloseForm", "The saved form could not be reopened: " & desiredName
    End If
    DoCmd.Close acForm, desiredName, acSaveNo
    If AT_FormIsOpen(desiredName) Then
        Err.Raise vbObjectError + 757, "AT_SaveAndCloseForm", "The verification window did not close: " & desiredName
    End If
    If Not AT_FormExists(desiredName) Then
        Err.Raise vbObjectError + 758, "AT_SaveAndCloseForm", "The form disappeared after verification: " & desiredName
    End If
    Exit Sub

EH:
    savedErrNumber = Err.Number: savedErrSource = Err.Source: eDesc = Err.Description
    On Error Resume Next
    Err.Clear
    If finalObjectWasCreated Then
        If AT_FormIsOpen(desiredName) Then DoCmd.Close acForm, desiredName, acSaveNo
        If AT_FormExists(desiredName) Then DoCmd.DeleteObject acForm, desiredName
    ElseIf Not temporaryExistedBefore Then
        If AT_FormIsOpen(temporaryName) Then DoCmd.Close acForm, temporaryName, acSaveNo
        If AT_FormExists(temporaryName) Then
            ' This temporary object belongs to the current CreateForm call; never
            ' close or delete a persisted object that existed before this call.
            DoCmd.DeleteObject acForm, temporaryName
        End If
    End If
    cleanupNumber = Err.Number: cleanupDescription = Err.Description
    On Error GoTo 0
    If cleanupNumber <> 0 Then eDesc = eDesc & " | Temporary-form cleanup failed (" & cleanupNumber & "): " & cleanupDescription
    Err.Raise savedErrNumber, "AT_SaveAndCloseForm(" & desiredName & ")", eDesc
End Sub

Private Function AT_RequiredGeneratedFormNames() As Variant
    AT_RequiredGeneratedFormNames = Array( _
        "frmMain", "frmOwnerCreate", "frmAccountCreate", "frmOwnerSearch", _
        "frmOwnerStatementPick", "frmOwnerEdit", "frmChargeCreate", "frmReceiptCreate", _
        "frmPaymentCreate", "frmTransferCreate", "frmAdjustmentCreate", "frmCreditApply", _
        "frmReceiptCorrection", "frmReceiptAllocations", "frmCampaignTargets", "frmUnitOccupancy", _
        "frmProjectSettings", "frmRateAdmin", "frmFinancialPeriod", "frmFinancialClosing", _
        "frmBuildingCreate", "frmUnitCreate", "frmOwnershipTransfer", "frmUnitExemption", _
        "frmExemptionReviewAction", "frmCampaignCreate", "frmExpenseCreate", "frmExpensePay", _
        "frmReceiptReprint", "frmOwnerDirectory", "frmReceiptList", "frmPaymentList", _
        "frmTransferList", "frmOwnerBalances", "frmAccountBalances", "frmAccountMovements", _
        "frmPendingAdjustments", "frmCreditLedger", "frmChargeBalances", "frmOwnerStatementList", _
        "frmBuildingDirectory", "frmUnitDirectory", "frmCampaignDirectory", "frmExpenseDirectory", _
        "frmDraftReceiptAllocations", "frmExemptionReview", "frmRateDirectory", "frmClosingAccounts")
End Function

Private Sub AT_VerifyRequiredGeneratedForms()
    Dim requiredNames As Variant, formName As String, i As Long
    Dim openedHere As Boolean, f As Form, c As Control
    Dim savedErrNumber As Long, savedErrSource As String, eDesc As String
    On Error GoTo EH
    requiredNames = AT_RequiredGeneratedFormNames()
    For i = LBound(requiredNames) To UBound(requiredNames)
        formName = CStr(requiredNames(i))
        If Not AT_FormExists(formName) Then
            Err.Raise vbObjectError + 759, "AT_VerifyRequiredGeneratedForms", "Required form is missing: " & formName
        End If
        openedHere = Not AT_FormIsOpen(formName)
        If openedHere Then
            DoCmd.OpenForm formName, acNormal, , , acFormReadOnly, acHidden
            If Not AT_FormIsOpen(formName) Then
                Err.Raise vbObjectError + 760, "AT_VerifyRequiredGeneratedForms", "Saved form could not be opened: " & formName
            End If
        End If
        Set f = Forms(formName)
        For Each c In f.Controls
            If Len(Trim$(c.Name)) = 0 Then
                Err.Raise vbObjectError + 761, "AT_VerifyRequiredGeneratedForms", "An unnamed control exists in form: " & formName
            End If
        Next c
        If openedHere Then DoCmd.Close acForm, formName, acSaveNo
        If AT_FormIsOpen(formName) And openedHere Then
            Err.Raise vbObjectError + 762, "AT_VerifyRequiredGeneratedForms", "Form remained open after verification: " & formName
        End If
    Next i
    Exit Sub
EH:
    savedErrNumber = Err.Number: savedErrSource = Err.Source: eDesc = Err.Description
    On Error Resume Next
    If openedHere And AT_FormIsOpen(formName) Then DoCmd.Close acForm, formName, acSaveNo
    On Error GoTo 0
    Err.Raise savedErrNumber, savedErrSource, eDesc
End Sub

Private Sub AT_CreateMainForm()
    Dim f As Form, actualName As String, x As Long, y As Long
    If AT_FormExists("frmMain") Then Exit Sub
    Set f = CreateForm: actualName = f.Name
    f.Caption = AT_U("0041,0072,0063,0068,006F,006E,0020,0054,0065,0063,0068,0020,007C,0020,0646,0638,0627,0645,0020,062A,062D,0635,064A,0644,0020,0645,0633,0627,0647,0645,0627,062A,0020,0627,0644,0645,062C,0627,0631,064A"): f.Width = 11800
    f.Section(acDetail).Height = 27000: f.Section(acDetail).BackColor = RGB(248, 250, 252)
    On Error Resume Next: f.RightToLeft = True: On Error GoTo 0
    AT_AddLabel actualName, "lblTitle", AT_U("0646,0638,0627,0645,0020,0625,062F,0627,0631,0629,0020,0645,0633,0627,0647,0645,0627,062A,0020,0645,0634,0631,0648,0639,0020,0627,0644,0645,062C,0627,0631,064A"), 450, 300, 10500, 500, 18, True
    AT_AddLabel actualName, "lblSubtitle", AT_U("064A,0628,062F,0623,0020,0643,0644,0020,0645,0644,0641,0020,062D,0627,0631,0629,0020,0645,0633,062A,0642,0644,064B,0627,0020,0645,0646,0020,0627,0644,0635,0641,0631,061B,0020,0643,0644,0020,0627,0639,062A,0645,0627,062F,0020,0645,0627,0644,064A,0020,064A,062A,0637,0644,0628,0020,062A,0623,0643,064A,062F,064B,0627,0020,0635,0631,064A,062D,064B,0627,002E"), 500, 850, 10500, 350, 10, False
    x = 650: y = 1500
    AT_AddButton actualName, "cmdOwnerCreate", AT_U("062A,0633,062C,064A,0644,0020,0645,0627,0644,0643,0020,062C,062F,064A,062F"), x, y, 4600, 650, "=AT_UI_OpenForm(""frmOwnerCreate"")"
    AT_AddButton actualName, "cmdOwnerDirectory", AT_U("0633,062C,0644,0020,0627,0644,0645,0644,0627,0643"), 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmOwnerDirectory"")"
    y = y + 850
    AT_AddButton actualName, "cmdChargeCreate", AT_U("0625,0646,0634,0627,0621,0020,0627,0633,062A,062D,0642,0627,0642"), x, y, 4600, 650, "=AT_UI_OpenForm(""frmChargeCreate"")"
    AT_AddButton actualName, "cmdReceiptCreate", AT_U("0642,0628,0636,0020,002F,0020,0625,064A,0635,0627,0644,0020,062C,062F,064A,062F"), 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmReceiptCreate"")"
    y = y + 850
    AT_AddButton actualName, "cmdReceiptList", AT_U("0633,062C,0644,0020,0627,0644,0625,064A,0635,0627,0644,0627,062A"), x, y, 4600, 650, "=AT_UI_OpenForm(""frmReceiptList"")"
    AT_AddButton actualName, "cmdReceiptAllocations", AT_U("062A,062E,0635,064A,0635,0020,0642,0628,0636,0020,064A,062F,0648,064A"), 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmReceiptAllocations"")"
    y = y + 850
    AT_AddButton actualName, "cmdPaymentCreate", AT_U("062A,0633,062C,064A,0644,0020,0645,0635,0631,0648,0641,0020,002F,0020,0635,0631,0641"), 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmExpenseCreate"")"
    y = y + 850
    AT_AddButton actualName, "cmdPaymentList", AT_U("0633,062C,0644,0020,0633,0646,062F,0627,062A,0020,0627,0644,0635,0631,0641"), x, y, 4600, 650, "=AT_UI_OpenForm(""frmPaymentList"")"
    AT_AddButton actualName, "cmdTransferCreate", AT_U("0627,0644,062A,062D,0648,064A,0644,0627,062A,0020,0628,064A,0646,0020,0627,0644,062D,0633,0627,0628,0627,062A"), 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmTransferCreate"")"
    y = y + 850
    AT_AddButton actualName, "cmdAdjustment", AT_U("0637,0644,0628,0627,062A,0020,0627,0644,062E,0635,0645,0020,0648,0627,0644,0625,0639,0641,0627,0621"), x, y, 4600, 650, "=AT_UI_OpenForm(""frmAdjustmentCreate"")"
    AT_AddButton actualName, "cmdTransferList", AT_U("0633,062C,0644,0020,0627,0644,062A,062D,0648,064A,0644,0627,062A"), 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmTransferList"")"
    y = y + 850
    AT_AddButton actualName, "cmdIntegrity", AT_U("0641,062D,0635,0020,0633,0644,0627,0645,0629,0020,0627,0644,0628,064A,0627,0646,0627,062A"), x, y, 4600, 650, "=AT_UI_RunIntegrity()"
    AT_AddButton actualName, "cmdOwnerBalances", AT_U("0643,0634,0641,0020,0623,0631,0635,062F,0629,0020,0627,0644,0645,0644,0627,0643,0020,0648,0627,0644,0645,062A,0623,062E,0631,0627,062A"), 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmOwnerBalances"")"
    y = y + 850
    AT_AddButton actualName, "cmdAccountBalances", AT_U("0623,0631,0635,062F,0629,0020,0627,0644,062D,0633,0627,0628,0627,062A"), x, y, 4600, 650, "=AT_UI_OpenForm(""frmAccountBalances"")"
    AT_AddButton actualName, "cmdAccountMovements", AT_U("062D,0631,0643,0629,0020,0627,0644,062D,0633,0627,0628,0627,062A"), 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmAccountMovements"")"
    y = y + 850
    AT_AddButton actualName, "cmdPendingAdjustments", AT_U("0627,0644,062E,0635,0648,0645,0627,062A,0020,0627,0644,0645,0639,0644,0642,0629"), x, y, 4600, 650, "=AT_UI_OpenForm(""frmPendingAdjustments"")"
    AT_AddButton actualName, "cmdOwnerEdit", AT_U("062A,0639,062F,064A,0644,0020,0628,064A,0627,0646,0627,062A,0020,0645,0627,0644,0643"), 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmOwnerEdit"")"
    y = y + 850
    AT_AddButton actualName, "cmdCreditApply", AT_U("062A,0637,0628,064A,0642,0020,0631,0635,064A,062F,0020,062F,0627,0626,0646,0020,0639,0644,0649,0020,0627,0633,062A,062D,0642,0627,0642"), x, y, 4600, 650, "=AT_UI_OpenForm(""frmCreditApply"")"
    AT_AddButton actualName, "cmdCreditLedger", AT_U("062F,0641,062A,0631,0020,0627,0644,0623,0631,0635,062F,0629,0020,0627,0644,062F,0627,0626,0646,0629"), 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmCreditLedger"")"
    y = y + 850
    AT_AddButton actualName, "cmdReverseAllocation", AT_U("0639,0643,0633,0020,062A,062E,0635,064A,0635,0020,0625,064A,0635,0627,0644"), x, y, 4600, 650, "=AT_UI_OpenForm(""frmReceiptCorrection"")"
    AT_AddButton actualName, "cmdChargeBalances", AT_U("0643,0634,0641,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,0627,062A"), 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmChargeBalances"")"
    y = y + 850
    AT_AddButton actualName, "cmdOwnerSearch", AT_U("0628,062D,062B,0020,0639,0646,0020,0645,0627,0644,0643"), x, y, 4600, 650, "=AT_UI_OpenForm(""frmOwnerSearch"")"
    AT_AddButton actualName, "cmdOwnerStatement", AT_U("0643,0634,0641,0020,062D,0633,0627,0628,0020,0645,0627,0644,0643"), 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmOwnerStatementPick"")"
    y = y + 850
    AT_AddButton actualName, "cmdAccountCreate", AT_U("0625,0636,0627,0641,0629,0020,0635,0646,062F,0648,0642,0020,002F,0020,0628,0646,0643,0020,002F,0020,0645,062D,0641,0638,0629"), x, y, 4600, 650, "=AT_UI_OpenForm(""frmAccountCreate"")"
    AT_AddButton actualName, "cmdAccountBalancesSecond", AT_U("0627,0633,062A,0639,0631,0627,0636,0020,0623,0631,0635,062F,0629,0020,0627,0644,062D,0633,0627,0628,0627,062A"), 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmAccountBalances"")"
    y = y + 850
    AT_AddButton actualName, "cmdBuildingCreate", AT_U("062A,0633,062C,064A,0644,0020,0645,0628,0646,0649"), x, y, 4600, 650, "=AT_UI_OpenForm(""frmBuildingCreate"")"
    AT_AddButton actualName, "cmdBuildingList", AT_U("062F,0644,064A,0644,0020,0627,0644,0645,0628,0627,0646,064A,0020,0648,0627,0644,0634,0642,0642"), 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmBuildingDirectory"")"
    y = y + 850
    AT_AddButton actualName, "cmdUnitCreate", AT_U("062A,0633,062C,064A,0644,0020,0634,0642,0629"), x, y, 4600, 650, "=AT_UI_OpenForm(""frmUnitCreate"")"
    AT_AddButton actualName, "cmdUnitOccupancy", AT_U("062A,0633,062C,064A,0644,0020,0627,0644,0623,0633,0631,0629,0020,002F,0020,0627,0644,0625,0634,063A,0627,0644"), 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmUnitOccupancy"")"
    y = y + 850
    AT_AddButton actualName, "cmdOwnershipTransfer", AT_U("0646,0642,0644,0020,0645,0644,0643,064A,0629,0020,0645,0628,0646,0649"), x, y, 4600, 650, "=AT_UI_OpenForm(""frmOwnershipTransfer"")"
    AT_AddButton actualName, "cmdCampaignCreate", AT_U("0625,0639,062F,0627,062F,0020,062D,0645,0644,0629,0020,0625,0636,0627,0641,064A,0629"), 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmCampaignCreate"")"
    y = y + 850
    AT_AddButton actualName, "cmdCampaignTargets", AT_U("062A,062D,062F,064A,062F,0020,0623,0647,062F,0627,0641,0020,0627,0644,062D,0645,0644,0629"), x, y, 4600, 650, "=AT_UI_OpenForm(""frmCampaignTargets"")"
    AT_AddButton actualName, "cmdCampaignList", AT_U("0633,062C,0644,0020,0627,0644,062D,0645,0644,0627,062A"), 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmCampaignDirectory"")"
    y = y + 850
    AT_AddButton actualName, "cmdExemptionRequest", AT_U("0637,0644,0628,0020,0625,0639,0641,0627,0621,0020,0634,0642,0629"), x, y, 4600, 650, "=AT_UI_OpenForm(""frmUnitExemption"")"
    AT_AddButton actualName, "cmdExemptionReview", AT_U("0645,0631,0627,062C,0639,0629,0020,0637,0644,0628,0627,062A,0020,0627,0644,0625,0639,0641,0627,0621"), 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmExemptionReview"")"
    y = y + 850
    AT_AddButton actualName, "cmdExpenseCreate", AT_U("062A,0633,062C,064A,0644,0020,0645,0635,0631,0648,0641"), x, y, 4600, 650, "=AT_UI_OpenForm(""frmExpenseCreate"")"
    AT_AddButton actualName, "cmdExpensePay", AT_U("0633,062F,0627,062F,0020,0645,0635,0631,0648,0641,0020,0645,0633,062C,0644"), 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmExpensePay"")"
    y = y + 850
    AT_AddButton actualName, "cmdExpenseList", AT_U("0633,062C,0644,0020,0627,0644,0645,0635,0631,0648,0641,0627,062A"), x, y, 4600, 650, "=AT_UI_OpenForm(""frmExpenseDirectory"")"
    AT_AddButton actualName, "cmdPrintReceipt", AT_U("0625,0639,0627,062F,0629,0020,0637,0628,0627,0639,0629,0020,0625,064A,0635,0627,0644"), 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmReceiptReprint"")"
    y = y + 850
    AT_AddButton actualName, "cmdProjectSettings", AT_U("0625,0639,062F,0627,062F,0627,062A,0020,0627,0644,0645,0634,0631,0648,0639,0020,0648,0627,0644,062D,0627,0631,0629"), x, y, 4600, 650, "=AT_UI_OpenForm(""frmProjectSettings"")"
    AT_AddButton actualName, "cmdRateAdmin", AT_U("0625,062F,0627,0631,0629,0020,0623,0633,0639,0627,0631,0020,0627,0644,0645,0633,0627,0647,0645,0629"), 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmRateAdmin"")"
    y = y + 850
    AT_AddButton actualName, "cmdFinancialPeriod", AT_U("0627,0644,0641,062A,0631,0627,062A,0020,0627,0644,0645,0627,0644,064A,0629"), x, y, 4600, 650, "=AT_UI_OpenForm(""frmFinancialPeriod"")"
    AT_AddButton actualName, "cmdFinancialClosing", AT_U("0627,0644,0625,0642,0641,0627,0644,0020,0648,0645,0637,0627,0628,0642,0629,0020,0627,0644,0623,0631,0635,062F,0629"), 5600, y, 4600, 650, "=AT_UI_OpenForm(""frmFinancialClosing"")"
    AT_AddLabel actualName, "lblFooter", AT_U("062A,0646,0628,064A,0647,003A,0020,064A,0644,0632,0645,0020,062A,062C,0645,064A,0639,0020,0056,0042,0041,0020,0648,0627,062E,062A,0628,0627,0631,0020,0627,0644,0646,0638,0627,0645,0020,0648,0627,0644,0637,0628,0627,0639,0629,0020,0648,0627,0644,0646,0633,062E,0020,0627,0644,0627,062D,062A,064A,0627,0637,064A,0020,062F,0627,062E,0644,0020,0041,0063,0063,0065,0073,0073,0020,0642,0628,0644,0020,062A,062D,0635,064A,0644,0020,0623,0645,0648,0627,0644,0020,062D,0642,064A,0642,064A,0629,002E"), 500, y + 850, 10500, 400, 9, False
    AT_SaveAndCloseForm actualName, "frmMain"
End Sub

Private Sub AT_CreateOwnerEntryForm()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmOwnerCreate") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("0625,0636,0627,0641,0629,0020,0645,0627,0644,0643,0020,0648,0627,0639,062A,0645,0627,062F,0647"), 8300
    y = 500
    AT_AddFieldPair n, "txtFullName", AT_U("0627,0633,0645,0020,0627,0644,0645,0627,0644,0643"), y, 6500: y = y + 650
    AT_AddFieldPair n, "txtPhone", AT_U("0631,0642,0645,0020,0627,0644,0647,0627,062A,0641"), y, 6500: y = y + 650
    AT_AddFieldPair n, "txtIdentityNo", AT_U("0631,0642,0645,0020,0627,0644,0647,0648,064A,0629,0020,0028,0627,062E,062A,064A,0627,0631,064A,0029"), y, 6500: y = y + 650
    AT_AddFieldPair n, "txtNotes", AT_U("0645,0644,0627,062D,0638,0627,062A"), y, 6500: y = y + 650
    AT_AddFieldPair n, "txtOwnerID", AT_U("0627,0644,0645,0639,0631,0641,0020,0627,0644,062F,0627,062E,0644,064A,0020,0644,0644,0645,0633,0648,062F,0629"), y, 6500: f.Controls("txtOwnerID").Locked = True: y = y + 850
    AT_AddButton n, "cmdSaveOwner", AT_U("062D,0641,0638,0020,0643,0645,0633,0648,062F,0629"), 3000, y, 2200, 600, "=AT_UI_SaveOwnerDraft()"
    AT_AddButton n, "cmdApproveOwner", AT_U("0645,0631,0627,062C,0639,0629,0020,0648,0627,0639,062A,0645,0627,062F"), 5500, y, 2500, 600, "=AT_UI_ApproveOwner()"
    AT_SaveAndCloseForm n, "frmOwnerCreate"
End Sub

Private Sub AT_CreateAccountEntryForm()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmAccountCreate") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("0625,0636,0627,0641,0629,0020,062D,0633,0627,0628,0020,0645,0627,0644,064A,0020,062F,0648,0646,0020,0631,0635,064A,062F,0020,0627,0641,062A,062A,0627,062D,064A"), 5200: y = 650
    AT_AddFieldPair n, "txtAccountCode", AT_U("0631,0645,0632,0020,0627,0644,062D,0633,0627,0628"), y, 6500: y = y + 750
    AT_AddFieldPair n, "txtAccountName", AT_U("0627,0633,0645,0020,0627,0644,062D,0633,0627,0628"), y, 6500: y = y + 750
    Call AT_AddCombo(n, "cboAccountType", AT_U("0646,0648,0639,0020,0627,0644,062D,0633,0627,0628"), y, AT_U("0053,0045,004C,0045,0043,0054,0020,0027,0043,0041,0053,0048,0027,0020,0041,0053,0020,0054,0079,0070,0065,0043,006F,0064,0065,002C,0020,0027,0635,0646,062F,0648,0642,0020,0646,0642,062F,064A,0027,0020,0041,0053,0020,0054,0079,0070,0065,004E,0061,006D,0065,0020,0055,004E,0049,004F,004E,0020,0053,0045,004C,0045,0043,0054,0020,0027,0042,0041,004E,004B,0027,002C,0027,062D,0633,0627,0628,0020,0628,0646,0643,064A,0027,0020,0055,004E,0049,004F,004E,0020,0053,0045,004C,0045,0043,0054,0020,0027,0045,0057,0041,004C,004C,0045,0054,0027,002C,0027,0645,062D,0641,0638,0629,0020,0625,0644,0643,062A,0631,0648,0646,064A,0629,0027,0020,0055,004E,0049,004F,004E,0020,0053,0045,004C,0045,0043,0054,0020,0027,004F,0054,0048,0045,0052,0027,002C,0027,0623,062E,0631,0649,0027,003B"), 2, "0;4")
    y = y + 750
    AT_AddFieldPair n, "txtAccountID", AT_U("0627,0644,0645,0639,0631,0641,0020,0627,0644,062F,0627,062E,0644,064A"), y, 6500: f.Controls("txtAccountID").Locked = True: y = y + 900
    AT_AddButton n, "cmdSaveAccount", AT_U("0625,0646,0634,0627,0621,0020,0627,0644,062D,0633,0627,0628,0020,0628,0631,0635,064A,062F,0020,0635,0641,0631"), 3600, y, 3400, 650, "=AT_UI_SaveAccount()"
    AT_SaveAndCloseForm n, "frmAccountCreate"
End Sub

Private Sub AT_CreateOwnerSearchForm()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmOwnerSearch") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("0628,062D,062B,0020,0639,0646,0020,0645,0627,0644,0643"), 3600: y = 650
    AT_AddFieldPair n, "txtSearchTerm", AT_U("0627,0644,0627,0633,0645,0020,0623,0648,0020,0627,0644,0631,0642,0645,0020,0623,0648,0020,0627,0644,0647,0627,062A,0641"), y, 6500: y = y + 1000
    AT_AddButton n, "cmdSearchOwners", AT_U("0628,062D,062B"), 4200, y, 2200, 650, "=AT_UI_SearchOwners()"
    AT_SaveAndCloseForm n, "frmOwnerSearch"
End Sub

Private Sub AT_CreateOwnerStatementPicker()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmOwnerStatementPick") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("0627,062E,062A,064A,0627,0631,0020,0645,0627,0644,0643,0020,0644,0639,0631,0636,0020,0643,0634,0641,0020,0627,0644,062D,0633,0627,0628"), 3600: y = 650
    Call AT_AddCombo(n, "cboOwnerID", AT_U("0627,0644,0645,0627,0644,0643"), y, "SELECT OwnerID, OwnerCode & ' - ' & FullName AS OwnerLabel FROM tblOwners WHERE AdminStatus='ACTIVE' ORDER BY FullName;", 2, "0;5")
    y = y + 1000
    AT_AddButton n, "cmdOpenOwnerStatement", AT_U("0639,0631,0636,0020,0643,0634,0641,0020,0627,0644,062D,0633,0627,0628"), 3900, y, 2800, 650, "=AT_UI_OpenOwnerStatement()"
    AT_SaveAndCloseForm n, "frmOwnerStatementPick"
End Sub

Private Sub AT_CreateOwnerEditForm()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmOwnerEdit") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("062A,0639,062F,064A,0644,0020,0628,064A,0627,0646,0627,062A,0020,0627,0644,0645,0627,0644,0643,0020,0645,0639,0020,0633,062C,0644,0020,0627,0644,062A,062F,0642,064A,0642"), 8500: y = 450
    Call AT_AddCombo(n, "cboOwnerID", AT_U("0627,0644,0645,0627,0644,0643"), y, "SELECT OwnerID, OwnerCode & ' - ' & FullName AS OwnerLabel FROM tblOwners WHERE AdminStatus='ACTIVE' ORDER BY FullName;", 2, "0;4"): y = y + 650
    AT_AddFieldPair n, "txtFullName", AT_U("0627,0633,0645,0020,0627,0644,0645,0627,0644,0643"), y, 6500: y = y + 650
    AT_AddFieldPair n, "txtPhone", AT_U("0631,0642,0645,0020,0627,0644,0647,0627,062A,0641"), y, 6500: y = y + 650
    AT_AddFieldPair n, "txtIdentityNo", AT_U("0631,0642,0645,0020,0627,0644,0647,0648,064A,0629,0020,0028,0627,062E,062A,064A,0627,0631,064A,0029"), y, 6500: y = y + 650
    AT_AddFieldPair n, "txtNotes", AT_U("0645,0644,0627,062D,0638,0627,062A"), y, 6500: y = y + 650
    AT_AddFieldPair n, "txtEditReason", AT_U("0633,0628,0628,0020,0627,0644,062A,0639,062F,064A,0644,0020,0028,0625,0644,0632,0627,0645,064A,0029"), y, 6500: y = y + 850
    AT_AddButton n, "cmdLoadOwner", AT_U("062A,062D,0645,064A,0644,0020,0628,064A,0627,0646,0627,062A,0020,0627,0644,0645,0627,0644,0643"), 2800, y, 2500, 600, "=AT_UI_LoadOwnerEdit()"
    AT_AddButton n, "cmdSaveOwnerEdit", AT_U("062D,0641,0638,0020,0627,0644,062A,0639,062F,064A,0644"), 5600, y, 2200, 600, "=AT_UI_SaveOwnerEdit()"
    AT_SaveAndCloseForm n, "frmOwnerEdit"
End Sub

Private Sub AT_CreateChargeEntryForm()
    Dim f As Form, n As String, y As Long, c As Control
    If AT_FormExists("frmChargeCreate") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("0625,0646,0634,0627,0621,0020,0627,0633,062A,062D,0642,0627,0642,0020,0648,0627,0639,062A,0645,0627,062F,0647"), 7600: y = 500
    Set c = AT_AddCombo(n, "cboOwnerID", AT_U("0627,0644,0645,0627,0644,0643"), y, "SELECT OwnerID, OwnerCode & ' - ' & FullName AS OwnerLabel FROM tblOwners WHERE AdminStatus='ACTIVE' ORDER BY FullName;", 2, "0;4")
    y = y + 650
    Set c = AT_AddCombo(n, "cboRateID", AT_U("0633,0639,0631,0020,0627,0644,0645,0633,0627,0647,0645,0629"), y, "SELECT RateID, RateCode, RateAmount FROM tblContributionRates WHERE IsActive=True ORDER BY EffectiveFrom DESC;", 3, "0;2;1.5"): y = y + 650
    AT_AddFieldPair n, "txtChargeDate", AT_U("062A,0627,0631,064A,062E,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642"), y, 6500: y = y + 650
    AT_AddFieldPair n, "txtReason", AT_U("0627,0644,0628,064A,0627,0646,0020,002F,0020,0627,0644,0633,0628,0628"), y, 6500: y = y + 650
    AT_AddFieldPair n, "txtChargeID", AT_U("0645,0639,0631,0641,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642"), y, 6500: f.Controls("txtChargeID").Locked = True: y = y + 850
    AT_AddButton n, "cmdSaveCharge", AT_U("062D,0641,0638,0020,0643,0645,0633,0648,062F,0629"), 3000, y, 2200, 600, "=AT_UI_SaveChargeDraft()"
    AT_AddButton n, "cmdApproveCharge", AT_U("0627,0639,062A,0645,0627,062F,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642"), 5500, y, 2500, 600, "=AT_UI_ApproveCharge()"
    AT_SaveAndCloseForm n, "frmChargeCreate"
End Sub

Private Sub AT_CreateReceiptEntryForm()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmReceiptCreate") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("062A,0633,062C,064A,0644,0020,0642,0628,0636,0020,0648,0625,0635,062F,0627,0631,0020,0625,064A,0635,0627,0644"), 10200: y = 350
    Call AT_AddCombo(n, "cboOwnerID", AT_U("0627,0644,0645,0627,0644,0643"), y, "SELECT OwnerID, OwnerCode & ' - ' & FullName AS OwnerLabel FROM tblOwners WHERE AdminStatus='ACTIVE' ORDER BY FullName;", 2, "0;4"): y = y + 550
    AT_AddFieldPair n, "txtAmount", AT_U("0627,0644,0645,0628,0644,063A,0020,0628,0627,0644,0631,064A,0627,0644"), y, 6500: y = y + 550
    Call AT_AddCombo(n, "cboPaymentMethodID", AT_U("0637,0631,064A,0642,0629,0020,0627,0644,062F,0641,0639"), y, "SELECT PaymentMethodID, MethodName FROM tblPaymentMethods WHERE IsActive=True ORDER BY MethodName;", 2, "0;4"): y = y + 550
    Call AT_AddCombo(n, "cboAccountID", AT_U("0627,0644,062D,0633,0627,0628,0020,0627,0644,0630,064A,0020,0627,0633,062A,0644,0645,0020,0627,0644,0645,0627,0644"), y, "SELECT AccountID, AccountCode & ' - ' & AccountName FROM tblAccounts WHERE IsActive=True ORDER BY AccountName;", 2, "0;4"): y = y + 550
    AT_AddFieldPair n, "txtPayerName", AT_U("0627,0633,0645,0020,0627,0644,062F,0627,0641,0639,0020,0028,0627,062E,062A,064A,0627,0631,064A,0029"), y, 6500: y = y + 550
    AT_AddFieldPair n, "txtCollectorName", AT_U("0627,0633,0645,0020,0627,0644,0645,062D,0635,0644"), y, 6500: y = y + 550
    AT_AddFieldPair n, "txtTransactionDate", AT_U("062A,0627,0631,064A,062E,0020,0627,0644,062D,0631,0643,0629"), y, 6500: y = y + 550
    AT_AddFieldPair n, "txtNotes", AT_U("0645,0644,0627,062D,0638,0627,062A"), y, 6500: y = y + 550
    AT_AddFieldPair n, "txtReceiptID", AT_U("0627,0644,0645,0639,0631,0641,0020,0627,0644,062F,0627,062E,0644,064A,0020,0644,0644,0625,064A,0635,0627,0644"), y, 6500: f.Controls("txtReceiptID").Locked = True: y = y + 800
    AT_AddButton n, "cmdSaveReceipt", AT_U("062D,0641,0638,0020,0643,0645,0633,0648,062F,0629"), 2500, y, 2200, 600, "=AT_UI_SaveReceiptDraft()"
    AT_AddButton n, "cmdApproveReceipt", AT_U("062A,0623,0643,064A,062F,0020,0627,0644,0642,0628,0636,0020,0648,0627,0639,062A,0645,0627,062F"), 5000, y, 3000, 600, "=AT_UI_ApproveReceipt()"
    AT_AddButton n, "cmdAttachReceipt", AT_U("0625,0631,0641,0627,0642,0020,0645,0644,0641"), 8000, y, 2100, 600, "=AT_UI_AttachReceiptFile()"
    AT_SaveAndCloseForm n, "frmReceiptCreate"
End Sub

Private Sub AT_CreatePaymentEntryForm()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmPaymentCreate") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("0625,0646,0634,0627,0621,0020,0633,0646,062F,0020,0635,0631,0641,0020,002F,0020,0631,062F,0020,0631,0635,064A,062F,0020,062F,0627,0626,0646"), 11200: y = 350
    Call AT_AddCombo(n, "cboAccountID", AT_U("062D,0633,0627,0628,0020,0627,0644,0635,0631,0641"), y, "SELECT AccountID, AccountCode & ' - ' & AccountName FROM tblAccounts WHERE IsActive=True ORDER BY AccountName;", 2, "0;4"): y = y + 550
    Call AT_AddCombo(n, "cboExpenseCategoryID", AT_U("0641,0626,0629,0020,0627,0644,0645,0635,0631,0648,0641"), y, "SELECT ExpenseCategoryID, CategoryName FROM tblExpenseCategories WHERE IsActive=True ORDER BY CategoryName;", 2, "0;4"): y = y + 550
    AT_AddFieldPair n, "txtAmount", AT_U("0627,0644,0645,0628,0644,063A,0020,0628,0627,0644,0631,064A,0627,0644"), y, 6500: y = y + 550
    Call AT_AddCombo(n, "cboPaymentMethodID", AT_U("0637,0631,064A,0642,0629,0020,0627,0644,062F,0641,0639"), y, "SELECT PaymentMethodID, MethodName FROM tblPaymentMethods WHERE IsActive=True ORDER BY MethodName;", 2, "0;4"): y = y + 550
    AT_AddFieldPair n, "txtBeneficiary", AT_U("0627,0644,0645,0633,062A,0641,064A,062F"), y, 6500: y = y + 550
    AT_AddFieldPair n, "txtRecipient", AT_U("0627,0633,0645,0020,0627,0644,0645,0633,062A,0644,0645"), y, 6500: y = y + 550
    AT_AddFieldPair n, "txtTransactionDate", AT_U("062A,0627,0631,064A,062E,0020,0627,0644,062D,0631,0643,0629"), y, 6500: y = y + 550
    AT_AddFieldPair n, "txtNotes", AT_U("0645,0644,0627,062D,0638,0627,062A"), y, 6500: y = y + 550
    AT_AddFieldPair n, "txtPaymentID", AT_U("0627,0644,0645,0639,0631,0641,0020,0627,0644,062F,0627,062E,0644,064A,0020,0644,0644,0633,0646,062F"), y, 6500: f.Controls("txtPaymentID").Locked = True: y = y + 650
    AT_AddFieldPair n, "txtRefundCreditMovementID", AT_U("0645,0639,0631,0641,0020,062D,0631,0643,0629,0020,0627,0644,062F,0627,0626,0646,0020,0644,0644,0631,062F,0020,0028,0627,062E,062A,064A,0627,0631,064A,0029"), y, 6500: y = y + 800
    AT_AddButton n, "cmdSavePayment", AT_U("062D,0641,0638,0020,0643,0645,0633,0648,062F,0629"), 2500, y, 2200, 600, "=AT_UI_SavePaymentDraft()"
    AT_AddButton n, "cmdApprovePayment", AT_U("062A,0623,0643,064A,062F,0020,0627,0644,0635,0631,0641,0020,0648,0627,0639,062A,0645,0627,062F"), 5000, y, 3000, 600, "=AT_UI_ApprovePayment()"
    AT_AddButton n, "cmdAttachPayment", AT_U("0625,0631,0641,0627,0642,0020,0645,0644,0641"), 8100, y, 2200, 600, "=AT_UI_AttachPaymentFile()"
    AT_SaveAndCloseForm n, "frmPaymentCreate"
End Sub

Private Sub AT_CreateTransferEntryForm()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmTransferCreate") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("062A,062D,0648,064A,0644,0020,0628,064A,0646,0020,0627,0644,062D,0633,0627,0628,0627,062A"), 9300: y = 350
    Call AT_AddCombo(n, "cboSourceAccountID", AT_U("062D,0633,0627,0628,0020,0627,0644,0645,0635,062F,0631"), y, "SELECT AccountID, AccountCode & ' - ' & AccountName FROM tblAccounts WHERE IsActive=True ORDER BY AccountName;", 2, "0;4"): y = y + 550
    Call AT_AddCombo(n, "cboDestinationAccountID", AT_U("062D,0633,0627,0628,0020,0627,0644,0648,062C,0647,0629"), y, "SELECT AccountID, AccountCode & ' - ' & AccountName FROM tblAccounts WHERE IsActive=True ORDER BY AccountName;", 2, "0;4"): y = y + 550
    AT_AddFieldPair n, "txtAmount", AT_U("0645,0628,0644,063A,0020,0627,0644,062A,062D,0648,064A,0644"), y, 6500: y = y + 550
    AT_AddFieldPair n, "txtFeeAmount", AT_U("0631,0633,0648,0645,0020,0627,0644,062A,062D,0648,064A,0644"), y, 6500: y = y + 550
    AT_AddFieldPair n, "txtTransactionDate", AT_U("062A,0627,0631,064A,062E,0020,0627,0644,062D,0631,0643,0629"), y, 6500: y = y + 550
    AT_AddFieldPair n, "txtNotes", AT_U("0645,0644,0627,062D,0638,0627,062A"), y, 6500: y = y + 550
    AT_AddFieldPair n, "txtTransferID", AT_U("0627,0644,0645,0639,0631,0641,0020,0627,0644,062F,0627,062E,0644,064A,0020,0644,0644,062A,062D,0648,064A,0644"), y, 6500: f.Controls("txtTransferID").Locked = True: y = y + 550
    AT_AddFieldPair n, "txtTransferRef", AT_U("0645,0631,062C,0639,0020,062A,0623,0643,064A,062F,0020,0627,0644,0648,0635,0648,0644"), y, 6500: y = y + 750
    AT_AddButton n, "cmdSaveTransfer", AT_U("062D,0641,0638,0020,0643,0645,0633,0648,062F,0629"), 800, y, 1900, 550, "=AT_UI_SaveTransferDraft()"
    AT_AddButton n, "cmdApproveTransfer", AT_U("0627,0639,062A,0645,0627,062F,0020,0648,062D,062C,0632"), 2850, y, 1900, 550, "=AT_UI_ApproveTransfer()"
    AT_AddButton n, "cmdConfirmTransfer", AT_U("062A,0623,0643,064A,062F,0020,0627,0644,0648,0635,0648,0644"), 4900, y, 2200, 550, "=AT_UI_ConfirmTransfer()"
    AT_SaveAndCloseForm n, "frmTransferCreate"
End Sub

Private Sub AT_CreateAdjustmentEntryForm()
    Dim f As Form, n As String, y As Long, c As Control
    If AT_FormExists("frmAdjustmentCreate") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("0637,0644,0628,0020,062E,0635,0645,0020,0623,0648,0020,0625,0639,0641,0627,0621"), 6500: y = 500
    AT_AddFieldPair n, "txtChargeID", AT_U("0645,0639,0631,0641,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642"), y, 6500: y = y + 650
    AT_AddFieldPair n, "txtDiscountAmount", AT_U("0645,0628,0644,063A,0020,0627,0644,062E,0635,0645,0020,0028,0645,0648,062C,0628,0020,0641,064A,0020,0627,0644,0625,062F,062E,0627,0644,0029"), y, 6500: y = y + 650
    Set c = AT_AddCombo(n, "cboAdjustmentType", AT_U("0646,0648,0639,0020,0627,0644,062A,0639,062F,064A,0644"), y, AT_U("0053,0045,004C,0045,0043,0054,0020,0027,0044,0049,0053,0043,004F,0055,004E,0054,0027,0020,0041,0053,0020,0054,0079,0070,0065,0043,006F,0064,0065,002C,0020,0027,062E,0635,0645,0027,0020,0041,0053,0020,0054,0079,0070,0065,004E,0061,006D,0065,0020,0055,004E,0049,004F,004E,0020,0053,0045,004C,0045,0043,0054,0020,0027,0045,0058,0045,004D,0050,0054,0049,004F,004E,0027,002C,0027,0625,0639,0641,0627,0621,0027,003B"), 2, "0;4"): y = y + 650
    AT_AddFieldPair n, "txtReason", AT_U("0633,0628,0628,0020,0627,0644,0637,0644,0628"), y, 6500: y = y + 650
    AT_AddFieldPair n, "txtAdjustmentID", AT_U("0645,0639,0631,0641,0020,0627,0644,0637,0644,0628"), y, 6500: f.Controls("txtAdjustmentID").Locked = True: y = y + 850
    AT_AddButton n, "cmdRequestAdjustment", AT_U("062D,0641,0638,0020,0627,0644,0637,0644,0628"), 2500, y, 2200, 600, "=AT_UI_RequestAdjustment()"
    AT_AddButton n, "cmdApproveAdjustment", AT_U("0627,0639,062A,0645,0627,062F,0020,0627,0644,0637,0644,0628"), 5000, y, 2200, 600, "=AT_UI_ApproveAdjustment()"
    AT_SaveAndCloseForm n, "frmAdjustmentCreate"
End Sub

Private Sub AT_CreateCreditApplicationForm()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmCreditApply") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("062A,0637,0628,064A,0642,0020,0631,0635,064A,062F,0020,062F,0627,0626,0646,0020,0639,0644,0649,0020,0627,0633,062A,062D,0642,0627,0642"), 6500: y = 450
    Call AT_AddCombo(n, "cboOwnerID", AT_U("0627,0644,0645,0627,0644,0643"), y, "SELECT OwnerID, OwnerCode & ' - ' & FullName AS OwnerLabel FROM tblOwners WHERE AdminStatus='ACTIVE' ORDER BY FullName;", 2, "0;4"): y = y + 650
    Call AT_AddCombo(n, "cboCreditMovementID", AT_U("0645,0635,062F,0631,0020,0627,0644,0631,0635,064A,062F,0020,0627,0644,062F,0627,0626,0646"), y, "SELECT CM.CreditMovementID, O.OwnerCode & ' - ' & O.FullName AS CreditOwner, CM.SignedAmount FROM (tblOwnerCreditMovements AS CM INNER JOIN tblOwners AS O ON CM.OwnerID=O.OwnerID) INNER JOIN tblFinancialDocuments AS D ON CM.DocumentID=D.DocumentID WHERE CM.SignedAmount>0 AND CM.SourceCreditMovementID Is Null AND D.StatusCode='APPROVED' ORDER BY CM.CreditMovementID DESC;", 3, "0;3.5;1.2"): y = y + 650
    Call AT_AddCombo(n, "cboChargeID", AT_U("0627,0644,0627,0633,062A,062D,0642,0627,0642"), y, "SELECT C.ChargeID, C.ChargeCode & ' - ' & O.FullName AS ChargeOwner, C.GrossAmount FROM (tblOwnerCharges AS C INNER JOIN tblOwners AS O ON C.OwnerID=O.OwnerID) INNER JOIN tblFinancialDocuments AS D ON C.DocumentID=D.DocumentID WHERE C.ChargeStatus='ACTIVE' AND D.StatusCode='APPROVED' ORDER BY C.ChargeDate, C.ChargeID;", 3, "0;3.5;1.2"): y = y + 650
    AT_AddFieldPair n, "txtAmount", AT_U("0645,0628,0644,063A,0020,0627,0644,062A,0637,0628,064A,0642"), y, 6500: y = y + 650
    AT_AddFieldPair n, "txtReason", AT_U("0633,0628,0628,0020,0627,0644,062A,0637,0628,064A,0642"), y, 6500: y = y + 850
    AT_AddButton n, "cmdApplyCredit", AT_U("062A,0623,0643,064A,062F,0020,062A,0637,0628,064A,0642,0020,0627,0644,0631,0635,064A,062F"), 3800, y, 3000, 600, "=AT_UI_ApplyCredit()"
    AT_SaveAndCloseForm n, "frmCreditApply"
End Sub

Private Sub AT_CreateReceiptCorrectionForm()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmReceiptCorrection") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("0639,0643,0633,0020,062A,062E,0635,064A,0635,0020,0625,064A,0635,0627,0644"), 5200: y = 550
    AT_AddFieldPair n, "txtAllocationID", AT_U("0645,0639,0631,0641,0020,0633,0637,0631,0020,0627,0644,062A,062E,0635,064A,0635"), y, 6500: y = y + 700
    AT_AddFieldPair n, "txtAmount", AT_U("0627,0644,0645,0628,0644,063A,0020,0627,0644,0645,0631,0627,062F,0020,0639,0643,0633,0647"), y, 6500: y = y + 700
    AT_AddFieldPair n, "txtReason", AT_U("0633,0628,0628,0020,0627,0644,0639,0643,0633,0020,0028,0625,0644,0632,0627,0645,064A,0029"), y, 6500: y = y + 900
    AT_AddButton n, "cmdReverseAllocation", AT_U("062A,0623,0643,064A,062F,0020,0627,0644,0639,0643,0633"), 3800, y, 2500, 600, "=AT_UI_ReverseAllocation()"
    AT_SaveAndCloseForm n, "frmReceiptCorrection"
End Sub


Private Sub AT_CreateReceiptAllocationForm()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmReceiptAllocations") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("062A,062E,0635,064A,0635,0020,0645,0628,0644,063A,0020,0627,0644,0625,064A,0635,0627,0644,0020,0639,0644,0649,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,0627,062A"), 9000: y = 400
    Call AT_AddCombo(n, "cboReceiptID", AT_U("0627,0644,0625,064A,0635,0627,0644,0020,0627,0644,0645,0633,0648,062F,0629"), y, "SELECT R.ReceiptID, R.ReceiptID & ' - ' & O.FullName AS ReceiptLabel, R.Amount FROM (tblReceipts AS R INNER JOIN tblFinancialDocuments AS D ON R.DocumentID=D.DocumentID) INNER JOIN tblOwners AS O ON R.OwnerID=O.OwnerID WHERE D.StatusCode='DRAFT' ORDER BY R.ReceiptID DESC;", 3, "0;4;1.5"): y = y + 750
    Call AT_AddCombo(n, "cboChargeID", AT_U("0627,0644,0627,0633,062A,062D,0642,0627,0642,0020,0644,0644,0645,0627,0644,0643,0020,0646,0641,0633,0647"), y, "SELECT C.ChargeID, C.ChargeCode & ' - ' & O.FullName AS ChargeLabel, C.NetAmount FROM (tblOwnerCharges AS C INNER JOIN tblOwners AS O ON C.OwnerID=O.OwnerID) INNER JOIN tblFinancialDocuments AS D ON C.DocumentID=D.DocumentID WHERE C.ChargeStatus='ACTIVE' AND D.StatusCode='APPROVED' ORDER BY C.ChargeDate,C.ChargeID;", 3, "0;4;1.5"): y = y + 750
    AT_AddFieldPair n, "txtAllocatedAmount", AT_U("0645,0628,0644,063A,0020,0627,0644,062A,062E,0635,064A,0635,0020,0628,0627,0644,0631,064A,0627,0644"), y, 6500: y = y + 800
    AT_AddLabel n, "lblAllocationHelp", AT_U("0627,0644,062A,062E,0635,064A,0635,0020,0627,0644,064A,062F,0648,064A,0020,064A,0633,062A,0628,062F,0644,0020,0627,0644,062A,0648,0632,064A,0639,0020,0627,0644,062A,0644,0642,0627,0626,064A,0020,0644,0630,0644,0643,0020,0627,0644,0625,064A,0635,0627,0644,002E,0020,0623,064A,0020,0641,0627,0626,0636,0020,064A,0628,0642,0649,0020,0631,0635,064A,062F,064B,0627,0020,062F,0627,0626,0646,064B,0627,002E"), 450, y, 8000, 500, 9, False: y = y + 650
    AT_AddButton n, "cmdSetAllocation", AT_U("062D,0641,0638,0020,002F,0020,062A,0639,062F,064A,0644,0020,0647,0630,0627,0020,0627,0644,062A,062E,0635,064A,0635"), 800, y, 3000, 600, "=AT_UI_SaveReceiptAllocation()"
    AT_AddButton n, "cmdClearAllocations", AT_U("0645,0633,062D,0020,062A,062E,0635,064A,0635,0627,062A,0020,0627,0644,0625,064A,0635,0627,0644"), 4000, y, 2300, 600, "=AT_UI_ClearReceiptAllocations()"
    y = y + 750
    AT_AddButton n, "cmdShowDraftAllocations", AT_U("0639,0631,0636,0020,0627,0644,062A,062E,0635,064A,0635,0627,062A,0020,0627,0644,0645,0639,0644,0642,0629"), 2800, y, 3000, 600, "=AT_UI_OpenForm(""frmDraftReceiptAllocations"")"
    AT_SaveAndCloseForm n, "frmReceiptAllocations"
End Sub

Private Sub AT_CreateCampaignTargetForm()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmCampaignTargets") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("062A,062D,062F,064A,062F,0020,0646,0637,0627,0642,0020,062D,0645,0644,0629,0020,0627,0644,0645,0628,0627,0646,064A,0020,0623,0648,0020,0627,0644,0634,0642,0642"), 9000: y = 400
    Call AT_AddCombo(n, "cboCampaignID", AT_U("0627,0644,062D,0645,0644,0629,0020,0627,0644,0645,0633,0648,062F,0629"), y, "SELECT CampaignID, CampaignCode & ' - ' & CampaignName AS CampaignLabel, TargetScope FROM tblContributionCampaigns WHERE StatusCode='DRAFT' ORDER BY CreatedAt DESC;", 3, "0;4;1.5"): y = y + 750
    Call AT_AddCombo(n, "cboTargetKind", AT_U("0646,0648,0639,0020,0627,0644,0647,062F,0641"), y, AT_U("0053,0045,004C,0045,0043,0054,0020,0027,0042,0055,0049,004C,0044,0049,004E,0047,0027,0020,0041,0053,0020,0054,0061,0072,0067,0065,0074,0043,006F,0064,0065,002C,0027,0645,0628,0646,0649,0027,0020,0041,0053,0020,0054,0061,0072,0067,0065,0074,004E,0061,006D,0065,0020,0055,004E,0049,004F,004E,0020,0053,0045,004C,0045,0043,0054,0020,0027,0055,004E,0049,0054,0027,002C,0027,0634,0642,0629,0027,003B"), 2, "0;4"): y = y + 750
    Call AT_AddCombo(n, "cboBuildingID", AT_U("0627,0644,0645,0628,0646,0649,0020,0028,0644,0646,0637,0627,0642,0020,0627,0644,0645,0628,0627,0646,064A,0029"), y, "SELECT BuildingID, BuildingCode & ' - ' & BuildingName AS BuildingLabel FROM tblBuildings WHERE IsActive=True ORDER BY BuildingCode;", 2, "0;4"): y = y + 750
    Call AT_AddCombo(n, "cboUnitID", AT_U("0627,0644,0634,0642,0629,0020,0028,0644,0646,0637,0627,0642,0020,0627,0644,0634,0642,0642,0029"), y, "SELECT U.UnitID, B.BuildingCode & ' / ' & U.UnitCode AS UnitLabel FROM tblUnits AS U INNER JOIN tblBuildings AS B ON U.BuildingID=B.BuildingID WHERE U.IsActive=True AND B.IsActive=True ORDER BY B.BuildingCode,U.UnitCode;", 2, "0;4"): y = y + 750
    AT_AddButton n, "cmdSaveCampaignTarget", AT_U("0625,0636,0627,0641,0629,0020,0627,0644,0647,062F,0641,0020,0644,0644,062D,0645,0644,0629"), 2800, y, 3200, 600, "=AT_UI_SaveCampaignTarget()"
    AT_SaveAndCloseForm n, "frmCampaignTargets"
End Sub

Private Sub AT_CreateOccupancyForm()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmUnitOccupancy") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("062A,0633,062C,064A,0644,0020,0627,0644,0623,0633,0631,0629,0020,0648,0633,062C,0644,0020,0625,0634,063A,0627,0644,0020,0627,0644,0634,0642,0629"), 7800: y = 450
    Call AT_AddCombo(n, "cboUnitID", AT_U("0627,0644,0634,0642,0629"), y, "SELECT U.UnitID, B.BuildingCode & ' / ' & U.UnitCode AS UnitLabel FROM tblUnits AS U INNER JOIN tblBuildings AS B ON U.BuildingID=B.BuildingID WHERE U.IsActive=True ORDER BY B.BuildingCode,U.UnitCode;", 2, "0;4"): y = y + 700
    AT_AddFieldPair n, "txtFamilyName", AT_U("0627,0633,0645,0020,0627,0644,0623,0633,0631,0629"), y, 6500: y = y + 700
    AT_AddFieldPair n, "txtHouseholdCount", AT_U("0639,062F,062F,0020,0623,0641,0631,0627,062F,0020,0627,0644,0623,0633,0631,0629"), y, 6500: y = y + 700
    AT_AddFieldPair n, "txtStartDate", AT_U("062A,0627,0631,064A,062E,0020,0628,062F,0627,064A,0629,0020,0627,0644,0625,0634,063A,0627,0644"), y, 6500: y = y + 700
    AT_AddFieldPair n, "txtNotes", AT_U("0645,0644,0627,062D,0638,0627,062A"), y, 6500: y = y + 700
    AT_AddFieldPair n, "txtOccupancyID", AT_U("0627,0644,0645,0639,0631,0641,0020,0627,0644,062F,0627,062E,0644,064A"), y, 6500: f.Controls("txtOccupancyID").Locked = True: y = y + 800
    AT_AddButton n, "cmdSaveOccupancy", AT_U("062D,0641,0638,0020,0633,062C,0644,0020,0627,0644,0625,0634,063A,0627,0644"), 3000, y, 2500, 600, "=AT_UI_SaveOccupancy()"
    AT_SaveAndCloseForm n, "frmUnitOccupancy"
End Sub


Private Sub AT_CreateConfigurationForms()
    Dim f As Form, n As String, y As Long
    If Not AT_FormExists("frmProjectSettings") Then
        Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("0625,0639,062F,0627,062F,0627,062A,0020,0627,0644,0645,0634,0631,0648,0639,0020,0648,0627,0644,062D,0627,0631,0629"), 9000: y = 400
        AT_AddFieldPair n, "txtProjectName", AT_U("0627,0633,0645,0020,0627,0644,0645,0634,0631,0648,0639"), y, 6500: y = y + 650
        AT_AddFieldPair n, "txtNeighborhoodName", AT_U("0627,0633,0645,0020,0627,0644,062D,0627,0631,0629"), y, 6500: y = y + 650
        AT_AddFieldPair n, "txtDefaultCollector", AT_U("0627,0633,0645,0020,0627,0644,0645,062D,0635,0651,0644,0020,0627,0644,0627,0641,062A,0631,0627,0636,064A"), y, 6500: y = y + 650
        AT_AddFieldPair n, "txtContactInfo", AT_U("0628,064A,0627,0646,0627,062A,0020,0627,0644,062A,0648,0627,0635,0644,0020,002F,0020,0645,0644,0627,062D,0638,0629,0020,0627,0644,0645,0633,062A,0646,062F"), y, 6500: y = y + 850
        AT_AddButton n, "cmdLoadProjectSettings", AT_U("062A,062D,0645,064A,0644,0020,0627,0644,0625,0639,062F,0627,062F,0627,062A,0020,0627,0644,062D,0627,0644,064A,0629"), 1200, y, 2800, 600, "=AT_UI_LoadProjectSettings()"
        AT_AddButton n, "cmdSaveProjectSettings", AT_U("062D,0641,0638,0020,0627,0644,0625,0639,062F,0627,062F,0627,062A"), 4300, y, 2200, 600, "=AT_UI_SaveProjectSettings()"
        AT_SaveAndCloseForm n, "frmProjectSettings"
    End If
    If Not AT_FormExists("frmRateAdmin") Then
        Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("0625,062F,0627,0631,0629,0020,0623,0633,0639,0627,0631,0020,0627,0644,0645,0633,0627,0647,0645,0629,0020,0630,0627,062A,0020,062A,0627,0631,064A,062E,0020,0627,0644,0633,0631,064A,0627,0646"), 8200: y = 450
        AT_AddFieldPair n, "txtRateCode", AT_U("0631,0645,0632,0020,0627,0644,0633,0639,0631,0020,0627,0644,062C,062F,064A,062F"), y, 6500: y = y + 650
        AT_AddFieldPair n, "txtRateAmount", AT_U("0642,064A,0645,0629,0020,0627,0644,0645,0633,0627,0647,0645,0629,0020,0644,0643,0644,0020,0634,0642,0629,0020,0628,0627,0644,0631,064A,0627,0644"), y, 6500: y = y + 650
        AT_AddFieldPair n, "txtEffectiveFrom", AT_U("062A,0627,0631,064A,062E,0020,0628,062F,0627,064A,0629,0020,0627,0644,0633,0631,064A,0627,0646"), y, 6500: y = y + 650
        AT_AddFieldPair n, "txtRateID", AT_U("0627,0644,0645,0639,0631,0641,0020,0627,0644,062F,0627,062E,0644,064A,0020,0644,0644,0633,0639,0631"), y, 6500: f.Controls("txtRateID").Locked = True: y = y + 800
        AT_AddButton n, "cmdCreateRate", AT_U("062D,0641,0638,0020,0627,0644,0633,0639,0631,0020,0627,0644,062C,062F,064A,062F"), 3000, y, 2600, 600, "=AT_UI_CreateRate()"
        AT_SaveAndCloseForm n, "frmRateAdmin"
    End If
    AT_CreateReadOnlyGridForm "frmRateDirectory", AT_U("0633,062C,0644,0020,0623,0633,0639,0627,0631,0020,0627,0644,0645,0633,0627,0647,0645,0629,0020,0627,0644,062A,0627,0631,064A,062E,064A,0629"), "SELECT RateCode,RateAmount,EffectiveFrom,EffectiveTo,IsActive FROM tblContributionRates ORDER BY EffectiveFrom DESC;"
End Sub

Private Sub AT_CreateFinancialClosingForms()
    Dim f As Form, n As String, y As Long
    If Not AT_FormExists("frmFinancialPeriod") Then
        Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("0625,062F,0627,0631,0629,0020,0627,0644,0641,062A,0631,0627,062A,0020,0627,0644,0645,0627,0644,064A,0629"), 9000: y = 350
        AT_AddFieldPair n, "txtPeriodCode", AT_U("0631,0645,0632,0020,0627,0644,0641,062A,0631,0629"), y, 6500: y = y + 600
        AT_AddFieldPair n, "txtPeriodName", AT_U("0627,0633,0645,0020,0627,0644,0641,062A,0631,0629"), y, 6500: y = y + 600
        AT_AddFieldPair n, "txtPeriodStart", AT_U("062A,0627,0631,064A,062E,0020,0627,0644,0628,062F,0627,064A,0629"), y, 6500: y = y + 600
        AT_AddFieldPair n, "txtPeriodEnd", AT_U("062A,0627,0631,064A,062E,0020,0627,0644,0646,0647,0627,064A,0629"), y, 6500: y = y + 700
        AT_AddFieldPair n, "txtPeriodID", AT_U("0627,0644,0645,0639,0631,0641,0020,0627,0644,062F,0627,062E,0644,064A,0020,0644,0644,0641,062A,0631,0629"), y, 6500: f.Controls("txtPeriodID").Locked = True: y = y + 700
        AT_AddButton n, "cmdCreatePeriod", AT_U("0625,0646,0634,0627,0621,0020,0641,062A,0631,0629,0020,0645,0627,0644,064A,0629"), 3000, y, 2500, 600, "=AT_UI_CreatePeriod()": y = y + 800
        Call AT_AddCombo(n, "cboClosedPeriodID", AT_U("0627,0644,0641,062A,0631,0629,0020,0627,0644,0645,063A,0644,0642,0629,0020,0644,0625,0639,0627,062F,0629,0020,0641,062A,062D,0647,0627"), y, "SELECT PeriodID,PeriodCode & ' - ' & PeriodName AS PeriodLabel FROM tblFinancialPeriods WHERE StatusCode='CLOSED' ORDER BY PeriodStart DESC;", 2, "0;4"): y = y + 650
        AT_AddFieldPair n, "txtReopenReason", AT_U("0633,0628,0628,0020,0625,0639,0627,062F,0629,0020,0627,0644,0641,062A,062D,0020,0028,0625,0644,0632,0627,0645,064A,0029"), y, 6500: y = y + 700
        AT_AddButton n, "cmdReopenPeriod", AT_U("0625,0639,0627,062F,0629,0020,0641,062A,062D,0020,0627,0644,0641,062A,0631,0629"), 3000, y, 2500, 600, "=AT_UI_ReopenPeriod()"
        AT_SaveAndCloseForm n, "frmFinancialPeriod"
    End If
    If Not AT_FormExists("frmFinancialClosing") Then
        Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("0627,0644,0625,0642,0641,0627,0644,0020,0648,0645,0637,0627,0628,0642,0629,0020,0627,0644,0623,0631,0635,062F,0629"), 9200: y = 350
        Call AT_AddCombo(n, "cboPeriodID", AT_U("0627,0644,0641,062A,0631,0629,0020,0627,0644,0645,0641,062A,0648,062D,0629"), y, "SELECT PeriodID,PeriodCode & ' - ' & PeriodName AS PeriodLabel FROM tblFinancialPeriods WHERE StatusCode='OPEN' ORDER BY PeriodStart;", 2, "0;4"): y = y + 650
        AT_AddFieldPair n, "txtClosingID", AT_U("0645,0639,0631,0641,0020,0627,0644,0625,0642,0641,0627,0644"), y, 6500: f.Controls("txtClosingID").Locked = True: y = y + 700
        AT_AddButton n, "cmdBeginClosing", AT_U("0628,062F,0621,0020,0627,0644,0625,0642,0641,0627,0644"), 3200, y, 2200, 600, "=AT_UI_BeginClosing()": y = y + 800
        Call AT_AddCombo(n, "cboAccountID", AT_U("0627,0644,062D,0633,0627,0628"), y, "SELECT AccountID,AccountCode & ' - ' & AccountName FROM tblAccounts WHERE IsActive=True ORDER BY AccountName;", 2, "0;4"): y = y + 650
        AT_AddFieldPair n, "txtActualBalance", AT_U("0627,0644,0631,0635,064A,062F,0020,0627,0644,0641,0639,0644,064A"), y, 6500: y = y + 600
        AT_AddFieldPair n, "txtClosingNotes", AT_U("0645,0644,0627,062D,0638,0629,0020,0627,0644,0645,0637,0627,0628,0642,0629"), y, 6500: y = y + 700
        AT_AddButton n, "cmdSaveActual", AT_U("062D,0641,0638,0020,0627,0644,0631,0635,064A,062F,0020,0627,0644,0641,0639,0644,064A"), 2800, y, 2500, 600, "=AT_UI_SaveClosingActual()": y = y + 800
        AT_AddFieldPair n, "txtExceptionReason", AT_U("0633,0628,0628,0020,0627,0633,062A,062B,0646,0627,0621,0020,0627,0644,0641,0631,0648,0642,0627,062A,0020,0028,0644,0644,0645,062F,064A,0631,0020,0627,0644,0623,0639,0644,0649,0020,0641,0642,0637,0029"), y, 6500: y = y + 750
        AT_AddButton n, "cmdFinalizeClosing", AT_U("0645,0631,0627,062C,0639,0629,0020,0648,0625,063A,0644,0627,0642,0020,0627,0644,0641,062A,0631,0629"), 2800, y, 3000, 600, "=AT_UI_FinalizeClosing()": y = y + 750
        AT_AddButton n, "cmdShowClosingAccounts", AT_U("0639,0631,0636,0020,062A,0641,0627,0635,064A,0644,0020,0623,0631,0635,062F,0629,0020,0627,0644,0625,0642,0641,0627,0644"), 2800, y, 3000, 600, "=AT_UI_OpenForm(""frmClosingAccounts"")"
        AT_SaveAndCloseForm n, "frmFinancialClosing"
    End If
    AT_CreateReadOnlyGridForm "frmClosingAccounts", AT_U("062A,0641,0627,0635,064A,0644,0020,0645,0637,0627,0628,0642,0629,0020,0627,0644,062D,0633,0627,0628,0627,062A,0020,0641,064A,0020,0627,0644,0625,0642,0641,0627,0644"), "SELECT C.ClosingID,A.AccountCode,A.AccountName,X.BookBalance,X.ActualBalance,X.DifferenceAmount,X.Notes FROM (tblFinancialClosingAccounts AS X INNER JOIN tblAccounts AS A ON X.AccountID=A.AccountID) INNER JOIN tblFinancialClosings AS C ON X.ClosingID=C.ClosingID ORDER BY C.ClosingID DESC,A.AccountName;"
End Sub

Private Sub AT_CreateReadOnlyGridForm(ByVal desiredName As String, ByVal captionText As String, ByVal sourceSQL As String)
    Dim f As Form, n As String
    If AT_FormExists(desiredName) Then Exit Sub
    Set f = CreateForm: n = f.Name
    f.Caption = captionText: f.RecordSource = sourceSQL: f.DefaultView = acDefViewDatasheet
    f.AllowEdits = False: f.AllowAdditions = False: f.AllowDeletions = False: f.DataEntry = False
    f.RecordSelectors = True: f.NavigationButtons = True
    AT_SaveAndCloseForm n, desiredName
End Sub


Public Function AT_UI_SaveBuilding() As Boolean
    On Error GoTo EH
    Dim f As Form, ownerID As Long, buildingID As Long
    Set f = Forms("frmBuildingCreate")
    ownerID = CLng(Nz(f.Controls("cboOwnerID").Value, 0))
    If MsgBox(AT_U("0633,064A,062A,0645,0020,062A,0633,062C,064A,0644,0020,0627,0644,0645,0628,0646,0649,0020,0648,0631,0628,0637,0647,0020,0628,0627,0644,0645,0627,0644,0643,0020,0627,0644,0645,062D,062F,062F,0020,0645,0639,0020,062D,0641,0638,0020,062A,0627,0631,064A,062E,0020,0628,062F,0627,064A,0629,0020,0627,0644,0645,0644,0643,064A,0629,002E,0020,0647,0644,0020,062A,0631,064A,062F,0020,0627,0644,0645,062A,0627,0628,0639,0629,061F"), vbQuestion + vbYesNo, AT_U("062A,0623,0643,064A,062F,0020,062A,0633,062C,064A,0644,0020,0627,0644,0645,0628,0646,0649")) <> vbYes Then Exit Function
    buildingID = AT_CreateBuilding(CStr(Nz(f.Controls("txtBuildingCode").Value, "")), CStr(Nz(f.Controls("txtBuildingName").Value, "")), CStr(Nz(f.Controls("txtBuildingAddress").Value, "")), ownerID, AT_UI_DateValue(f.Controls("txtEffectiveDate").Value), CStr(Nz(f.Controls("txtTransferRef").Value, "")), CStr(Nz(f.Controls("txtBuildingNotes").Value, "")))
    f.Controls("txtBuildingID").Value = buildingID
    MsgBox AT_U("062A,0645,0020,062A,0633,062C,064A,0644,0020,0627,0644,0645,0628,0646,0649,0020,0628,0627,0644,0645,0639,0631,0641,0020") & buildingID & AT_U("002E,0020,0623,0636,0641,0020,0627,0644,0634,0642,0642,0020,0645,0646,0020,0634,0627,0634,0629,0020,062A,0633,062C,064A,0644,0020,0627,0644,0634,0642,0629,002E"), vbInformation, "Archon Tech"
    AT_UI_SaveBuilding = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,062A,0633,062C,064A,0644,0020,0627,0644,0645,0628,0646,0649,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SaveUnit() As Boolean
    On Error GoTo EH
    Dim f As Form, buildingID As Long, unitID As Long
    Set f = Forms("frmUnitCreate"): buildingID = CLng(Nz(f.Controls("cboBuildingID").Value, 0))
    unitID = AT_AddUnit(buildingID, CStr(Nz(f.Controls("txtUnitCode").Value, "")), CStr(Nz(f.Controls("txtFloorNo").Value, "")), True, CStr(Nz(f.Controls("txtUnitNotes").Value, "")))
    f.Controls("txtUnitID").Value = unitID
    MsgBox AT_U("062A,0645,0020,062A,0633,062C,064A,0644,0020,0627,0644,0634,0642,0629,0020,0628,0627,0644,0645,0639,0631,0641,0020") & unitID & AT_U("002E,0020,0648,0647,064A,0020,0645,0624,0647,0644,0629,0020,0644,0644,0645,0633,0627,0647,0645,0629,0020,0627,0641,062A,0631,0627,0636,064A,064B,0627,0020,0645,0627,0020,0644,0645,0020,064A,064F,0639,062A,0645,062F,0020,0625,0639,0641,0627,0624,0647,0627,002E"), vbInformation, "Archon Tech"
    AT_UI_SaveUnit = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,062A,0633,062C,064A,0644,0020,0627,0644,0634,0642,0629,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_TransferOwnership() As Boolean
    On Error GoTo EH
    Dim f As Form, buildingID As Long, ownerID As Long
    Set f = Forms("frmOwnershipTransfer")
    buildingID = CLng(Nz(f.Controls("cboBuildingID").Value, 0)): ownerID = CLng(Nz(f.Controls("cboNewOwnerID").Value, 0))
    If MsgBox(AT_U("0646,0642,0644,0020,0627,0644,0645,0644,0643,064A,0629,0020,0644,0627,0020,064A,0646,0642,0644,0020,062F,064A,0648,0646,0020,0627,0644,0645,0627,0644,0643,0020,0627,0644,0633,0627,0628,0642,002E,0020,0633,064A,062A,0645,0020,062D,0641,0638,0020,0633,062C,0644,0020,0627,0644,0645,0644,0643,064A,0629,0020,0627,0644,062A,0627,0631,064A,062E,064A,002E,0020,0647,0644,0020,062A,0631,064A,062F,0020,0627,0644,062A,0646,0641,064A,0630,061F"), vbQuestion + vbYesNo, AT_U("062A,0623,0643,064A,062F,0020,0646,0642,0644,0020,0627,0644,0645,0644,0643,064A,0629")) <> vbYes Then Exit Function
    AT_TransferBuildingOwnership buildingID, ownerID, AT_UI_DateValue(f.Controls("txtTransferDate").Value), CStr(Nz(f.Controls("txtTransferDocument").Value, "")), CStr(Nz(f.Controls("txtTransferReason").Value, ""))
    MsgBox AT_U("062A,0645,0020,062D,0641,0638,0020,0646,0642,0644,0020,0627,0644,0645,0644,0643,064A,0629,002E,0020,0628,0642,064A,062A,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,0627,062A,0020,0627,0644,062A,0627,0631,064A,062E,064A,0629,0020,0639,0644,0649,0020,0623,0635,062D,0627,0628,0647,0627,0020,0648,0642,062A,0020,0625,0646,0634,0627,0626,0647,0627,002E"), vbInformation, "Archon Tech"
    AT_UI_TransferOwnership = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0646,0642,0644,0020,0627,0644,0645,0644,0643,064A,0629,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_RequestExemption() As Boolean
    On Error GoTo EH
    Dim f As Form, unitID As Long, requestID As Long, dEnd As Variant
    Set f = Forms("frmUnitExemption"): unitID = CLng(Nz(f.Controls("cboUnitID").Value, 0))
    dEnd = Null: If Len(Trim$(Nz(f.Controls("txtEffectiveTo").Value, ""))) > 0 Then dEnd = AT_UI_DateValue(f.Controls("txtEffectiveTo").Value)
    requestID = AT_RequestUnitExemption(unitID, AT_UI_DateValue(f.Controls("txtEffectiveFrom").Value), dEnd, CStr(Nz(f.Controls("txtExemptionReason").Value, "")), CStr(Nz(f.Controls("txtEvidenceReference").Value, "")))
    f.Controls("txtExemptionRequestID").Value = requestID
    MsgBox AT_U("0633,064F,062C,0644,0020,0637,0644,0628,0020,0627,0644,0625,0639,0641,0627,0621,0020,0631,0642,0645,0020") & requestID & AT_U("0020,0628,0627,0646,062A,0638,0627,0631,0020,0642,0631,0627,0631,0020,0627,0644,0645,0633,0624,0648,0644,002E,0020,0644,0645,0020,064A,062A,063A,064A,0631,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,0020,0627,0644,0645,0627,0644,064A,0020,0628,0639,062F,002E"), vbInformation, "Archon Tech"
    AT_UI_RequestExemption = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,062A,0633,062C,064A,0644,0020,0637,0644,0628,0020,0627,0644,0625,0639,0641,0627,0621,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_DecideExemption() As Boolean
    On Error GoTo EH
    Dim f As Form, reqID As Long, approveIt As Boolean
    Set f = Forms("frmExemptionReviewAction"): reqID = CLng(Nz(f.Controls("txtRequestID").Value, 0))
    approveIt = (Nz(f.Controls("cboDecision").Value, "REJECT") = "APPROVE")
    If MsgBox(AT_U("0647,0644,0020,062A,0624,0643,062F,0020,0642,0631,0627,0631,0020") & IIf(approveIt, AT_U("0627,0639,062A,0645,0627,062F"), AT_U("0631,0641,0636")) & AT_U("0020,0637,0644,0628,0020,0627,0644,0625,0639,0641,0627,0621,061F,0020,064A,062C,0628,0020,062A,0648,062B,064A,0642,0020,0627,0644,0633,0628,0628,002E"), vbQuestion + vbYesNo, AT_U("062A,0623,0643,064A,062F,0020,0627,0644,0642,0631,0627,0631")) <> vbYes Then Exit Function
    AT_DecideUnitExemption reqID, approveIt, CStr(Nz(f.Controls("txtDecisionReason").Value, ""))
    MsgBox AT_U("062A,0645,0020,062D,0641,0638,0020,0627,0644,0642,0631,0627,0631,002E,0020,0644,0627,0020,064A,0646,0641,0630,0020,0627,0644,0637,0644,0628,0020,0628,0623,062B,0631,0020,0631,062C,0639,064A,0020,0639,0644,0649,0020,0627,0633,062A,062D,0642,0627,0642,0627,062A,0020,0645,0639,062A,0645,062F,0629,002E"), vbInformation, "Archon Tech"
    AT_UI_DecideExemption = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,062D,0641,0638,0020,0642,0631,0627,0631,0020,0627,0644,0625,0639,0641,0627,0621,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SaveCampaign() As Boolean
    On Error GoTo EH
    Dim f As Form, campaignID As Long, typeID As Long, scopeCode As String
    Set f = Forms("frmCampaignCreate")
    typeID = CLng(Nz(f.Controls("cboContributionTypeID").Value, 0)): scopeCode = CStr(Nz(f.Controls("cboTargetScope").Value, "ALL"))
    campaignID = AT_CreateCampaign(CStr(Nz(f.Controls("txtCampaignCode").Value, "")), CStr(Nz(f.Controls("txtCampaignName").Value, "")), typeID, scopeCode, CCur(Val(Nz(f.Controls("txtAmountPerUnit").Value, "0"))), 0, AT_UI_OptionalDateValue(f.Controls("txtPeriodStart").Value), AT_UI_OptionalDateValue(f.Controls("txtPeriodEnd").Value), AT_UI_DateValue(f.Controls("txtDueDate").Value), CStr(Nz(f.Controls("txtCampaignReason").Value, "")))
    f.Controls("txtCampaignID").Value = campaignID
    MsgBox AT_U("062A,0645,0020,0625,0646,0634,0627,0621,0020,0627,0644,062D,0645,0644,0629,0020,0628,0627,0644,0645,0639,0631,0641,0020") & campaignID & AT_U("002E,0020,0646,0637,0627,0642,0020,0028,0627,0644,0645,0628,0627,0646,064A,002F,0627,0644,0634,0642,0642,0029,0020,063A,064A,0631,0020,0627,0644,0634,0627,0645,0644,0020,064A,062A,0637,0644,0628,0020,0625,0636,0627,0641,0629,0020,0623,0647,062F,0627,0641,0020,0642,0628,0644,0020,0625,0646,0634,0627,0621,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,0627,062A,002E"), vbInformation, "Archon Tech"
    AT_UI_SaveCampaign = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0625,0646,0634,0627,0621,0020,0627,0644,062D,0645,0644,0629,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_GenerateCampaign() As Boolean
    On Error GoTo EH
    Dim campaignID As Long
    campaignID = CLng(Nz(Forms("frmCampaignCreate").Controls("txtCampaignID").Value, 0))
    If campaignID <= 0 Then Err.Raise vbObjectError + 920, , AT_U("0623,0646,0634,0626,0020,0627,0644,062D,0645,0644,0629,0020,0623,0648,0644,064B,0627,002E")
    If MsgBox(AT_U("0633,064A,062A,0645,0020,0625,0646,0634,0627,0621,0020,0645,0633,0648,062F,0627,062A,0020,0627,0633,062A,062D,0642,0627,0642,0020,0644,0643,0644,0020,0645,0627,0644,0643,0020,062D,0633,0628,0020,0627,0644,0634,0642,0642,0020,0627,0644,0645,0633,062A,0647,062F,0641,0629,002E,0020,0644,0627,0020,0623,062B,0631,0020,0645,0627,0644,064A,0020,0642,0628,0644,0020,0627,0644,0627,0639,062A,0645,0627,062F,002E"), vbQuestion + vbYesNo, AT_U("0625,0646,0634,0627,0621,0020,0645,0633,0648,062F,0627,062A,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642")) <> vbYes Then Exit Function
    MsgBox AT_U("062A,0645,0020,0625,0646,0634,0627,0621,0020") & AT_GenerateCampaignCharges(campaignID, Date) & AT_U("0020,0645,0633,0648,062F,0629,002F,0645,0633,0648,062F,0627,062A,0020,0644,0644,0645,0627,0644,0643,064A,0646,0020,0627,0644,0645,062D,062F,062F,064A,0646,002E"), vbInformation, "Archon Tech"
    AT_UI_GenerateCampaign = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0625,0646,0634,0627,0621,0020,0627,0633,062A,062D,0642,0627,0642,0627,062A,0020,0627,0644,062D,0645,0644,0629,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ApproveCampaign() As Boolean
    On Error GoTo EH
    Dim campaignID As Long
    campaignID = CLng(Nz(Forms("frmCampaignCreate").Controls("txtCampaignID").Value, 0))
    If campaignID <= 0 Then Err.Raise vbObjectError + 921, , AT_U("0627,062E,062A,0631,0020,062D,0645,0644,0629,0020,0645,0648,062C,0648,062F,0629,002E")
    If MsgBox(AT_U("0633,064A,062A,0645,0020,0627,0639,062A,0645,0627,062F,0020,0627,0633,062A,062D,0642,0627,0642,0627,062A,0020,0627,0644,062D,0645,0644,0629,060C,0020,0648,0633,064A,0638,0647,0631,0020,0623,062B,0631,0647,0627,0020,0641,064A,0020,0623,0631,0635,062F,0629,0020,0627,0644,0645,0644,0627,0643,002E,0020,0647,0644,0020,0631,0627,062C,0639,062A,0020,062A,0641,0627,0635,064A,0644,0647,0627,061F"), vbQuestion + vbYesNo, AT_U("062A,0623,0643,064A,062F,0020,0627,0639,062A,0645,0627,062F,0020,0627,0644,062D,0645,0644,0629")) <> vbYes Then Exit Function
    MsgBox AT_U("062A,0645,0020,0627,0639,062A,0645,0627,062F,0020") & AT_ApproveCampaignCharges(campaignID) & AT_U("0020,0627,0633,062A,062D,0642,0627,0642,064B,0627,002E"), vbInformation, "Archon Tech"
    AT_UI_ApproveCampaign = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0627,0639,062A,0645,0627,062F,0020,0627,0633,062A,062D,0642,0627,0642,0627,062A,0020,0627,0644,062D,0645,0644,0629,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SaveExpense() As Boolean
    On Error GoTo EH
    Dim f As Form, expenseID As Long
    Set f = Forms("frmExpenseCreate")
    expenseID = AT_CreateExpenseDraft(AT_UI_DateValue(f.Controls("txtExpenseDate").Value), CLng(Nz(f.Controls("cboExpenseCategoryID").Value, 0)), CStr(Nz(f.Controls("txtBeneficiary").Value, "")), CStr(Nz(f.Controls("txtExpenseDescription").Value, "")), CCur(Val(Nz(f.Controls("txtExpenseAmount").Value, "0"))))
    f.Controls("txtExpenseID").Value = expenseID
    MsgBox AT_U("0633,064F,062C,0644,0020,0627,0644,0645,0635,0631,0648,0641,0020,0643,0645,0633,0648,062F,0629,0020,0631,0642,0645,0020") & expenseID & AT_U("002E,0020,0644,0645,0020,064A,062D,062F,062B,0020,0635,0631,0641,0020,0645,0627,0644,064A,002E"), vbInformation, "Archon Tech"
    AT_UI_SaveExpense = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,062A,0633,062C,064A,0644,0020,0627,0644,0645,0635,0631,0648,0641,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_SaveExpensePayment() As Boolean
    On Error GoTo EH
    Dim f As Form, paymentID As Long, expenseID As Long
    Set f = Forms("frmExpensePay"): expenseID = CLng(Nz(f.Controls("cboExpenseID").Value, 0))
    paymentID = AT_CreateExpensePaymentDraft(expenseID, CLng(Nz(f.Controls("cboAccountID").Value, 0)), CLng(Nz(f.Controls("cboPaymentMethodID").Value, 0)), CStr(Nz(f.Controls("txtRecipient").Value, "")), CStr(Nz(f.Controls("txtPaymentNotes").Value, "")))
    f.Controls("txtPaymentID").Value = paymentID
    MsgBox AT_U("0623,0646,0634,0626,0020,0633,0646,062F,0020,0627,0644,0635,0631,0641,0020,0643,0645,0633,0648,062F,0629,0020,0631,0642,0645,0020") & paymentID & AT_U("002E,0020,0628,0639,062F,0020,0627,0644,062F,0641,0639,0020,0627,0644,0641,0639,0644,064A,0020,0623,0643,0651,062F,0647,0020,0645,0646,0020,0646,0641,0633,0020,0627,0644,0634,0627,0634,0629,002E"), vbInformation, "Archon Tech"
    AT_UI_SaveExpensePayment = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0625,0646,0634,0627,0621,0020,0633,0646,062F,0020,0627,0644,0635,0631,0641,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ApproveExpensePayment() As Boolean
    On Error GoTo EH
    Dim f As Form, paymentID As Long
    Set f = Forms("frmExpensePay"): paymentID = CLng(Nz(f.Controls("txtPaymentID").Value, 0))
    If paymentID <= 0 Then Err.Raise vbObjectError + 922, , AT_U("0623,0646,0634,0626,0020,0633,0646,062F,0020,0627,0644,0635,0631,0641,0020,0623,0648,0644,064B,0627,002E")
    If MsgBox(AT_U("0644,0627,0020,062A,0624,0643,062F,0020,0625,0644,0627,0020,0628,0639,062F,0020,062A,0633,0644,064A,0645,0020,0627,0644,0645,0628,0644,063A,0020,0641,0639,0644,064A,064B,0627,0020,0644,0644,0645,0633,062A,0641,064A,062F,002E,0020,0647,0644,0020,062A,0645,0020,0627,0644,062F,0641,0639,061F"), vbQuestion + vbYesNo, AT_U("062A,0623,0643,064A,062F,0020,0627,0644,062F,0641,0639,0020,0627,0644,0641,0639,0644,064A")) <> vbYes Then Exit Function
    AT_ApprovePayment paymentID, Environ$("Username"), True
    MsgBox AT_U("062A,0645,0020,0627,0639,062A,0645,0627,062F,0020,0633,0646,062F,0020,0627,0644,0635,0631,0641,0020,0648,062A,0633,062C,064A,0644,0020,062D,0631,0643,0629,0020,0627,0644,062D,0633,0627,0628,002E"), vbInformation, "Archon Tech"
    AT_UI_ApproveExpensePayment = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,0020,0627,0639,062A,0645,0627,062F,0020,0633,0646,062F,0020,0627,0644,0635,0631,0641,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Public Function AT_UI_ReprintReceipt() As Boolean
    On Error GoTo EH
    Dim receiptID As Long
    receiptID = CLng(Nz(Forms("frmReceiptReprint").Controls("txtReceiptID").Value, 0))
    If receiptID <= 0 Then Err.Raise vbObjectError + 923, , AT_U("0623,062F,062E,0644,0020,0627,0644,0645,0639,0631,0641,0020,0627,0644,062F,0627,062E,0644,064A,0020,0644,0644,0625,064A,0635,0627,0644,002E")
    AT_PrintReceipt receiptID
    AT_UI_ReprintReceipt = True: Exit Function
EH:
    MsgBox AT_U("062A,0639,0630,0631,062A,0020,0625,0639,0627,062F,0629,0020,0637,0628,0627,0639,0629,0020,0627,0644,0625,064A,0635,0627,0644,003A,0020") & Err.Description, vbExclamation, "Archon Tech"
End Function

Private Sub AT_CreatePropertyForms()
    Dim f As Form, n As String, y As Long
    If Not AT_FormExists("frmBuildingCreate") Then
        Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("062A,0633,062C,064A,0644,0020,0645,0628,0646,0649,0020,0648,0631,0628,0637,0647,0020,0628,0645,0627,0644,0643,0647,0020,0627,0644,062D,0627,0644,064A"), 8200: y = 400
        AT_AddFieldPair n, "txtBuildingCode", AT_U("0631,0645,0632,0020,0627,0644,0645,0628,0646,0649"), y, 6500: y = y + 600
        AT_AddFieldPair n, "txtBuildingName", AT_U("0627,0633,0645,002F,0648,0635,0641,0020,0627,0644,0645,0628,0646,0649"), y, 6500: y = y + 600
        AT_AddFieldPair n, "txtBuildingAddress", AT_U("0639,0646,0648,0627,0646,0020,0627,0644,0645,0628,0646,0649"), y, 6500: y = y + 600
        Call AT_AddCombo(n, "cboOwnerID", AT_U("0627,0644,0645,0627,0644,0643,0020,0627,0644,062D,0627,0644,064A"), y, "SELECT OwnerID, OwnerCode & ' - ' & FullName FROM tblOwners WHERE AdminStatus='ACTIVE' ORDER BY FullName;", 2, "0;5"): y = y + 650
        AT_AddFieldPair n, "txtEffectiveDate", AT_U("062A,0627,0631,064A,062E,0020,0628,062F,0627,064A,0629,0020,0627,0644,0645,0644,0643,064A,0629,0020,0028,0627,0641,062A,0631,0627,0636,064A,0020,0627,0644,064A,0648,0645,0029"), y, 6500: y = y + 600
        AT_AddFieldPair n, "txtTransferRef", AT_U("0645,0631,062C,0639,0020,0645,0633,062A,0646,062F,0020,0627,0644,0645,0644,0643,064A,0629"), y, 6500: y = y + 600
        AT_AddFieldPair n, "txtBuildingNotes", AT_U("0645,0644,0627,062D,0638,0627,062A"), y, 6500: y = y + 600
        AT_AddFieldPair n, "txtBuildingID", AT_U("0627,0644,0645,0639,0631,0641,0020,0627,0644,062F,0627,062E,0644,064A"), y, 6500: f.Controls("txtBuildingID").Locked = True: y = y + 750
        AT_AddButton n, "cmdSaveBuilding", AT_U("062D,0641,0638,0020,0627,0644,0645,0628,0646,0649"), 3800, y, 2500, 600, "=AT_UI_SaveBuilding()"
        AT_SaveAndCloseForm n, "frmBuildingCreate"
    End If
    If Not AT_FormExists("frmUnitCreate") Then
        Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("062A,0633,062C,064A,0644,0020,0634,0642,0629,0020,062F,0627,062E,0644,0020,0645,0628,0646,0649"), 5600: y = 500
        Call AT_AddCombo(n, "cboBuildingID", AT_U("0627,0644,0645,0628,0646,0649"), y, "SELECT BuildingID, BuildingCode & ' - ' & BuildingName FROM tblBuildings WHERE IsActive=True ORDER BY BuildingCode;", 2, "0;5"): y = y + 700
        AT_AddFieldPair n, "txtUnitCode", AT_U("0631,0642,0645,002F,0631,0645,0632,0020,0627,0644,0634,0642,0629,0020,062F,0627,062E,0644,0020,0627,0644,0645,0628,0646,0649"), y, 6500: y = y + 600
        AT_AddFieldPair n, "txtFloorNo", AT_U("0627,0644,0637,0627,0628,0642"), y, 6500: y = y + 600
        AT_AddFieldPair n, "txtUnitNotes", AT_U("0645,0644,0627,062D,0638,0627,062A"), y, 6500: y = y + 600
        AT_AddFieldPair n, "txtUnitID", AT_U("0627,0644,0645,0639,0631,0641,0020,0627,0644,062F,0627,062E,0644,064A"), y, 6500: f.Controls("txtUnitID").Locked = True: y = y + 750
        AT_AddButton n, "cmdSaveUnit", AT_U("062D,0641,0638,0020,0627,0644,0634,0642,0629,0020,0627,0644,0645,0624,0647,0644,0629,0020,0627,0641,062A,0631,0627,0636,064A,064B,0627"), 3200, y, 3500, 600, "=AT_UI_SaveUnit()"
        AT_SaveAndCloseForm n, "frmUnitCreate"
    End If
    If Not AT_FormExists("frmOwnershipTransfer") Then
        Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("0646,0642,0644,0020,0645,0644,0643,064A,0629,0020,0645,0628,0646,0649,0020,0645,0639,0020,0627,0644,0627,062D,062A,0641,0627,0638,0020,0628,0627,0644,062A,0627,0631,064A,062E"), 6500: y = 500
        Call AT_AddCombo(n, "cboBuildingID", AT_U("0627,0644,0645,0628,0646,0649"), y, "SELECT BuildingID, BuildingCode & ' - ' & BuildingName FROM qryAT_BuildingCurrentOwners ORDER BY BuildingCode;", 2, "0;5"): y = y + 700
        Call AT_AddCombo(n, "cboNewOwnerID", AT_U("0627,0644,0645,0627,0644,0643,0020,0627,0644,062C,062F,064A,062F"), y, "SELECT OwnerID, OwnerCode & ' - ' & FullName FROM tblOwners WHERE AdminStatus='ACTIVE' ORDER BY FullName;", 2, "0;5"): y = y + 700
        AT_AddFieldPair n, "txtTransferDate", AT_U("062A,0627,0631,064A,062E,0020,0633,0631,064A,0627,0646,0020,0627,0644,0646,0642,0644"), y, 6500: y = y + 600
        AT_AddFieldPair n, "txtTransferDocument", AT_U("0645,0631,062C,0639,0020,0645,0633,062A,0646,062F,0020,0627,0644,0646,0642,0644"), y, 6500: y = y + 600
        AT_AddFieldPair n, "txtTransferReason", AT_U("0633,0628,0628,0020,0627,0644,0646,0642,0644,0020,0028,0625,0644,0632,0627,0645,064A,0029"), y, 6500: y = y + 800
        AT_AddButton n, "cmdTransferOwnership", AT_U("062A,0623,0643,064A,062F,0020,0646,0642,0644,0020,0627,0644,0645,0644,0643,064A,0629"), 3200, y, 3500, 600, "=AT_UI_TransferOwnership()"
        AT_SaveAndCloseForm n, "frmOwnershipTransfer"
    End If
    If Not AT_FormExists("frmUnitExemption") Then
        Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("0637,0644,0628,0020,0625,0639,0641,0627,0621,0020,0634,0642,0629"), 6500: y = 450
        Call AT_AddCombo(n, "cboUnitID", AT_U("0627,0644,0634,0642,0629"), y, "SELECT U.UnitID, B.BuildingCode & ' / ' & U.UnitCode FROM tblUnits AS U INNER JOIN tblBuildings AS B ON U.BuildingID=B.BuildingID WHERE U.IsActive=True ORDER BY B.BuildingCode,U.UnitCode;", 2, "0;5"): y = y + 700
        AT_AddFieldPair n, "txtEffectiveFrom", AT_U("0628,062F,0627,064A,0629,0020,0627,0644,0625,0639,0641,0627,0621"), y, 6500: y = y + 600
        AT_AddFieldPair n, "txtEffectiveTo", AT_U("0646,0647,0627,064A,0629,0020,0627,0644,0625,0639,0641,0627,0621,0020,0028,0627,062E,062A,064A,0627,0631,064A,0029"), y, 6500: y = y + 600
        AT_AddFieldPair n, "txtExemptionReason", AT_U("0633,0628,0628,0020,0627,0644,0625,0639,0641,0627,0621"), y, 6500: y = y + 600
        AT_AddFieldPair n, "txtEvidenceReference", AT_U("0645,0631,062C,0639,0020,0627,0644,0625,062B,0628,0627,062A"), y, 6500: y = y + 600
        AT_AddFieldPair n, "txtExemptionRequestID", AT_U("0631,0642,0645,0020,0627,0644,0637,0644,0628"), y, 6500: f.Controls("txtExemptionRequestID").Locked = True: y = y + 700
        AT_AddButton n, "cmdRequestExemption", AT_U("0625,0631,0633,0627,0644,0020,0627,0644,0637,0644,0628,0020,0644,0644,0645,0631,0627,062C,0639,0629"), 3000, y, 3700, 600, "=AT_UI_RequestExemption()"
        AT_SaveAndCloseForm n, "frmUnitExemption"
    End If
    If Not AT_FormExists("frmExemptionReviewAction") Then
        Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("0627,0644,0628,062A,0020,0641,064A,0020,0637,0644,0628,0020,0625,0639,0641,0627,0621,0020,0634,0642,0629"), 4000: y = 500
        AT_AddFieldPair n, "txtRequestID", AT_U("0631,0642,0645,0020,0637,0644,0628,0020,0627,0644,0625,0639,0641,0627,0621"), y, 6500: y = y + 700
        Call AT_AddCombo(n, "cboDecision", AT_U("0627,0644,0642,0631,0627,0631"), y, AT_U("0053,0045,004C,0045,0043,0054,0020,0027,0041,0050,0050,0052,004F,0056,0045,0027,002C,0027,0645,0648,0627,0641,0642,0629,0027,0020,0055,004E,0049,004F,004E,0020,0053,0045,004C,0045,0043,0054,0020,0027,0052,0045,004A,0045,0043,0054,0027,002C,0027,0631,0641,0636,0027,003B"), 2, "0;3"): y = y + 700
        AT_AddFieldPair n, "txtDecisionReason", AT_U("0633,0628,0628,0020,0627,0644,0642,0631,0627,0631"), y, 6500: y = y + 900
        AT_AddButton n, "cmdDecideExemption", AT_U("062D,0641,0638,0020,0642,0631,0627,0631,0020,0627,0644,0625,0639,0641,0627,0621"), 3300, y, 3200, 600, "=AT_UI_DecideExemption()"
        AT_SaveAndCloseForm n, "frmExemptionReviewAction"
    End If
End Sub

Private Sub AT_CreateCampaignForms()
    Dim f As Form, n As String, y As Long
    If AT_FormExists("frmCampaignCreate") Then Exit Sub
    Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("0625,0639,062F,0627,062F,0020,062D,0645,0644,0629,0020,0645,0633,0627,0647,0645,0629,0020,0625,0636,0627,0641,064A,0629"), 9300: y = 350
    AT_AddFieldPair n, "txtCampaignCode", AT_U("0631,0645,0632,0020,0627,0644,062D,0645,0644,0629"), y, 6500: y = y + 550
    AT_AddFieldPair n, "txtCampaignName", AT_U("0627,0633,0645,0020,0627,0644,062D,0645,0644,0629"), y, 6500: y = y + 550
    Call AT_AddCombo(n, "cboContributionTypeID", AT_U("0646,0648,0639,0020,0627,0644,0645,0633,0627,0647,0645,0629"), y, "SELECT ContributionTypeID, TypeName FROM tblContributionTypes WHERE IsActive=True ORDER BY TypeName;", 2, "0;4"): y = y + 650
    Call AT_AddCombo(n, "cboTargetScope", AT_U("0646,0637,0627,0642,0020,0627,0644,0627,0633,062A,0647,062F,0627,0641"), y, AT_U("0053,0045,004C,0045,0043,0054,0020,0027,0041,004C,004C,0027,002C,0027,0643,0627,0645,0644,0020,0627,0644,062D,064A,0027,0020,0055,004E,0049,004F,004E,0020,0053,0045,004C,0045,0043,0054,0020,0027,0042,0055,0049,004C,0044,0049,004E,0047,0053,0027,002C,0027,0645,0628,0627,0646,064D,0020,0645,062D,062F,062F,0629,0027,0020,0055,004E,0049,004F,004E,0020,0053,0045,004C,0045,0043,0054,0020,0027,0055,004E,0049,0054,0053,0027,002C,0027,0634,0642,0642,0020,0645,062D,062F,062F,0629,0027,003B"), 2, "0;4"): y = y + 650
    AT_AddFieldPair n, "txtAmountPerUnit", AT_U("0642,064A,0645,0629,0020,0627,0644,0645,0633,0627,0647,0645,0629,0020,0644,0643,0644,0020,0634,0642,0629"), y, 6500: y = y + 550
    AT_AddFieldPair n, "txtPeriodStart", AT_U("0628,062F,0627,064A,0629,0020,0627,0644,0641,062A,0631,0629,0020,0028,0627,062E,062A,064A,0627,0631,064A,0029"), y, 6500: y = y + 550
    AT_AddFieldPair n, "txtPeriodEnd", AT_U("0646,0647,0627,064A,0629,0020,0627,0644,0641,062A,0631,0629,0020,0028,0627,062E,062A,064A,0627,0631,064A,0029"), y, 6500: y = y + 550
    AT_AddFieldPair n, "txtDueDate", AT_U("062A,0627,0631,064A,062E,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642"), y, 6500: y = y + 550
    AT_AddFieldPair n, "txtCampaignReason", AT_U("0633,0628,0628,002F,0648,0635,0641,0020,0627,0644,062D,0645,0644,0629"), y, 6500: y = y + 550
    AT_AddFieldPair n, "txtCampaignID", AT_U("0627,0644,0645,0639,0631,0641,0020,0627,0644,062F,0627,062E,0644,064A"), y, 6500: f.Controls("txtCampaignID").Locked = True: y = y + 650
    AT_AddButton n, "cmdSaveCampaign", AT_U("062D,0641,0638,0020,0627,0644,062D,0645,0644,0629"), 2500, y, 2200, 600, "=AT_UI_SaveCampaign()"
    AT_AddButton n, "cmdGenerateCampaign", AT_U("0625,0646,0634,0627,0621,0020,0627,0644,0645,0633,0648,062F,0627,062A"), 4900, y, 2300, 600, "=AT_UI_GenerateCampaign()"
    AT_AddButton n, "cmdApproveCampaign", AT_U("0627,0639,062A,0645,0627,062F,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,0627,062A"), 7400, y, 2300, 600, "=AT_UI_ApproveCampaign()"
    AT_SaveAndCloseForm n, "frmCampaignCreate"
End Sub

Private Sub AT_CreateExpenseForms()
    Dim f As Form, n As String, y As Long
    If Not AT_FormExists("frmExpenseCreate") Then
        Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("062A,0633,062C,064A,0644,0020,0645,0635,0631,0648,0641,0020,0642,0628,0644,0020,0627,0644,062F,0641,0639"), 5600: y = 450
        AT_AddFieldPair n, "txtExpenseDate", AT_U("062A,0627,0631,064A,062E,0020,0627,0644,0645,0635,0631,0648,0641"), y, 6500: y = y + 650
        Call AT_AddCombo(n, "cboExpenseCategoryID", AT_U("062A,0635,0646,064A,0641,0020,0627,0644,0645,0635,0631,0648,0641"), y, "SELECT ExpenseCategoryID, CategoryName FROM tblExpenseCategories WHERE IsActive=True ORDER BY CategoryName;", 2, "0;4"): y = y + 650
        AT_AddFieldPair n, "txtBeneficiary", AT_U("0627,0644,0645,0633,062A,0641,064A,062F"), y, 6500: y = y + 650
        AT_AddFieldPair n, "txtExpenseDescription", AT_U("0648,0635,0641,0020,0627,0644,0645,0635,0631,0648,0641"), y, 6500: y = y + 650
        AT_AddFieldPair n, "txtExpenseAmount", AT_U("0627,0644,0645,0628,0644,063A,0020,0628,0627,0644,0631,064A,0627,0644"), y, 6500: y = y + 650
        AT_AddFieldPair n, "txtExpenseID", AT_U("0627,0644,0645,0639,0631,0641,0020,0627,0644,062F,0627,062E,0644,064A"), y, 6500: f.Controls("txtExpenseID").Locked = True: y = y + 750
        AT_AddButton n, "cmdSaveExpense", AT_U("062D,0641,0638,0020,0627,0644,0645,0635,0631,0648,0641,0020,0643,0645,0633,0648,062F,0629"), 3300, y, 3500, 600, "=AT_UI_SaveExpense()"
        AT_SaveAndCloseForm n, "frmExpenseCreate"
    End If
    If Not AT_FormExists("frmExpensePay") Then
        Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("0633,062F,0627,062F,0020,0645,0635,0631,0648,0641,0020,0645,0633,062C,0644"), 7000: y = 350
        Call AT_AddCombo(n, "cboExpenseID", AT_U("0627,0644,0645,0635,0631,0648,0641,0020,063A,064A,0631,0020,0627,0644,0645,062F,0641,0648,0639"), y, "SELECT ExpenseID, ExpenseNo & ' - ' & BeneficiaryName & ' - ' & Format(Amount,'#,##0') FROM tblExpenses WHERE StatusCode='DRAFT' ORDER BY ExpenseDate;", 2, "0;7"): y = y + 650
        Call AT_AddCombo(n, "cboAccountID", AT_U("062D,0633,0627,0628,0020,0627,0644,0635,0631,0641"), y, "SELECT AccountID, AccountName FROM tblAccounts WHERE IsActive=True ORDER BY AccountName;", 2, "0;4"): y = y + 650
        Call AT_AddCombo(n, "cboPaymentMethodID", AT_U("0637,0631,064A,0642,0629,0020,0627,0644,062F,0641,0639"), y, "SELECT PaymentMethodID, MethodName FROM tblPaymentMethods WHERE IsActive=True ORDER BY MethodName;", 2, "0;4"): y = y + 650
        AT_AddFieldPair n, "txtRecipient", AT_U("0627,0633,0645,0020,0645,0633,062A,0644,0645,0020,0627,0644,0645,0628,0644,063A"), y, 6500: y = y + 650
        AT_AddFieldPair n, "txtPaymentNotes", AT_U("0645,0644,0627,062D,0638,0627,062A,0020,0633,0646,062F,0020,0627,0644,0635,0631,0641"), y, 6500: y = y + 650
        AT_AddFieldPair n, "txtPaymentID", AT_U("0631,0642,0645,0020,0633,0646,062F,0020,0627,0644,0635,0631,0641,0020,0627,0644,062F,0627,062E,0644,064A"), y, 6500: f.Controls("txtPaymentID").Locked = True: y = y + 650
        AT_AddButton n, "cmdMakeExpensePayment", AT_U("0625,0646,0634,0627,0621,0020,0633,0646,062F,0020,0627,0644,0635,0631,0641"), 1700, y, 2500, 600, "=AT_UI_SaveExpensePayment()"
        AT_AddButton n, "cmdConfirmExpensePayment", AT_U("062A,0623,0643,064A,062F,0020,0627,0644,062F,0641,0639,0020,0627,0644,0641,0639,0644,064A"), 4600, y, 2900, 600, "=AT_UI_ApproveExpensePayment()"
        AT_AddButton n, "cmdAttachExpensePayment", AT_U("0625,0631,0641,0627,0642,0020,0645,0644,0641"), 7800, y, 2400, 600, "=AT_UI_AttachExpensePaymentFile()"
        AT_SaveAndCloseForm n, "frmExpensePay"
    End If
    If Not AT_FormExists("frmReceiptReprint") Then
        Set f = CreateForm: n = f.Name: AT_SetFormBase f, AT_U("0625,0639,0627,062F,0629,0020,0637,0628,0627,0639,0629,0020,0625,064A,0635,0627,0644,0020,0633,0627,0628,0642"), 2500: y = 450
        AT_AddFieldPair n, "txtReceiptID", AT_U("0627,0644,0645,0639,0631,0641,0020,0627,0644,062F,0627,062E,0644,064A,0020,0644,0644,0625,064A,0635,0627,0644"), y, 6500: y = y + 950
        AT_AddButton n, "cmdReprintReceipt", AT_U("0645,0639,0627,064A,0646,0629,002F,0637,0628,0627,0639,0629,0020,0627,0644,0625,064A,0635,0627,0644"), 3300, y, 3300, 600, "=AT_UI_ReprintReceipt()"
        AT_SaveAndCloseForm n, "frmReceiptReprint"
    End If
End Sub

Private Sub AT_SetFormBase(ByVal f As Form, ByVal captionText As String, ByVal heightValue As Long)
    f.Caption = captionText: f.Width = 10500: f.Section(acDetail).Height = heightValue
    f.Section(acDetail).BackColor = RGB(248, 250, 252): f.NavigationButtons = False: f.RecordSelectors = False
    ' Unbound entry controls must remain editable; read-only datasheets set their own lock flags.
    On Error Resume Next: f.RightToLeft = True: On Error GoTo 0
End Sub

Private Sub AT_AddLabel(ByVal formName As String, ByVal controlName As String, ByVal captionText As String, ByVal leftValue As Long, ByVal topValue As Long, ByVal widthValue As Long, ByVal heightValue As Long, Optional ByVal fontSizeValue As Long = 10, Optional ByVal isBold As Boolean = False)
    Dim c As Control
    Set c = CreateControl(formName, acLabel, acDetail, , , leftValue, topValue, widthValue, heightValue)
    c.Name = controlName: c.Caption = captionText: c.FontName = "Tahoma": c.FontSize = fontSizeValue: c.FontBold = isBold: c.TextAlign = 3
End Sub

Private Sub AT_AddFieldPair(ByVal formName As String, ByVal controlName As String, ByVal captionText As String, ByVal topValue As Long, ByVal widthValue As Long)
    Dim c As Control
    AT_AddLabel formName, "lbl_" & controlName, captionText, 550, topValue, 2150, 360, 10, False
    Set c = CreateControl(formName, acTextBox, acDetail, , , 2900, topValue, widthValue, 380)
    c.Name = controlName: c.FontName = "Tahoma": c.FontSize = 10: c.DecimalPlaces = 0: c.TextAlign = 3
End Sub

Private Function AT_AddCombo(ByVal formName As String, ByVal controlName As String, ByVal captionText As String, ByVal topValue As Long, ByVal rowSource As String, ByVal columnCountValue As Long, ByVal columnWidthsValue As String) As Control
    Dim c As Control
    AT_AddLabel formName, "lbl_" & controlName, captionText, 550, topValue, 2150, 360, 10, False
    Set c = CreateControl(formName, acComboBox, acDetail, , , 2900, topValue, 6500, 380)
    c.Name = controlName: c.RowSourceType = "Table/Query": c.RowSource = rowSource
    c.ColumnCount = columnCountValue: c.BoundColumn = 1: c.ColumnWidths = columnWidthsValue: c.LimitToList = True
    c.FontName = "Tahoma": c.FontSize = 10: c.TextAlign = 3
    Set AT_AddCombo = c
End Function

Private Sub AT_AddButton(ByVal formName As String, ByVal controlName As String, ByVal captionText As String, ByVal leftValue As Long, ByVal topValue As Long, ByVal widthValue As Long, ByVal heightValue As Long, ByVal onClickExpression As String)
    Dim c As Control
    Set c = CreateControl(formName, acCommandButton, acDetail, , , leftValue, topValue, widthValue, heightValue)
    c.Name = controlName: c.Caption = captionText: c.OnClick = onClickExpression
    c.FontName = "Tahoma": c.FontSize = 10: c.FontBold = True
    c.BackColor = RGB(16, 185, 129): c.ForeColor = RGB(255, 255, 255)
End Sub
