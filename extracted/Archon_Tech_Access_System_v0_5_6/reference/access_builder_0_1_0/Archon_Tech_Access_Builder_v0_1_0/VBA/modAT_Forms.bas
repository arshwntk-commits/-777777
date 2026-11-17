Attribute VB_Name = "modAT_Forms"
Option Compare Database
Option Explicit

Private Const AT_AC_FORM As Long = 2
Private Const AT_AC_LABEL As Long = 100
Private Const AT_AC_TEXTBOX As Long = 109
Private Const AT_AC_COMBOBOX As Long = 111
Private Const AT_AC_BUTTON As Long = 104
Private Const AT_AC_DETAIL As Long = 0
Private Const AT_AC_SAVE_YES As Long = 1
Private Const AT_AC_NEW_RECORD As Long = 2

Public Sub AT_BuildForms()
    On Error GoTo BuildFail
    AT_LogBuild "Forms", "Started", "Creating generated forms"

    AT_CreateHomeForm

    AT_CreateBoundForm "frmAT_Owners", "tblAT_Owners", AT_U("0633,062C,0644") & " " & AT_U("0627,0644,0645,0644,0627,0643"), Array( _
        Array("OwnerCode", AT_U("0631,0642,0645") & " " & AT_U("0627,0644,0645,0627,0644,0643"), "TEXT"), Array("FullName", AT_U("0627,0633,0645") & " " & AT_U("0627,0644,0645,0627,0644,0643"), "TEXT"), Array("Phone", AT_U("0631,0642,0645") & " " & AT_U("0627,0644,0647,0627,062A,0641"), "TEXT"), _
        Array("AltPhone", AT_U("0647,0627,062A,0641") & " " & AT_U("0628,062F,064A,0644"), "TEXT"), Array("Address", AT_U("0627,0644,0639,0646,0648,0627,0646"), "TEXT"), Array("OwnerStatus", AT_U("062D,0627,0644,0629") & " " & AT_U("0627,0644,0645,0627,0644,0643"), "TEXT"), Array("Notes", AT_U("0645,0644,0627,062D,0638,0627,062A"), "TEXT")), _
        Array(Array("NEW_OWNER", AT_U("0645,0627,0644,0643") & " " & AT_U("062C,062F,064A,062F")), Array("SAVE_RECORD", AT_U("062D,0641,0638") & " " & AT_U("0627,0644,0633,062C,0644")), Array("OWNER_BALANCES", AT_U("0623,0631,0635,062F,0629") & " " & AT_U("0627,0644,0645,0644,0627,0643"))))

    AT_CreateBoundForm "frmAT_Units", "tblAT_Units", AT_U("0627,0644,0634,0642,0642") & " " & AT_U("0648,0627,0644,0648,062D,062F,0627,062A"), Array( _
        Array("UnitCode", AT_U("0631,0642,0645") & " " & AT_U("0627,0644,0648,062D,062F,0629"), "TEXT"), Array("OwnerID", AT_U("0627,0644,0645,0627,0644,0643"), "OWNER"), Array("BuildingNo", AT_U("0631,0642,0645") & " " & AT_U("0627,0644,0645,0628,0646,0649"), "TEXT"), _
        Array("ApartmentNo", AT_U("0631,0642,0645") & " " & AT_U("0627,0644,0634,0642,0629"), "TEXT"), Array("FloorNo", AT_U("0627,0644,0637,0627,0628,0642"), "TEXT"), Array("ContributionAmount", AT_U("0642,064A,0645,0629") & " " & AT_U("0627,0644,0645,0633,0627,0647,0645,0629"), "TEXT"), _
        Array("IsActive", AT_U("0648,062D,062F,0629") & " " & AT_U("0646,0634,0637,0629"), "TEXT"), Array("Notes", AT_U("0645,0644,0627,062D,0638,0627,062A"), "TEXT")), _
        Array(Array("NEW_UNIT", AT_U("0648,062D,062F,0629") & " " & AT_U("062C,062F,064A,062F,0629")), Array("SAVE_RECORD", AT_U("062D,0641,0638") & " " & AT_U("0627,0644,0633,062C,0644")), Array("OPEN_OWNERS", AT_U("0633,062C,0644") & " " & AT_U("0627,0644,0645,0644,0627,0643"))))

    AT_CreateBoundForm "frmAT_Charges", "tblAT_Charges", AT_U("0627,0644,0627,0633,062A,062D,0642,0627,0642,0627,062A"), Array( _
        Array("ChargeCode", AT_U("0631,0642,0645") & " " & AT_U("0627,0644,0627,0633,062A,062D,0642,0627,0642"), "TEXT"), Array("OwnerID", AT_U("0627,0644,0645,0627,0644,0643"), "OWNER"), Array("UnitID", AT_U("0627,0644,0648,062D,062F,0629"), "UNIT"), _
        Array("ChargeDate", AT_U("062A,0627,0631,064A,062E") & " " & AT_U("0627,0644,0627,0633,062A,062D,0642,0627,0642"), "TEXT"), Array("DueDate", AT_U("062A,0627,0631,064A,062E") & " " & AT_U("0627,0644,0627,0633,062A,062D,0642,0627,0642") & " " & AT_U("0627,0644,0646,0647,0627,0626,064A"), "TEXT"), Array("Description", AT_U("0627,0644,0628,064A,0627,0646"), "TEXT"), _
        Array("OriginalAmount", AT_U("0627,0644,0645,0628,0644,063A") & " " & AT_U("0627,0644,0623,0635,0644,064A"), "TEXT"), Array("CurrentAmount", AT_U("0627,0644,0645,0628,0644,063A") & " " & AT_U("0627,0644,062D,0627,0644,064A"), "TEXT"), Array("Status", AT_U("0627,0644,062D,0627,0644,0629"), "TEXT")), _
        Array(Array("NEW_CHARGE", AT_U("0627,0633,062A,062D,0642,0627,0642") & " " & AT_U("064A,062F,0648,064A")), Array("GENERATE_INITIAL_CHARGES", AT_U("0625,0646,0634,0627,0621") & " " & AT_U("0627,0633,062A,062D,0642,0627,0642,0627,062A") & " " & AT_U("0627,0644,0648,062D,062F,0627,062A")), Array("SAVE_RECORD", AT_U("062D,0641,0638") & " " & AT_U("0627,0644,0633,062C,0644")), Array("OWNER_BALANCES", AT_U("0623,0631,0635,062F,0629") & " " & AT_U("0627,0644,0645,0644,0627,0643"))))

    AT_CreateBoundForm "frmAT_Receipts", "tblAT_Receipts", AT_U("0625,064A,0635,0627,0644,0627,062A") & " " & AT_U("0627,0644,062A,062D,0635,064A,0644"), Array( _
        Array("ReceiptNo", AT_U("0631,0642,0645") & " " & AT_U("0627,0644,0625,064A,0635,0627,0644") & " (" & AT_U("0628,0639,062F") & " " & AT_U("0627,0644,0627,0639,062A,0645,0627,062F") & ")", "TEXT"), Array("OwnerID", AT_U("0627,0644,0645,0627,0644,0643"), "OWNER"), Array("ReceiptDate", AT_U("062A,0627,0631,064A,062E") & " " & AT_U("0627,0644,0625,064A,0635,0627,0644"), "TEXT"), _
        Array("Amount", AT_U("0627,0644,0645,0628,0644,063A") & " " & AT_U("0627,0644,0645,0633,062A,0644,0645"), "TEXT"), Array("PaymentMethodID", AT_U("0637,0631,064A,0642,0629") & " " & AT_U("0627,0644,062F,0641,0639"), "METHOD"), Array("AccountID", AT_U("0627,0644,062D,0633,0627,0628") & " " & AT_U("0627,0644,0645,0633,062A,0644,0645"), "ACCOUNT"), _
        Array("PayerName", AT_U("0627,0633,0645") & " " & AT_U("0627,0644,062F,0627,0641,0639"), "TEXT"), Array("CollectorName", AT_U("0627,0633,0645") & " " & AT_U("0627,0644,0645,062D,0635,0644"), "TEXT"), Array("Status", AT_U("0627,0644,062D,0627,0644,0629"), "TEXT"), Array("Notes", AT_U("0645,0644,0627,062D,0638,0627,062A"), "TEXT")), _
        Array(Array("NEW_RECEIPT", AT_U("0625,064A,0635,0627,0644") & " " & AT_U("062C,062F,064A,062F")), Array("SAVE_RECORD", AT_U("062D,0641,0638") & " " & AT_U("0627,0644,0645,0633,0648,062F,0629")), Array("APPROVE_RECEIPT", AT_U("0627,0639,062A,0645,0627,062F") & " " & AT_U("0627,0644,0625,064A,0635,0627,0644")), Array("REVERSE_RECEIPT", AT_U("0639,0643,0633") & " " & AT_U("0627,0644,0625,064A,0635,0627,0644")), Array("APPLY_OWNER_CREDIT", AT_U("0627,0633,062A,062E,062F,0627,0645") & " " & AT_U("0631,0635,064A,062F") & " " & AT_U("0627,0644,0645,0627,0644,0643"))))

    AT_CreateBoundForm "frmAT_Expenses", "tblAT_Expenses", AT_U("062A,0633,062C,064A,0644") & " " & AT_U("0627,0644,0645,0635,0631,0648,0641,0627,062A"), Array( _
        Array("ExpenseNo", AT_U("0631,0642,0645") & " " & AT_U("0627,0644,0645,0635,0631,0648,0641"), "TEXT"), Array("CategoryID", AT_U("0627,0644,062A,0635,0646,064A,0641"), "CATEGORY"), Array("ExpenseDate", AT_U("062A,0627,0631,064A,062E") & " " & AT_U("0627,0644,0645,0635,0631,0648,0641"), "TEXT"), _
        Array("Beneficiary", AT_U("0627,0644,0645,0633,062A,0641,064A,062F"), "TEXT"), Array("Description", AT_U("062A,0641,0627,0635,064A,0644") & " " & AT_U("0627,0644,0645,0635,0631,0648,0641"), "TEXT"), Array("Amount", AT_U("0627,0644,0645,0628,0644,063A"), "TEXT"), Array("Status", AT_U("0627,0644,062D,0627,0644,0629"), "TEXT")), _
        Array(Array("NEW_EXPENSE", AT_U("0645,0635,0631,0648,0641") & " " & AT_U("062C,062F,064A,062F")), Array("SAVE_RECORD", AT_U("062D,0641,0638") & " " & AT_U("0627,0644,0645,0635,0631,0648,0641")), Array("OPEN_VOUCHERS", AT_U("0633,0646,062F,0627,062A") & " " & AT_U("0627,0644,0635,0631,0641"))))

    AT_CreateBoundForm "frmAT_Vouchers", "tblAT_PaymentVouchers", AT_U("0633,0646,062F,0627,062A") & " " & AT_U("0627,0644,0635,0631,0641") & " " & AT_U("0648,0631,062F") & " " & AT_U("0627,0644,0631,0635,064A,062F"), Array( _
        Array("VoucherNo", AT_U("0631,0642,0645") & " " & AT_U("0627,0644,0633,0646,062F") & " (" & AT_U("0628,0639,062F") & " " & AT_U("0627,0644,0627,0639,062A,0645,0627,062F") & ")", "TEXT"), Array("VoucherDate", AT_U("062A,0627,0631,064A,062E") & " " & AT_U("0627,0644,0633,0646,062F"), "TEXT"), Array("VoucherType", AT_U("0646,0648,0639") & " " & AT_U("0627,0644,0633,0646,062F") & " EXPENSE / OWNER_REFUND", "TEXT"), _
        Array("ExpenseID", AT_U("0627,0644,0645,0635,0631,0648,0641") & " " & AT_U("0627,0644,0645,0631,062A,0628,0637"), "EXPENSE"), Array("OwnerID", AT_U("0627,0644,0645,0627,0644,0643") & " " & AT_U("0639,0646,062F") & " " & AT_U("0631,062F") & " " & AT_U("0627,0644,0631,0635,064A,062F"), "OWNER"), Array("PayeeName", AT_U("0627,0644,0645,0633,062A,0641,064A,062F"), "TEXT"), _
        Array("Amount", AT_U("0627,0644,0645,0628,0644,063A"), "TEXT"), Array("PaymentMethodID", AT_U("0637,0631,064A,0642,0629") & " " & AT_U("0627,0644,062F,0641,0639"), "METHOD"), Array("AccountID", AT_U("062D,0633,0627,0628") & " " & AT_U("0627,0644,0635,0631,0641"), "ACCOUNT"), _
        Array("Status", AT_U("0627,0644,062D,0627,0644,0629"), "TEXT"), Array("Notes", AT_U("0645,0644,0627,062D,0638,0627,062A"), "TEXT")), _
        Array(Array("NEW_VOUCHER", AT_U("0633,0646,062F") & " " & AT_U("062C,062F,064A,062F")), Array("SAVE_RECORD", AT_U("062D,0641,0638") & " " & AT_U("0627,0644,0645,0633,0648,062F,0629")), Array("APPROVE_VOUCHER", AT_U("0627,0639,062A,0645,0627,062F") & " " & AT_U("0627,0644,0635,0631,0641"))))

    AT_CreateBoundForm "frmAT_Transfers", "tblAT_Transfers", AT_U("0627,0644,062A,062D,0648,064A,0644,0627,062A") & " " & AT_U("0627,0644,062F,0627,062E,0644,064A,0629"), Array( _
        Array("TransferNo", AT_U("0631,0642,0645") & " " & AT_U("0627,0644,062A,062D,0648,064A,0644"), "TEXT"), Array("TransferDate", AT_U("062A,0627,0631,064A,062E") & " " & AT_U("0627,0644,062A,062D,0648,064A,0644"), "TEXT"), Array("FromAccountID", AT_U("0645,0646") & " " & AT_U("062D,0633,0627,0628"), "ACCOUNT"), _
        Array("ToAccountID", AT_U("0625,0644,0649") & " " & AT_U("062D,0633,0627,0628"), "ACCOUNT"), Array("Amount", AT_U("0645,0628,0644,063A") & " " & AT_U("0627,0644,062A,062D,0648,064A,0644"), "TEXT"), Array("FeeAmount", AT_U("0631,0633,0648,0645") & " " & AT_U("0627,0644,062A,062D,0648,064A,0644"), "TEXT"), _
        Array("Status", AT_U("0627,0644,062D,0627,0644,0629"), "TEXT"), Array("Notes", AT_U("0645,0644,0627,062D,0638,0627,062A"), "TEXT")), _
        Array(Array("NEW_TRANSFER", AT_U("062A,062D,0648,064A,0644") & " " & AT_U("062C,062F,064A,062F")), Array("SAVE_RECORD", AT_U("062D,0641,0638") & " " & AT_U("0627,0644,0645,0633,0648,062F,0629")), Array("START_TRANSFER", AT_U("0628,062F,0621") & " " & AT_U("0648,062D,062C,0632") & " " & AT_U("0627,0644,062A,062D,0648,064A,0644")), Array("COMPLETE_TRANSFER", AT_U("062A,0623,0643,064A,062F") & " " & AT_U("0648,0635,0648,0644") & " " & AT_U("0627,0644,062A,062D,0648,064A,0644"))))

    AT_CreateBoundForm "frmAT_Accounts", "tblAT_FinancialAccounts", AT_U("0627,0644,062D,0633,0627,0628,0627,062A") & " " & AT_U("0627,0644,0645,0627,0644,064A,0629"), Array( _
        Array("AccountCode", AT_U("0631,0645,0632") & " " & AT_U("0627,0644,062D,0633,0627,0628"), "TEXT"), Array("AccountName", AT_U("0627,0633,0645") & " " & AT_U("0627,0644,062D,0633,0627,0628"), "TEXT"), Array("AccountType", AT_U("0646,0648,0639") & " " & AT_U("0627,0644,062D,0633,0627,0628"), "TEXT"), _
        Array("InstitutionName", AT_U("0627,0644,0628,0646,0643") & " / " & AT_U("0627,0644,0645,062D,0641,0638,0629"), "TEXT"), Array("AccountReference", AT_U("0645,0631,062C,0639") & " " & AT_U("0627,0644,062D,0633,0627,0628"), "TEXT"), Array("IsActive", AT_U("0646,0634,0637"), "TEXT")), _
        Array(Array("NEW_ACCOUNT", AT_U("062D,0633,0627,0628") & " " & AT_U("062C,062F,064A,062F")), Array("SAVE_RECORD", AT_U("062D,0641,0638") & " " & AT_U("0627,0644,062D,0633,0627,0628")), Array("ACCOUNT_BALANCES", AT_U("0623,0631,0635,062F,0629") & " " & AT_U("0627,0644,062D,0633,0627,0628,0627,062A"))))

    AT_CreateBoundForm "frmAT_PaymentMethods", "tblAT_PaymentMethods", AT_U("0637,0631,0642") & " " & AT_U("0627,0644,062F,0641,0639"), Array( _
        Array("MethodName", AT_U("0627,0633,0645") & " " & AT_U("0637,0631,064A,0642,0629") & " " & AT_U("0627,0644,062F,0641,0639"), "TEXT"), Array("IsActive", AT_U("0646,0634,0637,0629"), "TEXT"), Array("Notes", AT_U("0645,0644,0627,062D,0638,0627,062A"), "TEXT")), _
        Array(Array("NEW_METHOD", AT_U("0637,0631,064A,0642,0629") & " " & AT_U("062C,062F,064A,062F,0629")), Array("SAVE_RECORD", AT_U("062D,0641,0638"))))

    AT_CreateBoundForm "frmAT_ExpenseCategories", "tblAT_ExpenseCategories", AT_U("062A,0635,0646,064A,0641,0627,062A") & " " & AT_U("0627,0644,0645,0635,0631,0648,0641,0627,062A"), Array( _
        Array("CategoryName", AT_U("0627,0633,0645") & " " & AT_U("0627,0644,062A,0635,0646,064A,0641"), "TEXT"), Array("IsActive", AT_U("0646,0634,0637"), "TEXT")), _
        Array(Array("NEW_CATEGORY", AT_U("062A,0635,0646,064A,0641") & " " & AT_U("062C,062F,064A,062F")), Array("SAVE_RECORD", AT_U("062D,0641,0638"))))

    AT_CreateBoundForm "frmAT_Discounts", "tblAT_DiscountRequests", AT_U("0637,0644,0628,0627,062A") & " " & AT_U("0627,0644,062E,0635,0645") & " " & AT_U("0648,0627,0644,0625,0639,0641,0627,0621"), Array( _
        Array("DiscountNo", AT_U("0631,0642,0645") & " " & AT_U("0627,0644,0637,0644,0628"), "TEXT"), Array("ChargeID", AT_U("0627,0644,0627,0633,062A,062D,0642,0627,0642"), "CHARGE"), Array("RequestDate", AT_U("062A,0627,0631,064A,062E") & " " & AT_U("0627,0644,0637,0644,0628"), "TEXT"), _
        Array("RequestedAmount", AT_U("0642,064A,0645,0629") & " " & AT_U("0627,0644,062E,0635,0645") & " " & AT_U("0627,0644,0645,0637,0644,0648,0628,0629"), "TEXT"), Array("Reason", AT_U("0633,0628,0628") & " " & AT_U("0627,0644,0637,0644,0628"), "TEXT"), Array("Status", AT_U("0627,0644,062D,0627,0644,0629"), "TEXT"), _
        Array("AmountBefore", AT_U("0627,0644,0645,0628,0644,063A") & " " & AT_U("0642,0628,0644") & " " & AT_U("0627,0644,0627,0639,062A,0645,0627,062F"), "TEXT"), Array("AmountAfter", AT_U("0627,0644,0645,0628,0644,063A") & " " & AT_U("0628,0639,062F") & " " & AT_U("0627,0644,0627,0639,062A,0645,0627,062F"), "TEXT")), _
        Array(Array("NEW_DISCOUNT", AT_U("0637,0644,0628") & " " & AT_U("062E,0635,0645") & " " & AT_U("062C,062F,064A,062F")), Array("SAVE_RECORD", AT_U("062D,0641,0638") & " " & AT_U("0627,0644,0637,0644,0628")), Array("APPROVE_DISCOUNT", AT_U("0627,0639,062A,0645,0627,062F") & " " & AT_U("0627,0644,062E,0635,0645"))))

    AT_LogBuild "Forms", "Success", "Core data-entry forms created; existing forms preserved"
    Exit Sub
BuildFail:
    AT_LogBuild "Forms", "Failed", Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "AT_BuildForms", Err.Description
End Sub

Private Sub AT_CreateHomeForm()
    Dim frm As Object, tmpName As String, ctl As Object
    If AT_FormExists("frmAT_Home") Then Exit Sub
    Set frm = CreateForm
    tmpName = frm.Name
    frm.Caption = "Archon Tech | " & AT_U("0646,0638,0627,0645") & " " & AT_U("062A,062D,0635,064A,0644") & " " & AT_U("0631,0633,0648,0645") & " " & AT_U("0627,0644,0645,062C,0627,0631,064A")
    frm.Width = 12500
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.DividingLines = False
    frm.AutoCenter = True
    SetControlIfProperty frm, "RightToLeft", True
    frm.Section(AT_AC_DETAIL).Height = 7800
    frm.Section(AT_AC_DETAIL).BackColor = RGB(248, 250, 252)

    Set ctl = CreateControl(tmpName, AT_AC_LABEL, AT_AC_DETAIL, , "ARCHON TECH", 700, 350, 8500, 650)
    ctl.FontSize = 24
    ctl.FontBold = True
    Set ctl = CreateControl(tmpName, AT_AC_LABEL, AT_AC_DETAIL, , AT_U("0646,0638,0627,0645") & " " & AT_U("0625,062F,0627,0631,0629") & " " & AT_U("062A,062D,0635,064A,0644") & " " & AT_U("0645,0633,0627,0647,0645,0627,062A") & " " & AT_U("0645,0634,0631,0648,0639") & " " & AT_U("0627,0644,0645,062C,0627,0631,064A"), 700, 1050, 9500, 450)
    ctl.FontSize = 15
    ctl.FontBold = True

    AT_AddHomeButton tmpName, "cmdOwners", AT_U("0625,062F,0627,0631,0629") & " " & AT_U("0627,0644,0645,0644,0627,0643"), "OPEN_OWNERS", 700, 1900
    AT_AddHomeButton tmpName, "cmdUnits", AT_U("0627,0644,0648,062D,062F,0627,062A") & " " & AT_U("0648,0627,0644,0634,0642,0642"), "OPEN_UNITS", 3900, 1900
    AT_AddHomeButton tmpName, "cmdCharges", AT_U("0627,0644,0627,0633,062A,062D,0642,0627,0642,0627,062A"), "OPEN_CHARGES", 7100, 1900
    AT_AddHomeButton tmpName, "cmdReceipts", AT_U("0625,064A,0635,0627,0644,0627,062A") & " " & AT_U("0627,0644,062A,062D,0635,064A,0644"), "OPEN_RECEIPTS", 700, 2750
    AT_AddHomeButton tmpName, "cmdExpenses", AT_U("0627,0644,0645,0635,0631,0648,0641,0627,062A"), "OPEN_EXPENSES", 3900, 2750
    AT_AddHomeButton tmpName, "cmdVouchers", AT_U("0633,0646,062F,0627,062A") & " " & AT_U("0627,0644,0635,0631,0641"), "OPEN_VOUCHERS", 7100, 2750
    AT_AddHomeButton tmpName, "cmdTransfers", AT_U("0627,0644,062A,062D,0648,064A,0644,0627,062A"), "OPEN_TRANSFERS", 700, 3600
    AT_AddHomeButton tmpName, "cmdAccounts", AT_U("0627,0644,062D,0633,0627,0628,0627,062A") & " " & AT_U("0627,0644,0645,0627,0644,064A,0629"), "OPEN_ACCOUNTS", 3900, 3600
    AT_AddHomeButton tmpName, "cmdDiscounts", AT_U("0637,0644,0628,0627,062A") & " " & AT_U("0627,0644,062E,0635,0645"), "OPEN_DISCOUNTS", 7100, 3600
    AT_AddHomeButton tmpName, "cmdOwnerBalances", AT_U("0623,0631,0635,062F,0629") & " " & AT_U("0627,0644,0645,0644,0627,0643"), "OWNER_BALANCES", 700, 4650
    AT_AddHomeButton tmpName, "cmdAccountBalances", AT_U("0623,0631,0635,062F,0629") & " " & AT_U("0627,0644,062D,0633,0627,0628,0627,062A"), "ACCOUNT_BALANCES", 3900, 4650
    AT_AddHomeButton tmpName, "cmdSelfTest", AT_U("0641,062D,0635") & " " & AT_U("0627,0644,0646,0638,0627,0645"), "RUN_TESTS", 7100, 4650
    AT_AddHomeButton tmpName, "cmdArrearsReport", AT_U("062A,0642,0631,064A,0631") & " " & AT_U("0627,0644,0645,062A,0623,062E,0631,0627,062A"), "OPEN_ARREARS_REPORT", 700, 5450
    AT_AddHomeButton tmpName, "cmdReceiptReport", AT_U("062A,0642,0631,064A,0631") & " " & AT_U("0627,0644,0625,064A,0635,0627,0644,0627,062A"), "OPEN_RECEIPT_REPORT", 3900, 5450
    AT_AddHomeButton tmpName, "cmdExpensesReport", AT_U("062A,0642,0631,064A,0631") & " " & AT_U("0627,0644,0645,0635,0631,0648,0641,0627,062A"), "OPEN_EXPENSE_REPORT", 7100, 5450
    AT_AddHomeButton tmpName, "cmdMethods", AT_U("0637,0631,0642") & " " & AT_U("0627,0644,062F,0641,0639"), "OPEN_METHODS", 700, 6300
    AT_AddHomeButton tmpName, "cmdCategories", AT_U("062A,0635,0646,064A,0641,0627,062A") & " " & AT_U("0627,0644,0645,0635,0631,0648,0641,0627,062A"), "OPEN_CATEGORIES", 3900, 6300
    AT_AddHomeButton tmpName, "cmdSearchOwners", AT_U("0628,062D,062B") & " " & AT_U("0639,0646") & " " & AT_U("0645,0627,0644,0643"), "SEARCH_OWNER", 7100, 6300
    Set ctl = CreateControl(tmpName, AT_AC_LABEL, AT_AC_DETAIL, , AT_U("0627,0644,0625,0635,062F,0627,0631") & " " & AT_VERSION & " | " & AT_U("062A,064F,0633,062C,0644") & " " & AT_U("0627,0644,0623,0631,0635,062F,0629") & " " & AT_U("0645,0646") & " " & AT_U("0627,0644,062D,0631,0643,0627,062A") & " " & AT_U("0627,0644,0645,0639,062A,0645,062F,0629") & " " & AT_U("0641,0642,0637"), 700, 7100, 11000, 350)
    ctl.FontSize = 9

    DoCmd.Save AT_AC_FORM, tmpName
    DoCmd.Close AT_AC_FORM, tmpName, AT_AC_SAVE_YES
    DoCmd.Rename "frmAT_Home", AT_AC_FORM, tmpName
End Sub

Private Sub AT_AddHomeButton(ByVal formName As String, ByVal controlName As String, ByVal captionText As String, ByVal actionName As String, ByVal leftPos As Long, ByVal topPos As Long)
    Dim ctl As Object
    Set ctl = CreateControl(formName, AT_AC_BUTTON, AT_AC_DETAIL, , captionText, leftPos, topPos, 2700, 600)
    ctl.Name = controlName
    ctl.FontSize = 12
    ctl.BackColor = RGB(16, 185, 129)
    ctl.ForeColor = RGB(255, 255, 255)
    ctl.OnClick = "=AT_FormAction(" & Chr$(34) & actionName & Chr$(34) & ")"
End Sub

Private Sub AT_CreateBoundForm(ByVal formName As String, ByVal tableName As String, ByVal captionText As String, ByVal fieldDefs As Variant, ByVal actions As Variant)
    Dim frm As Object, tmpName As String
    Dim i As Long, fieldSpec As Variant
    Dim y As Long, x As Long, col As Long, fieldName As String
    Dim formHeight As Long
    If AT_FormExists(formName) Then Exit Sub
    Set frm = CreateForm
    tmpName = frm.Name
    frm.RecordSource = tableName
    frm.Caption = captionText
    frm.Width = 14500
    frm.DefaultView = 0
    frm.AllowAdditions = True
    frm.AllowEdits = True
    frm.AllowDeletions = False
    frm.NavigationButtons = True
    frm.RecordSelectors = True
    frm.AutoCenter = True
    SetControlIfProperty frm, "RightToLeft", True
    frm.DividingLines = False
    frm.Section(AT_AC_DETAIL).BackColor = RGB(248, 250, 252)
    frm.Section(AT_AC_DETAIL).Height = 8000

    SetControlIfProperty frm, "FontName", "Tahoma"
    y = 450
    For i = LBound(fieldDefs) To UBound(fieldDefs)
        fieldSpec = fieldDefs(i)
        fieldName = CStr(fieldSpec(0))
        col = i Mod 2
        x = IIf(col = 0, 500, 7300)
        If col = 0 And i > 0 Then y = y + 760
        AT_AddFieldControl tmpName, fieldName, CStr(fieldSpec(1)), CStr(fieldSpec(2)), x, y
    Next i
    formHeight = y + 1050
    AT_AddFormButtons tmpName, actions, 500, formHeight
    frm.Section(AT_AC_DETAIL).Height = formHeight + 1300

    DoCmd.Save AT_AC_FORM, tmpName
    DoCmd.Close AT_AC_FORM, tmpName, AT_AC_SAVE_YES
    DoCmd.Rename formName, AT_AC_FORM, tmpName
End Sub

Private Sub SetControlIfProperty(ByVal obj As Object, ByVal propertyName As String, ByVal propertyValue As Variant)
    On Error Resume Next
    CallByName obj, propertyName, VbLet, propertyValue
    On Error GoTo 0
End Sub

Private Sub AT_AddFieldControl(ByVal formName As String, ByVal fieldName As String, ByVal labelText As String, ByVal fieldKind As String, ByVal leftPos As Long, ByVal topPos As Long)
    Dim lbl As Object, ctl As Object
    Dim ctlType As Long, ctlName As String, rowSource As String
    ctlName = "txt" & fieldName
    ctlType = AT_AC_TEXTBOX
    Select Case UCase$(fieldKind)
        Case "OWNER": ctlType = AT_AC_COMBOBOX: ctlName = "cbo" & fieldName: rowSource = "SELECT OwnerID, OwnerCode & ' - ' & FullName FROM tblAT_Owners ORDER BY FullName;"
        Case "UNIT": ctlType = AT_AC_COMBOBOX: ctlName = "cbo" & fieldName: rowSource = "SELECT UnitID, UnitCode FROM tblAT_Units ORDER BY UnitCode;"
        Case "METHOD": ctlType = AT_AC_COMBOBOX: ctlName = "cbo" & fieldName: rowSource = "SELECT PaymentMethodID, MethodName FROM tblAT_PaymentMethods WHERE IsActive=True ORDER BY MethodName;"
        Case "ACCOUNT": ctlType = AT_AC_COMBOBOX: ctlName = "cbo" & fieldName: rowSource = "SELECT AccountID, AccountName FROM tblAT_FinancialAccounts WHERE IsActive=True ORDER BY AccountName;"
        Case "CATEGORY": ctlType = AT_AC_COMBOBOX: ctlName = "cbo" & fieldName: rowSource = "SELECT CategoryID, CategoryName FROM tblAT_ExpenseCategories WHERE IsActive=True ORDER BY CategoryName;"
        Case "EXPENSE": ctlType = AT_AC_COMBOBOX: ctlName = "cbo" & fieldName: rowSource = "SELECT ExpenseID, Nz(ExpenseNo,'" & AT_U("0645,0633,0648,062F,0629") & "') & ' - ' & Beneficiary FROM tblAT_Expenses WHERE Status<>'Paid' ORDER BY ExpenseDate DESC;"
        Case "CHARGE": ctlType = AT_AC_COMBOBOX: ctlName = "cbo" & fieldName: rowSource = "SELECT ChargeID, ChargeCode & ' - ' & Description FROM tblAT_Charges ORDER BY ChargeDate DESC;"
    End Select

    Set lbl = CreateControl(formName, AT_AC_LABEL, AT_AC_DETAIL, , labelText, leftPos, topPos, 2450, 300)
    lbl.FontName = "Tahoma"
    lbl.FontSize = 10
    lbl.TextAlign = 3
    SetControlIfProperty lbl, "RightToLeft", True
    Set ctl = CreateControl(formName, ctlType, AT_AC_DETAIL, , , leftPos + 2500, topPos - 30, 3600, 360)
    ctl.Name = ctlName
    ctl.ControlSource = fieldName
    ctl.FontName = "Tahoma"
    ctl.FontSize = 10
    ctl.TextAlign = 3
    SetControlIfProperty ctl, "RightToLeft", True
    If ctlType = AT_AC_COMBOBOX Then
        ctl.RowSourceType = "Table/Query"
        ctl.RowSource = rowSource
        ctl.ColumnCount = 2
        ctl.BoundColumn = 1
        ctl.ColumnWidths = "0cm;7cm"
        ctl.LimitToList = True
    End If
    If fieldName = "OwnerCode" Or fieldName = "UnitCode" Or fieldName = "ReceiptNo" Or fieldName = "ChargeCode" Or fieldName = "VoucherNo" Or fieldName = "TransferNo" Or fieldName = "DiscountNo" Or fieldName = "Status" Or fieldName = "AmountBefore" Or fieldName = "AmountAfter" Or fieldName = "AccountCode" Then
        ctl.Locked = True
    End If
End Sub

Private Sub AT_AddFormButtons(ByVal formName As String, ByVal actions As Variant, ByVal leftPos As Long, ByVal topPos As Long)
    Dim i As Long, item As Variant, ctl As Object
    Dim x As Long, rowNo As Long
    For i = LBound(actions) To UBound(actions)
        item = actions(i)
        rowNo = i \ 3
        x = leftPos + (i Mod 3) * 3000
        Set ctl = CreateControl(formName, AT_AC_BUTTON, AT_AC_DETAIL, , CStr(item(1)), x, topPos + rowNo * 650, 2700, 500)
        ctl.Name = "cmdAT_" & CStr(item(0))
        ctl.FontName = "Tahoma"
        ctl.FontSize = 10
        ctl.OnClick = "=AT_FormAction(" & Chr$(34) & CStr(item(0)) & Chr$(34) & ")"
    Next i
End Sub

Public Function AT_FormAction(ByVal actionName As String) As Boolean
    Dim fName As String
    Dim recordID As Long
    Dim resultCount As Long
    Dim amountApplied As Currency
    On Error GoTo ActionFail
    fName = Screen.ActiveForm.Name

    Select Case UCase$(actionName)
        Case "OPEN_OWNERS": DoCmd.OpenForm "frmAT_Owners"
        Case "OPEN_UNITS": DoCmd.OpenForm "frmAT_Units"
        Case "OPEN_CHARGES": DoCmd.OpenForm "frmAT_Charges"
        Case "OPEN_RECEIPTS": DoCmd.OpenForm "frmAT_Receipts"
        Case "OPEN_EXPENSES": DoCmd.OpenForm "frmAT_Expenses"
        Case "OPEN_VOUCHERS": DoCmd.OpenForm "frmAT_Vouchers"
        Case "OPEN_TRANSFERS": DoCmd.OpenForm "frmAT_Transfers"
        Case "OPEN_ACCOUNTS": DoCmd.OpenForm "frmAT_Accounts"
        Case "OPEN_DISCOUNTS": DoCmd.OpenForm "frmAT_Discounts"
        Case "OPEN_METHODS": DoCmd.OpenForm "frmAT_PaymentMethods"
        Case "OPEN_CATEGORIES": DoCmd.OpenForm "frmAT_ExpenseCategories"
        Case "SEARCH_OWNER": AT_SearchOwners
        Case "OWNER_BALANCES": DoCmd.OpenQuery "qAT_OwnerBalances"
        Case "ACCOUNT_BALANCES": DoCmd.OpenQuery "qAT_AccountAvailableBalances"
        Case "RUN_TESTS": Call AT_RunSelfTests
        Case "OPEN_ARREARS_REPORT": Call AT_OpenReport("rptAT_Arrears")
        Case "OPEN_RECEIPT_REPORT": Call AT_OpenReport("rptAT_ReceiptRegister")
        Case "OPEN_EXPENSE_REPORT": Call AT_OpenReport("rptAT_ExpenseRegister")
        Case "SAVE_RECORD"
            If Forms(fName).Dirty Then Forms(fName).Dirty = False
            MsgBox AT_U("062A,0645") & " " & AT_U("062D,0641,0638") & " " & AT_U("0627,0644,0633,062C,0644") & ".", vbInformation, "Archon Tech"
        Case "NEW_OWNER", "NEW_UNIT", "NEW_CHARGE", "NEW_RECEIPT", "NEW_EXPENSE", "NEW_VOUCHER", "NEW_TRANSFER", "NEW_DISCOUNT", "NEW_ACCOUNT", "NEW_METHOD", "NEW_CATEGORY"
            DoCmd.GoToRecord , , AT_AC_NEW_RECORD
            AT_InitializeNewRecord fName, UCase$(actionName)
        Case "GENERATE_INITIAL_CHARGES"
            resultCount = AT_CreateInitialCharges()
            MsgBox AT_U("062A,0645") & " " & AT_U("0625,0646,0634,0627,0621") & " " & resultCount & " " & AT_U("0627,0633,062A,062D,0642,0627,0642,064B,0627") & " " & AT_U("062C,062F,064A,062F,064B,0627") & ". " & AT_U("0644,0627") & " " & AT_U("062A,064F,0646,0634,0623") & " " & AT_U("0627,0633,062A,062D,0642,0627,0642,0627,062A") & " " & AT_U("0645,0643,0631,0631,0629") & " " & AT_U("0644,0644,0648,062D,062F,0629") & " " & AT_U("0646,0641,0633,0647,0627") & ".", vbInformation, "Archon Tech"
            Forms(fName).Requery
        Case "APPROVE_RECEIPT"
            recordID = AT_CurrentRecordID(fName, "ReceiptID")
            If recordID > 0 Then AT_ApproveReceipt recordID Else MsgBox AT_U("0627,062D,0641,0638") & " " & AT_U("0627,0644,0625,064A,0635,0627,0644") & " " & AT_U("0623,0648,0644,064B,0627") & ".", vbExclamation
            Forms(fName).Requery
        Case "REVERSE_RECEIPT"
            recordID = AT_CurrentRecordID(fName, "ReceiptID")
            If recordID > 0 Then
                If AT_ReverseReceipt(recordID, InputBox(AT_U("0623,062F,062E,0644") & " " & AT_U("0633,0628,0628") & " " & AT_U("0639,0643,0633") & " " & AT_U("0627,0644,0625,064A,0635,0627,0644") & ":", AT_U("0633,0628,0628") & " " & AT_U("0627,0644,0639,0643,0633"))) Then MsgBox AT_U("062A,0645") & " " & AT_U("0639,0643,0633") & " " & AT_U("0627,0644,0625,064A,0635,0627,0644") & " " & AT_U("0645,062D,0627,0633,0628,064A,064B,0627") & ".", vbInformation
            Else
                MsgBox AT_U("0627,062E,062A,0631") & " " & AT_U("0625,064A,0635,0627,0644,064B,0627") & " " & AT_U("0645,062D,0641,0648,0638,064B,0627") & ".", vbExclamation
            End If
            Forms(fName).Requery
        Case "APPLY_OWNER_CREDIT"
            recordID = AT_CurrentRecordID(fName, "OwnerID")
            If recordID <= 0 And fName = "frmAT_Receipts" Then recordID = CLng(Nz(Forms(fName)!OwnerID, 0))
            If recordID > 0 Then
                amountApplied = AT_ApplyOwnerCredit(recordID)
                MsgBox AT_U("0625,062C,0645,0627,0644,064A") & " " & AT_U("0627,0644,0631,0635,064A,062F") & " " & AT_U("0627,0644,0645,0633,062A,062E,062F,0645") & ": " & FormatCurrency(amountApplied), vbInformation, "Archon Tech"
            Else
                MsgBox AT_U("0627,062E,062A,0631") & " " & AT_U("0627,0644,0645,0627,0644,0643") & " " & AT_U("0623,0648,0644,064B,0627") & ".", vbExclamation
            End If
        Case "APPROVE_VOUCHER"
            recordID = AT_CurrentRecordID(fName, "VoucherID")
            If recordID > 0 Then AT_ApproveVoucher recordID Else MsgBox AT_U("0627,062D,0641,0638") & " " & AT_U("0627,0644,0633,0646,062F") & " " & AT_U("0623,0648,0644,064B,0627") & ".", vbExclamation
            Forms(fName).Requery
        Case "START_TRANSFER"
            recordID = AT_CurrentRecordID(fName, "TransferID")
            If recordID > 0 Then AT_StartTransfer recordID Else MsgBox AT_U("0627,062D,0641,0638") & " " & AT_U("0627,0644,062A,062D,0648,064A,0644") & " " & AT_U("0623,0648,0644,064B,0627") & ".", vbExclamation
            Forms(fName).Requery
        Case "COMPLETE_TRANSFER"
            recordID = AT_CurrentRecordID(fName, "TransferID")
            If recordID > 0 Then AT_CompleteTransfer recordID Else MsgBox AT_U("0627,062D,0641,0638") & " " & AT_U("0627,0644,062A,062D,0648,064A,0644") & " " & AT_U("0623,0648,0644,064B,0627") & ".", vbExclamation
            Forms(fName).Requery
        Case "APPROVE_DISCOUNT"
            recordID = AT_CurrentRecordID(fName, "DiscountRequestID")
            If recordID > 0 Then AT_ApproveDiscount recordID Else MsgBox AT_U("0627,062D,0641,0638") & " " & AT_U("0627,0644,0637,0644,0628") & " " & AT_U("0623,0648,0644,064B,0627") & ".", vbExclamation
            Forms(fName).Requery
        Case Else
            MsgBox AT_U("0627,0644,0625,062C,0631,0627,0621") & " " & AT_U("063A,064A,0631") & " " & AT_U("0645,0639,0631,0648,0641") & ": " & actionName, vbExclamation, "Archon Tech"
    End Select
    AT_FormAction = True
    Exit Function
ActionFail:
    MsgBox AT_U("062A,0639,0630,0631") & " " & AT_U("062A,0646,0641,064A,0630") & " " & AT_U("0627,0644,0625,062C,0631,0627,0621") & " " & actionName & ": " & Err.Description, vbCritical, "Archon Tech"
    AT_FormAction = False
End Function

Public Sub AT_SearchOwners()
    Dim rawText As String, normalized As String, whereText As String
    rawText = InputBox(AT_U("0627,0643,062A,0628") & " " & AT_U("0627,0633,0645") & " " & AT_U("0627,0644,0645,0627,0644,0643") & " " & AT_U("0623,0648") & " " & AT_U("062C,0632,0621,064B,0627") & " " & AT_U("0645,0646,0647") & " " & AT_U("0623,0648") & " " & AT_U("0631,0642,0645") & " " & AT_U("0627,0644,0647,0627,062A,0641") & " " & AT_U("0623,0648") & " " & AT_U("0631,0642,0645") & " " & AT_U("0627,0644,0645,0627,0644,0643") & ":", AT_U("0628,062D,062B") & " " & AT_U("0639,0646") & " " & AT_U("0645,0627,0644,0643"))
    If Len(Trim$(rawText)) = 0 Then Exit Sub
    normalized = AT_NormalizeArabic(rawText)
    whereText = "(AT_NormalizeArabic([FullName]) Like " & AT_SqlText("*" & normalized & "*") & _
        " OR [OwnerCode] Like " & AT_SqlText("*" & Replace(rawText, "'", "''") & "*") & _
        " OR [Phone] Like " & AT_SqlText("*" & Replace(rawText, "'", "''") & "*" ) & ")"
    DoCmd.OpenForm "frmAT_Owners", , , whereText
End Sub

Private Sub AT_InitializeNewRecord(ByVal formName As String, ByVal actionName As String)
    On Error Resume Next
    Select Case actionName
        Case "NEW_OWNER"
            Forms(formName)!OwnerCode = AT_NextCode("OW-", "tblAT_Owners", "OwnerCode", 4)
            Forms(formName)!OwnerStatus = "Active"
            Forms(formName)!CreatedAt = Now()
        Case "NEW_UNIT"
            Forms(formName)!UnitCode = AT_NextCode("UN-", "tblAT_Units", "UnitCode", 4)
            Forms(formName)!ContributionAmount = CCur(Val(AT_GetSetting("ContributionPerUnit", CStr(AT_DEFAULT_CONTRIBUTION))))
            Forms(formName)!IsActive = True
            Forms(formName)!CreatedAt = Now()
        Case "NEW_CHARGE"
            Forms(formName)!ChargeCode = AT_NextCode("CH-", "tblAT_Charges", "ChargeCode", 6)
            Forms(formName)!ChargeDate = Date()
            Forms(formName)!DueDate = Date()
            Forms(formName)!OriginalAmount = CCur(Val(AT_GetSetting("ContributionPerUnit", CStr(AT_DEFAULT_CONTRIBUTION))))
            Forms(formName)!CurrentAmount = Forms(formName)!OriginalAmount
            Forms(formName)!Status = "Open"
            Forms(formName)!SourceType = "MANUAL"
            Forms(formName)!CreatedAt = Now()
            Forms(formName)!CreatedBy = AT_UserName()
        Case "NEW_RECEIPT"
            Forms(formName)!ReceiptDate = Date()
            Forms(formName)!Status = "Draft"
            Forms(formName)!CreatedAt = Now()
        Case "NEW_EXPENSE"
            Forms(formName)!ExpenseDate = Date()
            Forms(formName)!Status = "Draft"
            Forms(formName)!CreatedAt = Now()
            Forms(formName)!CreatedBy = AT_UserName()
        Case "NEW_VOUCHER"
            Forms(formName)!VoucherDate = Date()
            Forms(formName)!VoucherType = "EXPENSE"
            Forms(formName)!Status = "Draft"
            Forms(formName)!CreatedAt = Now()
        Case "NEW_TRANSFER"
            Forms(formName)!TransferDate = Date()
            Forms(formName)!FeeAmount = 0
            Forms(formName)!Status = "Draft"
            Forms(formName)!CreatedAt = Now()
        Case "NEW_DISCOUNT"
            Forms(formName)!DiscountNo = AT_NextCode("DS-", "tblAT_DiscountRequests", "DiscountNo", 6)
            Forms(formName)!RequestDate = Date()
            Forms(formName)!Status = "Pending"
            Forms(formName)!CreatedAt = Now()
        Case "NEW_ACCOUNT"
            Forms(formName)!AccountCode = AT_NextCode("ACC-", "tblAT_FinancialAccounts", "AccountCode", 3)
            Forms(formName)!IsActive = True
            Forms(formName)!CreatedAt = Now()
        Case "NEW_METHOD"
            Forms(formName)!IsActive = True
        Case "NEW_CATEGORY"
            Forms(formName)!IsActive = True
    End Select
    On Error GoTo 0
End Sub
