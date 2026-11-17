Attribute VB_Name = "modArchonFinance"
Option Compare Database
Option Explicit

' Financial services for Archon Access.
' Critical rule: caller owns the transaction. Number-series helpers NEVER begin/commit transactions.
' All financial effects are inserted only by approval services. Do not expose ledger tables for direct editing.

Private Function AT_CurrentActor() As String
    AT_CurrentActor = Environ$("Username")
    If Len(Trim$(AT_CurrentActor)) = 0 Then AT_CurrentActor = "UNKNOWN-WINDOWS-USER"
End Function

Private Function AT_SqlDate(ByVal d As Date) As String
    AT_SqlDate = "#" & Format$(d, "mm\/dd\/yyyy hh:nn:ss") & "#"
End Function

Private Function AT_GetText(ByVal rs As DAO.Recordset, ByVal fieldName As String, Optional ByVal fallback As String = "") As String
    If IsNull(rs.Fields(fieldName).Value) Then AT_GetText = fallback Else AT_GetText = CStr(rs.Fields(fieldName).Value)
End Function

Private Function AT_GetCurrency(ByVal rs As DAO.Recordset, ByVal fieldName As String) As Currency
    If IsNull(rs.Fields(fieldName).Value) Then AT_GetCurrency = 0 Else AT_GetCurrency = CCur(rs.Fields(fieldName).Value)
End Function

Private Sub AT_RaiseOriginal(ByVal sourceName As String, ByVal eNum As Long, ByVal eDesc As String)
    Err.Raise eNum, sourceName, eDesc
End Sub

Private Function AT_Exists(ByVal tableName As String, ByVal criteria As String) As Boolean
    AT_Exists = (DCount("*", tableName, criteria) > 0)
End Function

Public Sub AT_RequirePermission(ByVal permissionCode As String)
    Dim db As DAO.Database, rs As DAO.Recordset, userID As Long, permID As Long, allowed As Boolean
    If Not AT_Exists("tblUsers", "UserName='" & Replace(AT_CurrentActor(), "'", "''") & "' AND IsActive=True") Then _
        Err.Raise vbObjectError + 190, "AT_RequirePermission", "لا يوجد حساب مستخدم نشط مرتبط بمستخدم Windows الحالي."
    userID = CLng(Nz(DLookup("UserID", "tblUsers", "UserName='" & Replace(AT_CurrentActor(), "'", "''") & "' AND IsActive=True"), 0))
    permID = CLng(Nz(DLookup("PermissionID", "tblPermissions", "PermissionCode='" & Replace(permissionCode, "'", "''") & "' AND IsActive=True"), 0))
    If userID = 0 Or permID = 0 Then Err.Raise vbObjectError + 191, "AT_RequirePermission", "تعذر تحديد المستخدم أو الصلاحية المطلوبة: " & permissionCode
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT RoleID FROM tblUserRoleAssignments WHERE UserID=" & userID & " AND IsActive=True", dbOpenSnapshot)
    Do While Not rs.EOF
        If AT_Exists("tblRoles", "RoleID=" & CLng(rs!RoleID) & " AND IsActive=True") Then
            If AT_Exists("tblRolePermissions", "RoleID=" & CLng(rs!RoleID) & " AND PermissionID=" & permID) Then allowed = True
        End If
        rs.MoveNext
        If allowed Then Exit Do
    Loop
    rs.Close
    If Not allowed Then Err.Raise vbObjectError + 192, "AT_RequirePermission", "ليست لديك الصلاحية لتنفيذ العملية: " & permissionCode
End Sub

Private Function AT_CreateDocument(ByVal documentType As String, ByVal transactionDate As Date, ByVal reason As String) As Long
    Dim db As DAO.Database, rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblFinancialDocuments", dbOpenDynaset)
    rs.AddNew
    rs!DocumentType = documentType
    rs!TransactionDate = transactionDate
    rs!StatusCode = "DRAFT"
    rs!CreatedAt = Now()
    rs!CreatedByUserName = AT_CurrentActor()
    rs!Reason = reason
    rs.Update
    rs.Bookmark = rs.LastModified
    AT_CreateDocument = CLng(rs!DocumentID)
    rs.Close
End Function

Private Function AT_FormatSeriesNumber(ByVal prefix As String, ByVal numericValue As Long, ByVal padLength As Long) As String
    Dim digits As String
    If numericValue < 1 Or padLength < 1 Then Err.Raise vbObjectError + 204, "AT_FormatSeriesNumber", "قيمة الترقيم أو طول التعبئة غير صالح."
    digits = CStr(numericValue)
    ' Minimum width, not fixed width: never truncate digits after the pad length is exceeded.
    If Len(digits) < padLength Then digits = String$(padLength - Len(digits), "0") & digits
    AT_FormatSeriesNumber = prefix & digits
End Function

Private Function AT_NextSeriesValueInTransaction(ByVal db As DAO.Database, ByVal seriesCode As String) As String
    Dim rs As DAO.Recordset, n As Long, prefix As String, pad As Long
    Set rs = db.OpenRecordset("SELECT * FROM tblNumberSeries WHERE SeriesCode='" & Replace(seriesCode, "'", "''") & "'", dbOpenDynaset)
    If rs.EOF Then
        rs.Close
        Err.Raise vbObjectError + 200, "AT_NextSeriesValueInTransaction", "سلسلة الترقيم غير مهيأة: " & seriesCode
    End If
    n = CLng(rs!NextValue): prefix = AT_GetText(rs, "Prefix"): pad = CLng(Nz(rs!PadLength, 6))
    If n < 1 Then rs.Close: Err.Raise vbObjectError + 201, "AT_NextSeriesValueInTransaction", "قيمة عداد الترقيم غير صالحة."
    rs.Edit: rs!NextValue = n + 1: rs.Update: rs.Close
    AT_NextSeriesValueInTransaction = AT_FormatSeriesNumber(prefix, n, pad)
End Function

Private Function AT_AssignOfficialNumberInTransaction(ByVal db As DAO.Database, ByVal seriesCode As String, ByVal documentID As Long, ByVal actor As String) As String
    Dim numberText As String, rs As DAO.Recordset, n As Long, prefix As String, pad As Long
    Set rs = db.OpenRecordset("SELECT * FROM tblNumberSeries WHERE SeriesCode='" & Replace(seriesCode, "'", "''") & "'", dbOpenDynaset)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 202, "AT_AssignOfficialNumberInTransaction", "سلسلة الترقيم غير مهيأة: " & seriesCode
    n = CLng(rs!NextValue): prefix = AT_GetText(rs, "Prefix"): pad = CLng(Nz(rs!PadLength, 6))
    If n < 1 Then rs.Close: Err.Raise vbObjectError + 203, "AT_AssignOfficialNumberInTransaction", "قيمة عداد الترقيم غير صالحة."
    numberText = AT_FormatSeriesNumber(prefix, n, pad)
    rs.Edit: rs!NextValue = n + 1: rs.Update: rs.Close
    Set rs = db.OpenRecordset("tblOfficialDocumentNumbers", dbOpenDynaset)
    rs.AddNew: rs!DocumentID = documentID: rs!OfficialNumber = numberText: rs!SeriesCode = seriesCode
    rs!NumericValue = n: rs!AssignedAt = Now(): rs!AssignedByUserName = actor: rs.Update: rs.Close
    AT_AssignOfficialNumberInTransaction = numberText
End Function

Private Sub AT_TransitionDraftToPosting(ByVal db As DAO.Database, ByVal documentID As Long)
    db.Execute "UPDATE tblFinancialDocuments SET StatusCode='POSTING' WHERE DocumentID=" & documentID & " AND StatusCode='DRAFT'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 210, "AT_TransitionDraftToPosting", "المستند ليس مسودة أو سبق ترحيله؛ ألغيت العملية لمنع التكرار."
End Sub

Private Sub AT_CompleteApproval(ByVal db As DAO.Database, ByVal documentID As Long, ByVal actor As String)
    db.Execute "UPDATE tblFinancialDocuments SET StatusCode='APPROVED', ApprovedAt=Now(), ApprovedByUserName='" & Replace(actor, "'", "''") & "' WHERE DocumentID=" & documentID & " AND StatusCode='POSTING'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 211, "AT_CompleteApproval", "تعذر تثبيت اعتماد المستند."
End Sub

Private Sub AT_InsertAccountMovement(ByVal db As DAO.Database, ByVal accountID As Long, ByVal documentID As Long, ByVal movementDate As Date, ByVal signedAmount As Currency, ByVal movementType As String, ByVal lineNo As Long)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("tblAccountMovements", dbOpenDynaset)
    rs.AddNew: rs!AccountID = accountID: rs!DocumentID = documentID
    rs!MovementDate = movementDate: rs!SignedAmount = signedAmount
    rs!MovementType = movementType: rs!LineNo = lineNo: rs!CreatedAt = Now()
    rs.Update: rs.Close
End Sub

Public Function AT_CreateAccount(ByVal accountCode As String, ByVal accountName As String, ByVal accountType As String) As Long
    On Error GoTo EH
    AT_RequirePermission "SETTINGS_ADMIN"
    accountCode = UCase$(Trim$(accountCode)): accountName = Trim$(accountName): accountType = UCase$(Trim$(accountType))
    If Len(accountCode) = 0 Or Len(accountName) = 0 Then Err.Raise vbObjectError + 195, , "رمز الحساب واسمه مطلوبان."
    If Not (accountType = "CASH" Or accountType = "BANK" Or accountType = "EWALLET" Or accountType = "OTHER") Then _
        Err.Raise vbObjectError + 196, , "نوع الحساب غير صالح."
    If AT_Exists("tblAccounts", "AccountCode='" & Replace(accountCode, "'", "''") & "'") Then _
        Err.Raise vbObjectError + 197, , "رمز الحساب مستخدم مسبقًا."
    Dim db As DAO.Database, rs As DAO.Recordset, accountID As Long
    Set db = CurrentDb: DBEngine(0)(0).BeginTrans
    Set rs = db.OpenRecordset("tblAccounts", dbOpenDynaset)
    rs.AddNew: rs!AccountCode = accountCode: rs!AccountName = accountName
    rs!AccountType = accountType: rs!IsActive = True: rs!CreatedAt = Now(): rs.Update
    rs.Bookmark = rs.LastModified: accountID = CLng(rs!AccountID): rs.Close
    DBEngine(0)(0).CommitTrans
    AT_Audit "ACCOUNT_CREATE", "tblAccounts", accountID, "SUCCESS", "Account created with zero opening balance; type=" & accountType
    AT_CreateAccount = accountID
    Exit Function
EH:
    Dim eNum As Long, eDesc As String: eNum = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "ACCOUNT_CREATE", "tblAccounts", 0, "FAILED", eDesc
    AT_RaiseOriginal "AT_CreateAccount", eNum, eDesc
End Function

Public Function AT_AccountBookBalance(ByVal accountID As Long, Optional ByVal asOfDate As Date = 0) As Currency
    Dim sql As String, rs As DAO.Recordset
    sql = "SELECT Sum(M.SignedAmount) AS BalanceValue FROM tblAccountMovements AS M INNER JOIN tblFinancialDocuments AS D ON M.DocumentID=D.DocumentID WHERE M.AccountID=" & accountID & " AND D.StatusCode='APPROVED'"
    If asOfDate <> 0 Then sql = sql & " AND M.MovementDate<=" & AT_SqlDate(asOfDate)
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    AT_AccountBookBalance = AT_GetCurrency(rs, "BalanceValue")
    rs.Close
End Function

Public Function AT_AccountReservedAmount(ByVal accountID As Long, Optional ByVal exceptReservationID As Long = 0) As Currency
    Dim criteria As String
    criteria = "AccountID=" & accountID & " AND StatusCode='ACTIVE'"
    If exceptReservationID > 0 Then criteria = criteria & " AND ReservationID<>" & exceptReservationID
    AT_AccountReservedAmount = Nz(DSum("Amount", "tblAccountReservations", criteria), 0)
End Function

Public Function AT_AccountAvailableBalance(ByVal accountID As Long, Optional ByVal exceptReservationID As Long = 0) As Currency
    AT_AccountAvailableBalance = AT_AccountBookBalance(accountID) - AT_AccountReservedAmount(accountID, exceptReservationID)
End Function

Public Function AT_OwnerAvailableCredit(ByVal ownerID As Long) As Currency
    Dim sql As String, rs As DAO.Recordset
    sql = "SELECT Sum(C.SignedAmount) AS CreditBalance FROM tblOwnerCreditMovements AS C INNER JOIN tblFinancialDocuments AS D ON C.DocumentID=D.DocumentID WHERE C.OwnerID=" & ownerID & " AND D.StatusCode='APPROVED'"
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    AT_OwnerAvailableCredit = AT_GetCurrency(rs, "CreditBalance")
    rs.Close
End Function

Private Function AT_CreditMovementRawAvailable(ByVal creditMovementID As Long) As Currency
    Dim originalAmount As Currency, sql As String, rs As DAO.Recordset, sourceDocID As Long
    If Not AT_Exists("tblOwnerCreditMovements", "CreditMovementID=" & creditMovementID & " AND SourceCreditMovementID Is Null") Then Exit Function
    originalAmount = Nz(DLookup("SignedAmount", "tblOwnerCreditMovements", "CreditMovementID=" & creditMovementID), 0)
    sourceDocID = CLng(Nz(DLookup("DocumentID", "tblOwnerCreditMovements", "CreditMovementID=" & creditMovementID), 0))
    If originalAmount <= 0 Or sourceDocID = 0 Then Exit Function
    If Not AT_Exists("tblFinancialDocuments", "DocumentID=" & sourceDocID & " AND StatusCode='APPROVED'") Then Exit Function
    sql = "SELECT Sum(C.SignedAmount) AS UsedAmount FROM tblOwnerCreditMovements AS C INNER JOIN tblFinancialDocuments AS D ON C.DocumentID=D.DocumentID WHERE C.SourceCreditMovementID=" & creditMovementID & " AND D.StatusCode='APPROVED'"
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    AT_CreditMovementRawAvailable = originalAmount + AT_GetCurrency(rs, "UsedAmount")
    rs.Close
End Function

Public Function AT_CreditMovementAvailable(ByVal creditMovementID As Long) As Currency
    Dim originalAmount As Currency, sql As String, rs As DAO.Recordset, sourceDocID As Long
    If Not AT_Exists("tblOwnerCreditMovements", "CreditMovementID=" & creditMovementID & " AND SourceCreditMovementID Is Null") Then Exit Function
    originalAmount = Nz(DLookup("SignedAmount", "tblOwnerCreditMovements", "CreditMovementID=" & creditMovementID), 0)
    sourceDocID = CLng(Nz(DLookup("DocumentID", "tblOwnerCreditMovements", "CreditMovementID=" & creditMovementID), 0))
    If originalAmount <= 0 Or sourceDocID = 0 Then Exit Function
    If Not AT_Exists("tblFinancialDocuments", "DocumentID=" & sourceDocID & " AND StatusCode='APPROVED'") Then Exit Function
    sql = "SELECT Sum(C.SignedAmount) AS UsedAmount FROM tblOwnerCreditMovements AS C INNER JOIN tblFinancialDocuments AS D ON C.DocumentID=D.DocumentID WHERE C.SourceCreditMovementID=" & creditMovementID & " AND D.StatusCode='APPROVED'"
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    AT_CreditMovementAvailable = originalAmount + AT_GetCurrency(rs, "UsedAmount")
    rs.Close
    If AT_CreditMovementAvailable < 0 Then AT_CreditMovementAvailable = 0
End Function

Public Function AT_ChargeRemaining(ByVal chargeID As Long) As Currency
    Dim gross As Currency, adjustments As Currency, settled As Currency
    Dim rs As DAO.Recordset, sql As String
    If Not AT_Exists("tblOwnerCharges", "ChargeID=" & chargeID & " AND ChargeStatus='ACTIVE'") Then Exit Function
    gross = Nz(DLookup("GrossAmount", "tblOwnerCharges", "ChargeID=" & chargeID), 0)
    sql = "SELECT Sum(A.SignedAmount) AS AdjustmentTotal FROM tblOwnerAdjustments AS A INNER JOIN tblFinancialDocuments AS D ON A.DocumentID=D.DocumentID WHERE A.ChargeID=" & chargeID & " AND D.StatusCode='APPROVED'"
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot): adjustments = AT_GetCurrency(rs, "AdjustmentTotal"): rs.Close
    sql = "SELECT Sum(S.SignedAmount) AS SettlementTotal FROM tblChargeSettlementLines AS S INNER JOIN tblFinancialDocuments AS D ON S.DocumentID=D.DocumentID WHERE S.ChargeID=" & chargeID & " AND D.StatusCode='APPROVED'"
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot): settled = AT_GetCurrency(rs, "SettlementTotal"): rs.Close
    AT_ChargeRemaining = gross + adjustments - settled
End Function

Public Function AT_CreateOwnerDraft(ByVal fullName As String, ByVal phone As String, ByVal houseCount As Long, ByVal apartmentCount As Long, Optional ByVal notes As String = "") As Long
    On Error GoTo EH
    AT_RequirePermission "OWNER_EDIT"
    If Len(Trim$(fullName)) = 0 Then Err.Raise vbObjectError + 220, , "اسم المالك مطلوب."
    If houseCount < 0 Or apartmentCount < 0 Then Err.Raise vbObjectError + 221, , "عدد المنازل والشقق لا يمكن أن يكون سالبًا."
    Dim db As DAO.Database, rs As DAO.Recordset, ownerID As Long
    Set db = CurrentDb: DBEngine(0)(0).BeginTrans
    Set rs = db.OpenRecordset("tblOwners", dbOpenDynaset)
    rs.AddNew: rs!FullName = Trim$(fullName): rs!FullNameSearch = AT_NormalizeArabicSearch(fullName)
    rs!Phone = Trim$(phone): rs!HouseCount = houseCount: rs!ApartmentCount = apartmentCount
    rs!Notes = notes: rs!AdminStatus = "DRAFT": rs!DataStatus = IIf(Len(Trim$(phone)) = 0, "INCOMPLETE", "COMPLETE")
    rs!ReviewStatus = "NOT_REVIEWED": rs!CreatedAt = Now(): rs!UpdatedAt = Now(): rs.Update
    rs.Bookmark = rs.LastModified: ownerID = CLng(rs!OwnerID): rs.Close
    DBEngine(0)(0).CommitTrans
    AT_Audit "OWNER_DRAFT", "tblOwners", ownerID, "SUCCESS", "Owner draft saved; no official code assigned."
    AT_CreateOwnerDraft = ownerID: Exit Function
EH:
    Dim eNum As Long, eDesc As String: eNum = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "OWNER_DRAFT", "tblOwners", 0, "FAILED", eDesc
    AT_RaiseOriginal "AT_CreateOwnerDraft", eNum, eDesc
End Function

Public Sub AT_ApproveOwner(ByVal ownerID As Long, ByVal approver As String)
    On Error GoTo EH
    AT_RequirePermission "OWNER_APPROVE"
    approver = AT_CurrentActor()
    Dim db As DAO.Database, rs As DAO.Recordset, newCode As String
    Set db = CurrentDb: DBEngine(0)(0).BeginTrans
    db.Execute "UPDATE tblOwners SET AdminStatus='APPROVING' WHERE OwnerID=" & ownerID & " AND AdminStatus='DRAFT'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 230, , "المالك غير موجود كمسودة أو سبق اعتماده."
    newCode = AT_NextSeriesValueInTransaction(db, "OWNER")
    db.Execute "UPDATE tblOwners SET OwnerCode='" & Replace(newCode, "'", "''") & "', AdminStatus='ACTIVE', UpdatedAt=Now() WHERE OwnerID=" & ownerID & " AND AdminStatus='APPROVING'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 231, , "تعذر تثبيت رقم المالك."
    DBEngine(0)(0).CommitTrans
    AT_Audit "OWNER_APPROVE", "tblOwners", ownerID, "SUCCESS", "Assigned " & newCode, "Approved by " & approver
    Exit Sub
EH:
    Dim eNum As Long, eDesc As String: eNum = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "OWNER_APPROVE", "tblOwners", ownerID, "FAILED", eDesc
    AT_RaiseOriginal "AT_ApproveOwner", eNum, eDesc
End Sub

Public Sub AT_UpdateOwnerDetails(ByVal ownerID As Long, ByVal fullName As String, ByVal phone As String, ByVal houseCount As Long, ByVal apartmentCount As Long, ByVal notes As String, ByVal reason As String)
    On Error GoTo EH
    AT_RequirePermission "OWNER_EDIT"
    If ownerID <= 0 Or Len(Trim$(fullName)) = 0 Then Err.Raise vbObjectError + 225, , "معرف المالك والاسم مطلوبان."
    If houseCount < 0 Or apartmentCount < 0 Then Err.Raise vbObjectError + 226, , "عدد المنازل والشقق لا يمكن أن يكون سالبًا."
    If Len(Trim$(reason)) = 0 Then Err.Raise vbObjectError + 227, , "سبب تعديل بيانات المالك مطلوب للتدقيق."

    Dim db As DAO.Database, rs As DAO.Recordset, userID As Long, auditID As Long
    Dim oldName As String, oldPhone As String, oldHouses As String, oldApartments As String, oldNotes As String
    Dim newName As String, newPhone As String, newHouses As String, newApartments As String, newNotes As String, dataStatus As String
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tblOwners WHERE OwnerID=" & ownerID & " AND AdminStatus='ACTIVE'", dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 228, , "المالك غير موجود أو غير نشط."
    oldName = AT_GetText(rs, "FullName"): oldPhone = AT_GetText(rs, "Phone")
    oldHouses = CStr(Nz(rs!HouseCount, 0)): oldApartments = CStr(Nz(rs!ApartmentCount, 0)): oldNotes = AT_GetText(rs, "Notes")
    rs.Close
    newName = Trim$(fullName): newPhone = Trim$(phone): newHouses = CStr(houseCount): newApartments = CStr(apartmentCount): newNotes = notes
    dataStatus = IIf(Len(newPhone) = 0 Or apartmentCount <= 0, "INCOMPLETE", "COMPLETE")
    userID = CLng(Nz(DLookup("UserID", "tblUsers", "UserName='" & Replace(AT_CurrentActor(), "'", "''") & "' AND IsActive=True"), 0))

    DBEngine(0)(0).BeginTrans
    db.Execute "UPDATE tblOwners SET FullName='" & Replace(newName, "'", "''") & "', FullNameSearch='" & Replace(AT_NormalizeArabicSearch(newName), "'", "''") & "', Phone='" & Replace(newPhone, "'", "''") & "', HouseCount=" & houseCount & ", ApartmentCount=" & apartmentCount & ", Notes='" & Replace(notes, "'", "''") & "', DataStatus='" & dataStatus & "', UpdatedAt=Now(), UpdatedByUserID=" & IIf(userID > 0, CStr(userID), "Null") & " WHERE OwnerID=" & ownerID & " AND AdminStatus='ACTIVE'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 229, , "لم يُحدّث سجل المالك؛ ربما تغيرت حالته."

    Set rs = db.OpenRecordset("tblAuditEvents", dbOpenDynaset)
    rs.AddNew: rs!EventAt = Now(): rs!UserName = AT_CurrentActor(): rs!ActionCode = "OWNER_UPDATE": rs!EntityName = "tblOwners"
    rs!EntityID = ownerID: rs!ResultCode = "SUCCESS": rs!Reason = reason: rs!Detail = "Updated owner profile fields": rs.Update
    rs.Bookmark = rs.LastModified: auditID = CLng(rs!AuditEventID): rs.Close
    If oldName <> newName Then AT_AddAuditChange db, auditID, "FullName", oldName, newName
    If oldPhone <> newPhone Then AT_AddAuditChange db, auditID, "Phone", oldPhone, newPhone
    If oldHouses <> newHouses Then AT_AddAuditChange db, auditID, "HouseCount", oldHouses, newHouses
    If oldApartments <> newApartments Then AT_AddAuditChange db, auditID, "ApartmentCount", oldApartments, newApartments
    If oldNotes <> newNotes Then AT_AddAuditChange db, auditID, "Notes", oldNotes, newNotes
    DBEngine(0)(0).CommitTrans
    Exit Sub
EH:
    Dim eNum As Long, eDesc As String: eNum = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "OWNER_UPDATE", "tblOwners", ownerID, "FAILED", eDesc, reason
    AT_RaiseOriginal "AT_UpdateOwnerDetails", eNum, eDesc
End Sub

Private Sub AT_AddAuditChange(ByVal db As DAO.Database, ByVal auditEventID As Long, ByVal fieldName As String, ByVal oldValue As String, ByVal newValue As String)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("tblAuditChanges", dbOpenDynaset)
    rs.AddNew: rs!AuditEventID = auditEventID: rs!FieldName = fieldName
    rs!OldValue = oldValue: rs!NewValue = newValue: rs.Update: rs.Close
End Sub

Public Function AT_NormalizeArabicSearch(ByVal sourceText As String) As String
    Dim s As String, marks As Variant, i As Long
    s = Trim$(sourceText)
    s = Replace(s, "أ", "ا"): s = Replace(s, "إ", "ا"): s = Replace(s, "آ", "ا")
    s = Replace(s, "ى", "ي"): s = Replace(s, "ة", "ه"): s = Replace(s, ChrW(&H640), "")
    marks = Array(&H64B, &H64C, &H64D, &H64E, &H64F, &H650, &H651, &H652, &H653, &H654, &H655, &H670)
    For i = LBound(marks) To UBound(marks): s = Replace(s, ChrW(marks(i)), ""): Next i
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    AT_NormalizeArabicSearch = LCase$(s)
End Function

Public Function AT_CreateChargeForOwner(ByVal ownerID As Long, ByVal chargeDate As Date, ByVal rateID As Long, Optional ByVal reason As String = "") As Long
    On Error GoTo EH
    AT_RequirePermission "CHARGE_DRAFT"
    Dim db As DAO.Database, rr As DAO.Recordset, rc As DAO.Recordset, ownerApartments As Long
    Dim rateAmount As Currency, docID As Long, chargeID As Long, gross As Currency
    If Not AT_Exists("tblOwners", "OwnerID=" & ownerID & " AND AdminStatus='ACTIVE'") Then Err.Raise vbObjectError + 240, , "المالك غير موجود أو غير نشط."
    Set rr = CurrentDb.OpenRecordset("SELECT * FROM tblContributionRates WHERE RateID=" & rateID & " AND IsActive=True", dbOpenSnapshot)
    If rr.EOF Then rr.Close: Err.Raise vbObjectError + 241, , "معدل المساهمة غير فعال أو غير موجود."
    If CDate(rr!EffectiveFrom) > chargeDate Then rr.Close: Err.Raise vbObjectError + 242, , "تاريخ الاستحقاق يسبق بداية سريان السعر."
    If Not IsNull(rr!EffectiveTo) Then If CDate(rr!EffectiveTo) < chargeDate Then rr.Close: Err.Raise vbObjectError + 243, , "انتهى سريان معدل المساهمة في تاريخ الاستحقاق."
    rateAmount = CCur(rr!RateAmount): rr.Close
    ownerApartments = CLng(Nz(DLookup("ApartmentCount", "tblOwners", "OwnerID=" & ownerID), 0))
    If ownerApartments <= 0 Then Err.Raise vbObjectError + 244, , "عدد الشقق يجب أن يكون أكبر من صفر لإنشاء الاستحقاق."
    gross = CCur(ownerApartments) * rateAmount
    Set db = CurrentDb: DBEngine(0)(0).BeginTrans
    docID = AT_CreateDocument("CHARGE", chargeDate, reason)
    Set rc = db.OpenRecordset("tblOwnerCharges", dbOpenDynaset)
    rc.AddNew: rc!DocumentID = docID: rc!OwnerID = ownerID: rc!ChargeDate = chargeDate
    rc!ApartmentCountAtCharge = ownerApartments: rc!RateAmount = rateAmount: rc!GrossAmount = gross
    rc!ChargeStatus = "DRAFT": rc!CreatedAt = Now(): rc.Update: rc.Bookmark = rc.LastModified
    chargeID = CLng(rc!ChargeID): rc.Close
    DBEngine(0)(0).CommitTrans
    AT_Audit "CHARGE_DRAFT", "tblOwnerCharges", chargeID, "SUCCESS", "Gross=" & CStr(gross)
    AT_CreateChargeForOwner = chargeID: Exit Function
EH:
    Dim eNum As Long, eDesc As String: eNum = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "CHARGE_DRAFT", "tblOwnerCharges", 0, "FAILED", eDesc
    AT_RaiseOriginal "AT_CreateChargeForOwner", eNum, eDesc
End Function

Public Sub AT_ApproveCharge(ByVal chargeID As Long, ByVal approver As String)
    On Error GoTo EH
    AT_RequirePermission "CHARGE_APPROVE"
    approver = AT_CurrentActor()
    Dim db As DAO.Database, rs As DAO.Recordset, docID As Long, codeText As String
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT C.DocumentID, C.ChargeStatus FROM tblOwnerCharges AS C WHERE C.ChargeID=" & chargeID, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 250, , "الاستحقاق غير موجود."
    docID = CLng(rs!DocumentID): If rs!ChargeStatus <> "DRAFT" Then rs.Close: Err.Raise vbObjectError + 251, , "الاستحقاق ليس مسودة."
    rs.Close
    DBEngine(0)(0).BeginTrans: AT_TransitionDraftToPosting db, docID
    codeText = AT_NextSeriesValueInTransaction(db, "CHARGE")
    db.Execute "UPDATE tblOwnerCharges SET ChargeCode='" & codeText & "', ChargeStatus='ACTIVE' WHERE ChargeID=" & chargeID & " AND ChargeStatus='DRAFT'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 252, , "تعذر اعتماد الاستحقاق."
    AT_CompleteApproval db, docID, approver
    DBEngine(0)(0).CommitTrans
    AT_Audit "CHARGE_APPROVE", "tblOwnerCharges", chargeID, "SUCCESS", "Assigned " & codeText
    Exit Sub
EH:
    Dim eNum As Long, eDesc As String: eNum = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "CHARGE_APPROVE", "tblOwnerCharges", chargeID, "FAILED", eDesc
    AT_RaiseOriginal "AT_ApproveCharge", eNum, eDesc
End Sub

Public Function AT_CreateReceiptDraft(ByVal ownerID As Long, ByVal amount As Currency, ByVal paymentMethodID As Long, ByVal accountID As Long, ByVal payerName As String, ByVal collectorName As String, Optional ByVal notes As String = "", Optional ByVal transactionDate As Date = 0) As Long
    On Error GoTo EH
    AT_RequirePermission "RECEIPT_DRAFT"
    If amount <= 0 Or amount <> Fix(amount) Then Err.Raise vbObjectError + 301, , "المبلغ يجب أن يكون عددًا صحيحًا موجبًا بالريال اليمني."
    If Not AT_Exists("tblOwners", "OwnerID=" & ownerID & " AND AdminStatus='ACTIVE'") Then Err.Raise vbObjectError + 302, , "المالك غير موجود أو غير نشط."
    If Not AT_Exists("tblAccounts", "AccountID=" & accountID & " AND IsActive=True") Then Err.Raise vbObjectError + 303, , "الحساب غير موجود أو غير نشط."
    If Not AT_Exists("tblPaymentMethods", "PaymentMethodID=" & paymentMethodID & " AND IsActive=True") Then Err.Raise vbObjectError + 304, , "طريقة الدفع غير موجودة أو غير نشطة."
    If Len(Trim$(collectorName)) = 0 Then collectorName = Nz(DLookup("DefaultCollectorName", "tblProjectSettings"), "")
    If Len(Trim$(collectorName)) = 0 Then Err.Raise vbObjectError + 305, , "اسم المحصل مطلوب؛ اضبط اسم المحصل الافتراضي أو أدخله في الإيصال."
    If Len(Trim$(payerName)) = 0 Then payerName = Nz(DLookup("FullName", "tblOwners", "OwnerID=" & ownerID), "")
    If transactionDate = 0 Then transactionDate = Date
    Dim db As DAO.Database, rs As DAO.Recordset, docID As Long, receiptID As Long
    Set db = CurrentDb: DBEngine(0)(0).BeginTrans
    docID = AT_CreateDocument("RECEIPT", transactionDate, notes)
    Set rs = db.OpenRecordset("tblReceipts", dbOpenDynaset)
    rs.AddNew: rs!DocumentID = docID: rs!OwnerID = ownerID: rs!Amount = amount
    rs!PaymentMethodID = paymentMethodID: rs!AccountID = accountID: rs!PayerName = Trim$(payerName)
    rs!CollectorName = Trim$(collectorName): rs!Notes = notes
    rs!PrintedOwnerName = Nz(DLookup("FullName", "tblOwners", "OwnerID=" & ownerID), "")
    rs!PrintedOwnerCode = Nz(DLookup("OwnerCode", "tblOwners", "OwnerID=" & ownerID), "")
    rs!PrintedNeighborhoodName = Nz(DLookup("NeighborhoodName", "tblProjectSettings"), "")
    rs!CreatedAt = Now(): rs.Update: rs.Bookmark = rs.LastModified: receiptID = CLng(rs!ReceiptID): rs.Close
    DBEngine(0)(0).CommitTrans
    AT_Audit "RECEIPT_DRAFT", "tblReceipts", receiptID, "SUCCESS", "Draft created; no financial effect."
    AT_CreateReceiptDraft = receiptID: Exit Function
EH:
    Dim eNum As Long, eDesc As String: eNum = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "RECEIPT_DRAFT", "tblReceipts", 0, "FAILED", eDesc
    AT_RaiseOriginal "AT_CreateReceiptDraft", eNum, eDesc
End Function

Public Sub AT_ApproveReceipt(ByVal receiptID As Long, ByVal approver As String, Optional ByVal confirmedActualReceipt As Boolean = False)
    On Error GoTo EH
    AT_RequirePermission "RECEIPT_APPROVE"
    approver = AT_CurrentActor()
    If Not confirmedActualReceipt Then Err.Raise vbObjectError + 321, , "يجب تأكيد استلام المال فعليًا قبل الاعتماد."
    Dim db As DAO.Database, rs As DAO.Recordset, docID As Long, ownerID As Long, amount As Currency, accountID As Long, docDate As Date
    Dim remaining As Currency, allocated As Currency, chargeID As Long, lineNo As Long, officialNo As String
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT R.*, D.StatusCode, D.TransactionDate FROM tblReceipts AS R INNER JOIN tblFinancialDocuments AS D ON R.DocumentID=D.DocumentID WHERE R.ReceiptID=" & receiptID, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 322, , "الإيصال غير موجود."
    docID = CLng(rs!DocumentID): ownerID = CLng(rs!OwnerID): amount = CCur(rs!Amount): accountID = CLng(rs!AccountID): docDate = CDate(rs!TransactionDate)
    If rs!StatusCode <> "DRAFT" Then rs.Close: Err.Raise vbObjectError + 323, , "الإيصال ليس مسودة."
    rs.Close
    DBEngine(0)(0).BeginTrans: AT_TransitionDraftToPosting db, docID
    If Not AT_Exists("tblAccounts", "AccountID=" & accountID & " AND IsActive=True") Then Err.Raise vbObjectError + 324, , "الحساب غير نشط."
    If Not AT_Exists("tblOwners", "OwnerID=" & ownerID & " AND AdminStatus='ACTIVE'") Then Err.Raise vbObjectError + 325, , "المالك غير نشط."
    If AT_Exists("tblReceiptAllocationLines", "ReceiptID=" & receiptID) Then Err.Raise vbObjectError + 326, , "توجد تخصيصات سابقة للمسودة؛ يلزم فحص يدوي بدل إعادة الترحيل."
    remaining = amount: lineNo = 0
    Set rs = db.OpenRecordset("SELECT ChargeID FROM tblOwnerCharges WHERE OwnerID=" & ownerID & " AND ChargeStatus='ACTIVE' ORDER BY ChargeDate, ChargeID", dbOpenSnapshot)
    Do While Not rs.EOF And remaining > 0
        chargeID = CLng(rs!ChargeID)
        Dim due As Currency
        due = AT_ChargeRemaining(chargeID)
        If due > 0 Then
            allocated = due: If remaining < due Then allocated = remaining
            lineNo = lineNo + 1
            Dim rLine As DAO.Recordset
            Set rLine = db.OpenRecordset("tblReceiptAllocationLines", dbOpenDynaset)
            rLine.AddNew: rLine!ReceiptID = receiptID: rLine!ChargeID = chargeID: rLine!AllocatedAmount = allocated: rLine!LineNo = lineNo: rLine.Update: rLine.Close
            Dim rSettle As DAO.Recordset
            Set rSettle = db.OpenRecordset("tblChargeSettlementLines", dbOpenDynaset)
            rSettle.AddNew: rSettle!ChargeID = chargeID: rSettle!DocumentID = docID: rSettle!ReceiptID = receiptID
            rSettle!SignedAmount = allocated: rSettle!SettlementType = "RECEIPT": rSettle!CreatedAt = Now(): rSettle.Update: rSettle.Close
            remaining = remaining - allocated
        End If
        rs.MoveNext
    Loop
    rs.Close
    If remaining > 0 Then
        Dim rCredit As DAO.Recordset
        Set rCredit = db.OpenRecordset("tblOwnerCreditMovements", dbOpenDynaset)
        rCredit.AddNew: rCredit!OwnerID = ownerID: rCredit!DocumentID = docID: rCredit!SignedAmount = remaining
        rCredit!CreditType = "RECEIPT_EXCESS": rCredit!SourceReceiptID = receiptID: rCredit!Reason = "زيادة قبض على المستحقات": rCredit!CreatedAt = Now(): rCredit.Update: rCredit.Close
    End If
    AT_InsertAccountMovement db, accountID, docID, docDate, amount, "RECEIPT", 1
    officialNo = AT_AssignOfficialNumberInTransaction(db, "RECEIPT", docID, approver)
    AT_CompleteApproval db, docID, approver
    DBEngine(0)(0).CommitTrans
    AT_Audit "RECEIPT_APPROVE", "tblReceipts", receiptID, "SUCCESS", "Official=" & officialNo & "; excess credit=" & CStr(remaining)
    Exit Sub
EH:
    Dim eNum As Long, eDesc As String: eNum = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "RECEIPT_APPROVE", "tblReceipts", receiptID, "FAILED", eDesc
    AT_RaiseOriginal "AT_ApproveReceipt", eNum, eDesc
End Sub

Public Function AT_CreatePaymentDraft(ByVal accountID As Long, ByVal expenseCategoryID As Long, ByVal amount As Currency, ByVal paymentMethodID As Long, ByVal beneficiaryName As String, Optional ByVal recipientName As String = "", Optional ByVal notes As String = "", Optional ByVal transactionDate As Date = 0, Optional ByVal refundCreditMovementID As Long = 0, Optional ByVal refundDocumentID As Long = 0) As Long
    On Error GoTo EH
    AT_RequirePermission "PAYMENT_DRAFT"
    If amount <= 0 Or amount <> Fix(amount) Then Err.Raise vbObjectError + 330, , "المبلغ يجب أن يكون عددًا صحيحًا موجبًا بالريال اليمني."
    If Len(Trim$(beneficiaryName)) = 0 Then Err.Raise vbObjectError + 331, , "اسم المستفيد مطلوب."
    If Not AT_Exists("tblAccounts", "AccountID=" & accountID & " AND IsActive=True") Then Err.Raise vbObjectError + 332, , "الحساب غير موجود أو غير نشط."
    If Not AT_Exists("tblExpenseCategories", "ExpenseCategoryID=" & expenseCategoryID & " AND IsActive=True") Then Err.Raise vbObjectError + 333, , "فئة المصروف غير موجودة أو غير نشطة."
    If Not AT_Exists("tblPaymentMethods", "PaymentMethodID=" & paymentMethodID & " AND IsActive=True") Then Err.Raise vbObjectError + 334, , "طريقة الدفع غير موجودة أو غير نشطة."
    If refundCreditMovementID > 0 Then
        If AT_CreditMovementAvailable(refundCreditMovementID) < amount Then Err.Raise vbObjectError + 335, , "رصيد المالك الدائن المتاح غير كافٍ لهذا الرد."
        If refundDocumentID <= 0 Then refundDocumentID = CLng(Nz(DLookup("DocumentID", "tblOwnerCreditMovements", "CreditMovementID=" & refundCreditMovementID), 0))
        If refundDocumentID <= 0 Then Err.Raise vbObjectError + 336, , "تعذر ربط الرد بمستند مصدر الرصيد الدائن."
    End If
    If transactionDate = 0 Then transactionDate = Date
    Dim db As DAO.Database, rs As DAO.Recordset, docID As Long, paymentID As Long
    Set db = CurrentDb: DBEngine(0)(0).BeginTrans
    docID = AT_CreateDocument("PAYMENT", transactionDate, notes)
    Set rs = db.OpenRecordset("tblPayments", dbOpenDynaset)
    rs.AddNew: rs!DocumentID = docID: rs!AccountID = accountID: rs!ExpenseCategoryID = expenseCategoryID
    rs!BeneficiaryName = Trim$(beneficiaryName): rs!Amount = amount: rs!PaymentMethodID = paymentMethodID
    rs!RecipientName = recipientName: rs!Notes = notes
    If refundCreditMovementID > 0 Then rs!RefundOfCreditMovementID = refundCreditMovementID
    If refundDocumentID > 0 Then rs!RefundOfDocumentID = refundDocumentID
    rs.Update: rs.Bookmark = rs.LastModified: paymentID = CLng(rs!PaymentID): rs.Close
    DBEngine(0)(0).CommitTrans
    AT_Audit "PAYMENT_DRAFT", "tblPayments", paymentID, "SUCCESS", "Draft created; no financial effect."
    AT_CreatePaymentDraft = paymentID: Exit Function
EH:
    Dim eNum As Long, eDesc As String: eNum = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "PAYMENT_DRAFT", "tblPayments", 0, "FAILED", eDesc
    AT_RaiseOriginal "AT_CreatePaymentDraft", eNum, eDesc
End Function

Public Sub AT_ApprovePayment(ByVal paymentID As Long, ByVal approver As String, Optional ByVal confirmedPaid As Boolean = False)
    On Error GoTo EH
    AT_RequirePermission "PAYMENT_APPROVE"
    approver = AT_CurrentActor()
    If Not confirmedPaid Then Err.Raise vbObjectError + 341, , "يجب تأكيد أن المبلغ دُفع فعليًا قبل الاعتماد."
    Dim db As DAO.Database, rs As DAO.Recordset, docID As Long, accountID As Long, amount As Currency, docDate As Date, creditMovementID As Long, officialNo As String
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT P.*, D.StatusCode, D.TransactionDate FROM tblPayments AS P INNER JOIN tblFinancialDocuments AS D ON P.DocumentID=D.DocumentID WHERE P.PaymentID=" & paymentID, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 342, , "سند الصرف غير موجود."
    docID = CLng(rs!DocumentID): accountID = CLng(rs!AccountID): amount = CCur(rs!Amount): docDate = CDate(rs!TransactionDate)
    creditMovementID = CLng(Nz(rs!RefundOfCreditMovementID, 0))
    If rs!StatusCode <> "DRAFT" Then rs.Close: Err.Raise vbObjectError + 343, , "سند الصرف ليس مسودة."
    rs.Close
    DBEngine(0)(0).BeginTrans: AT_TransitionDraftToPosting db, docID
    If amount <= 0 Or amount <> Fix(amount) Then Err.Raise vbObjectError + 344, , "المبلغ غير صالح."
    If Not AT_Exists("tblAccounts", "AccountID=" & accountID & " AND IsActive=True") Then Err.Raise vbObjectError + 348, , "الحساب غير نشط."
    If AT_AccountAvailableBalance(accountID) < amount Then Err.Raise vbObjectError + 345, , "الرصيد المتاح غير كافٍ بعد احتساب الحجوزات؛ لا يسمح النظام باستثناء."
    If creditMovementID > 0 Then
        If AT_CreditMovementAvailable(creditMovementID) < amount Then Err.Raise vbObjectError + 346, , "الرصيد الدائن المتاح غير كافٍ للرد أو استُخدم في عملية أخرى."
        Dim srcOwnerID As Long
        srcOwnerID = CLng(Nz(DLookup("OwnerID", "tblOwnerCreditMovements", "CreditMovementID=" & creditMovementID), 0))
        If srcOwnerID = 0 Then Err.Raise vbObjectError + 347, , "حركة الرصيد الدائن المصدر غير موجودة."
        Dim cr As DAO.Recordset
        Set cr = db.OpenRecordset("tblOwnerCreditMovements", dbOpenDynaset)
        cr.AddNew: cr!OwnerID = srcOwnerID: cr!DocumentID = docID: cr!SignedAmount = -amount
        cr!CreditType = "CASH_REFUND": cr!SourceCreditMovementID = creditMovementID: cr!Reason = "رد نقدي معتمد عبر سند صرف " & paymentID: cr!CreatedAt = Now(): cr.Update: cr.Close
    End If
    AT_InsertAccountMovement db, accountID, docID, docDate, -amount, IIf(creditMovementID > 0, "CREDIT_REFUND", "PAYMENT"), 1
    officialNo = AT_AssignOfficialNumberInTransaction(db, "PAYMENT", docID, approver)
    AT_CompleteApproval db, docID, approver
    DBEngine(0)(0).CommitTrans
    AT_Audit "PAYMENT_APPROVE", "tblPayments", paymentID, "SUCCESS", "Official=" & officialNo
    Exit Sub
EH:
    Dim eNum As Long, eDesc As String: eNum = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "PAYMENT_APPROVE", "tblPayments", paymentID, "FAILED", eDesc
    AT_RaiseOriginal "AT_ApprovePayment", eNum, eDesc
End Sub

Public Function AT_CreateTransferDraft(ByVal sourceAccountID As Long, ByVal destinationAccountID As Long, ByVal amount As Currency, Optional ByVal feeAmount As Currency = 0, Optional ByVal notes As String = "", Optional ByVal transactionDate As Date = 0) As Long
    On Error GoTo EH
    AT_RequirePermission "TRANSFER_DRAFT"
    If sourceAccountID = destinationAccountID Then Err.Raise vbObjectError + 360, , "لا يمكن تحويل المبلغ إلى الحساب نفسه."
    If amount <= 0 Or amount <> Fix(amount) Then Err.Raise vbObjectError + 361, , "مبلغ التحويل يجب أن يكون عددًا صحيحًا موجبًا."
    If feeAmount < 0 Or feeAmount <> Fix(feeAmount) Then Err.Raise vbObjectError + 362, , "رسوم التحويل غير صالحة."
    If Not AT_Exists("tblAccounts", "AccountID=" & sourceAccountID & " AND IsActive=True") Or Not AT_Exists("tblAccounts", "AccountID=" & destinationAccountID & " AND IsActive=True") Then Err.Raise vbObjectError + 363, , "حساب المصدر أو الوجهة غير نشط."
    If transactionDate = 0 Then transactionDate = Date
    Dim db As DAO.Database, rs As DAO.Recordset, docID As Long, transferID As Long
    Set db = CurrentDb: DBEngine(0)(0).BeginTrans
    docID = AT_CreateDocument("TRANSFER", transactionDate, notes)
    Set rs = db.OpenRecordset("tblAccountTransfers", dbOpenDynaset)
    rs.AddNew: rs!DocumentID = docID: rs!SourceAccountID = sourceAccountID: rs!DestinationAccountID = destinationAccountID
    rs!Amount = amount: rs!FeeAmount = feeAmount: rs!Notes = notes: rs.Update: rs.Bookmark = rs.LastModified
    transferID = CLng(rs!TransferID): rs.Close
    DBEngine(0)(0).CommitTrans
    AT_Audit "TRANSFER_DRAFT", "tblAccountTransfers", transferID, "SUCCESS", "Draft created; no balance effect."
    AT_CreateTransferDraft = transferID: Exit Function
EH:
    Dim eNum As Long, eDesc As String: eNum = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "TRANSFER_DRAFT", "tblAccountTransfers", 0, "FAILED", eDesc
    AT_RaiseOriginal "AT_CreateTransferDraft", eNum, eDesc
End Function

Public Sub AT_ApproveTransfer(ByVal transferID As Long, ByVal approver As String)
    On Error GoTo EH
    AT_RequirePermission "TRANSFER_APPROVE"
    approver = AT_CurrentActor()
    Dim db As DAO.Database, rs As DAO.Recordset, docID As Long, sourceID As Long, destID As Long, totalReserved As Currency, officialNo As String, r As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT T.*, D.StatusCode FROM tblAccountTransfers AS T INNER JOIN tblFinancialDocuments AS D ON T.DocumentID=D.DocumentID WHERE T.TransferID=" & transferID, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 370, , "التحويل غير موجود."
    docID = CLng(rs!DocumentID): sourceID = CLng(rs!SourceAccountID): destID = CLng(rs!DestinationAccountID): totalReserved = CCur(rs!Amount) + CCur(Nz(rs!FeeAmount, 0))
    If rs!StatusCode <> "DRAFT" Then rs.Close: Err.Raise vbObjectError + 371, , "التحويل ليس مسودة."
    rs.Close
    DBEngine(0)(0).BeginTrans: AT_TransitionDraftToPosting db, docID
    If Not AT_Exists("tblAccounts", "AccountID=" & sourceID & " AND IsActive=True") Or Not AT_Exists("tblAccounts", "AccountID=" & destID & " AND IsActive=True") Then Err.Raise vbObjectError + 374, , "حساب المصدر أو الوجهة غير نشط عند الاعتماد."
    If AT_AccountAvailableBalance(sourceID) < totalReserved Then Err.Raise vbObjectError + 372, , "الرصيد المتاح لا يغطي مبلغ التحويل ورسومه."
    Set r = db.OpenRecordset("tblAccountReservations", dbOpenDynaset)
    r.AddNew: r!DocumentID = docID: r!AccountID = sourceID: r!Amount = totalReserved: r!StatusCode = "ACTIVE": r!CreatedAt = Now(): r!Reason = "حجز مبلغ تحويل حتى تأكيد وصوله": r.Update: r.Close
    officialNo = AT_AssignOfficialNumberInTransaction(db, "TRANSFER", docID, approver)
    db.Execute "UPDATE tblFinancialDocuments SET StatusCode='PENDING', ApprovedAt=Now(), ApprovedByUserName='" & Replace(approver, "'", "''") & "' WHERE DocumentID=" & docID & " AND StatusCode='POSTING'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 373, , "تعذر تثبيت التحويل المعلق."
    DBEngine(0)(0).CommitTrans
    AT_Audit "TRANSFER_PENDING", "tblAccountTransfers", transferID, "SUCCESS", "Official=" & officialNo & "; reserved=" & CStr(totalReserved)
    Exit Sub
EH:
    Dim eNum As Long, eDesc As String: eNum = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "TRANSFER_PENDING", "tblAccountTransfers", transferID, "FAILED", eDesc
    AT_RaiseOriginal "AT_ApproveTransfer", eNum, eDesc
End Sub

Public Sub AT_ConfirmTransfer(ByVal transferID As Long, ByVal approver As String, ByVal confirmationReference As String)
    On Error GoTo EH
    AT_RequirePermission "TRANSFER_APPROVE"
    approver = AT_CurrentActor()
    If Len(Trim$(confirmationReference)) = 0 Then Err.Raise vbObjectError + 380, , "مرجع تأكيد التحويل مطلوب."
    Dim db As DAO.Database, rs As DAO.Recordset, docID As Long, sourceID As Long, destID As Long, amount As Currency, fee As Currency, reservationID As Long, reserveAmount As Currency
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT T.*, D.StatusCode FROM tblAccountTransfers AS T INNER JOIN tblFinancialDocuments AS D ON T.DocumentID=D.DocumentID WHERE T.TransferID=" & transferID, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 381, , "التحويل غير موجود."
    docID = CLng(rs!DocumentID): sourceID = CLng(rs!SourceAccountID): destID = CLng(rs!DestinationAccountID): amount = CCur(rs!Amount): fee = CCur(Nz(rs!FeeAmount, 0))
    If rs!StatusCode <> "PENDING" Then rs.Close: Err.Raise vbObjectError + 382, , "التحويل ليس معلقًا بانتظار التأكيد."
    rs.Close
    DBEngine(0)(0).BeginTrans
    Set rs = db.OpenRecordset("SELECT ReservationID, Amount FROM tblAccountReservations WHERE DocumentID=" & docID & " AND StatusCode='ACTIVE'", dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 383, , "حجز التحويل غير موجود؛ أوقف التأكيد للتحقيق."
    reservationID = CLng(rs!ReservationID): reserveAmount = CCur(rs!Amount): rs.Close
    If reserveAmount <> amount + fee Then Err.Raise vbObjectError + 387, , "قيمة الحجز لا تطابق مبلغ التحويل والرسوم؛ أوقف العملية للتحقيق."
    If AT_AccountAvailableBalance(sourceID, reservationID) < amount + fee Then Err.Raise vbObjectError + 384, , "الرصيد الفعلي بعد استبعاد الحجز غير كافٍ؛ لم يُرحّل التحويل."
    db.Execute "UPDATE tblFinancialDocuments SET StatusCode='POSTING' WHERE DocumentID=" & docID & " AND StatusCode='PENDING'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 385, , "تم تأكيد التحويل مسبقًا أو تغيرت حالته."
    AT_InsertAccountMovement db, sourceID, docID, CDate(DLookup("TransactionDate", "tblFinancialDocuments", "DocumentID=" & docID)), -amount, "TRANSFER_OUT", 1
    If fee > 0 Then AT_InsertAccountMovement db, sourceID, docID, CDate(DLookup("TransactionDate", "tblFinancialDocuments", "DocumentID=" & docID)), -fee, "TRANSFER_FEE", 2
    AT_InsertAccountMovement db, destID, docID, CDate(DLookup("TransactionDate", "tblFinancialDocuments", "DocumentID=" & docID)), amount, "TRANSFER_IN", 3
    db.Execute "UPDATE tblAccountReservations SET StatusCode='RELEASED', ReleasedAt=Now() WHERE ReservationID=" & reservationID & " AND StatusCode='ACTIVE'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 386, , "تعذر تحرير حجز التحويل."
    db.Execute "UPDATE tblAccountTransfers SET ConfirmedAt=Now(), ConfirmedByUserName='" & Replace(approver, "'", "''") & "', ConfirmationRef='" & Replace(confirmationReference, "'", "''") & "' WHERE TransferID=" & transferID, dbFailOnError
    AT_CompleteApproval db, docID, approver
    DBEngine(0)(0).CommitTrans
    AT_Audit "TRANSFER_CONFIRMED", "tblAccountTransfers", transferID, "SUCCESS", "Transfer and fee posted separately; reference=" & confirmationReference
    Exit Sub
EH:
    Dim eNum As Long, eDesc As String: eNum = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "TRANSFER_CONFIRMED", "tblAccountTransfers", transferID, "FAILED", eDesc
    AT_RaiseOriginal "AT_ConfirmTransfer", eNum, eDesc
End Sub

Public Sub AT_CancelPendingTransfer(ByVal transferID As Long, ByVal reason As String, ByVal actor As String)
    On Error GoTo EH
    AT_RequirePermission "DOCUMENT_CANCEL"
    actor = AT_CurrentActor()
    If Len(Trim$(reason)) = 0 Then Err.Raise vbObjectError + 390, , "سبب الإلغاء مطلوب."
    Dim db As DAO.Database, rs As DAO.Recordset, docID As Long
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT T.DocumentID, D.StatusCode FROM tblAccountTransfers AS T INNER JOIN tblFinancialDocuments AS D ON T.DocumentID=D.DocumentID WHERE T.TransferID=" & transferID, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 391, , "التحويل غير موجود."
    docID = CLng(rs!DocumentID): If rs!StatusCode <> "PENDING" Then rs.Close: Err.Raise vbObjectError + 392, , "يسمح بإلغاء التحويل المعلق فقط."
    rs.Close
    DBEngine(0)(0).BeginTrans
    db.Execute "UPDATE tblFinancialDocuments SET StatusCode='CANCELLED', CancelReason='" & Replace(reason, "'", "''") & "' WHERE DocumentID=" & docID & " AND StatusCode='PENDING'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 393, , "تغيرت حالة التحويل؛ لم يُلغَ."
    db.Execute "UPDATE tblAccountReservations SET StatusCode='RELEASED', ReleasedAt=Now(), Reason='" & Replace(reason, "'", "''") & "' WHERE DocumentID=" & docID & " AND StatusCode='ACTIVE'", dbFailOnError
    DBEngine(0)(0).CommitTrans
    AT_Audit "TRANSFER_CANCEL", "tblAccountTransfers", transferID, "SUCCESS", "Reservation released", reason
    Exit Sub
EH:
    Dim eNum As Long, eDesc As String: eNum = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "TRANSFER_CANCEL", "tblAccountTransfers", transferID, "FAILED", eDesc, reason
    AT_RaiseOriginal "AT_CancelPendingTransfer", eNum, eDesc
End Sub

Public Function AT_RequestChargeAdjustment(ByVal chargeID As Long, ByVal signedAmount As Currency, ByVal reason As String, ByVal requester As String, Optional ByVal adjustmentType As String = "DISCOUNT") As Long
    On Error GoTo EH
    AT_RequirePermission "ADJUSTMENT_REQUEST"
    requester = AT_CurrentActor()
    If signedAmount >= 0 Or signedAmount <> Fix(signedAmount) Then Err.Raise vbObjectError + 400, , "يجب أن يكون الخصم مبلغًا صحيحًا سالب الإشارة."
    If Len(Trim$(reason)) = 0 Then Err.Raise vbObjectError + 401, , "سبب طلب الخصم مطلوب."
    If AT_ChargeRemaining(chargeID) + signedAmount < 0 Then Err.Raise vbObjectError + 402, , "الخصم أكبر من المتبقي؛ يلزم مسار تصحيح/رد منفصل."
    Dim db As DAO.Database, rs As DAO.Recordset, docID As Long, adjID As Long
    Set db = CurrentDb: DBEngine(0)(0).BeginTrans
    docID = AT_CreateDocument("ADJUSTMENT", Date, reason)
    Set rs = db.OpenRecordset("tblOwnerAdjustments", dbOpenDynaset)
    rs.AddNew: rs!DocumentID = docID: rs!ChargeID = chargeID: rs!AdjustmentType = adjustmentType
    rs!SignedAmount = signedAmount: rs!Reason = reason: rs!RequestedByUserName = requester: rs.Update: rs.Bookmark = rs.LastModified: adjID = CLng(rs!AdjustmentID): rs.Close
    DBEngine(0)(0).CommitTrans
    AT_Audit "ADJUSTMENT_REQUEST", "tblOwnerAdjustments", adjID, "SUCCESS", "No financial effect until approved.", reason
    AT_RequestChargeAdjustment = adjID: Exit Function
EH:
    Dim eNum As Long, eDesc As String: eNum = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "ADJUSTMENT_REQUEST", "tblOwnerAdjustments", 0, "FAILED", eDesc, reason
    AT_RaiseOriginal "AT_RequestChargeAdjustment", eNum, eDesc
End Function

Public Sub AT_ApproveChargeAdjustment(ByVal adjustmentID As Long, ByVal approver As String)
    On Error GoTo EH
    AT_RequirePermission "ADJUSTMENT_APPROVE"
    approver = AT_CurrentActor()
    Dim db As DAO.Database, rs As DAO.Recordset, docID As Long, chargeID As Long, signedAmount As Currency
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT A.*, D.StatusCode FROM tblOwnerAdjustments AS A INNER JOIN tblFinancialDocuments AS D ON A.DocumentID=D.DocumentID WHERE A.AdjustmentID=" & adjustmentID, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 410, , "طلب الخصم غير موجود."
    docID = CLng(rs!DocumentID): chargeID = CLng(rs!ChargeID): signedAmount = CCur(rs!SignedAmount)
    If rs!StatusCode <> "DRAFT" Then rs.Close: Err.Raise vbObjectError + 411, , "طلب الخصم ليس مسودة."
    rs.Close
    DBEngine(0)(0).BeginTrans: AT_TransitionDraftToPosting db, docID
    If signedAmount >= 0 Or AT_ChargeRemaining(chargeID) + signedAmount < 0 Then Err.Raise vbObjectError + 412, , "الخصم غير صالح أو يخفض المستحق إلى أقل من المسدد."
    db.Execute "UPDATE tblOwnerAdjustments SET ApprovedByUserName='" & Replace(approver, "'", "''") & "', ApprovedAt=Now() WHERE AdjustmentID=" & adjustmentID, dbFailOnError
    AT_CompleteApproval db, docID, approver
    DBEngine(0)(0).CommitTrans
    AT_Audit "ADJUSTMENT_APPROVE", "tblOwnerAdjustments", adjustmentID, "SUCCESS", "Approved signed amount=" & CStr(signedAmount)
    Exit Sub
EH:
    Dim eNum As Long, eDesc As String: eNum = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "ADJUSTMENT_APPROVE", "tblOwnerAdjustments", adjustmentID, "FAILED", eDesc
    AT_RaiseOriginal "AT_ApproveChargeAdjustment", eNum, eDesc
End Sub

Public Sub AT_ApplyOwnerCreditToCharge(ByVal ownerID As Long, ByVal creditMovementID As Long, ByVal chargeID As Long, ByVal amount As Currency, ByVal approver As String, ByVal reason As String)
    On Error GoTo EH
    AT_RequirePermission "CREDIT_APPLY"
    approver = AT_CurrentActor()
    If amount <= 0 Or amount <> Fix(amount) Then Err.Raise vbObjectError + 420, , "مبلغ التطبيق يجب أن يكون عددًا صحيحًا موجبًا."
    If Len(Trim$(reason)) = 0 Then Err.Raise vbObjectError + 421, , "سبب تطبيق الرصيد مطلوب."
    Dim db As DAO.Database, rs As DAO.Recordset, sourceOwnerID As Long, chargeOwnerID As Long, docID As Long, appID As Long
    sourceOwnerID = CLng(Nz(DLookup("OwnerID", "tblOwnerCreditMovements", "CreditMovementID=" & creditMovementID), 0))
    chargeOwnerID = CLng(Nz(DLookup("OwnerID", "tblOwnerCharges", "ChargeID=" & chargeID), 0))
    If sourceOwnerID <> ownerID Or chargeOwnerID <> ownerID Then Err.Raise vbObjectError + 422, , "يجب أن يعود الرصيد والاستحقاق إلى المالك نفسه."
    If AT_CreditMovementAvailable(creditMovementID) < amount Then Err.Raise vbObjectError + 423, , "الرصيد الدائن المتاح غير كافٍ."
    If AT_ChargeRemaining(chargeID) < amount Then Err.Raise vbObjectError + 424, , "المتبقي من الاستحقاق أقل من مبلغ التطبيق."
    Set db = CurrentDb: DBEngine(0)(0).BeginTrans
    ' Recheck both balances inside the transaction before making any ledger entries.
    If AT_CreditMovementAvailable(creditMovementID) < amount Then Err.Raise vbObjectError + 425, , "الرصيد الدائن تغير ولم يعد كافيًا."
    If AT_ChargeRemaining(chargeID) < amount Then Err.Raise vbObjectError + 426, , "المتبقي من الاستحقاق تغير ولم يعد كافيًا."
    docID = AT_CreateDocument("CREDIT_APPLICATION", Date, reason): AT_TransitionDraftToPosting db, docID
    Set rs = db.OpenRecordset("tblCreditApplications", dbOpenDynaset)
    rs.AddNew: rs!DocumentID = docID: rs!OwnerID = ownerID: rs!Amount = amount: rs!ApprovedByUserName = approver: rs!ApprovedAt = Now(): rs!Reason = reason
    rs.Update: rs.Bookmark = rs.LastModified: appID = CLng(rs!CreditApplicationID): rs.Close
    Set rs = db.OpenRecordset("tblCreditApplicationLines", dbOpenDynaset)
    rs.AddNew: rs!CreditApplicationID = appID: rs!CreditMovementID = creditMovementID: rs!ChargeID = chargeID: rs!Amount = amount: rs.Update: rs.Close
    Set rs = db.OpenRecordset("tblChargeSettlementLines", dbOpenDynaset)
    rs.AddNew: rs!ChargeID = chargeID: rs!DocumentID = docID: rs!CreditApplicationID = appID: rs!SignedAmount = amount: rs!SettlementType = "CREDIT_APPLICATION": rs!CreatedAt = Now(): rs.Update: rs.Close
    Set rs = db.OpenRecordset("tblOwnerCreditMovements", dbOpenDynaset)
    rs.AddNew: rs!OwnerID = ownerID: rs!DocumentID = docID: rs!SignedAmount = -amount: rs!CreditType = "CREDIT_APPLIED": rs!SourceCreditMovementID = creditMovementID: rs!Reason = reason: rs!CreatedAt = Now(): rs.Update: rs.Close
    AT_CompleteApproval db, docID, approver
    DBEngine(0)(0).CommitTrans
    AT_Audit "CREDIT_APPLY", "tblCreditApplications", appID, "SUCCESS", "Amount=" & CStr(amount), reason
    Exit Sub
EH:
    Dim eNum As Long, eDesc As String: eNum = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "CREDIT_APPLY", "tblCreditApplications", 0, "FAILED", eDesc, reason
    AT_RaiseOriginal "AT_ApplyOwnerCreditToCharge", eNum, eDesc
End Sub

Public Sub AT_ReverseReceiptAllocation(ByVal allocationID As Long, ByVal amount As Currency, ByVal approver As String, ByVal reason As String)
    On Error GoTo EH
    AT_RequirePermission "ADJUSTMENT_APPROVE"
    approver = AT_CurrentActor()
    If amount <= 0 Or amount <> Fix(amount) Then Err.Raise vbObjectError + 430, , "مبلغ العكس يجب أن يكون عددًا صحيحًا موجبًا."
    If Len(Trim$(reason)) = 0 Then Err.Raise vbObjectError + 431, , "سبب العكس مطلوب."
    Dim db As DAO.Database, rs As DAO.Recordset, receiptID As Long, chargeID As Long, ownerID As Long, originalAmount As Currency, alreadyReversed As Currency, docID As Long, correctionID As Long
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT A.AllocationID, A.ReceiptID, A.ChargeID, A.AllocatedAmount, R.OwnerID, D.StatusCode FROM (tblReceiptAllocationLines AS A INNER JOIN tblReceipts AS R ON A.ReceiptID=R.ReceiptID) INNER JOIN tblFinancialDocuments AS D ON R.DocumentID=D.DocumentID WHERE A.AllocationID=" & allocationID, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 432, , "سطر التخصيص غير موجود."
    receiptID = CLng(rs!ReceiptID): chargeID = CLng(rs!ChargeID): ownerID = CLng(rs!OwnerID): originalAmount = CCur(rs!AllocatedAmount)
    If rs!StatusCode <> "APPROVED" Then rs.Close: Err.Raise vbObjectError + 433, , "لا يمكن عكس تخصيص إيصال غير معتمد."
    rs.Close
    Dim sql As String
    sql = "SELECT Sum(L.Amount) AS ReversedAmount FROM tblReceiptCorrectionLines AS L INNER JOIN tblReceiptCorrections AS C ON L.ReceiptCorrectionID=C.ReceiptCorrectionID INNER JOIN tblFinancialDocuments AS D ON C.DocumentID=D.DocumentID WHERE L.OriginalAllocationID=" & allocationID & " AND D.StatusCode='APPROVED'"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot): alreadyReversed = AT_GetCurrency(rs, "ReversedAmount"): rs.Close
    If amount > originalAmount - alreadyReversed Then Err.Raise vbObjectError + 434, , "مبلغ العكس يتجاوز الجزء غير المعكوس من التخصيص الأصلي."
    DBEngine(0)(0).BeginTrans
    sql = "SELECT Sum(L.Amount) AS ReversedAmount FROM tblReceiptCorrectionLines AS L INNER JOIN tblReceiptCorrections AS C ON L.ReceiptCorrectionID=C.ReceiptCorrectionID INNER JOIN tblFinancialDocuments AS D ON C.DocumentID=D.DocumentID WHERE L.OriginalAllocationID=" & allocationID & " AND D.StatusCode='APPROVED'"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot): alreadyReversed = AT_GetCurrency(rs, "ReversedAmount"): rs.Close
    If amount > originalAmount - alreadyReversed Then Err.Raise vbObjectError + 435, , "مبلغ العكس تجاوز الجزء المتبقي أثناء بدء العملية؛ لم يُرحّل."
    docID = AT_CreateDocument("RECEIPT_CORRECTION", Date, reason): AT_TransitionDraftToPosting db, docID
    Set rs = db.OpenRecordset("tblReceiptCorrections", dbOpenDynaset)
    rs.AddNew: rs!DocumentID = docID: rs!OriginalReceiptID = receiptID: rs!CorrectionType = "ALLOCATION_REVERSAL": rs!Reason = reason
    rs.Update: rs.Bookmark = rs.LastModified: correctionID = CLng(rs!ReceiptCorrectionID): rs.Close
    Set rs = db.OpenRecordset("tblReceiptCorrectionLines", dbOpenDynaset)
    rs.AddNew: rs!ReceiptCorrectionID = correctionID: rs!OriginalAllocationID = allocationID: rs!Amount = amount: rs!CreatedAt = Now(): rs.Update: rs.Close
    Set rs = db.OpenRecordset("tblChargeSettlementLines", dbOpenDynaset)
    rs.AddNew: rs!ChargeID = chargeID: rs!DocumentID = docID: rs!ReceiptID = receiptID: rs!ReceiptCorrectionID = correctionID
    rs!SignedAmount = -amount: rs!SettlementType = "RECEIPT_REVERSAL": rs!CreatedAt = Now(): rs.Update: rs.Close
    ' Cash remains in the account. Reversed allocation therefore becomes owner credit unless a separate cash refund is approved.
    Set rs = db.OpenRecordset("tblOwnerCreditMovements", dbOpenDynaset)
    rs.AddNew: rs!OwnerID = ownerID: rs!DocumentID = docID: rs!SignedAmount = amount
    rs!CreditType = "RECEIPT_ALLOCATION_REVERSED": rs!SourceReceiptID = receiptID
    rs!Reason = "رصيد دائن ناتج عن عكس تخصيص الإيصال: " & reason: rs!CreatedAt = Now(): rs.Update: rs.Close
    AT_CompleteApproval db, docID, approver
    DBEngine(0)(0).CommitTrans
    AT_Audit "RECEIPT_ALLOCATION_REVERSE", "tblReceiptCorrectionLines", correctionID, "SUCCESS", "Reversed=" & CStr(amount), reason
    Exit Sub
EH:
    Dim eNum As Long, eDesc As String: eNum = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "RECEIPT_ALLOCATION_REVERSE", "tblReceiptCorrectionLines", allocationID, "FAILED", eDesc, reason
    AT_RaiseOriginal "AT_ReverseReceiptAllocation", eNum, eDesc
End Sub

Public Sub AT_CancelDraftDocument(ByVal documentID As Long, ByVal reason As String)
    On Error GoTo EH
    AT_RequirePermission "DOCUMENT_CANCEL"
    If Len(Trim$(reason)) = 0 Then Err.Raise vbObjectError + 440, , "سبب الإلغاء مطلوب."
    Dim db As DAO.Database
    Set db = CurrentDb: DBEngine(0)(0).BeginTrans
    db.Execute "UPDATE tblFinancialDocuments SET StatusCode='CANCELLED', CancelReason='" & Replace(reason, "'", "''") & "' WHERE DocumentID=" & documentID & " AND StatusCode='DRAFT'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 441, , "يسمح بالإلغاء للمسودات فقط؛ لا يمكن حذف أو إلغاء مستند مرحّل بهذه الطريقة."
    DBEngine(0)(0).CommitTrans
    AT_Audit "DOCUMENT_CANCEL", "tblFinancialDocuments", documentID, "SUCCESS", "Draft cancelled", reason
    Exit Sub
EH:
    Dim eNum As Long, eDesc As String: eNum = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "DOCUMENT_CANCEL", "tblFinancialDocuments", documentID, "FAILED", eDesc, reason
    AT_RaiseOriginal "AT_CancelDraftDocument", eNum, eDesc
End Sub

Public Sub AT_RunIntegrityReview()
    On Error GoTo EH
    AT_RequirePermission "INTEGRITY_RUN"
    Dim db As DAO.Database, rs As DAO.Recordset, runID As Long, accountID As Long, chargeID As Long, balance As Currency, remaining As Currency
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblIntegrityRuns", dbOpenDynaset)
    rs.AddNew: rs!StartedAt = Now(): rs!RunByUserName = AT_CurrentActor(): rs!ResultSummary = "Running": rs.Update: rs.Bookmark = rs.LastModified: runID = CLng(rs!IntegrityRunID): rs.Close
    ' Findings are flags only. This routine intentionally never edits financial data.
    db.Execute "INSERT INTO tblIntegrityFindings (IntegrityRunID,Severity,EntityName,EntityID,FindingCode,FindingText,ResolutionStatus) SELECT " & runID & ",'ERROR','tblOwners',OwnerID,'OWNER_CODE_MISSING','Active owner has no official code','OPEN' FROM tblOwners WHERE AdminStatus='ACTIVE' AND (OwnerCode Is Null OR OwnerCode='')", dbFailOnError
    db.Execute "INSERT INTO tblIntegrityFindings (IntegrityRunID,Severity,EntityName,EntityID,FindingCode,FindingText,ResolutionStatus) SELECT " & runID & ",'ERROR','tblFinancialDocuments',D.DocumentID,'APPROVED_NO_NUMBER','Approved official document has no official number','OPEN' FROM tblFinancialDocuments AS D WHERE D.StatusCode='APPROVED' AND D.DocumentType In ('RECEIPT','PAYMENT','TRANSFER') AND NOT EXISTS (SELECT DocumentID FROM tblOfficialDocumentNumbers AS N WHERE N.DocumentID=D.DocumentID)", dbFailOnError
    db.Execute "INSERT INTO tblIntegrityFindings (IntegrityRunID,Severity,EntityName,EntityID,FindingCode,FindingText,ResolutionStatus) SELECT " & runID & ",'ERROR','tblAccountMovements',M.AccountMovementID,'ORPHAN_MOVEMENT','Movement references missing document','OPEN' FROM tblAccountMovements AS M LEFT JOIN tblFinancialDocuments AS D ON M.DocumentID=D.DocumentID WHERE D.DocumentID Is Null", dbFailOnError
    Set rs = db.OpenRecordset("SELECT AccountID FROM tblAccounts WHERE IsActive=True", dbOpenSnapshot)
    Do While Not rs.EOF
        accountID = CLng(rs!AccountID): balance = AT_AccountBookBalance(accountID)
        If balance < 0 Then AT_AddIntegrityFinding db, runID, "ERROR", "tblAccounts", accountID, "NEGATIVE_ACCOUNT_BALANCE", "الرصيد الدفتري للحساب سالب: " & CStr(balance)
        If AT_AccountReservedAmount(accountID) > balance Then AT_AddIntegrityFinding db, runID, "ERROR", "tblAccounts", accountID, "RESERVATIONS_EXCEED_BALANCE", "إجمالي الحجوزات النشطة يتجاوز الرصيد الدفتري."
        rs.MoveNext
    Loop
    rs.Close
    Set rs = db.OpenRecordset("SELECT ChargeID FROM tblOwnerCharges WHERE ChargeStatus='ACTIVE'", dbOpenSnapshot)
    Do While Not rs.EOF
        chargeID = CLng(rs!ChargeID): remaining = AT_ChargeRemaining(chargeID)
        If remaining < 0 Then AT_AddIntegrityFinding db, runID, "ERROR", "tblOwnerCharges", chargeID, "NEGATIVE_REMAINING", "المتبقي من الاستحقاق سالب؛ راجع التسويات والخصومات."
        rs.MoveNext
    Loop
    rs.Close
    Set rs = db.OpenRecordset("SELECT CreditMovementID FROM tblOwnerCreditMovements WHERE SignedAmount>0 AND SourceCreditMovementID Is Null", dbOpenSnapshot)
    Do While Not rs.EOF
        If AT_CreditMovementRawAvailable(CLng(rs!CreditMovementID)) < 0 Then AT_AddIntegrityFinding db, runID, "ERROR", "tblOwnerCreditMovements", CLng(rs!CreditMovementID), "NEGATIVE_CREDIT_AVAILABLE", "الرصيد المتاح من حركة الدائن أصبح سالبًا."
        rs.MoveNext
    Loop
    rs.Close
    db.Execute "UPDATE tblIntegrityRuns SET CompletedAt=Now(), ResultSummary='Integrity checks completed; findings require human review. No auto-correction performed.' WHERE IntegrityRunID=" & runID, dbFailOnError
    AT_Audit "INTEGRITY_REVIEW", "tblIntegrityRuns", runID, "SUCCESS", "Flags generated without auto-correction."
    MsgBox "اكتمل فحص السلامة. راجع tblIntegrityFindings؛ لم تُعدّل أي بيانات مالية تلقائيًا.", vbInformation, "Archon Tech"
    Exit Sub
EH:
    Dim eNum As Long, eDesc As String: eNum = Err.Number: eDesc = Err.Description
    AT_Audit "INTEGRITY_REVIEW", "tblIntegrityRuns", 0, "FAILED", eDesc
    MsgBox "تعذر إكمال فحص السلامة: " & eNum & " - " & eDesc, vbExclamation, "Archon Tech"
End Sub

Private Sub AT_AddIntegrityFinding(ByVal db As DAO.Database, ByVal runID As Long, ByVal severity As String, ByVal entityName As String, ByVal entityID As Long, ByVal findingCode As String, ByVal findingText As String)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("tblIntegrityFindings", dbOpenDynaset)
    rs.AddNew: rs!IntegrityRunID = runID: rs!Severity = severity: rs!EntityName = entityName
    rs!EntityID = entityID: rs!FindingCode = findingCode: rs!FindingText = findingText: rs!ResolutionStatus = "OPEN"
    rs.Update: rs.Close
End Sub
