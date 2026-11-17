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

Private Sub AT_RaiseOriginal(ByVal sourceName As String, ByVal savedErrNumber As Long, ByVal eDesc As String)
    Err.Raise savedErrNumber, sourceName, eDesc
End Sub

Private Function AT_Exists(ByVal tableName As String, ByVal criteria As String) As Boolean
    AT_Exists = (DCount("*", tableName, criteria) > 0)
End Function

Public Sub AT_RequirePermission(ByVal permissionCode As String)
    Dim db As DAO.Database, rs As DAO.Recordset, userID As Long, permID As Long, allowed As Boolean
    If Not AT_Exists("tblUsers", "UserName='" & Replace(AT_CurrentActor(), "'", "''") & "' AND IsActive=True") Then _
        Err.Raise vbObjectError + 190, "AT_RequirePermission", AT_U("0644,0627,0020,064A,0648,062C,062F,0020,062D,0633,0627,0628,0020,0645,0633,062A,062E,062F,0645,0020,0646,0634,0637,0020,0645,0631,062A,0628,0637,0020,0628,0645,0633,062A,062E,062F,0645,0020,0057,0069,006E,0064,006F,0077,0073,0020,0627,0644,062D,0627,0644,064A,002E")
    userID = CLng(Nz(DLookup("UserID", "tblUsers", "UserName='" & Replace(AT_CurrentActor(), "'", "''") & "' AND IsActive=True"), 0))
    permID = CLng(Nz(DLookup("PermissionID", "tblPermissions", "PermissionCode='" & Replace(permissionCode, "'", "''") & "' AND IsActive=True"), 0))
    If userID = 0 Or permID = 0 Then Err.Raise vbObjectError + 191, "AT_RequirePermission", AT_U("062A,0639,0630,0631,0020,062A,062D,062F,064A,062F,0020,0627,0644,0645,0633,062A,062E,062F,0645,0020,0623,0648,0020,0627,0644,0635,0644,0627,062D,064A,0629,0020,0627,0644,0645,0637,0644,0648,0628,0629,003A,0020") & permissionCode
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
    If Not allowed Then Err.Raise vbObjectError + 192, "AT_RequirePermission", AT_U("0644,064A,0633,062A,0020,0644,062F,064A,0643,0020,0627,0644,0635,0644,0627,062D,064A,0629,0020,0644,062A,0646,0641,064A,0630,0020,0627,0644,0639,0645,0644,064A,0629,003A,0020") & permissionCode
End Sub

Private Function AT_CreateDocument(ByVal documentType As String, ByVal transactionDate As Date, ByVal reason As String) As Long
    Dim db As DAO.Database, rs As DAO.Recordset
    AT_EnsureDateOpen transactionDate
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
    If numericValue < 1 Or padLength < 1 Then Err.Raise vbObjectError + 204, "AT_FormatSeriesNumber", AT_U("0642,064A,0645,0629,0020,0627,0644,062A,0631,0642,064A,0645,0020,0623,0648,0020,0637,0648,0644,0020,0627,0644,062A,0639,0628,0626,0629,0020,063A,064A,0631,0020,0635,0627,0644,062D,002E")
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
        Err.Raise vbObjectError + 200, "AT_NextSeriesValueInTransaction", AT_U("0633,0644,0633,0644,0629,0020,0627,0644,062A,0631,0642,064A,0645,0020,063A,064A,0631,0020,0645,0647,064A,0623,0629,003A,0020") & seriesCode
    End If
    n = CLng(rs!NextValue): prefix = AT_GetText(rs, "Prefix"): pad = CLng(Nz(rs!PadLength, 6))
    If n < 1 Then rs.Close: Err.Raise vbObjectError + 201, "AT_NextSeriesValueInTransaction", AT_U("0642,064A,0645,0629,0020,0639,062F,0627,062F,0020,0627,0644,062A,0631,0642,064A,0645,0020,063A,064A,0631,0020,0635,0627,0644,062D,0629,002E")
    rs.Edit: rs!NextValue = n + 1: rs.Update: rs.Close
    AT_NextSeriesValueInTransaction = AT_FormatSeriesNumber(prefix, n, pad)
End Function

Private Function AT_AssignOfficialNumberInTransaction(ByVal db As DAO.Database, ByVal seriesCode As String, ByVal documentID As Long, ByVal actor As String) As String
    Dim numberText As String, rs As DAO.Recordset, n As Long, prefix As String, pad As Long
    Set rs = db.OpenRecordset("SELECT * FROM tblNumberSeries WHERE SeriesCode='" & Replace(seriesCode, "'", "''") & "'", dbOpenDynaset)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 202, "AT_AssignOfficialNumberInTransaction", AT_U("0633,0644,0633,0644,0629,0020,0627,0644,062A,0631,0642,064A,0645,0020,063A,064A,0631,0020,0645,0647,064A,0623,0629,003A,0020") & seriesCode
    n = CLng(rs!NextValue): prefix = AT_GetText(rs, "Prefix"): pad = CLng(Nz(rs!PadLength, 6))
    If n < 1 Then rs.Close: Err.Raise vbObjectError + 203, "AT_AssignOfficialNumberInTransaction", AT_U("0642,064A,0645,0629,0020,0639,062F,0627,062F,0020,0627,0644,062A,0631,0642,064A,0645,0020,063A,064A,0631,0020,0635,0627,0644,062D,0629,002E")
    numberText = AT_FormatSeriesNumber(prefix, n, pad)
    rs.Edit: rs!NextValue = n + 1: rs.Update: rs.Close
    Set rs = db.OpenRecordset("tblOfficialDocumentNumbers", dbOpenDynaset)
    rs.AddNew: rs!DocumentID = documentID: rs!OfficialNumber = numberText: rs!SeriesCode = seriesCode
    rs!NumericValue = n: rs!AssignedAt = Now(): rs!AssignedByUserName = actor: rs.Update: rs.Close
    AT_AssignOfficialNumberInTransaction = numberText
End Function

Private Sub AT_TransitionDraftToPosting(ByVal db As DAO.Database, ByVal documentID As Long)
    db.Execute "UPDATE tblFinancialDocuments SET StatusCode='POSTING' WHERE DocumentID=" & documentID & " AND StatusCode='DRAFT'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 210, "AT_TransitionDraftToPosting", AT_U("0627,0644,0645,0633,062A,0646,062F,0020,0644,064A,0633,0020,0645,0633,0648,062F,0629,0020,0623,0648,0020,0633,0628,0642,0020,062A,0631,062D,064A,0644,0647,061B,0020,0623,0644,063A,064A,062A,0020,0627,0644,0639,0645,0644,064A,0629,0020,0644,0645,0646,0639,0020,0627,0644,062A,0643,0631,0627,0631,002E")
End Sub

Private Sub AT_CompleteApproval(ByVal db As DAO.Database, ByVal documentID As Long, ByVal actor As String)
    db.Execute "UPDATE tblFinancialDocuments SET StatusCode='APPROVED', ApprovedAt=Now(), ApprovedByUserName='" & Replace(actor, "'", "''") & "' WHERE DocumentID=" & documentID & " AND StatusCode='POSTING'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 211, "AT_CompleteApproval", AT_U("062A,0639,0630,0631,0020,062A,062B,0628,064A,062A,0020,0627,0639,062A,0645,0627,062F,0020,0627,0644,0645,0633,062A,0646,062F,002E")
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
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "SETTINGS_ADMIN"
    accountCode = UCase$(Trim$(accountCode)): accountName = Trim$(accountName): accountType = UCase$(Trim$(accountType))
    If Len(accountCode) = 0 Or Len(accountName) = 0 Then Err.Raise vbObjectError + 195, , AT_U("0631,0645,0632,0020,0627,0644,062D,0633,0627,0628,0020,0648,0627,0633,0645,0647,0020,0645,0637,0644,0648,0628,0627,0646,002E")
    If Not (accountType = "CASH" Or accountType = "BANK" Or accountType = "EWALLET" Or accountType = "OTHER") Then _
        Err.Raise vbObjectError + 196, , AT_U("0646,0648,0639,0020,0627,0644,062D,0633,0627,0628,0020,063A,064A,0631,0020,0635,0627,0644,062D,002E")
    If AT_Exists("tblAccounts", "AccountCode='" & Replace(accountCode, "'", "''") & "'") Then _
        Err.Raise vbObjectError + 197, , AT_U("0631,0645,0632,0020,0627,0644,062D,0633,0627,0628,0020,0645,0633,062A,062E,062F,0645,0020,0645,0633,0628,0642,064B,0627,002E")
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
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "ACCOUNT_CREATE", "tblAccounts", 0, "FAILED", eDesc
    AT_RaiseOriginal "AT_CreateAccount", savedErrNumber, eDesc
End Function

Public Function AT_AccountBookBalance(ByVal accountID As Long, Optional ByVal asOfDate As Date = 0) As Currency
    Dim sql As String, rs As DAO.Recordset
    sql = "SELECT Sum(M.SignedAmount) AS BalanceValue FROM tblAccountMovements AS M INNER JOIN tblFinancialDocuments AS D ON M.DocumentID=D.DocumentID WHERE M.AccountID=" & accountID & " AND D.StatusCode='APPROVED'"
    ' The optional parameter is a calendar date: include every movement before the next midnight.
    If asOfDate <> 0 Then sql = sql & " AND M.MovementDate<" & AT_SqlDate(DateAdd("d", 1, DateValue(asOfDate)))
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
    gross = Nz(DLookup("NetAmount", "tblOwnerCharges", "ChargeID=" & chargeID), Nz(DLookup("GrossAmount", "tblOwnerCharges", "ChargeID=" & chargeID), 0))
    sql = "SELECT Sum(A.SignedAmount) AS AdjustmentTotal FROM tblOwnerAdjustments AS A INNER JOIN tblFinancialDocuments AS D ON A.DocumentID=D.DocumentID WHERE A.ChargeID=" & chargeID & " AND D.StatusCode='APPROVED'"
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot): adjustments = AT_GetCurrency(rs, "AdjustmentTotal"): rs.Close
    sql = "SELECT Sum(S.SignedAmount) AS SettlementTotal FROM tblChargeSettlementLines AS S INNER JOIN tblFinancialDocuments AS D ON S.DocumentID=D.DocumentID WHERE S.ChargeID=" & chargeID & " AND D.StatusCode='APPROVED'"
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot): settled = AT_GetCurrency(rs, "SettlementTotal"): rs.Close
    AT_ChargeRemaining = gross + adjustments - settled
End Function

Public Function AT_CreateOwnerDraft(ByVal fullName As String, ByVal phone As String, ByVal houseCount As Long, ByVal apartmentCount As Long, Optional ByVal notes As String = "", Optional ByVal identityNo As String = "") As Long
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "OWNER_EDIT"
    If Len(Trim$(fullName)) = 0 Then Err.Raise vbObjectError + 220, , AT_U("0627,0633,0645,0020,0627,0644,0645,0627,0644,0643,0020,0645,0637,0644,0648,0628,002E")
    ' HouseCount and ApartmentCount are derived summary caches, not user-entered source data.
    Dim db As DAO.Database, rs As DAO.Recordset, ownerID As Long
    Set db = CurrentDb: DBEngine(0)(0).BeginTrans
    Set rs = db.OpenRecordset("tblOwners", dbOpenDynaset)
    rs.AddNew: rs!FullName = Trim$(fullName): rs!FullNameSearch = AT_NormalizeArabicSearch(fullName)
    rs!Phone = Trim$(phone): rs!IdentityNo = Trim$(identityNo): rs!HouseCount = 0: rs!ApartmentCount = 0
    rs!Notes = notes: rs!AdminStatus = "DRAFT": rs!DataStatus = IIf(Len(Trim$(phone)) = 0, "INCOMPLETE", "COMPLETE")
    rs!ReviewStatus = "NOT_REVIEWED": rs!ActivityStatus = "ACTIVE": rs!CreatedAt = Now(): rs!UpdatedAt = Now(): rs.Update
    rs.Bookmark = rs.LastModified: ownerID = CLng(rs!OwnerID): rs.Close
    DBEngine(0)(0).CommitTrans
    AT_Audit "OWNER_DRAFT", "tblOwners", ownerID, "SUCCESS", "Owner draft saved; no official code assigned."
    AT_CreateOwnerDraft = ownerID: Exit Function
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "OWNER_DRAFT", "tblOwners", 0, "FAILED", eDesc
    AT_RaiseOriginal "AT_CreateOwnerDraft", savedErrNumber, eDesc
End Function

Public Sub AT_ApproveOwner(ByVal ownerID As Long, ByVal approver As String)
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "OWNER_APPROVE"
    approver = AT_CurrentActor()
    Dim db As DAO.Database, rs As DAO.Recordset, newCode As String
    Set db = CurrentDb: DBEngine(0)(0).BeginTrans
    db.Execute "UPDATE tblOwners SET AdminStatus='APPROVING' WHERE OwnerID=" & ownerID & " AND AdminStatus='DRAFT'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 230, , AT_U("0627,0644,0645,0627,0644,0643,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0020,0643,0645,0633,0648,062F,0629,0020,0623,0648,0020,0633,0628,0642,0020,0627,0639,062A,0645,0627,062F,0647,002E")
    newCode = AT_NextSeriesValueInTransaction(db, "OWNER")
    db.Execute "UPDATE tblOwners SET OwnerCode='" & Replace(newCode, "'", "''") & "', AdminStatus='ACTIVE', UpdatedAt=Now() WHERE OwnerID=" & ownerID & " AND AdminStatus='APPROVING'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 231, , AT_U("062A,0639,0630,0631,0020,062A,062B,0628,064A,062A,0020,0631,0642,0645,0020,0627,0644,0645,0627,0644,0643,002E")
    DBEngine(0)(0).CommitTrans
    AT_Audit "OWNER_APPROVE", "tblOwners", ownerID, "SUCCESS", "Assigned " & newCode, "Approved by " & approver
    Exit Sub
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "OWNER_APPROVE", "tblOwners", ownerID, "FAILED", eDesc
    AT_RaiseOriginal "AT_ApproveOwner", savedErrNumber, eDesc
End Sub

Public Sub AT_UpdateOwnerDetails(ByVal ownerID As Long, ByVal fullName As String, ByVal phone As String, ByVal houseCount As Long, ByVal apartmentCount As Long, ByVal notes As String, ByVal reason As String, Optional ByVal identityNo As String = "")
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "OWNER_EDIT"
    If ownerID <= 0 Or Len(Trim$(fullName)) = 0 Then Err.Raise vbObjectError + 225, , AT_U("0645,0639,0631,0641,0020,0627,0644,0645,0627,0644,0643,0020,0648,0627,0644,0627,0633,0645,0020,0645,0637,0644,0648,0628,0627,0646,002E")
    If Len(Trim$(reason)) = 0 Then Err.Raise vbObjectError + 227, , AT_U("0633,0628,0628,0020,062A,0639,062F,064A,0644,0020,0628,064A,0627,0646,0627,062A,0020,0627,0644,0645,0627,0644,0643,0020,0645,0637,0644,0648,0628,0020,0644,0644,062A,062F,0642,064A,0642,002E")

    Dim db As DAO.Database, rs As DAO.Recordset, userID As Long, auditID As Long
    Dim oldName As String, oldPhone As String, oldIdentity As String, oldNotes As String
    Dim newName As String, newPhone As String, newIdentity As String, newNotes As String, dataStatus As String
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tblOwners WHERE OwnerID=" & ownerID & " AND AdminStatus='ACTIVE'", dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 228, , AT_U("0627,0644,0645,0627,0644,0643,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0020,0623,0648,0020,063A,064A,0631,0020,0646,0634,0637,002E")
    oldName = AT_GetText(rs, "FullName"): oldPhone = AT_GetText(rs, "Phone"): oldIdentity = AT_GetText(rs, "IdentityNo"): oldNotes = AT_GetText(rs, "Notes")
    rs.Close
    newName = Trim$(fullName): newPhone = Trim$(phone): newIdentity = Trim$(identityNo): newNotes = notes
    dataStatus = IIf(Len(newPhone) = 0, "INCOMPLETE", "COMPLETE")
    userID = CLng(Nz(DLookup("UserID", "tblUsers", "UserName='" & Replace(AT_CurrentActor(), "'", "''") & "' AND IsActive=True"), 0))

    DBEngine(0)(0).BeginTrans
    db.Execute "UPDATE tblOwners SET FullName='" & Replace(newName, "'", "''") & "', FullNameSearch='" & Replace(AT_NormalizeArabicSearch(newName), "'", "''") & "', Phone='" & Replace(newPhone, "'", "''") & "', IdentityNo='" & Replace(newIdentity, "'", "''") & "', Notes='" & Replace(notes, "'", "''") & "', DataStatus='" & dataStatus & "', UpdatedAt=Now(), UpdatedByUserID=" & IIf(userID > 0, CStr(userID), "Null") & " WHERE OwnerID=" & ownerID & " AND AdminStatus='ACTIVE'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 229, , AT_U("0644,0645,0020,064A,064F,062D,062F,0651,062B,0020,0633,062C,0644,0020,0627,0644,0645,0627,0644,0643,061B,0020,0631,0628,0645,0627,0020,062A,063A,064A,0631,062A,0020,062D,0627,0644,062A,0647,002E")
    AT_RefreshOwnerSummaryCounts ownerID

    Set rs = db.OpenRecordset("tblAuditEvents", dbOpenDynaset)
    rs.AddNew: rs!EventAt = Now(): rs!UserName = AT_CurrentActor(): rs!ActionCode = "OWNER_UPDATE": rs!EntityName = "tblOwners"
    rs!EntityID = ownerID: rs!ResultCode = "SUCCESS": rs!Reason = reason: rs!Detail = "Updated owner profile fields": rs.Update
    rs.Bookmark = rs.LastModified: auditID = CLng(rs!AuditEventID): rs.Close
    If oldName <> newName Then AT_AddAuditChange db, auditID, "FullName", oldName, newName
    If oldPhone <> newPhone Then AT_AddAuditChange db, auditID, "Phone", oldPhone, newPhone
    If oldIdentity <> newIdentity Then AT_AddAuditChange db, auditID, "IdentityNo", oldIdentity, newIdentity
    If oldNotes <> newNotes Then AT_AddAuditChange db, auditID, "Notes", oldNotes, newNotes
    DBEngine(0)(0).CommitTrans
    Exit Sub
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "OWNER_UPDATE", "tblOwners", ownerID, "FAILED", eDesc, reason
    AT_RaiseOriginal "AT_UpdateOwnerDetails", savedErrNumber, eDesc
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
    s = Replace(s, AT_U("0623"), AT_U("0627")): s = Replace(s, AT_U("0625"), AT_U("0627")): s = Replace(s, AT_U("0622"), AT_U("0627"))
    s = Replace(s, AT_U("0649"), AT_U("064A")): s = Replace(s, AT_U("0629"), AT_U("0647")): s = Replace(s, ChrW(&H640), "")
    marks = Array(&H64B, &H64C, &H64D, &H64E, &H64F, &H650, &H651, &H652, &H653, &H654, &H655, &H670)
    For i = LBound(marks) To UBound(marks): s = Replace(s, ChrW(marks(i)), ""): Next i
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    AT_NormalizeArabicSearch = LCase$(s)
End Function

Public Function AT_CreateChargeForOwner(ByVal ownerID As Long, ByVal chargeDate As Date, ByVal rateID As Long, Optional ByVal reason As String = "") As Long
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "CHARGE_DRAFT"
    Dim db As DAO.Database, rr As DAO.Recordset, ru As DAO.Recordset, rc As DAO.Recordset, rl As DAO.Recordset
    Dim ownerApartments As Long, rateAmount As Currency, docID As Long, chargeID As Long, gross As Currency
    Dim ownerCode As String, ownerName As String
    If Not AT_Exists("tblOwners", "OwnerID=" & ownerID & " AND AdminStatus='ACTIVE'") Then Err.Raise vbObjectError + 240, , AT_U("0627,0644,0645,0627,0644,0643,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0020,0623,0648,0020,063A,064A,0631,0020,0646,0634,0637,002E")
    If chargeDate = 0 Then chargeDate = Date
    Set rr = CurrentDb.OpenRecordset("SELECT * FROM tblContributionRates WHERE RateID=" & rateID & " AND IsActive=True", dbOpenSnapshot)
    If rr.EOF Then rr.Close: Err.Raise vbObjectError + 241, , AT_U("0645,0639,062F,0644,0020,0627,0644,0645,0633,0627,0647,0645,0629,0020,063A,064A,0631,0020,0641,0639,0627,0644,0020,0623,0648,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,002E")
    If CDate(rr!EffectiveFrom) > chargeDate Then rr.Close: Err.Raise vbObjectError + 242, , AT_U("062A,0627,0631,064A,062E,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,0020,064A,0633,0628,0642,0020,0628,062F,0627,064A,0629,0020,0633,0631,064A,0627,0646,0020,0627,0644,0633,0639,0631,002E")
    If Not IsNull(rr!EffectiveTo) Then If CDate(rr!EffectiveTo) < chargeDate Then rr.Close: Err.Raise vbObjectError + 243, , AT_U("0627,0646,062A,0647,0649,0020,0633,0631,064A,0627,0646,0020,0645,0639,062F,0644,0020,0627,0644,0645,0633,0627,0647,0645,0629,0020,0641,064A,0020,062A,0627,0631,064A,062E,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,002E")
    rateAmount = CCur(rr!RateAmount): rr.Close
    If rateAmount <= 0 Then Err.Raise vbObjectError + 245, , AT_U("0645,0639,062F,0644,0020,0627,0644,0645,0633,0627,0647,0645,0629,0020,064A,062C,0628,0020,0623,0646,0020,064A,0643,0648,0646,0020,0645,0648,062C,0628,064B,0627,002E")
    If DCount("*", "tblOwnerCharges", "OwnerID=" & ownerID & " AND ChargeDate=#" & Format$(chargeDate, "mm\/dd\/yyyy") & "# AND ChargeStatus In ('DRAFT','ACTIVE')") > 0 Then Err.Raise vbObjectError + 246, , AT_U("064A,0648,062C,062F,0020,0627,0633,062A,062D,0642,0627,0642,0020,0645,0633,062C,0644,0020,0644,0647,0630,0627,0020,0627,0644,0645,0627,0644,0643,0020,0641,064A,0020,0627,0644,062A,0627,0631,064A,062E,0020,0646,0641,0633,0647,061B,0020,0631,0627,062C,0639,0020,0627,0644,0633,062C,0644,0020,0642,0628,0644,0020,0627,0644,062A,0643,0631,0627,0631,002E")
    Set ru = CurrentDb.OpenRecordset("SELECT U.UnitID,U.UnitCode,U.BuildingID,B.BuildingCode FROM ((tblUnits AS U INNER JOIN tblBuildings AS B ON U.BuildingID=B.BuildingID) INNER JOIN tblBuildingOwnership AS BO ON B.BuildingID=BO.BuildingID) WHERE BO.OwnerID=" & ownerID & " AND BO.StartDate<=#" & Format$(chargeDate, "mm\/dd\/yyyy") & "# AND (BO.EndDate Is Null OR BO.EndDate>=#" & Format$(chargeDate, "mm\/dd\/yyyy") & "#) AND U.IsActive=True AND U.IsContributionEligible=True AND B.IsActive=True ORDER BY B.BuildingCode,U.UnitCode", dbOpenSnapshot)
    If ru.EOF Then ru.Close: Err.Raise vbObjectError + 244, , AT_U("0644,0627,0020,062A,0648,062C,062F,0020,0634,0642,0642,0020,0645,0624,0647,0644,0629,0020,0645,0631,062A,0628,0637,0629,0020,0628,0627,0644,0645,0627,0644,0643,0020,062D,0627,0644,064A,064B,0627,002E")
    Set db = CurrentDb: DBEngine(0)(0).BeginTrans
    docID = AT_CreateDocument("CHARGE", chargeDate, reason)
    ownerApartments = 0
    Do While Not ru.EOF
        If Not AT_UnitIsExempt(CLng(ru!UnitID), chargeDate) Then ownerApartments = ownerApartments + 1
        ru.MoveNext
    Loop
    ru.MoveFirst
    If ownerApartments <= 0 Then Err.Raise vbObjectError + 244, , AT_U("0644,0627,0020,062A,0648,062C,062F,0020,0634,0642,0642,0020,0645,0624,0647,0644,0629,0020,0628,0639,062F,0020,0627,0633,062A,0628,0639,0627,062F,0020,0627,0644,0625,0639,0641,0627,0621,0627,062A,0020,0627,0644,0645,0639,062A,0645,062F,0629,002E")
    gross = CCur(ownerApartments) * rateAmount
    ownerCode = Nz(DLookup("OwnerCode", "tblOwners", "OwnerID=" & ownerID), "")
    ownerName = Nz(DLookup("FullName", "tblOwners", "OwnerID=" & ownerID), "")
    Set rc = db.OpenRecordset("tblOwnerCharges", dbOpenDynaset)
    rc.AddNew: rc!DocumentID = docID: rc!OwnerID = ownerID: rc!ChargeDate = chargeDate
    rc!ApartmentCountAtCharge = ownerApartments: rc!RateAmount = rateAmount: rc!GrossAmount = gross
    rc!DiscountAmount = 0: rc!NetAmount = gross: rc!OwnerCodeAtCharge = ownerCode: rc!OwnerNameAtCharge = ownerName
    rc!ChargeStatus = "DRAFT": rc!CreatedAt = Now(): rc.Update: rc.Bookmark = rc.LastModified
    chargeID = CLng(rc!ChargeID): rc.Close
    Do While Not ru.EOF
        If Not AT_UnitIsExempt(CLng(ru!UnitID), chargeDate) Then
            Set rl = db.OpenRecordset("tblChargeLines", dbOpenDynaset)
            rl.AddNew: rl!ChargeID = chargeID: rl!UnitID = CLng(ru!UnitID): rl!BuildingID = CLng(ru!BuildingID)
            rl!OwnerIDSnapshot = ownerID: rl!UnitCodeSnapshot = CStr(ru!UnitCode): rl!BuildingCodeSnapshot = CStr(ru!BuildingCode)
            rl!CampaignID = Null: rl!RateAmount = rateAmount: rl!GrossAmount = rateAmount: rl!DiscountAmount = 0: rl!NetAmount = rateAmount: rl!CreatedAt = Now()
            rl.Update: rl.Close
        End If
        ru.MoveNext
    Loop
    ru.Close
    DBEngine(0)(0).CommitTrans
    AT_Audit "CHARGE_DRAFT", "tblOwnerCharges", chargeID, "SUCCESS", "Gross=" & CStr(gross) & "; units=" & ownerApartments
    AT_CreateChargeForOwner = chargeID: Exit Function
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: If Not ru Is Nothing Then ru.Close
    DBEngine(0)(0).Rollback
    AT_Audit "CHARGE_DRAFT", "tblOwnerCharges", 0, "FAILED", eDesc
    AT_RaiseOriginal "AT_CreateChargeForOwner", savedErrNumber, eDesc
End Function

Public Sub AT_ApproveCharge(ByVal chargeID As Long, ByVal approver As String)
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "CHARGE_APPROVE"
    approver = AT_CurrentActor()
    Dim db As DAO.Database, rs As DAO.Recordset, docID As Long, codeText As String
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT C.DocumentID, C.ChargeStatus FROM tblOwnerCharges AS C WHERE C.ChargeID=" & chargeID, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 250, , AT_U("0627,0644,0627,0633,062A,062D,0642,0627,0642,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,002E")
    docID = CLng(rs!DocumentID): If rs!ChargeStatus <> "DRAFT" Then rs.Close: Err.Raise vbObjectError + 251, , AT_U("0627,0644,0627,0633,062A,062D,0642,0627,0642,0020,0644,064A,0633,0020,0645,0633,0648,062F,0629,002E")
    rs.Close
    If DCount("*", "tblChargeLines", "ChargeID=" & chargeID) = 0 Then Err.Raise vbObjectError + 253, , AT_U("062A,0641,0627,0635,064A,0644,0020,0627,0644,0634,0642,0642,0020,0644,0644,0627,0633,062A,062D,0642,0627,0642,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0629,061B,0020,0644,0627,0020,064A,0645,0643,0646,0020,0627,0644,0627,0639,062A,0645,0627,062F,002E")
    If Abs(CCur(Nz(DLookup("NetAmount", "tblOwnerCharges", "ChargeID=" & chargeID), 0)) - CCur(Nz(DSum("NetAmount", "tblChargeLines", "ChargeID=" & chargeID), 0))) > 0.01 Then Err.Raise vbObjectError + 254, , AT_U("0625,062C,0645,0627,0644,064A,0020,062A,0641,0627,0635,064A,0644,0020,0627,0644,0634,0642,0642,0020,0644,0627,0020,064A,0637,0627,0628,0642,0020,0625,062C,0645,0627,0644,064A,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,002E")
    DBEngine(0)(0).BeginTrans: AT_TransitionDraftToPosting db, docID
    codeText = AT_NextSeriesValueInTransaction(db, "CHARGE")
    db.Execute "UPDATE tblOwnerCharges SET ChargeCode='" & codeText & "', ChargeStatus='ACTIVE' WHERE ChargeID=" & chargeID & " AND ChargeStatus='DRAFT'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 252, , AT_U("062A,0639,0630,0631,0020,0627,0639,062A,0645,0627,062F,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,002E")
    AT_CompleteApproval db, docID, approver
    DBEngine(0)(0).CommitTrans
    AT_Audit "CHARGE_APPROVE", "tblOwnerCharges", chargeID, "SUCCESS", "Assigned " & codeText
    Exit Sub
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "CHARGE_APPROVE", "tblOwnerCharges", chargeID, "FAILED", eDesc
    AT_RaiseOriginal "AT_ApproveCharge", savedErrNumber, eDesc
End Sub

Public Function AT_CreateReceiptDraft(ByVal ownerID As Long, ByVal amount As Currency, ByVal paymentMethodID As Long, ByVal accountID As Long, ByVal payerName As String, ByVal collectorName As String, Optional ByVal notes As String = "", Optional ByVal transactionDate As Date = 0) As Long
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "RECEIPT_DRAFT"
    If amount <= 0 Or amount <> Fix(amount) Then Err.Raise vbObjectError + 301, , AT_U("0627,0644,0645,0628,0644,063A,0020,064A,062C,0628,0020,0623,0646,0020,064A,0643,0648,0646,0020,0639,062F,062F,064B,0627,0020,0635,062D,064A,062D,064B,0627,0020,0645,0648,062C,0628,064B,0627,0020,0628,0627,0644,0631,064A,0627,0644,0020,0627,0644,064A,0645,0646,064A,002E")
    If Not AT_Exists("tblOwners", "OwnerID=" & ownerID & " AND AdminStatus='ACTIVE'") Then Err.Raise vbObjectError + 302, , AT_U("0627,0644,0645,0627,0644,0643,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0020,0623,0648,0020,063A,064A,0631,0020,0646,0634,0637,002E")
    If Not AT_Exists("tblAccounts", "AccountID=" & accountID & " AND IsActive=True") Then Err.Raise vbObjectError + 303, , AT_U("0627,0644,062D,0633,0627,0628,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0020,0623,0648,0020,063A,064A,0631,0020,0646,0634,0637,002E")
    If Not AT_Exists("tblPaymentMethods", "PaymentMethodID=" & paymentMethodID & " AND IsActive=True") Then Err.Raise vbObjectError + 304, , AT_U("0637,0631,064A,0642,0629,0020,0627,0644,062F,0641,0639,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0629,0020,0623,0648,0020,063A,064A,0631,0020,0646,0634,0637,0629,002E")
    If Len(Trim$(collectorName)) = 0 Then collectorName = Nz(DLookup("DefaultCollectorName", "tblProjectSettings"), "")
    If Len(Trim$(collectorName)) = 0 Then Err.Raise vbObjectError + 305, , AT_U("0627,0633,0645,0020,0627,0644,0645,062D,0635,0644,0020,0645,0637,0644,0648,0628,061B,0020,0627,0636,0628,0637,0020,0627,0633,0645,0020,0627,0644,0645,062D,0635,0644,0020,0627,0644,0627,0641,062A,0631,0627,0636,064A,0020,0623,0648,0020,0623,062F,062E,0644,0647,0020,0641,064A,0020,0627,0644,0625,064A,0635,0627,0644,002E")
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
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "RECEIPT_DRAFT", "tblReceipts", 0, "FAILED", eDesc
    AT_RaiseOriginal "AT_CreateReceiptDraft", savedErrNumber, eDesc
End Function

Public Function AT_SetReceiptAllocation(ByVal receiptID As Long, ByVal chargeID As Long, ByVal amount As Currency) As Long
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "RECEIPT_DRAFT"
    If amount <= 0 Or amount <> Fix(amount) Then Err.Raise vbObjectError + 315, , AT_U("0642,064A,0645,0629,0020,0627,0644,062A,062E,0635,064A,0635,0020,064A,062C,0628,0020,0623,0646,0020,062A,0643,0648,0646,0020,0639,062F,062F,064B,0627,0020,0635,062D,064A,062D,064B,0627,0020,0645,0648,062C,0628,064B,0627,002E")
    Dim db As DAO.Database, rs As DAO.Recordset, ownerID As Long, chargeOwnerID As Long
    Dim receiptAmount As Currency, currentForReceipt As Currency, currentForCharge As Currency, due As Currency, allocationID As Long
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT R.OwnerID,R.Amount,D.StatusCode FROM tblReceipts AS R INNER JOIN tblFinancialDocuments AS D ON R.DocumentID=D.DocumentID WHERE R.ReceiptID=" & receiptID, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 316, , AT_U("0627,0644,0625,064A,0635,0627,0644,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,002E")
    If rs!StatusCode <> "DRAFT" Then rs.Close: Err.Raise vbObjectError + 317, , AT_U("0644,0627,0020,064A,0645,0643,0646,0020,062A,0639,062F,064A,0644,0020,062A,062E,0635,064A,0635,0627,062A,0020,0625,064A,0635,0627,0644,0020,063A,064A,0631,0020,0645,0633,0648,062F,0629,002E")
    ownerID = CLng(rs!OwnerID): receiptAmount = CCur(rs!Amount): rs.Close
    chargeOwnerID = CLng(Nz(DLookup("OwnerID", "tblOwnerCharges", "ChargeID=" & chargeID & " AND ChargeStatus='ACTIVE'"), 0))
    If chargeOwnerID = 0 Or chargeOwnerID <> ownerID Then Err.Raise vbObjectError + 318, , AT_U("0627,0644,0627,0633,062A,062D,0642,0627,0642,0020,063A,064A,0631,0020,0646,0634,0637,0020,0623,0648,0020,0644,0627,0020,064A,0639,0648,062F,0020,0625,0644,0649,0020,0635,0627,062D,0628,0020,0627,0644,0625,064A,0635,0627,0644,002E")
    currentForReceipt = CCur(Nz(DSum("AllocatedAmount", "tblReceiptAllocationLines", "ReceiptID=" & receiptID), 0))
    currentForCharge = CCur(Nz(DSum("AllocatedAmount", "tblReceiptAllocationLines", "ReceiptID=" & receiptID & " AND ChargeID=" & chargeID), 0))
    If currentForReceipt - currentForCharge + amount > receiptAmount Then Err.Raise vbObjectError + 319, , AT_U("0645,062C,0645,0648,0639,0020,0627,0644,062A,062E,0635,064A,0635,0627,062A,0020,064A,062A,062C,0627,0648,0632,0020,0645,0628,0644,063A,0020,0627,0644,0625,064A,0635,0627,0644,002E")
    due = AT_ChargeRemaining(chargeID) + currentForCharge
    If amount > due Then Err.Raise vbObjectError + 320, , AT_U("0627,0644,0645,0628,0644,063A,0020,064A,062A,062C,0627,0648,0632,0020,0627,0644,0645,062A,0628,0642,064A,0020,0641,064A,0020,0647,0630,0627,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,002E")
    DBEngine(0)(0).BeginTrans
    Set rs = db.OpenRecordset("SELECT * FROM tblReceiptAllocationLines WHERE ReceiptID=" & receiptID & " AND ChargeID=" & chargeID, dbOpenDynaset)
    If rs.EOF Then
        rs.AddNew: rs!ReceiptID = receiptID: rs!ChargeID = chargeID: rs!AllocatedAmount = amount
        rs!LineNo = CLng(Nz(DMax("LineNo", "tblReceiptAllocationLines", "ReceiptID=" & receiptID), 0)) + 1
    Else
        rs.Edit: rs!AllocatedAmount = amount
    End If
    rs.Update: rs.Bookmark = rs.LastModified: allocationID = CLng(rs!AllocationID): rs.Close
    DBEngine(0)(0).CommitTrans
    AT_Audit "RECEIPT_ALLOCATION_DRAFT", "tblReceiptAllocationLines", allocationID, "SUCCESS", "ReceiptID=" & receiptID & "; ChargeID=" & chargeID
    AT_SetReceiptAllocation = allocationID: Exit Function
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "RECEIPT_ALLOCATION_DRAFT", "tblReceiptAllocationLines", 0, "FAILED", eDesc
    Err.Raise savedErrNumber, "AT_SetReceiptAllocation", eDesc
End Function

Public Sub AT_ClearReceiptDraftAllocations(ByVal receiptID As Long)
    On Error GoTo EH
    AT_RequirePermission "RECEIPT_DRAFT"
    Dim documentID As Long
    documentID = CLng(Nz(DLookup("DocumentID", "tblReceipts", "ReceiptID=" & receiptID), 0))
    If documentID = 0 Or Nz(DLookup("StatusCode", "tblFinancialDocuments", "DocumentID=" & documentID), "") <> "DRAFT" Then Err.Raise vbObjectError + 327, , AT_U("064A,0645,0643,0646,0020,0645,0633,062D,0020,062A,062E,0635,064A,0635,0627,062A,0020,0627,0644,0645,0633,0648,062F,0629,0020,0641,0642,0637,002E")
    CurrentDb.Execute "DELETE FROM tblReceiptAllocationLines WHERE ReceiptID=" & receiptID, dbFailOnError
    AT_Audit "RECEIPT_ALLOCATION_CLEAR", "tblReceiptAllocationLines", receiptID, "SUCCESS", "Draft allocations cleared."
    Exit Sub
EH:
    AT_Audit "RECEIPT_ALLOCATION_CLEAR", "tblReceiptAllocationLines", receiptID, "FAILED", Err.Description
    Err.Raise Err.Number, "AT_ClearReceiptDraftAllocations", Err.Description
End Sub

Public Sub AT_ApproveReceipt(ByVal receiptID As Long, ByVal approver As String, Optional ByVal confirmedActualReceipt As Boolean = False)
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "RECEIPT_APPROVE"
    approver = AT_CurrentActor()
    If Not confirmedActualReceipt Then Err.Raise vbObjectError + 321, , AT_U("064A,062C,0628,0020,062A,0623,0643,064A,062F,0020,0627,0633,062A,0644,0627,0645,0020,0627,0644,0645,0627,0644,0020,0641,0639,0644,064A,064B,0627,0020,0642,0628,0644,0020,0627,0644,0627,0639,062A,0645,0627,062F,002E")
    Dim db As DAO.Database, rs As DAO.Recordset, docID As Long, ownerID As Long, amount As Currency, accountID As Long, docDate As Date
    Dim remaining As Currency, allocated As Currency, chargeID As Long, lineNo As Long, officialNo As String
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT R.*, D.StatusCode, D.TransactionDate FROM tblReceipts AS R INNER JOIN tblFinancialDocuments AS D ON R.DocumentID=D.DocumentID WHERE R.ReceiptID=" & receiptID, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 322, , AT_U("0627,0644,0625,064A,0635,0627,0644,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,002E")
    docID = CLng(rs!DocumentID): ownerID = CLng(rs!OwnerID): amount = CCur(rs!Amount): accountID = CLng(rs!AccountID): docDate = CDate(rs!TransactionDate)
    If rs!StatusCode <> "DRAFT" Then rs.Close: Err.Raise vbObjectError + 323, , AT_U("0627,0644,0625,064A,0635,0627,0644,0020,0644,064A,0633,0020,0645,0633,0648,062F,0629,002E")
    rs.Close
    DBEngine(0)(0).BeginTrans: AT_TransitionDraftToPosting db, docID
    If Not AT_Exists("tblAccounts", "AccountID=" & accountID & " AND IsActive=True") Then Err.Raise vbObjectError + 324, , AT_U("0627,0644,062D,0633,0627,0628,0020,063A,064A,0631,0020,0646,0634,0637,002E")
    If Not AT_Exists("tblOwners", "OwnerID=" & ownerID & " AND AdminStatus='ACTIVE'") Then Err.Raise vbObjectError + 325, , AT_U("0627,0644,0645,0627,0644,0643,0020,063A,064A,0631,0020,0646,0634,0637,002E")
    remaining = amount: lineNo = 0
    If AT_Exists("tblReceiptAllocationLines", "ReceiptID=" & receiptID) Then
        Set rs = db.OpenRecordset("SELECT A.AllocationID,A.ChargeID,A.AllocatedAmount,C.OwnerID,C.ChargeStatus FROM tblReceiptAllocationLines AS A INNER JOIN tblOwnerCharges AS C ON A.ChargeID=C.ChargeID WHERE A.ReceiptID=" & receiptID & " ORDER BY A.LineNo", dbOpenSnapshot)
        Do While Not rs.EOF
            chargeID = CLng(rs!ChargeID): allocated = CCur(rs!AllocatedAmount)
            If CLng(rs!OwnerID) <> ownerID Or rs!ChargeStatus <> "ACTIVE" Then Err.Raise vbObjectError + 326, , AT_U("064A,0648,062C,062F,0020,062A,062E,0635,064A,0635,0020,064A,0634,064A,0631,0020,0625,0644,0649,0020,0627,0633,062A,062D,0642,0627,0642,0020,063A,064A,0631,0020,0635,0627,0644,062D,0020,0623,0648,0020,0644,0645,0627,0644,0643,0020,0645,062E,062A,0644,0641,002E")
            If allocated <= 0 Or allocated > AT_ChargeRemaining(chargeID) Then Err.Raise vbObjectError + 328, , AT_U("062A,062C,0627,0648,0632,0020,0623,062D,062F,0020,0627,0644,062A,062E,0635,064A,0635,0627,062A,0020,0627,0644,0631,0635,064A,062F,0020,0627,0644,0645,062A,0628,0642,064A,0020,0639,0646,062F,0020,0627,0644,0627,0639,062A,0645,0627,062F,002E")
            If allocated > remaining Then Err.Raise vbObjectError + 329, , AT_U("0645,062C,0645,0648,0639,0020,0627,0644,062A,062E,0635,064A,0635,0627,062A,0020,0627,0644,064A,062F,0648,064A,0629,0020,064A,062A,062C,0627,0648,0632,0020,0645,0628,0644,063A,0020,0627,0644,0625,064A,0635,0627,0644,002E")
            lineNo = lineNo + 1
            Dim rSettleManual As DAO.Recordset
            Set rSettleManual = db.OpenRecordset("tblChargeSettlementLines", dbOpenDynaset)
            rSettleManual.AddNew: rSettleManual!ChargeID = chargeID: rSettleManual!DocumentID = docID: rSettleManual!ReceiptID = receiptID
            rSettleManual!SignedAmount = allocated: rSettleManual!SettlementType = "RECEIPT": rSettleManual!CreatedAt = Now(): rSettleManual.Update: rSettleManual.Close
            remaining = remaining - allocated: rs.MoveNext
        Loop
        rs.Close
    Else
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
    End If
    If remaining > 0 Then
        Dim rCredit As DAO.Recordset
        Set rCredit = db.OpenRecordset("tblOwnerCreditMovements", dbOpenDynaset)
        rCredit.AddNew: rCredit!OwnerID = ownerID: rCredit!DocumentID = docID: rCredit!SignedAmount = remaining
        rCredit!CreditType = "RECEIPT_EXCESS": rCredit!SourceReceiptID = receiptID: rCredit!Reason = AT_U("0632,064A,0627,062F,0629,0020,0642,0628,0636,0020,0639,0644,0649,0020,0627,0644,0645,0633,062A,062D,0642,0627,062A"): rCredit!CreatedAt = Now(): rCredit.Update: rCredit.Close
    End If
    AT_InsertAccountMovement db, accountID, docID, docDate, amount, "RECEIPT", 1
    officialNo = AT_AssignOfficialNumberInTransaction(db, "RECEIPT", docID, approver)
    AT_CompleteApproval db, docID, approver
    DBEngine(0)(0).CommitTrans
    AT_Audit "RECEIPT_APPROVE", "tblReceipts", receiptID, "SUCCESS", "Official=" & officialNo & "; excess credit=" & CStr(remaining)
    Exit Sub
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "RECEIPT_APPROVE", "tblReceipts", receiptID, "FAILED", eDesc
    AT_RaiseOriginal "AT_ApproveReceipt", savedErrNumber, eDesc
End Sub

Public Function AT_CreatePaymentDraft(ByVal accountID As Long, ByVal expenseCategoryID As Long, ByVal amount As Currency, ByVal paymentMethodID As Long, ByVal beneficiaryName As String, Optional ByVal recipientName As String = "", Optional ByVal notes As String = "", Optional ByVal transactionDate As Date = 0, Optional ByVal refundCreditMovementID As Long = 0, Optional ByVal refundDocumentID As Long = 0) As Long
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "PAYMENT_DRAFT"
    If refundCreditMovementID <= 0 Then Err.Raise vbObjectError + 339, , AT_U("0633,062C,0644,0020,0627,0644,0645,0635,0631,0648,0641,0020,0623,0648,0644,064B,0627,060C,0020,062B,0645,0020,0623,0646,0634,0626,0020,0633,0646,062F,0020,0627,0644,0635,0631,0641,0020,0645,0646,0020,0634,0627,0634,0629,0020,0627,0644,0645,0635,0631,0648,0641,0627,062A,002E,0020,0627,0644,0635,0631,0641,0020,0627,0644,0645,0628,0627,0634,0631,0020,0645,062A,0627,062D,0020,0644,0631,062F,0020,0627,0644,0631,0635,064A,062F,0020,0627,0644,062F,0627,0626,0646,0020,0641,0642,0637,002E")
    If amount <= 0 Or amount <> Fix(amount) Then Err.Raise vbObjectError + 330, , AT_U("0627,0644,0645,0628,0644,063A,0020,064A,062C,0628,0020,0623,0646,0020,064A,0643,0648,0646,0020,0639,062F,062F,064B,0627,0020,0635,062D,064A,062D,064B,0627,0020,0645,0648,062C,0628,064B,0627,0020,0628,0627,0644,0631,064A,0627,0644,0020,0627,0644,064A,0645,0646,064A,002E")
    If Len(Trim$(beneficiaryName)) = 0 Then Err.Raise vbObjectError + 331, , AT_U("0627,0633,0645,0020,0627,0644,0645,0633,062A,0641,064A,062F,0020,0645,0637,0644,0648,0628,002E")
    If Not AT_Exists("tblAccounts", "AccountID=" & accountID & " AND IsActive=True") Then Err.Raise vbObjectError + 332, , AT_U("0627,0644,062D,0633,0627,0628,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0020,0623,0648,0020,063A,064A,0631,0020,0646,0634,0637,002E")
    If Not AT_Exists("tblExpenseCategories", "ExpenseCategoryID=" & expenseCategoryID & " AND IsActive=True") Then Err.Raise vbObjectError + 333, , AT_U("0641,0626,0629,0020,0627,0644,0645,0635,0631,0648,0641,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0629,0020,0623,0648,0020,063A,064A,0631,0020,0646,0634,0637,0629,002E")
    If Not AT_Exists("tblPaymentMethods", "PaymentMethodID=" & paymentMethodID & " AND IsActive=True") Then Err.Raise vbObjectError + 334, , AT_U("0637,0631,064A,0642,0629,0020,0627,0644,062F,0641,0639,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0629,0020,0623,0648,0020,063A,064A,0631,0020,0646,0634,0637,0629,002E")
    If refundCreditMovementID > 0 Then
        If AT_CreditMovementAvailable(refundCreditMovementID) < amount Then Err.Raise vbObjectError + 335, , AT_U("0631,0635,064A,062F,0020,0627,0644,0645,0627,0644,0643,0020,0627,0644,062F,0627,0626,0646,0020,0627,0644,0645,062A,0627,062D,0020,063A,064A,0631,0020,0643,0627,0641,064D,0020,0644,0647,0630,0627,0020,0627,0644,0631,062F,002E")
        If refundDocumentID <= 0 Then refundDocumentID = CLng(Nz(DLookup("DocumentID", "tblOwnerCreditMovements", "CreditMovementID=" & refundCreditMovementID), 0))
        If refundDocumentID <= 0 Then Err.Raise vbObjectError + 336, , AT_U("062A,0639,0630,0631,0020,0631,0628,0637,0020,0627,0644,0631,062F,0020,0628,0645,0633,062A,0646,062F,0020,0645,0635,062F,0631,0020,0627,0644,0631,0635,064A,062F,0020,0627,0644,062F,0627,0626,0646,002E")
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
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "PAYMENT_DRAFT", "tblPayments", 0, "FAILED", eDesc
    AT_RaiseOriginal "AT_CreatePaymentDraft", savedErrNumber, eDesc
End Function

Public Sub AT_ApprovePayment(ByVal paymentID As Long, ByVal approver As String, Optional ByVal confirmedPaid As Boolean = False)
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "PAYMENT_APPROVE"
    approver = AT_CurrentActor()
    If Not confirmedPaid Then Err.Raise vbObjectError + 341, , AT_U("064A,062C,0628,0020,062A,0623,0643,064A,062F,0020,0623,0646,0020,0627,0644,0645,0628,0644,063A,0020,062F,064F,0641,0639,0020,0641,0639,0644,064A,064B,0627,0020,0642,0628,0644,0020,0627,0644,0627,0639,062A,0645,0627,062F,002E")
    Dim db As DAO.Database, rs As DAO.Recordset, docID As Long, accountID As Long, amount As Currency, docDate As Date, creditMovementID As Long, expenseID As Long, officialNo As String
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT P.*, D.StatusCode, D.TransactionDate FROM tblPayments AS P INNER JOIN tblFinancialDocuments AS D ON P.DocumentID=D.DocumentID WHERE P.PaymentID=" & paymentID, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 342, , AT_U("0633,0646,062F,0020,0627,0644,0635,0631,0641,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,002E")
    docID = CLng(rs!DocumentID): accountID = CLng(rs!AccountID): amount = CCur(rs!Amount): docDate = CDate(rs!TransactionDate)
    creditMovementID = CLng(Nz(rs!RefundOfCreditMovementID, 0)): expenseID = CLng(Nz(rs!ExpenseID, 0))
    If expenseID = 0 And creditMovementID = 0 Then rs.Close: Err.Raise vbObjectError + 351, , AT_U("064A,062C,0628,0020,0631,0628,0637,0020,0633,0646,062F,0020,0627,0644,0635,0631,0641,0020,0628,0645,0635,0631,0648,0641,0020,0645,0633,062C,0644,060C,0020,0648,0644,0627,0020,064A,0633,0645,062D,0020,0628,0627,0644,0635,0631,0641,0020,0627,0644,0645,0628,0627,0634,0631,0020,0625,0644,0627,0020,0644,0631,062F,0020,0631,0635,064A,062F,0020,062F,0627,0626,0646,0020,0645,0648,062B,0642,002E")
    If rs!StatusCode <> "DRAFT" Then rs.Close: Err.Raise vbObjectError + 343, , AT_U("0633,0646,062F,0020,0627,0644,0635,0631,0641,0020,0644,064A,0633,0020,0645,0633,0648,062F,0629,002E")
    rs.Close
    DBEngine(0)(0).BeginTrans: AT_TransitionDraftToPosting db, docID
    If amount <= 0 Or amount <> Fix(amount) Then Err.Raise vbObjectError + 344, , AT_U("0627,0644,0645,0628,0644,063A,0020,063A,064A,0631,0020,0635,0627,0644,062D,002E")
    If expenseID > 0 Then
        If Not AT_Exists("tblExpenses", "ExpenseID=" & expenseID & " AND StatusCode='PAYMENT_DRAFT' AND Amount=" & Replace(CStr(amount), ",", ".")) Then Err.Raise vbObjectError + 349, , AT_U("062D,0627,0644,0629,0020,0627,0644,0645,0635,0631,0648,0641,0020,0623,0648,0020,0642,064A,0645,062A,0647,0020,0644,0627,0020,062A,0637,0627,0628,0642,0020,0633,0646,062F,0020,0627,0644,0635,0631,0641,0020,0627,0644,0645,0631,062A,0628,0637,002E")
    End If
    If Not AT_Exists("tblAccounts", "AccountID=" & accountID & " AND IsActive=True") Then Err.Raise vbObjectError + 348, , AT_U("0627,0644,062D,0633,0627,0628,0020,063A,064A,0631,0020,0646,0634,0637,002E")
    If AT_AccountAvailableBalance(accountID) < amount Then Err.Raise vbObjectError + 345, , AT_U("0627,0644,0631,0635,064A,062F,0020,0627,0644,0645,062A,0627,062D,0020,063A,064A,0631,0020,0643,0627,0641,064D,0020,0628,0639,062F,0020,0627,062D,062A,0633,0627,0628,0020,0627,0644,062D,062C,0648,0632,0627,062A,061B,0020,0644,0627,0020,064A,0633,0645,062D,0020,0627,0644,0646,0638,0627,0645,0020,0628,0627,0633,062A,062B,0646,0627,0621,002E")
    If creditMovementID > 0 Then
        If AT_CreditMovementAvailable(creditMovementID) < amount Then Err.Raise vbObjectError + 346, , AT_U("0627,0644,0631,0635,064A,062F,0020,0627,0644,062F,0627,0626,0646,0020,0627,0644,0645,062A,0627,062D,0020,063A,064A,0631,0020,0643,0627,0641,064D,0020,0644,0644,0631,062F,0020,0623,0648,0020,0627,0633,062A,064F,062E,062F,0645,0020,0641,064A,0020,0639,0645,0644,064A,0629,0020,0623,062E,0631,0649,002E")
        Dim srcOwnerID As Long
        srcOwnerID = CLng(Nz(DLookup("OwnerID", "tblOwnerCreditMovements", "CreditMovementID=" & creditMovementID), 0))
        If srcOwnerID = 0 Then Err.Raise vbObjectError + 347, , AT_U("062D,0631,0643,0629,0020,0627,0644,0631,0635,064A,062F,0020,0627,0644,062F,0627,0626,0646,0020,0627,0644,0645,0635,062F,0631,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0629,002E")
        Dim cr As DAO.Recordset
        Set cr = db.OpenRecordset("tblOwnerCreditMovements", dbOpenDynaset)
        cr.AddNew: cr!OwnerID = srcOwnerID: cr!DocumentID = docID: cr!SignedAmount = -amount
        cr!CreditType = "CASH_REFUND": cr!SourceCreditMovementID = creditMovementID: cr!Reason = AT_U("0631,062F,0020,0646,0642,062F,064A,0020,0645,0639,062A,0645,062F,0020,0639,0628,0631,0020,0633,0646,062F,0020,0635,0631,0641,0020") & paymentID: cr!CreatedAt = Now(): cr.Update: cr.Close
    End If
    AT_InsertAccountMovement db, accountID, docID, docDate, -amount, IIf(creditMovementID > 0, "CREDIT_REFUND", "PAYMENT"), 1
    officialNo = AT_AssignOfficialNumberInTransaction(db, "PAYMENT", docID, approver)
    db.Execute "UPDATE tblPayments SET VoucherNo='" & Replace(officialNo, "'", "''") & "' WHERE PaymentID=" & paymentID, dbFailOnError
    If expenseID > 0 Then db.Execute "UPDATE tblExpenses SET StatusCode='PAID', UpdatedAt=Now() WHERE ExpenseID=" & expenseID & " AND StatusCode='PAYMENT_DRAFT'", dbFailOnError
    If expenseID > 0 And db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 350, , AT_U("062A,0639,0630,0631,0020,062A,062B,0628,064A,062A,0020,062D,0627,0644,0629,0020,0627,0644,0645,0635,0631,0648,0641,0020,0645,062F,0641,0648,0639,064B,0627,002E")
    AT_CompleteApproval db, docID, approver
    DBEngine(0)(0).CommitTrans
    AT_Audit "PAYMENT_APPROVE", "tblPayments", paymentID, "SUCCESS", "Official=" & officialNo
    Exit Sub
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "PAYMENT_APPROVE", "tblPayments", paymentID, "FAILED", eDesc
    AT_RaiseOriginal "AT_ApprovePayment", savedErrNumber, eDesc
End Sub

Public Function AT_CreateTransferDraft(ByVal sourceAccountID As Long, ByVal destinationAccountID As Long, ByVal amount As Currency, Optional ByVal feeAmount As Currency = 0, Optional ByVal notes As String = "", Optional ByVal transactionDate As Date = 0) As Long
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "TRANSFER_DRAFT"
    If sourceAccountID = destinationAccountID Then Err.Raise vbObjectError + 360, , AT_U("0644,0627,0020,064A,0645,0643,0646,0020,062A,062D,0648,064A,0644,0020,0627,0644,0645,0628,0644,063A,0020,0625,0644,0649,0020,0627,0644,062D,0633,0627,0628,0020,0646,0641,0633,0647,002E")
    If amount <= 0 Or amount <> Fix(amount) Then Err.Raise vbObjectError + 361, , AT_U("0645,0628,0644,063A,0020,0627,0644,062A,062D,0648,064A,0644,0020,064A,062C,0628,0020,0623,0646,0020,064A,0643,0648,0646,0020,0639,062F,062F,064B,0627,0020,0635,062D,064A,062D,064B,0627,0020,0645,0648,062C,0628,064B,0627,002E")
    If feeAmount < 0 Or feeAmount <> Fix(feeAmount) Then Err.Raise vbObjectError + 362, , AT_U("0631,0633,0648,0645,0020,0627,0644,062A,062D,0648,064A,0644,0020,063A,064A,0631,0020,0635,0627,0644,062D,0629,002E")
    If Not AT_Exists("tblAccounts", "AccountID=" & sourceAccountID & " AND IsActive=True") Or Not AT_Exists("tblAccounts", "AccountID=" & destinationAccountID & " AND IsActive=True") Then Err.Raise vbObjectError + 363, , AT_U("062D,0633,0627,0628,0020,0627,0644,0645,0635,062F,0631,0020,0623,0648,0020,0627,0644,0648,062C,0647,0629,0020,063A,064A,0631,0020,0646,0634,0637,002E")
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
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "TRANSFER_DRAFT", "tblAccountTransfers", 0, "FAILED", eDesc
    AT_RaiseOriginal "AT_CreateTransferDraft", savedErrNumber, eDesc
End Function

Public Sub AT_ApproveTransfer(ByVal transferID As Long, ByVal approver As String)
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "TRANSFER_APPROVE"
    approver = AT_CurrentActor()
    Dim db As DAO.Database, rs As DAO.Recordset, docID As Long, sourceID As Long, destID As Long, totalReserved As Currency, officialNo As String, r As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT T.*, D.StatusCode FROM tblAccountTransfers AS T INNER JOIN tblFinancialDocuments AS D ON T.DocumentID=D.DocumentID WHERE T.TransferID=" & transferID, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 370, , AT_U("0627,0644,062A,062D,0648,064A,0644,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,002E")
    docID = CLng(rs!DocumentID): sourceID = CLng(rs!SourceAccountID): destID = CLng(rs!DestinationAccountID): totalReserved = CCur(rs!Amount) + CCur(Nz(rs!FeeAmount, 0))
    If rs!StatusCode <> "DRAFT" Then rs.Close: Err.Raise vbObjectError + 371, , AT_U("0627,0644,062A,062D,0648,064A,0644,0020,0644,064A,0633,0020,0645,0633,0648,062F,0629,002E")
    rs.Close
    DBEngine(0)(0).BeginTrans: AT_TransitionDraftToPosting db, docID
    If Not AT_Exists("tblAccounts", "AccountID=" & sourceID & " AND IsActive=True") Or Not AT_Exists("tblAccounts", "AccountID=" & destID & " AND IsActive=True") Then Err.Raise vbObjectError + 374, , AT_U("062D,0633,0627,0628,0020,0627,0644,0645,0635,062F,0631,0020,0623,0648,0020,0627,0644,0648,062C,0647,0629,0020,063A,064A,0631,0020,0646,0634,0637,0020,0639,0646,062F,0020,0627,0644,0627,0639,062A,0645,0627,062F,002E")
    If AT_AccountAvailableBalance(sourceID) < totalReserved Then Err.Raise vbObjectError + 372, , AT_U("0627,0644,0631,0635,064A,062F,0020,0627,0644,0645,062A,0627,062D,0020,0644,0627,0020,064A,063A,0637,064A,0020,0645,0628,0644,063A,0020,0627,0644,062A,062D,0648,064A,0644,0020,0648,0631,0633,0648,0645,0647,002E")
    Set r = db.OpenRecordset("tblAccountReservations", dbOpenDynaset)
    r.AddNew: r!DocumentID = docID: r!AccountID = sourceID: r!Amount = totalReserved: r!StatusCode = "ACTIVE": r!CreatedAt = Now(): r!Reason = AT_U("062D,062C,0632,0020,0645,0628,0644,063A,0020,062A,062D,0648,064A,0644,0020,062D,062A,0649,0020,062A,0623,0643,064A,062F,0020,0648,0635,0648,0644,0647"): r.Update: r.Close
    officialNo = AT_AssignOfficialNumberInTransaction(db, "TRANSFER", docID, approver)
    db.Execute "UPDATE tblFinancialDocuments SET StatusCode='PENDING', ApprovedAt=Now(), ApprovedByUserName='" & Replace(approver, "'", "''") & "' WHERE DocumentID=" & docID & " AND StatusCode='POSTING'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 373, , AT_U("062A,0639,0630,0631,0020,062A,062B,0628,064A,062A,0020,0627,0644,062A,062D,0648,064A,0644,0020,0627,0644,0645,0639,0644,0642,002E")
    DBEngine(0)(0).CommitTrans
    AT_Audit "TRANSFER_PENDING", "tblAccountTransfers", transferID, "SUCCESS", "Official=" & officialNo & "; reserved=" & CStr(totalReserved)
    Exit Sub
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "TRANSFER_PENDING", "tblAccountTransfers", transferID, "FAILED", eDesc
    AT_RaiseOriginal "AT_ApproveTransfer", savedErrNumber, eDesc
End Sub

Public Sub AT_ConfirmTransfer(ByVal transferID As Long, ByVal approver As String, ByVal confirmationReference As String)
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "TRANSFER_APPROVE"
    approver = AT_CurrentActor()
    If Len(Trim$(confirmationReference)) = 0 Then Err.Raise vbObjectError + 380, , AT_U("0645,0631,062C,0639,0020,062A,0623,0643,064A,062F,0020,0627,0644,062A,062D,0648,064A,0644,0020,0645,0637,0644,0648,0628,002E")
    Dim db As DAO.Database, rs As DAO.Recordset, docID As Long, sourceID As Long, destID As Long, amount As Currency, fee As Currency, reservationID As Long, reserveAmount As Currency
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT T.*, D.StatusCode FROM tblAccountTransfers AS T INNER JOIN tblFinancialDocuments AS D ON T.DocumentID=D.DocumentID WHERE T.TransferID=" & transferID, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 381, , AT_U("0627,0644,062A,062D,0648,064A,0644,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,002E")
    docID = CLng(rs!DocumentID): sourceID = CLng(rs!SourceAccountID): destID = CLng(rs!DestinationAccountID): amount = CCur(rs!Amount): fee = CCur(Nz(rs!FeeAmount, 0))
    If rs!StatusCode <> "PENDING" Then rs.Close: Err.Raise vbObjectError + 382, , AT_U("0627,0644,062A,062D,0648,064A,0644,0020,0644,064A,0633,0020,0645,0639,0644,0642,064B,0627,0020,0628,0627,0646,062A,0638,0627,0631,0020,0627,0644,062A,0623,0643,064A,062F,002E")
    rs.Close
    DBEngine(0)(0).BeginTrans
    Set rs = db.OpenRecordset("SELECT ReservationID, Amount FROM tblAccountReservations WHERE DocumentID=" & docID & " AND StatusCode='ACTIVE'", dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 383, , AT_U("062D,062C,0632,0020,0627,0644,062A,062D,0648,064A,0644,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,061B,0020,0623,0648,0642,0641,0020,0627,0644,062A,0623,0643,064A,062F,0020,0644,0644,062A,062D,0642,064A,0642,002E")
    reservationID = CLng(rs!ReservationID): reserveAmount = CCur(rs!Amount): rs.Close
    If reserveAmount <> amount + fee Then Err.Raise vbObjectError + 387, , AT_U("0642,064A,0645,0629,0020,0627,0644,062D,062C,0632,0020,0644,0627,0020,062A,0637,0627,0628,0642,0020,0645,0628,0644,063A,0020,0627,0644,062A,062D,0648,064A,0644,0020,0648,0627,0644,0631,0633,0648,0645,061B,0020,0623,0648,0642,0641,0020,0627,0644,0639,0645,0644,064A,0629,0020,0644,0644,062A,062D,0642,064A,0642,002E")
    If AT_AccountAvailableBalance(sourceID, reservationID) < amount + fee Then Err.Raise vbObjectError + 384, , AT_U("0627,0644,0631,0635,064A,062F,0020,0627,0644,0641,0639,0644,064A,0020,0628,0639,062F,0020,0627,0633,062A,0628,0639,0627,062F,0020,0627,0644,062D,062C,0632,0020,063A,064A,0631,0020,0643,0627,0641,064D,061B,0020,0644,0645,0020,064A,064F,0631,062D,0651,0644,0020,0627,0644,062A,062D,0648,064A,0644,002E")
    db.Execute "UPDATE tblFinancialDocuments SET StatusCode='POSTING' WHERE DocumentID=" & docID & " AND StatusCode='PENDING'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 385, , AT_U("062A,0645,0020,062A,0623,0643,064A,062F,0020,0627,0644,062A,062D,0648,064A,0644,0020,0645,0633,0628,0642,064B,0627,0020,0623,0648,0020,062A,063A,064A,0631,062A,0020,062D,0627,0644,062A,0647,002E")
    AT_InsertAccountMovement db, sourceID, docID, CDate(DLookup("TransactionDate", "tblFinancialDocuments", "DocumentID=" & docID)), -amount, "TRANSFER_OUT", 1
    If fee > 0 Then AT_InsertAccountMovement db, sourceID, docID, CDate(DLookup("TransactionDate", "tblFinancialDocuments", "DocumentID=" & docID)), -fee, "TRANSFER_FEE", 2
    AT_InsertAccountMovement db, destID, docID, CDate(DLookup("TransactionDate", "tblFinancialDocuments", "DocumentID=" & docID)), amount, "TRANSFER_IN", 3
    db.Execute "UPDATE tblAccountReservations SET StatusCode='RELEASED', ReleasedAt=Now() WHERE ReservationID=" & reservationID & " AND StatusCode='ACTIVE'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 386, , AT_U("062A,0639,0630,0631,0020,062A,062D,0631,064A,0631,0020,062D,062C,0632,0020,0627,0644,062A,062D,0648,064A,0644,002E")
    db.Execute "UPDATE tblAccountTransfers SET ConfirmedAt=Now(), ConfirmedByUserName='" & Replace(approver, "'", "''") & "', ConfirmationRef='" & Replace(confirmationReference, "'", "''") & "' WHERE TransferID=" & transferID, dbFailOnError
    AT_CompleteApproval db, docID, approver
    DBEngine(0)(0).CommitTrans
    AT_Audit "TRANSFER_CONFIRMED", "tblAccountTransfers", transferID, "SUCCESS", "Transfer and fee posted separately; reference=" & confirmationReference
    Exit Sub
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "TRANSFER_CONFIRMED", "tblAccountTransfers", transferID, "FAILED", eDesc
    AT_RaiseOriginal "AT_ConfirmTransfer", savedErrNumber, eDesc
End Sub

Public Sub AT_CancelPendingTransfer(ByVal transferID As Long, ByVal reason As String, ByVal actor As String)
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "DOCUMENT_CANCEL"
    actor = AT_CurrentActor()
    If Len(Trim$(reason)) = 0 Then Err.Raise vbObjectError + 390, , AT_U("0633,0628,0628,0020,0627,0644,0625,0644,063A,0627,0621,0020,0645,0637,0644,0648,0628,002E")
    Dim db As DAO.Database, rs As DAO.Recordset, docID As Long
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT T.DocumentID, D.StatusCode FROM tblAccountTransfers AS T INNER JOIN tblFinancialDocuments AS D ON T.DocumentID=D.DocumentID WHERE T.TransferID=" & transferID, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 391, , AT_U("0627,0644,062A,062D,0648,064A,0644,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,002E")
    docID = CLng(rs!DocumentID): If rs!StatusCode <> "PENDING" Then rs.Close: Err.Raise vbObjectError + 392, , AT_U("064A,0633,0645,062D,0020,0628,0625,0644,063A,0627,0621,0020,0627,0644,062A,062D,0648,064A,0644,0020,0627,0644,0645,0639,0644,0642,0020,0641,0642,0637,002E")
    rs.Close
    DBEngine(0)(0).BeginTrans
    db.Execute "UPDATE tblFinancialDocuments SET StatusCode='CANCELLED', CancelReason='" & Replace(reason, "'", "''") & "' WHERE DocumentID=" & docID & " AND StatusCode='PENDING'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 393, , AT_U("062A,063A,064A,0631,062A,0020,062D,0627,0644,0629,0020,0627,0644,062A,062D,0648,064A,0644,061B,0020,0644,0645,0020,064A,064F,0644,063A,064E,002E")
    db.Execute "UPDATE tblAccountReservations SET StatusCode='RELEASED', ReleasedAt=Now(), Reason='" & Replace(reason, "'", "''") & "' WHERE DocumentID=" & docID & " AND StatusCode='ACTIVE'", dbFailOnError
    DBEngine(0)(0).CommitTrans
    AT_Audit "TRANSFER_CANCEL", "tblAccountTransfers", transferID, "SUCCESS", "Reservation released", reason
    Exit Sub
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "TRANSFER_CANCEL", "tblAccountTransfers", transferID, "FAILED", eDesc, reason
    AT_RaiseOriginal "AT_CancelPendingTransfer", savedErrNumber, eDesc
End Sub

Public Function AT_RequestChargeAdjustment(ByVal chargeID As Long, ByVal signedAmount As Currency, ByVal reason As String, ByVal requester As String, Optional ByVal adjustmentType As String = "DISCOUNT") As Long
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "ADJUSTMENT_REQUEST"
    requester = AT_CurrentActor()
    If signedAmount >= 0 Or signedAmount <> Fix(signedAmount) Then Err.Raise vbObjectError + 400, , AT_U("064A,062C,0628,0020,0623,0646,0020,064A,0643,0648,0646,0020,0627,0644,062E,0635,0645,0020,0645,0628,0644,063A,064B,0627,0020,0635,062D,064A,062D,064B,0627,0020,0633,0627,0644,0628,0020,0627,0644,0625,0634,0627,0631,0629,002E")
    If Len(Trim$(reason)) = 0 Then Err.Raise vbObjectError + 401, , AT_U("0633,0628,0628,0020,0637,0644,0628,0020,0627,0644,062E,0635,0645,0020,0645,0637,0644,0648,0628,002E")
    If AT_ChargeRemaining(chargeID) + signedAmount < 0 Then Err.Raise vbObjectError + 402, , AT_U("0627,0644,062E,0635,0645,0020,0623,0643,0628,0631,0020,0645,0646,0020,0627,0644,0645,062A,0628,0642,064A,061B,0020,064A,0644,0632,0645,0020,0645,0633,0627,0631,0020,062A,0635,062D,064A,062D,002F,0631,062F,0020,0645,0646,0641,0635,0644,002E")
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
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "ADJUSTMENT_REQUEST", "tblOwnerAdjustments", 0, "FAILED", eDesc, reason
    AT_RaiseOriginal "AT_RequestChargeAdjustment", savedErrNumber, eDesc
End Function

Public Sub AT_ApproveChargeAdjustment(ByVal adjustmentID As Long, ByVal approver As String)
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "ADJUSTMENT_APPROVE"
    approver = AT_CurrentActor()
    Dim db As DAO.Database, rs As DAO.Recordset, docID As Long, chargeID As Long, signedAmount As Currency
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT A.*, D.StatusCode FROM tblOwnerAdjustments AS A INNER JOIN tblFinancialDocuments AS D ON A.DocumentID=D.DocumentID WHERE A.AdjustmentID=" & adjustmentID, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 410, , AT_U("0637,0644,0628,0020,0627,0644,062E,0635,0645,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,002E")
    docID = CLng(rs!DocumentID): chargeID = CLng(rs!ChargeID): signedAmount = CCur(rs!SignedAmount)
    If rs!StatusCode <> "DRAFT" Then rs.Close: Err.Raise vbObjectError + 411, , AT_U("0637,0644,0628,0020,0627,0644,062E,0635,0645,0020,0644,064A,0633,0020,0645,0633,0648,062F,0629,002E")
    rs.Close
    DBEngine(0)(0).BeginTrans: AT_TransitionDraftToPosting db, docID
    If signedAmount >= 0 Or AT_ChargeRemaining(chargeID) + signedAmount < 0 Then Err.Raise vbObjectError + 412, , AT_U("0627,0644,062E,0635,0645,0020,063A,064A,0631,0020,0635,0627,0644,062D,0020,0623,0648,0020,064A,062E,0641,0636,0020,0627,0644,0645,0633,062A,062D,0642,0020,0625,0644,0649,0020,0623,0642,0644,0020,0645,0646,0020,0627,0644,0645,0633,062F,062F,002E")
    db.Execute "UPDATE tblOwnerAdjustments SET ApprovedByUserName='" & Replace(approver, "'", "''") & "', ApprovedAt=Now() WHERE AdjustmentID=" & adjustmentID, dbFailOnError
    AT_CompleteApproval db, docID, approver
    DBEngine(0)(0).CommitTrans
    AT_Audit "ADJUSTMENT_APPROVE", "tblOwnerAdjustments", adjustmentID, "SUCCESS", "Approved signed amount=" & CStr(signedAmount)
    Exit Sub
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "ADJUSTMENT_APPROVE", "tblOwnerAdjustments", adjustmentID, "FAILED", eDesc
    AT_RaiseOriginal "AT_ApproveChargeAdjustment", savedErrNumber, eDesc
End Sub

Public Sub AT_ApplyOwnerCreditToCharge(ByVal ownerID As Long, ByVal creditMovementID As Long, ByVal chargeID As Long, ByVal amount As Currency, ByVal approver As String, ByVal reason As String)
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "CREDIT_APPLY"
    approver = AT_CurrentActor()
    If amount <= 0 Or amount <> Fix(amount) Then Err.Raise vbObjectError + 420, , AT_U("0645,0628,0644,063A,0020,0627,0644,062A,0637,0628,064A,0642,0020,064A,062C,0628,0020,0623,0646,0020,064A,0643,0648,0646,0020,0639,062F,062F,064B,0627,0020,0635,062D,064A,062D,064B,0627,0020,0645,0648,062C,0628,064B,0627,002E")
    If Len(Trim$(reason)) = 0 Then Err.Raise vbObjectError + 421, , AT_U("0633,0628,0628,0020,062A,0637,0628,064A,0642,0020,0627,0644,0631,0635,064A,062F,0020,0645,0637,0644,0648,0628,002E")
    Dim db As DAO.Database, rs As DAO.Recordset, sourceOwnerID As Long, chargeOwnerID As Long, docID As Long, appID As Long
    sourceOwnerID = CLng(Nz(DLookup("OwnerID", "tblOwnerCreditMovements", "CreditMovementID=" & creditMovementID), 0))
    chargeOwnerID = CLng(Nz(DLookup("OwnerID", "tblOwnerCharges", "ChargeID=" & chargeID), 0))
    If sourceOwnerID <> ownerID Or chargeOwnerID <> ownerID Then Err.Raise vbObjectError + 422, , AT_U("064A,062C,0628,0020,0623,0646,0020,064A,0639,0648,062F,0020,0627,0644,0631,0635,064A,062F,0020,0648,0627,0644,0627,0633,062A,062D,0642,0627,0642,0020,0625,0644,0649,0020,0627,0644,0645,0627,0644,0643,0020,0646,0641,0633,0647,002E")
    If AT_CreditMovementAvailable(creditMovementID) < amount Then Err.Raise vbObjectError + 423, , AT_U("0627,0644,0631,0635,064A,062F,0020,0627,0644,062F,0627,0626,0646,0020,0627,0644,0645,062A,0627,062D,0020,063A,064A,0631,0020,0643,0627,0641,064D,002E")
    If AT_ChargeRemaining(chargeID) < amount Then Err.Raise vbObjectError + 424, , AT_U("0627,0644,0645,062A,0628,0642,064A,0020,0645,0646,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,0020,0623,0642,0644,0020,0645,0646,0020,0645,0628,0644,063A,0020,0627,0644,062A,0637,0628,064A,0642,002E")
    Set db = CurrentDb: DBEngine(0)(0).BeginTrans
    ' Recheck both balances inside the transaction before making any ledger entries.
    If AT_CreditMovementAvailable(creditMovementID) < amount Then Err.Raise vbObjectError + 425, , AT_U("0627,0644,0631,0635,064A,062F,0020,0627,0644,062F,0627,0626,0646,0020,062A,063A,064A,0631,0020,0648,0644,0645,0020,064A,0639,062F,0020,0643,0627,0641,064A,064B,0627,002E")
    If AT_ChargeRemaining(chargeID) < amount Then Err.Raise vbObjectError + 426, , AT_U("0627,0644,0645,062A,0628,0642,064A,0020,0645,0646,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,0020,062A,063A,064A,0631,0020,0648,0644,0645,0020,064A,0639,062F,0020,0643,0627,0641,064A,064B,0627,002E")
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
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "CREDIT_APPLY", "tblCreditApplications", 0, "FAILED", eDesc, reason
    AT_RaiseOriginal "AT_ApplyOwnerCreditToCharge", savedErrNumber, eDesc
End Sub

Public Sub AT_ReverseReceiptAllocation(ByVal allocationID As Long, ByVal amount As Currency, ByVal approver As String, ByVal reason As String)
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "ADJUSTMENT_APPROVE"
    approver = AT_CurrentActor()
    If amount <= 0 Or amount <> Fix(amount) Then Err.Raise vbObjectError + 430, , AT_U("0645,0628,0644,063A,0020,0627,0644,0639,0643,0633,0020,064A,062C,0628,0020,0623,0646,0020,064A,0643,0648,0646,0020,0639,062F,062F,064B,0627,0020,0635,062D,064A,062D,064B,0627,0020,0645,0648,062C,0628,064B,0627,002E")
    If Len(Trim$(reason)) = 0 Then Err.Raise vbObjectError + 431, , AT_U("0633,0628,0628,0020,0627,0644,0639,0643,0633,0020,0645,0637,0644,0648,0628,002E")
    Dim db As DAO.Database, rs As DAO.Recordset, receiptID As Long, chargeID As Long, ownerID As Long, originalAmount As Currency, alreadyReversed As Currency, docID As Long, correctionID As Long
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT A.AllocationID, A.ReceiptID, A.ChargeID, A.AllocatedAmount, R.OwnerID, D.StatusCode FROM (tblReceiptAllocationLines AS A INNER JOIN tblReceipts AS R ON A.ReceiptID=R.ReceiptID) INNER JOIN tblFinancialDocuments AS D ON R.DocumentID=D.DocumentID WHERE A.AllocationID=" & allocationID, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 432, , AT_U("0633,0637,0631,0020,0627,0644,062A,062E,0635,064A,0635,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,002E")
    receiptID = CLng(rs!ReceiptID): chargeID = CLng(rs!ChargeID): ownerID = CLng(rs!OwnerID): originalAmount = CCur(rs!AllocatedAmount)
    If rs!StatusCode <> "APPROVED" Then rs.Close: Err.Raise vbObjectError + 433, , AT_U("0644,0627,0020,064A,0645,0643,0646,0020,0639,0643,0633,0020,062A,062E,0635,064A,0635,0020,0625,064A,0635,0627,0644,0020,063A,064A,0631,0020,0645,0639,062A,0645,062F,002E")
    rs.Close
    Dim sql As String
    sql = "SELECT Sum(L.Amount) AS ReversedAmount FROM tblReceiptCorrectionLines AS L INNER JOIN tblReceiptCorrections AS C ON L.ReceiptCorrectionID=C.ReceiptCorrectionID INNER JOIN tblFinancialDocuments AS D ON C.DocumentID=D.DocumentID WHERE L.OriginalAllocationID=" & allocationID & " AND D.StatusCode='APPROVED'"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot): alreadyReversed = AT_GetCurrency(rs, "ReversedAmount"): rs.Close
    If amount > originalAmount - alreadyReversed Then Err.Raise vbObjectError + 434, , AT_U("0645,0628,0644,063A,0020,0627,0644,0639,0643,0633,0020,064A,062A,062C,0627,0648,0632,0020,0627,0644,062C,0632,0621,0020,063A,064A,0631,0020,0627,0644,0645,0639,0643,0648,0633,0020,0645,0646,0020,0627,0644,062A,062E,0635,064A,0635,0020,0627,0644,0623,0635,0644,064A,002E")
    DBEngine(0)(0).BeginTrans
    sql = "SELECT Sum(L.Amount) AS ReversedAmount FROM tblReceiptCorrectionLines AS L INNER JOIN tblReceiptCorrections AS C ON L.ReceiptCorrectionID=C.ReceiptCorrectionID INNER JOIN tblFinancialDocuments AS D ON C.DocumentID=D.DocumentID WHERE L.OriginalAllocationID=" & allocationID & " AND D.StatusCode='APPROVED'"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot): alreadyReversed = AT_GetCurrency(rs, "ReversedAmount"): rs.Close
    If amount > originalAmount - alreadyReversed Then Err.Raise vbObjectError + 435, , AT_U("0645,0628,0644,063A,0020,0627,0644,0639,0643,0633,0020,062A,062C,0627,0648,0632,0020,0627,0644,062C,0632,0621,0020,0627,0644,0645,062A,0628,0642,064A,0020,0623,062B,0646,0627,0621,0020,0628,062F,0621,0020,0627,0644,0639,0645,0644,064A,0629,061B,0020,0644,0645,0020,064A,064F,0631,062D,0651,0644,002E")
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
    rs!Reason = AT_U("0631,0635,064A,062F,0020,062F,0627,0626,0646,0020,0646,0627,062A,062C,0020,0639,0646,0020,0639,0643,0633,0020,062A,062E,0635,064A,0635,0020,0627,0644,0625,064A,0635,0627,0644,003A,0020") & reason: rs!CreatedAt = Now(): rs.Update: rs.Close
    AT_CompleteApproval db, docID, approver
    DBEngine(0)(0).CommitTrans
    AT_Audit "RECEIPT_ALLOCATION_REVERSE", "tblReceiptCorrectionLines", correctionID, "SUCCESS", "Reversed=" & CStr(amount), reason
    Exit Sub
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "RECEIPT_ALLOCATION_REVERSE", "tblReceiptCorrectionLines", allocationID, "FAILED", eDesc, reason
    AT_RaiseOriginal "AT_ReverseReceiptAllocation", savedErrNumber, eDesc
End Sub

Public Sub AT_CancelDraftDocument(ByVal documentID As Long, ByVal reason As String)
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "DOCUMENT_CANCEL"
    If Len(Trim$(reason)) = 0 Then Err.Raise vbObjectError + 440, , AT_U("0633,0628,0628,0020,0627,0644,0625,0644,063A,0627,0621,0020,0645,0637,0644,0648,0628,002E")
    Dim db As DAO.Database, expenseID As Long, receiptID As Long
    Set db = CurrentDb: DBEngine(0)(0).BeginTrans
    expenseID = CLng(Nz(DLookup("ExpenseID", "tblPayments", "DocumentID=" & documentID), 0))
    receiptID = CLng(Nz(DLookup("ReceiptID", "tblReceipts", "DocumentID=" & documentID), 0))
    db.Execute "UPDATE tblFinancialDocuments SET StatusCode='CANCELLED', CancelReason='" & Replace(reason, "'", "''") & "' WHERE DocumentID=" & documentID & " AND StatusCode='DRAFT'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 441, , AT_U("064A,0633,0645,062D,0020,0628,0627,0644,0625,0644,063A,0627,0621,0020,0644,0644,0645,0633,0648,062F,0627,062A,0020,0641,0642,0637,061B,0020,0644,0627,0020,064A,0645,0643,0646,0020,062D,0630,0641,0020,0623,0648,0020,0625,0644,063A,0627,0621,0020,0645,0633,062A,0646,062F,0020,0645,0631,062D,0651,0644,0020,0628,0647,0630,0647,0020,0627,0644,0637,0631,064A,0642,0629,002E")
    If expenseID > 0 Then db.Execute "UPDATE tblExpenses SET StatusCode='DRAFT', UpdatedAt=Now() WHERE ExpenseID=" & expenseID & " AND StatusCode='PAYMENT_DRAFT'", dbFailOnError
    If receiptID > 0 Then db.Execute "DELETE FROM tblReceiptAllocationLines WHERE ReceiptID=" & receiptID, dbFailOnError
    DBEngine(0)(0).CommitTrans
    AT_Audit "DOCUMENT_CANCEL", "tblFinancialDocuments", documentID, "SUCCESS", "Draft cancelled", reason
    Exit Sub
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "DOCUMENT_CANCEL", "tblFinancialDocuments", documentID, "FAILED", eDesc, reason
    AT_RaiseOriginal "AT_CancelDraftDocument", savedErrNumber, eDesc
End Sub

Public Sub AT_RunIntegrityReview()
    Dim savedErrNumber As Long, eDesc As String
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
        If balance < 0 Then AT_AddIntegrityFinding db, runID, "ERROR", "tblAccounts", accountID, "NEGATIVE_ACCOUNT_BALANCE", AT_U("0627,0644,0631,0635,064A,062F,0020,0627,0644,062F,0641,062A,0631,064A,0020,0644,0644,062D,0633,0627,0628,0020,0633,0627,0644,0628,003A,0020") & CStr(balance)
        If AT_AccountReservedAmount(accountID) > balance Then AT_AddIntegrityFinding db, runID, "ERROR", "tblAccounts", accountID, "RESERVATIONS_EXCEED_BALANCE", AT_U("0625,062C,0645,0627,0644,064A,0020,0627,0644,062D,062C,0648,0632,0627,062A,0020,0627,0644,0646,0634,0637,0629,0020,064A,062A,062C,0627,0648,0632,0020,0627,0644,0631,0635,064A,062F,0020,0627,0644,062F,0641,062A,0631,064A,002E")
        rs.MoveNext
    Loop
    rs.Close
    Set rs = db.OpenRecordset("SELECT ChargeID FROM tblOwnerCharges WHERE ChargeStatus='ACTIVE'", dbOpenSnapshot)
    Do While Not rs.EOF
        chargeID = CLng(rs!ChargeID): remaining = AT_ChargeRemaining(chargeID)
        If remaining < 0 Then AT_AddIntegrityFinding db, runID, "ERROR", "tblOwnerCharges", chargeID, "NEGATIVE_REMAINING", AT_U("0627,0644,0645,062A,0628,0642,064A,0020,0645,0646,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,0020,0633,0627,0644,0628,061B,0020,0631,0627,062C,0639,0020,0627,0644,062A,0633,0648,064A,0627,062A,0020,0648,0627,0644,062E,0635,0648,0645,0627,062A,002E")
        rs.MoveNext
    Loop
    rs.Close
    Set rs = db.OpenRecordset("SELECT CreditMovementID FROM tblOwnerCreditMovements WHERE SignedAmount>0 AND SourceCreditMovementID Is Null", dbOpenSnapshot)
    Do While Not rs.EOF
        If AT_CreditMovementRawAvailable(CLng(rs!CreditMovementID)) < 0 Then AT_AddIntegrityFinding db, runID, "ERROR", "tblOwnerCreditMovements", CLng(rs!CreditMovementID), "NEGATIVE_CREDIT_AVAILABLE", AT_U("0627,0644,0631,0635,064A,062F,0020,0627,0644,0645,062A,0627,062D,0020,0645,0646,0020,062D,0631,0643,0629,0020,0627,0644,062F,0627,0626,0646,0020,0623,0635,0628,062D,0020,0633,0627,0644,0628,064B,0627,002E")
        rs.MoveNext
    Loop
    rs.Close
    db.Execute "UPDATE tblIntegrityRuns SET CompletedAt=Now(), ResultSummary='Integrity checks completed; findings require human review. No auto-correction performed.' WHERE IntegrityRunID=" & runID, dbFailOnError
    AT_Audit "INTEGRITY_REVIEW", "tblIntegrityRuns", runID, "SUCCESS", "Flags generated without auto-correction."
    MsgBox AT_U("0627,0643,062A,0645,0644,0020,0641,062D,0635,0020,0627,0644,0633,0644,0627,0645,0629,002E,0020,0631,0627,062C,0639,0020,0074,0062,006C,0049,006E,0074,0065,0067,0072,0069,0074,0079,0046,0069,006E,0064,0069,006E,0067,0073,061B,0020,0644,0645,0020,062A,064F,0639,062F,0651,0644,0020,0623,064A,0020,0628,064A,0627,0646,0627,062A,0020,0645,0627,0644,064A,0629,0020,062A,0644,0642,0627,0626,064A,064B,0627,002E"), vbInformation, "Archon Tech"
    Exit Sub
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    AT_Audit "INTEGRITY_REVIEW", "tblIntegrityRuns", 0, "FAILED", eDesc
    MsgBox AT_U("062A,0639,0630,0631,0020,0625,0643,0645,0627,0644,0020,0641,062D,0635,0020,0627,0644,0633,0644,0627,0645,0629,003A,0020") & savedErrNumber & " - " & eDesc, vbExclamation, "Archon Tech"
End Sub

Private Sub AT_AddIntegrityFinding(ByVal db As DAO.Database, ByVal runID As Long, ByVal severity As String, ByVal entityName As String, ByVal entityID As Long, ByVal findingCode As String, ByVal findingText As String)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("tblIntegrityFindings", dbOpenDynaset)
    rs.AddNew: rs!IntegrityRunID = runID: rs!Severity = severity: rs!EntityName = entityName
    rs!EntityID = entityID: rs!FindingCode = findingCode: rs!FindingText = findingText: rs!ResolutionStatus = "OPEN"
    rs.Update: rs.Close
End Sub
