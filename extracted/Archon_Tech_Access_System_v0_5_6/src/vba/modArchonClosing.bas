Attribute VB_Name = "modArchonClosing"
Option Compare Database
Option Explicit

Public Sub AT_EnsureDateOpen(ByVal transactionDate As Date)
    If transactionDate = 0 Then Exit Sub
    If DCount("*", "tblFinancialPeriods", "StatusCode='CLOSED' AND PeriodStart<=#" & Format$(DateValue(transactionDate), "mm\/dd\/yyyy") & "# AND PeriodEnd>=#" & Format$(DateValue(transactionDate), "mm\/dd\/yyyy") & "#") > 0 Then _
        Err.Raise vbObjectError + 1300, , AT_U("0627,0644,062A,0627,0631,064A,062E,0020,064A,0642,0639,0020,062F,0627,062E,0644,0020,0641,062A,0631,0629,0020,0645,0627,0644,064A,0629,0020,0645,063A,0644,0642,0629,061B,0020,0644,0627,0020,064A,0645,0643,0646,0020,062A,0633,062C,064A,0644,0020,062D,0631,0643,0629,0020,062C,062F,064A,062F,0629,0020,0641,064A,0647,0627,002E")
End Sub

Private Function AT_IsSuperAdmin() As Boolean
    Dim userID As Long, adminRoleID As Long
    userID = CLng(Nz(DLookup("UserID", "tblUsers", "UserName='" & Replace(Environ$("Username"), "'", "''") & "' AND IsActive=True"), 0))
    adminRoleID = CLng(Nz(DLookup("RoleID", "tblRoles", "RoleCode='ADMIN' AND IsActive=True"), 0))
    If userID > 0 And adminRoleID > 0 Then AT_IsSuperAdmin = (DCount("*", "tblUserRoleAssignments", "UserID=" & userID & " AND RoleID=" & adminRoleID & " AND IsActive=True") > 0)
End Function

Public Function AT_CreateFinancialPeriod(ByVal periodCode As String, ByVal periodName As String, ByVal startDate As Date, ByVal endDate As Date) As Long
    On Error GoTo EH
    AT_RequirePermission "FINANCIAL_CLOSE"
    periodCode = Trim$(periodCode): periodName = Trim$(periodName)
    If Len(periodCode) = 0 Or Len(periodName) = 0 Or startDate = 0 Or endDate = 0 Then Err.Raise vbObjectError + 1310, , AT_U("0631,0645,0632,0020,0627,0644,0641,062A,0631,0629,0020,0648,0627,0633,0645,0647,0627,0020,0648,062A,0627,0631,064A,062E,0627,0020,0627,0644,0628,062F,0627,064A,0629,0020,0648,0627,0644,0646,0647,0627,064A,0629,0020,0645,0637,0644,0648,0628,0629,002E")
    If DateValue(endDate) < DateValue(startDate) Then Err.Raise vbObjectError + 1311, , AT_U("062A,0627,0631,064A,062E,0020,0646,0647,0627,064A,0629,0020,0627,0644,0641,062A,0631,0629,0020,064A,0633,0628,0642,0020,0628,062F,0627,064A,062A,0647,0627,002E")
    If DCount("*", "tblFinancialPeriods", "PeriodCode='" & Replace(periodCode, "'", "''") & "'") > 0 Then Err.Raise vbObjectError + 1312, , AT_U("0631,0645,0632,0020,0627,0644,0641,062A,0631,0629,0020,0645,0633,062A,062E,062F,0645,0020,0645,0633,0628,0642,064B,0627,002E")
    If DCount("*", "tblFinancialPeriods", "StatusCode<>'CANCELLED' AND PeriodStart<=#" & Format$(endDate, "mm\/dd\/yyyy") & "# AND PeriodEnd>=#" & Format$(startDate, "mm\/dd\/yyyy") & "#") > 0 Then Err.Raise vbObjectError + 1313, , AT_U("062A,062A,062F,0627,062E,0644,0020,0627,0644,0641,062A,0631,0629,0020,0645,0639,0020,0641,062A,0631,0629,0020,0645,0627,0644,064A,0629,0020,0645,0648,062C,0648,062F,0629,002E")
    Dim rs As DAO.Recordset, periodID As Long
    Set rs = CurrentDb.OpenRecordset("tblFinancialPeriods", dbOpenDynaset)
    rs.AddNew: rs!PeriodCode = periodCode: rs!PeriodName = periodName: rs!PeriodStart = DateValue(startDate)
    rs!PeriodEnd = DateValue(endDate): rs!StatusCode = "OPEN": rs!CreatedAt = Now(): rs!CreatedByUserName = Environ$("Username")
    rs.Update: rs.Bookmark = rs.LastModified: periodID = CLng(rs!PeriodID): rs.Close
    AT_Audit "PERIOD_CREATE", "tblFinancialPeriods", periodID, "SUCCESS", periodCode
    AT_CreateFinancialPeriod = periodID: Exit Function
EH:
    AT_Audit "PERIOD_CREATE", "tblFinancialPeriods", 0, "FAILED", Err.Description
    Err.Raise Err.Number, "AT_CreateFinancialPeriod", Err.Description
End Function

Public Function AT_BeginFinancialClosing(ByVal periodID As Long) As Long
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "FINANCIAL_CLOSE"
    Dim rs As DAO.Recordset, closingID As Long, periodStart As Date, periodEnd As Date, db As DAO.Database
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tblFinancialPeriods WHERE PeriodID=" & periodID, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 1320, , AT_U("0627,0644,0641,062A,0631,0629,0020,0627,0644,0645,0627,0644,064A,0629,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0629,002E")
    If rs!StatusCode <> "OPEN" Then rs.Close: Err.Raise vbObjectError + 1321, , AT_U("0627,0644,0641,062A,0631,0629,0020,0644,064A,0633,062A,0020,0645,0641,062A,0648,062D,0629,0020,0644,0644,0625,0642,0641,0627,0644,002E")
    periodStart = CDate(rs!PeriodStart): periodEnd = CDate(rs!PeriodEnd): rs.Close
    If DCount("*", "tblFinancialClosings", "PeriodID=" & periodID & " AND StatusCode='CLOSED'") > 0 Then Err.Raise vbObjectError + 1322, , AT_U("0627,0644,0641,062A,0631,0629,0020,0644,062F,064A,0647,0627,0020,0625,0642,0641,0627,0644,0020,0633,0627,0628,0642,061B,0020,0644,0627,0020,062A,0646,0634,0626,0020,0625,0642,0641,0627,0644,064B,0627,0020,062B,0627,0646,064A,064B,0627,002E")
    DBEngine(0)(0).BeginTrans
    Set rs = db.OpenRecordset("tblFinancialClosings", dbOpenDynaset)
    rs.AddNew: rs!PeriodID = periodID: rs!PeriodStart = periodStart: rs!PeriodEnd = periodEnd: rs!StatusCode = "COUNTING": rs!Summary = "Awaiting actual balances for all active accounts.": rs.Update: rs.Bookmark = rs.LastModified: closingID = CLng(rs!ClosingID): rs.Close
    Set rs = db.OpenRecordset("SELECT AccountID FROM tblAccounts WHERE IsActive=True", dbOpenSnapshot)
    Do While Not rs.EOF
        db.Execute "INSERT INTO tblFinancialClosingAccounts (ClosingID,AccountID,BookBalance,Notes) VALUES (" & closingID & "," & CLng(rs!AccountID) & "," & CStr(AT_AccountBookBalance(CLng(rs!AccountID), DateValue(periodEnd))) & ",'Actual count required')", dbFailOnError
        rs.MoveNext
    Loop
    rs.Close: DBEngine(0)(0).CommitTrans
    AT_Audit "CLOSING_BEGIN", "tblFinancialClosings", closingID, "SUCCESS", "PeriodID=" & periodID
    AT_BeginFinancialClosing = closingID: Exit Function
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "CLOSING_BEGIN", "tblFinancialClosings", 0, "FAILED", eDesc
    Err.Raise savedErrNumber, "AT_BeginFinancialClosing", eDesc
End Function

Public Sub AT_SaveClosingActual(ByVal closingID As Long, ByVal accountID As Long, ByVal actualBalance As Currency, Optional ByVal notes As String = "")
    On Error GoTo EH
    AT_RequirePermission "FINANCIAL_CLOSE"
    If DCount("*", "tblFinancialClosings", "ClosingID=" & closingID & " AND StatusCode='COUNTING'") = 0 Then Err.Raise vbObjectError + 1330, , AT_U("0627,0644,0625,0642,0641,0627,0644,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0020,0623,0648,0020,0644,0645,0020,064A,0639,062F,0020,0641,064A,0020,0645,0631,062D,0644,0629,0020,0627,0644,0639,062F,002E")
    If DCount("*", "tblFinancialClosingAccounts", "ClosingID=" & closingID & " AND AccountID=" & accountID) = 0 Then Err.Raise vbObjectError + 1331, , AT_U("0627,0644,062D,0633,0627,0628,0020,0644,0627,0020,064A,0646,062A,0645,064A,0020,0625,0644,0649,0020,062A,0641,0627,0635,064A,0644,0020,0647,0630,0627,0020,0627,0644,0625,0642,0641,0627,0644,002E")
    CurrentDb.Execute "UPDATE tblFinancialClosingAccounts SET ActualBalance=" & Replace(CStr(actualBalance), ",", ".") & ", Notes='" & Replace(notes, "'", "''") & "' WHERE ClosingID=" & closingID & " AND AccountID=" & accountID, dbFailOnError
    AT_Audit "CLOSING_ACTUAL", "tblFinancialClosingAccounts", closingID, "SUCCESS", "AccountID=" & accountID & "; actual=" & CStr(actualBalance)
    Exit Sub
EH:
    AT_Audit "CLOSING_ACTUAL", "tblFinancialClosings", closingID, "FAILED", Err.Description
    Err.Raise Err.Number, "AT_SaveClosingActual", Err.Description
End Sub

Public Sub AT_FinalizeFinancialClosing(ByVal closingID As Long, Optional ByVal discrepancyExceptionReason As String = "")
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "FINANCIAL_CLOSE"
    Dim db As DAO.Database, rs As DAO.Recordset, periodID As Long, periodEnd As Date
    Dim diffCount As Long, missingCount As Long, pendingCount As Long, summaryText As String
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT C.PeriodID,C.PeriodEnd,C.StatusCode FROM tblFinancialClosings AS C WHERE C.ClosingID=" & closingID, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 1340, , AT_U("0633,062C,0644,0020,0627,0644,0625,0642,0641,0627,0644,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,002E")
    periodID = CLng(rs!PeriodID): periodEnd = CDate(rs!PeriodEnd)
    If rs!StatusCode <> "COUNTING" Then rs.Close: Err.Raise vbObjectError + 1341, , AT_U("0627,0644,0625,0642,0641,0627,0644,0020,0644,064A,0633,0020,0641,064A,0020,062D,0627,0644,0629,0020,0627,0644,0639,062F,002E")
    rs.Close
    missingCount = DCount("*", "tblFinancialClosingAccounts", "ClosingID=" & closingID & " AND ActualBalance Is Null")
    If missingCount > 0 Then Err.Raise vbObjectError + 1342, , AT_U("064A,062C,0628,0020,0625,062F,062E,0627,0644,0020,0627,0644,0631,0635,064A,062F,0020,0627,0644,0641,0639,0644,064A,0020,0644,0643,0644,0020,062D,0633,0627,0628,0020,0646,0634,0637,0020,0642,0628,0644,0020,0627,0644,0625,0642,0641,0627,0644,002E")
    db.Execute "UPDATE tblFinancialClosingAccounts SET DifferenceAmount=ActualBalance-BookBalance WHERE ClosingID=" & closingID, dbFailOnError
    diffCount = DCount("*", "tblFinancialClosingAccounts", "ClosingID=" & closingID & " AND Abs(DifferenceAmount)>0.01")
    pendingCount = DCount("*", "tblAccountReservations", "StatusCode='ACTIVE'")
    If pendingCount > 0 Then Err.Raise vbObjectError + 1343, , AT_U("062A,0648,062C,062F,0020,062A,062D,0648,064A,0644,0627,062A,0020,0623,0648,0020,062D,062C,0648,0632,0627,062A,0020,0645,0627,0644,064A,0629,0020,0645,0639,0644,0642,0629,061B,0020,0644,0627,0020,064A,0645,0643,0646,0020,0625,0642,0641,0627,0644,0020,0627,0644,0641,062A,0631,0629,0020,0642,0628,0644,0020,062A,0633,0648,064A,062A,0647,0627,0020,0623,0648,0020,062A,0648,062B,064A,0642,0647,0627,0020,062E,0627,0631,062C,0020,0647,0630,0647,0020,0627,0644,0634,0627,0634,0629,002E")
    If diffCount > 0 Then
        If Len(Trim$(discrepancyExceptionReason)) = 0 Or Not AT_IsSuperAdmin() Then Err.Raise vbObjectError + 1344, , AT_U("062A,0648,062C,062F,0020,0641,0631,0648,0642,0627,062A,0020,0641,0639,0644,064A,0629,002E,0020,064A,0645,0646,0639,0020,0627,0644,0625,0642,0641,0627,0644,0020,062D,062A,0649,0020,062A,0633,0648,064A,062A,0647,0627,061B,0020,0627,0644,0627,0633,062A,062B,0646,0627,0621,0020,0645,062A,0627,062D,0020,0644,0644,0645,062F,064A,0631,0020,0627,0644,0623,0639,0644,0649,0020,0641,0642,0637,0020,0645,0639,0020,0633,0628,0628,0020,0645,0648,062B,0642,002E")
    End If
    summaryText = "Account differences=" & diffCount & "; pending reservations=" & pendingCount
    DBEngine(0)(0).BeginTrans
    db.Execute "UPDATE tblFinancialClosings SET StatusCode='CLOSED', ClosedAt=Now(), ClosedByUserName='" & Replace(Environ$("Username"), "'", "''") & "', DiscrepancyException=" & IIf(diffCount > 0, "True", "False") & ", ExceptionReason='" & Replace(discrepancyExceptionReason, "'", "''") & "', Summary='" & Replace(summaryText, "'", "''") & "' WHERE ClosingID=" & closingID & " AND StatusCode='COUNTING'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 1345, , AT_U("062A,063A,064A,0631,062A,0020,062D,0627,0644,0629,0020,0627,0644,0625,0642,0641,0627,0644,0020,0642,0628,0644,0020,0627,0644,062A,062B,0628,064A,062A,002E")
    db.Execute "UPDATE tblFinancialPeriods SET StatusCode='CLOSED', ClosedAt=Now(), ClosedByUserName='" & Replace(Environ$("Username"), "'", "''") & "' WHERE PeriodID=" & periodID & " AND StatusCode='OPEN'", dbFailOnError
    If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 1346, , AT_U("062A,063A,064A,0631,062A,0020,062D,0627,0644,0629,0020,0627,0644,0641,062A,0631,0629,0020,0627,0644,0645,0627,0644,064A,0629,002E")
    DBEngine(0)(0).CommitTrans
    AT_Audit "CLOSING_FINALIZE", "tblFinancialClosings", closingID, "SUCCESS", summaryText, discrepancyExceptionReason
    Exit Sub
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "CLOSING_FINALIZE", "tblFinancialClosings", closingID, "FAILED", eDesc, discrepancyExceptionReason
    Err.Raise savedErrNumber, "AT_FinalizeFinancialClosing", eDesc
End Sub

Public Sub AT_ReopenFinancialPeriod(ByVal periodID As Long, ByVal reason As String)
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "FINANCIAL_REOPEN"
    If Not AT_IsSuperAdmin() Then Err.Raise vbObjectError + 1350, , AT_U("0625,0639,0627,062F,0629,0020,0641,062A,062D,0020,0627,0644,0641,062A,0631,0629,0020,0627,0644,0645,0627,0644,064A,0629,0020,0645,062D,0635,0648,0631,0629,0020,0628,0627,0644,0645,062F,064A,0631,0020,0627,0644,0623,0639,0644,0649,002E")
    If Len(Trim$(reason)) = 0 Then Err.Raise vbObjectError + 1351, , AT_U("0633,0628,0628,0020,0625,0639,0627,062F,0629,0020,0627,0644,0641,062A,062D,0020,0625,0644,0632,0627,0645,064A,002E")
    If DCount("*", "tblFinancialPeriods", "PeriodID=" & periodID & " AND StatusCode='CLOSED'") = 0 Then Err.Raise vbObjectError + 1352, , AT_U("0627,0644,0641,062A,0631,0629,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0629,0020,0623,0648,0020,0644,064A,0633,062A,0020,0645,063A,0644,0642,0629,002E")
    DBEngine(0)(0).BeginTrans
    CurrentDb.Execute "UPDATE tblFinancialPeriods SET StatusCode='OPEN', ReopenedAt=Now(), ReopenedByUserName='" & Replace(Environ$("Username"), "'", "''") & "', ReopenReason='" & Replace(reason, "'", "''") & "' WHERE PeriodID=" & periodID, dbFailOnError
    CurrentDb.Execute "UPDATE tblFinancialClosings SET StatusCode='REOPENED', ReopenedAt=Now(), ReopenedByUserName='" & Replace(Environ$("Username"), "'", "''") & "', ReopenReason='" & Replace(reason, "'", "''") & "' WHERE PeriodID=" & periodID & " AND StatusCode='CLOSED'", dbFailOnError
    DBEngine(0)(0).CommitTrans
    AT_Audit "PERIOD_REOPEN", "tblFinancialPeriods", periodID, "SUCCESS", "Period reopened.", reason
    Exit Sub
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "PERIOD_REOPEN", "tblFinancialPeriods", periodID, "FAILED", eDesc, reason
    Err.Raise savedErrNumber, "AT_ReopenFinancialPeriod", eDesc
End Sub
