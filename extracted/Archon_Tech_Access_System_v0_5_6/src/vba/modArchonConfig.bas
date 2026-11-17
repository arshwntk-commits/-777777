Attribute VB_Name = "modArchonConfig"
Option Compare Database
Option Explicit

' Central project/system settings and effective-dated contribution rates.

Public Sub AT_UpdateProjectSettings(ByVal projectName As String, ByVal neighborhoodName As String, _
                                    ByVal defaultCollectorName As String, Optional ByVal contactText As String = "")
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "SETTINGS_ADMIN"
    projectName = Trim$(projectName): neighborhoodName = Trim$(neighborhoodName)
    If Len(projectName) = 0 Or Len(neighborhoodName) = 0 Then Err.Raise vbObjectError + 1500, , AT_U("0627,0633,0645,0020,0627,0644,0645,0634,0631,0648,0639,0020,0648,0627,0633,0645,0020,0627,0644,062D,0627,0631,0629,0020,0645,0637,0644,0648,0628,0627,0646,002E")
    Dim db As DAO.Database, rs As DAO.Recordset, settingID As Long
    Set db = CurrentDb: DBEngine(0)(0).BeginTrans
    Set rs = db.OpenRecordset("SELECT * FROM tblProjectSettings ORDER BY ProjectSettingID", dbOpenDynaset)
    If rs.EOF Then
        rs.AddNew
    Else
        rs.MoveFirst: rs.Edit
    End If
    rs!ProjectName = projectName: rs!NeighborhoodName = neighborhoodName
    rs!DefaultCollectorName = Trim$(defaultCollectorName): rs!UpdatedAt = Now()
    rs.Update: rs.Bookmark = rs.LastModified: settingID = CLng(rs!ProjectSettingID): rs.Close
    Set rs = db.OpenRecordset("SELECT * FROM tblSystemSettings WHERE SettingKey='PROJECT_CONTACT'", dbOpenDynaset)
    If rs.EOF Then rs.AddNew Else rs.Edit
    rs!SettingKey = "PROJECT_CONTACT": rs!SettingValue = contactText
    rs!UpdatedAt = Now(): rs!UpdatedByUserName = Environ$("Username"): rs.Update: rs.Close
    DBEngine(0)(0).CommitTrans
    AT_Audit "PROJECT_SETTINGS", "tblProjectSettings", settingID, "SUCCESS", "Project and neighborhood display settings updated.", contactText
    Exit Sub
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "PROJECT_SETTINGS", "tblProjectSettings", 0, "FAILED", eDesc
    Err.Raise savedErrNumber, "AT_UpdateProjectSettings", eDesc
End Sub

Public Function AT_CreateContributionRate(ByVal rateCode As String, ByVal rateAmount As Currency, ByVal effectiveFrom As Date) As Long
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "SETTINGS_ADMIN"
    rateCode = UCase$(Trim$(rateCode))
    If Len(rateCode) = 0 Or effectiveFrom = 0 Then Err.Raise vbObjectError + 1510, , AT_U("0631,0645,0632,0020,0627,0644,0633,0639,0631,0020,0648,062A,0627,0631,064A,062E,0020,0628,062F,0627,064A,0629,0020,0627,0644,0633,0631,064A,0627,0646,0020,0645,0637,0644,0648,0628,0627,0646,002E")
    If rateAmount <= 0 Or rateAmount <> Fix(rateAmount) Then Err.Raise vbObjectError + 1511, , AT_U("0642,064A,0645,0629,0020,0627,0644,0645,0633,0627,0647,0645,0629,0020,064A,062C,0628,0020,0623,0646,0020,062A,0643,0648,0646,0020,0639,062F,062F,064B,0627,0020,0635,062D,064A,062D,064B,0627,0020,0645,0648,062C,0628,064B,0627,0020,0628,0627,0644,0631,064A,0627,0644,0020,0627,0644,064A,0645,0646,064A,002E")
    If DCount("*", "tblContributionRates", "RateCode='" & Replace(rateCode, "'", "''") & "'") > 0 Then Err.Raise vbObjectError + 1512, , AT_U("0631,0645,0632,0020,0627,0644,0633,0639,0631,0020,0645,0633,062A,062E,062F,0645,0020,0645,0633,0628,0642,064B,0627,002E")
    If DCount("*", "tblContributionRates", "EffectiveFrom>=#" & Format$(DateValue(effectiveFrom), "mm\/dd\/yyyy") & "#") > 0 Then Err.Raise vbObjectError + 1513, , AT_U("062A,0627,0631,064A,062E,0020,0627,0644,0633,0631,064A,0627,0646,0020,064A,062C,0628,0020,0623,0646,0020,064A,0643,0648,0646,0020,0628,0639,062F,0020,0622,062E,0631,0020,0633,0639,0631,0020,0645,0633,062C,0644,061B,0020,0644,0627,0020,064A,064F,0633,0645,062D,0020,0628,0625,0639,0627,062F,0629,0020,0643,062A,0627,0628,0629,0020,0627,0644,062A,0627,0631,064A,062E,0020,0627,0644,0633,0627,0628,0642,002E")

    Dim db As DAO.Database, rs As DAO.Recordset, rateID As Long, priorID As Long, priorStart As Date
    Set db = CurrentDb: DBEngine(0)(0).BeginTrans
    Set rs = db.OpenRecordset("SELECT TOP 1 RateID,EffectiveFrom,EffectiveTo FROM tblContributionRates ORDER BY EffectiveFrom DESC", dbOpenSnapshot)
    If Not rs.EOF Then
        priorID = CLng(rs!RateID): priorStart = CDate(rs!EffectiveFrom)
        If IsNull(rs!EffectiveTo) Then
            If DateValue(effectiveFrom) <= DateValue(priorStart) Then rs.Close: Err.Raise vbObjectError + 1514, , AT_U("0627,0644,0633,0639,0631,0020,0627,0644,062C,062F,064A,062F,0020,064A,062C,0628,0020,0623,0646,0020,064A,0628,062F,0623,0020,0628,0639,062F,0020,0627,0644,0633,0639,0631,0020,0627,0644,0633,0627,0628,0642,002E")
            db.Execute "UPDATE tblContributionRates SET EffectiveTo=#" & Format$(DateAdd("d", -1, DateValue(effectiveFrom)), "mm\/dd\/yyyy") & "# WHERE RateID=" & priorID & " AND EffectiveTo Is Null", dbFailOnError
        End If
    End If
    rs.Close
    Set rs = db.OpenRecordset("tblContributionRates", dbOpenDynaset)
    rs.AddNew: rs!RateCode = rateCode: rs!RateAmount = rateAmount: rs!EffectiveFrom = DateValue(effectiveFrom): rs!IsActive = True: rs!CreatedAt = Now()
    rs.Update: rs.Bookmark = rs.LastModified: rateID = CLng(rs!RateID): rs.Close
    db.Execute "UPDATE tblProjectSettings SET DefaultRate=" & Replace(CStr(rateAmount), ",", ".") & ",UpdatedAt=Now()", dbFailOnError
    DBEngine(0)(0).CommitTrans
    AT_Audit "RATE_CREATE", "tblContributionRates", rateID, "SUCCESS", "Code=" & rateCode & "; Amount=" & CStr(rateAmount) & "; EffectiveFrom=" & CStr(effectiveFrom)
    AT_CreateContributionRate = rateID: Exit Function
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "RATE_CREATE", "tblContributionRates", 0, "FAILED", eDesc
    Err.Raise savedErrNumber, "AT_CreateContributionRate", eDesc
End Function

Public Sub AT_SetUserPreference(ByVal preferenceKey As String, ByVal preferenceValue As String)
    On Error GoTo EH
    Dim userID As Long, db As DAO.Database, rs As DAO.Recordset
    userID = CLng(Nz(DLookup("UserID", "tblUsers", "UserName='" & Replace(Environ$("Username"), "'", "''") & "' AND IsActive=True"), 0))
    If userID = 0 Then Err.Raise vbObjectError + 1520, , AT_U("0627,0644,0645,0633,062A,062E,062F,0645,0020,0627,0644,062D,0627,0644,064A,0020,063A,064A,0631,0020,0645,0633,062C,0644,0020,0641,064A,0020,0627,0644,0646,0638,0627,0645,002E")
    If Len(Trim$(preferenceKey)) = 0 Then Err.Raise vbObjectError + 1521, , AT_U("0645,0641,062A,0627,062D,0020,0627,0644,062A,0641,0636,064A,0644,0020,0645,0637,0644,0648,0628,002E")
    Set db = CurrentDb: Set rs = db.OpenRecordset("SELECT * FROM tblUserPreferences WHERE UserID=" & userID & " AND PreferenceKey='" & Replace(preferenceKey, "'", "''") & "'", dbOpenDynaset)
    If rs.EOF Then rs.AddNew Else rs.Edit
    rs!UserID = userID: rs!PreferenceKey = preferenceKey: rs!PreferenceValue = preferenceValue: rs!UpdatedAt = Now(): rs.Update: rs.Close
    AT_Audit "USER_PREFERENCE", "tblUserPreferences", userID, "SUCCESS", preferenceKey
    Exit Sub
EH:
    Err.Raise Err.Number, "AT_SetUserPreference", Err.Description
End Sub
