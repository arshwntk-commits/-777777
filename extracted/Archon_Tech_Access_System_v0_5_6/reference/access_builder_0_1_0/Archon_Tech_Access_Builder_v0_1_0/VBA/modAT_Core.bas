Attribute VB_Name = "modAT_Core"
Option Compare Database
Option Explicit

' Archon Tech Access Builder - core utilities
' Compatible target: Access 2021 / Access 2024, 32-bit or 64-bit VBA.

Public Const AT_VERSION As String = "0.1.0"
Public Const AT_DEFAULT_CURRENCY As String = "YER"
Public Const AT_DEFAULT_CONTRIBUTION As Currency = 3000

Public Const AT_DB_TEXT As Long = 10
Public Const AT_DB_MEMO As Long = 12
Public Const AT_DB_LONG As Long = 4
Public Const AT_DB_CURRENCY As Long = 5
Public Const AT_DB_DOUBLE As Long = 7
Public Const AT_DB_DATE As Long = 8
Public Const AT_DB_BOOLEAN As Long = 1
Public Const AT_DB_AUTOINCR As Long = 16
Public Const AT_DB_FAIL_ON_ERROR As Long = 128
Public Const AT_DB_OPEN_DYNASET As Long = 2
Public Const AT_DB_OPEN_SNAPSHOT As Long = 4

Public Function AT_U(ByVal encodedCodePoints As String) As String
    Dim tokens As Variant, i As Long
    tokens = Split(encodedCodePoints, ",")
    For i = LBound(tokens) To UBound(tokens)
        AT_U = AT_U & ChrW(CLng(Val("&H" & Trim$(CStr(tokens(i))))))
    Next i
End Function

Public Function AT_TableExists(ByVal tableName As String) As Boolean
    Dim tdf As Object
    On Error GoTo CleanFail
    For Each tdf In CurrentDb.TableDefs
        If StrComp(tdf.Name, tableName, vbTextCompare) = 0 Then
            AT_TableExists = True
            Exit Function
        End If
    Next
CleanExit:
    Exit Function
CleanFail:
    AT_TableExists = False
    Resume CleanExit
End Function

Public Function AT_QueryExists(ByVal queryName As String) As Boolean
    Dim qdf As Object
    On Error GoTo CleanFail
    For Each qdf In CurrentDb.QueryDefs
        If StrComp(qdf.Name, queryName, vbTextCompare) = 0 Then
            AT_QueryExists = True
            Exit Function
        End If
    Next
CleanExit:
    Exit Function
CleanFail:
    AT_QueryExists = False
    Resume CleanExit
End Function

Public Function AT_FormExists(ByVal formName As String) As Boolean
    On Error GoTo CleanFail
    AT_FormExists = CurrentProject.AllForms(formName).Name = formName
    Exit Function
CleanFail:
    AT_FormExists = False
End Function

Public Function AT_FieldExists(ByVal tableName As String, ByVal fieldName As String) As Boolean
    Dim fld As Object
    On Error GoTo CleanFail
    For Each fld In CurrentDb.TableDefs(tableName).Fields
        If StrComp(fld.Name, fieldName, vbTextCompare) = 0 Then
            AT_FieldExists = True
            Exit Function
        End If
    Next
CleanExit:
    Exit Function
CleanFail:
    AT_FieldExists = False
    Resume CleanExit
End Function

Public Function AT_SqlText(ByVal value As String) As String
    AT_SqlText = "'" & Replace(Nz(value, vbNullString), "'", "''") & "'"
End Function

Public Function AT_SqlDate(ByVal value As Date) As String
    AT_SqlDate = "#" & Format$(value, "yyyy-mm-dd hh:nn:ss") & "#"
End Function

Public Function AT_UserName() As String
    AT_UserName = Environ$("USERNAME")
    If Len(AT_UserName) = 0 Then AT_UserName = "AccessUser"
End Function

Public Sub AT_ExecSQL(ByVal sqlText As String)
    CurrentDb.Execute sqlText, AT_DB_FAIL_ON_ERROR
End Sub

Public Sub AT_EnsureQuery(ByVal queryName As String, ByVal sqlText As String)
    Dim db As Object
    Dim qdf As Object
    Set db = CurrentDb
    If AT_QueryExists(queryName) Then
        Set qdf = db.QueryDefs(queryName)
        qdf.SQL = sqlText
    Else
        Set qdf = db.CreateQueryDef(queryName, sqlText)
    End If
End Sub

Public Function AT_NextCode(ByVal prefix As String, ByVal tableName As String, ByVal fieldName As String, Optional ByVal digits As Long = 6) As String
    Dim rs As Object
    Dim sqlText As String
    Dim candidate As String
    Dim suffix As String
    Dim maxValue As Long
    Dim currentValue As Long
    On Error GoTo CleanFail

    maxValue = 0
    If AT_TableExists(tableName) And AT_FieldExists(tableName, fieldName) Then
        sqlText = "SELECT [" & fieldName & "] FROM [" & tableName & "] WHERE [" & fieldName & "] Is Not Null"
        Set rs = CurrentDb.OpenRecordset(sqlText, AT_DB_OPEN_SNAPSHOT)
        Do While Not rs.EOF
            candidate = Nz(rs.Fields(0).Value, vbNullString)
            If Len(candidate) > Len(prefix) And StrComp(Left$(candidate, Len(prefix)), prefix, vbTextCompare) = 0 Then
                suffix = Mid$(candidate, Len(prefix) + 1)
                If IsNumeric(suffix) Then
                    currentValue = CLng(Val(suffix))
                    If currentValue > maxValue Then maxValue = currentValue
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If
    AT_NextCode = prefix & Format$(maxValue + 1, String$(digits, "0"))
    Exit Function
CleanFail:
    AT_NextCode = prefix & Format$(1, String$(digits, "0"))
End Function

Public Sub AT_LogBuild(ByVal stepName As String, ByVal result As String, Optional ByVal details As String = vbNullString)
    On Error Resume Next
    If Not AT_TableExists("tblAT_BuildLog") Then Exit Sub
    CurrentDb.Execute "INSERT INTO tblAT_BuildLog (EventTime, StepName, Result, Details, WindowsUser) VALUES (" & _
        AT_SqlDate(Now()) & "," & AT_SqlText(stepName) & "," & AT_SqlText(result) & "," & AT_SqlText(details) & "," & AT_SqlText(AT_UserName()) & ")", AT_DB_FAIL_ON_ERROR
End Sub

Public Sub AT_Audit(ByVal actionName As String, ByVal entityName As String, Optional ByVal entityID As Variant, Optional ByVal result As String = "Success", Optional ByVal details As String = vbNullString, Optional ByVal oldValue As String = vbNullString, Optional ByVal newValue As String = vbNullString)
    Dim idSql As String
    On Error Resume Next
    If IsNull(entityID) Or IsEmpty(entityID) Then
        idSql = "Null"
    ElseIf IsNumeric(entityID) Then
        idSql = CStr(CLng(entityID))
    Else
        idSql = AT_SqlText(CStr(entityID))
    End If
    CurrentDb.Execute "INSERT INTO tblAT_AuditLog (EventDateTime, UserName, ActionName, EntityName, EntityID, Result, Details, OldValue, NewValue) VALUES (" & _
        AT_SqlDate(Now()) & "," & AT_SqlText(AT_UserName()) & "," & AT_SqlText(actionName) & "," & AT_SqlText(entityName) & "," & idSql & "," & _
        AT_SqlText(result) & "," & AT_SqlText(details) & "," & AT_SqlText(oldValue) & "," & AT_SqlText(newValue) & ")", AT_DB_FAIL_ON_ERROR
End Sub

Public Function AT_GetSetting(ByVal settingKey As String, Optional ByVal defaultValue As String = vbNullString) As String
    Dim result As Variant
    On Error GoTo UseDefault
    result = DLookup("SettingValue", "tblAT_AppSettings", "SettingKey=" & AT_SqlText(settingKey))
    If IsNull(result) Then GoTo UseDefault
    AT_GetSetting = CStr(result)
    Exit Function
UseDefault:
    AT_GetSetting = defaultValue
End Function

Public Sub AT_SetSetting(ByVal settingKey As String, ByVal settingValue As String)
    Dim criteria As String
    criteria = "SettingKey=" & AT_SqlText(settingKey)
    If DCount("*", "tblAT_AppSettings", criteria) = 0 Then
        CurrentDb.Execute "INSERT INTO tblAT_AppSettings (SettingKey, SettingValue, UpdatedAt) VALUES (" & AT_SqlText(settingKey) & "," & AT_SqlText(settingValue) & "," & AT_SqlDate(Now()) & ")", AT_DB_FAIL_ON_ERROR
    Else
        CurrentDb.Execute "UPDATE tblAT_AppSettings SET SettingValue=" & AT_SqlText(settingValue) & ", UpdatedAt=" & AT_SqlDate(Now()) & " WHERE " & criteria, AT_DB_FAIL_ON_ERROR
    End If
End Sub

Public Function AT_AccountBalance(ByVal accountID As Long) As Currency
    AT_AccountBalance = CCur(Nz(DSum("AmountSigned", "tblAT_AccountMovements", "AccountID=" & accountID), 0))
End Function

Public Function AT_AccountAvailableBalance(ByVal accountID As Long, Optional ByVal excludeTransferID As Long = 0) As Currency
    Dim reserved As Currency
    Dim criteria As String
    criteria = "FromAccountID=" & accountID & " AND Status='Pending'"
    If excludeTransferID > 0 Then criteria = criteria & " AND TransferID<>" & excludeTransferID
    reserved = CCur(Nz(DSum("Amount", "tblAT_Transfers", criteria), 0)) + CCur(Nz(DSum("FeeAmount", "tblAT_Transfers", criteria), 0))
    AT_AccountAvailableBalance = AT_AccountBalance(accountID) - reserved
End Function

Public Function AT_OwnerDue(ByVal ownerID As Long) As Currency
    Dim rs As Object
    Dim totalCharges As Currency
    Dim totalAllocated As Currency
    Set rs = CurrentDb.OpenRecordset("SELECT Sum(CurrentAmount) AS TotalValue FROM tblAT_Charges WHERE OwnerID=" & ownerID & " AND Status<>'Cancelled'", AT_DB_OPEN_SNAPSHOT)
    If Not rs.EOF Then totalCharges = CCur(Nz(rs!TotalValue, 0))
    rs.Close
    Set rs = CurrentDb.OpenRecordset("SELECT Sum(A.AmountSigned) AS TotalValue FROM tblAT_ReceiptAllocations AS A INNER JOIN tblAT_Charges AS C ON A.ChargeID=C.ChargeID WHERE C.OwnerID=" & ownerID, AT_DB_OPEN_SNAPSHOT)
    If Not rs.EOF Then totalAllocated = CCur(Nz(rs!TotalValue, 0))
    rs.Close
    AT_OwnerDue = totalCharges - totalAllocated
End Function

Public Function AT_OwnerPaid(ByVal ownerID As Long) As Currency
    Dim rs As Object
    Set rs = CurrentDb.OpenRecordset("SELECT Sum(A.AmountSigned) AS PaidTotal FROM tblAT_ReceiptAllocations AS A INNER JOIN tblAT_Charges AS C ON A.ChargeID=C.ChargeID WHERE C.OwnerID=" & ownerID, AT_DB_OPEN_SNAPSHOT)
    If rs.EOF Then
        AT_OwnerPaid = 0
    Else
        AT_OwnerPaid = CCur(Nz(rs!PaidTotal, 0))
    End If
    rs.Close
End Function

Public Function AT_OwnerCredit(ByVal ownerID As Long) As Currency
    AT_OwnerCredit = CCur(Nz(DSum("AmountSigned", "tblAT_OwnerCreditMovements", "OwnerID=" & ownerID), 0))
End Function


Public Function AT_NormalizeArabic(ByVal inputText As Variant) As String
    Dim s As String, marks As Variant, i As Long
    If IsNull(inputText) Or IsEmpty(inputText) Then Exit Function
    s = CStr(inputText)
    s = Replace(s, ChrW(&H640), "") ' Tatweel
    marks = Array(ChrW(&H64B), ChrW(&H64C), ChrW(&H64D), ChrW(&H64E), ChrW(&H64F), ChrW(&H650), ChrW(&H651), ChrW(&H652), ChrW(&H670))
    For i = LBound(marks) To UBound(marks)
        s = Replace(s, CStr(marks(i)), vbNullString)
    Next i
    s = Replace(s, AT_U("0623"), AT_U("0627"))
    s = Replace(s, AT_U("0625"), AT_U("0627"))
    s = Replace(s, AT_U("0622"), AT_U("0627"))
    s = Replace(s, AT_U("0629"), AT_U("0647"))
    s = Replace(s, AT_U("0649"), AT_U("064A"))
    s = Trim$(s)
    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop
    AT_NormalizeArabic = LCase$(s)
End Function

Public Function AT_CurrentRecordID(ByVal formName As String, ByVal primaryKeyName As String) As Long
    On Error GoTo CleanFail
    If Forms(formName).NewRecord Then Exit Function
    AT_CurrentRecordID = CLng(Nz(Forms(formName).Recordset.Fields(primaryKeyName).Value, 0))
CleanExit:
    Exit Function
CleanFail:
    AT_CurrentRecordID = 0
    Resume CleanExit
End Function

Public Function AT_IsAccountActive(ByVal accountID As Long) As Boolean
    AT_IsAccountActive = (Nz(DLookup("IsActive", "tblAT_FinancialAccounts", "AccountID=" & accountID), False) = True)
End Function

Public Function AT_IsMethodActive(ByVal methodID As Long) As Boolean
    AT_IsMethodActive = (Nz(DLookup("IsActive", "tblAT_PaymentMethods", "PaymentMethodID=" & methodID), False) = True)
End Function

Public Sub AT_SetStartupForm(ByVal formName As String)
    Dim db As Object
    Dim prp As Object
    Set db = CurrentDb
    On Error Resume Next
    db.Properties("StartupForm") = formName
    If Err.Number <> 0 Then
        Err.Clear
        Set prp = db.CreateProperty("StartupForm", 10, formName)
        db.Properties.Append prp
    End If
    On Error GoTo 0
End Sub
