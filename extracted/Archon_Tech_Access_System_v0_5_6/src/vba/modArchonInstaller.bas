Attribute VB_Name = "modArchonInstaller"
Option Compare Database
Option Explicit

' Called by the Windows PowerShell bootstrap after it creates a blank FE and imports modules.
Public Sub AT_InstallSplitProject(ByVal backendPath As String)
    Dim savedErrNumber As Long, eDesc As String, savedErrSource As String
    Dim dbBE As DAO.Database, tdfBE As DAO.TableDef
    Dim workspace As DAO.Workspace, createdBackend As Boolean
    Dim backendFolder As String, attachmentsFolder As String, backupsFolder As String
    Dim dbVersion As Long, prop As DAO.Property
    On Error GoTo EH

    backendPath = Trim$(backendPath)
    If Len(backendPath) = 0 Then Err.Raise vbObjectError + 1100, , AT_U("0645,0643,0627,0646,0020,062D,0641,0638,0020,0645,0644,0641,0020,0627,0644,0628,064A,0627,0646,0627,062A,0020,063A,064A,0631,0020,0635,0627,0644,062D,002E")
    If Len(Dir$(backendPath)) > 0 Then Err.Raise vbObjectError + 1101, , AT_U("064A,0648,062C,062F,0020,0645,0644,0641,0020,0628,064A,0627,0646,0627,062A,0020,0628,0646,0641,0633,0020,0627,0644,0627,0633,0645,002E,0020,0644,0645,0020,064A,062A,0645,0020,0627,0644,0643,062A,0627,0628,0629,0020,0639,0644,064A,0647,002E")
    If CurrentProject.AllForms.Count > 0 Or CurrentDb.TableDefs.Count > 0 Then
        If AT_TableExistsInCurrent("tblOwners") Or AT_TableExistsInCurrent("tblFinancialDocuments") Then Err.Raise vbObjectError + 1102, , AT_U("0645,0644,0641,0020,0627,0644,0648,0627,062C,0647,0629,0020,0644,064A,0633,0020,0641,0627,0631,063A,0627,0020,0623,0648,0020,0633,0628,0642,0020,0628,0646,0627,0624,0647,002E,0020,0644,0645,0020,064A,062A,0645,0020,062A,063A,064A,064A,0631,0020,0627,0644,0645,0644,0641,0020,0627,0644,062D,0627,0644,064A,002E")
    End If

    backendFolder = Left$(backendPath, InStrRev(backendPath, "\") - 1)
    attachmentsFolder = backendFolder & "\Attachments"
    backupsFolder = backendFolder & "\Backups"
    AT_EnsureDirectory attachmentsFolder
    AT_EnsureDirectory backupsFolder

    Set workspace = DBEngine.Workspaces(0)
    Set dbBE = workspace.CreateDatabase(backendPath, dbLangGeneral, dbVersion120)
    createdBackend = True
    AT_CreateSchemaOnDatabase dbBE
    dbBE.Close: Set dbBE = Nothing

    AT_LinkBackendTables backendPath
    ' Seed after links exist because the initializer uses domain functions against CurrentDb.
    AT_SeedDefaults CurrentDb
    AT_SetSchemaVersion CurrentDb
    CurrentDb.Execute "UPDATE tblSystemSettings SET SettingValue='" & Replace(backendPath, "'", "''") & "', UpdatedAt=Now(), UpdatedByUserName='INSTALLER' WHERE SettingKey='BACKEND_PATH'", dbFailOnError
    If CurrentDb.RecordsAffected = 0 Then CurrentDb.Execute "INSERT INTO tblSystemSettings (SettingKey,SettingValue,UpdatedAt,UpdatedByUserName) VALUES ('BACKEND_PATH','" & Replace(backendPath, "'", "''") & "',Now(),'INSTALLER')", dbFailOnError

    AT_BuildInitialUI
    AT_BuildPrintableReports
    AT_SetStartupForm "frmMain"
    AT_Audit "INSTALL_SPLIT_PROJECT", "tblSchemaVersion", 0, "SUCCESS", "Split FE/BE installed; schema 0.5.0."
    MsgBox AT_U("062A,0645,0020,0625,0646,0634,0627,0621,0020,0645,0644,0641,0020,0627,0644,0628,064A,0627,0646,0627,062A,0020,0648,0631,0628,0637,0020,0645,0644,0641,0020,0627,0644,0648,0627,062C,0647,0629,0020,0648,0625,0646,0634,0627,0621,0020,0627,0644,0648,0627,062C,0647,0627,062A,0020,0648,0627,0644,0627,0633,062A,0639,0644,0627,0645,0627,062A,0020,0648,0627,0644,062A,0642,0627,0631,064A,0631,0627,062A,002E,0020,0644,0627,0020,062A,0633,062A,062E,062F,0645,0020,0628,064A,0627,0646,0627,062A,0020,062D,0642,064A,0642,064A,0629,0020,0642,0628,0644,0020,062A,0646,0641,064A,0630,0020,0645,062C,0645,0648,0639,0629,0020,0627,062E,062A,0628,0627,0631,0627,062A,0020,0627,0644,0642,0628,0648,0644,0020,0641,064A,0020,0625,0635,062F,0627,0631,0020,0041,0063,0063,0065,0073,0073,0020,0627,0644,0645,0633,062A,0647,062F,0641,002E"), vbInformation, "Archon Tech"
    Exit Sub
EH:

    savedErrNumber = Err.Number: eDesc = Err.Description: savedErrSource = Err.Source
    On Error Resume Next
    If Not dbBE Is Nothing Then dbBE.Close
    If createdBackend Then
        AT_RemoveBackendLinks
        If Len(Dir$(backendPath)) > 0 Then Kill backendPath
    End If
    MsgBox AT_U("0641,0634,0644,0020,0627,0644,062A,062B,0628,064A,062A,002E,0020,0627,0644,062E,0637,0623,0020,0631,0642,0645,0020") & savedErrNumber & " [" & savedErrSource & "] - " & eDesc, vbCritical, "Archon Tech"
    Err.Raise savedErrNumber, savedErrSource, eDesc
End Sub

' No-argument entry point for Access.Application.Run.
' The PowerShell bootstrap writes the backend path to a temporary UTF-16 file,
' avoiding COM argument marshaling differences between PowerShell and Access.
Public Function AT_InstallSplitProject_Run() As Boolean
    Dim savedErrNumber As Long, eDesc As String, savedErrSource As String
    Dim fso As Object, ts As Object, configPath As String, backendPath As String
    On Error GoTo EH

    configPath = CurrentProject.Path & "\.archon_backend_path.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(configPath) Then _
        Err.Raise vbObjectError + 1190, "AT_InstallSplitProject_Run", "Installer configuration file was not found."

    Set ts = fso.OpenTextFile(configPath, 1, False, -1) ' ForReading, Unicode
    backendPath = Trim$(ts.ReadAll)
    ts.Close
    Set ts = Nothing
    fso.DeleteFile configPath, True
    Set fso = Nothing

    If Len(backendPath) = 0 Then _
        Err.Raise vbObjectError + 1191, "AT_InstallSplitProject_Run", "Backend path in installer configuration is empty."

    AT_InstallSplitProject backendPath
    AT_InstallSplitProject_Run = True
    Exit Function
EH:

    savedErrNumber = Err.Number: eDesc = Err.Description: savedErrSource = Err.Source
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    Err.Raise savedErrNumber, savedErrSource, eDesc
End Function

Public Sub AT_LinkBackendTables(ByVal backendPath As String)
    Dim savedErrNumber As Long, eDesc As String, savedErrSource As String
    Dim dbFE As DAO.Database, dbBE As DAO.Database, tdf As DAO.TableDef, linkDef As DAO.TableDef
    On Error GoTo EH
    Set dbFE = CurrentDb: Set dbBE = DBEngine.Workspaces(0).OpenDatabase(backendPath, False, True)
    For Each tdf In dbBE.TableDefs
        If (tdf.Attributes And dbSystemObject) = 0 And Left$(tdf.Name, 4) <> "MSys" Then
            If AT_TableExistsInCurrent(tdf.Name) Then Err.Raise vbObjectError + 1110, , AT_U("0627,0644,062C,062F,0648,0644,0020,0645,0648,062C,0648,062F,0020,0645,0633,0628,0642,064B,0627,0020,0641,064A,0020,0645,0644,0641,0020,0627,0644,0648,0627,062C,0647,0629,003A,0020") & tdf.Name
            Set linkDef = dbFE.CreateTableDef(tdf.Name)
            linkDef.Connect = ";DATABASE=" & backendPath
            linkDef.SourceTableName = tdf.Name
            dbFE.TableDefs.Append linkDef
            dbFE.TableDefs.Refresh
        End If
    Next tdf
    dbBE.Close
    Exit Sub
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description: savedErrSource = Err.Source
    On Error Resume Next: If Not dbBE Is Nothing Then dbBE.Close
    Err.Raise savedErrNumber, "AT_LinkBackendTables", eDesc
End Sub

Public Sub AT_RemoveBackendLinks()
    On Error Resume Next
    Dim db As DAO.Database, i As Long, tdf As DAO.TableDef
    Set db = CurrentDb
    For i = db.TableDefs.Count - 1 To 0 Step -1
        Set tdf = db.TableDefs(i)
        If Len(tdf.Connect) > 0 Then db.TableDefs.Delete tdf.Name
    Next i
End Sub

Public Sub AT_SetStartupForm(ByVal formName As String)
    On Error GoTo CreateProp
    CurrentDb.Properties("StartupForm") = formName
    Exit Sub
CreateProp:
    Err.Clear
    Dim prop As DAO.Property
    Set prop = CurrentDb.CreateProperty("StartupForm", dbText, formName)
    CurrentDb.Properties.Append prop
End Sub

Public Sub AT_RelinkBackend(ByVal backendPath As String)
    On Error GoTo EH
    Dim db As DAO.Database, dbBE As DAO.Database, tdf As DAO.TableDef
    Dim tableNames() As String, oldConnects() As String, sourceNames() As String
    Dim linkCount As Long, changedCount As Long, i As Long, sourceTable As String
    Dim savedErrNumber As Long, eDesc As String, savedErrSource As String, rollbackErrors As String

    backendPath = Trim$(backendPath)
    If Len(backendPath) = 0 Or Len(Dir$(backendPath)) = 0 Then _
        Err.Raise vbObjectError + 1120, "AT_RelinkBackend", AT_U("0645,0644,0641,0020,0642,0627,0639,062F,0629,0020,0627,0644,0628,064A,0627,0646,0627,062A,0020,0627,0644,062E,0644,0641,064A,0629,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,002E")
    Set db = CurrentDb
    Set dbBE = DBEngine.Workspaces(0).OpenDatabase(backendPath, False, True)

    ' Preflight every linked table before changing any connection.
    For Each tdf In db.TableDefs
        If Len(tdf.Connect) > 0 Then
            linkCount = linkCount + 1
            sourceTable = tdf.SourceTableName
            If Len(sourceTable) = 0 Then sourceTable = tdf.Name
            If Not AT_TableExistsInDatabase(dbBE, sourceTable) Then _
                Err.Raise vbObjectError + 1121, "AT_RelinkBackend", "Missing backend table: " & sourceTable
        End If
    Next tdf
    If linkCount = 0 Then Err.Raise vbObjectError + 1122, "AT_RelinkBackend", AT_U("0644,0627,0020,062A,0648,062C,062F,0020,062C,062F,0627,0648,0644,0020,0645,0631,062A,0628,0637,0629,0020,0641,064A,0020,0645,0644,0641,0020,0627,0644,0648,0627,062C,0647,0629,002E")

    ReDim tableNames(1 To linkCount)
    ReDim oldConnects(1 To linkCount)
    ReDim sourceNames(1 To linkCount)
    i = 0
    For Each tdf In db.TableDefs
        If Len(tdf.Connect) > 0 Then
            i = i + 1
            tableNames(i) = tdf.Name
            oldConnects(i) = tdf.Connect
            sourceNames(i) = tdf.SourceTableName
        End If
    Next tdf
    dbBE.Close: Set dbBE = Nothing

    For i = 1 To linkCount
        Set tdf = db.TableDefs(tableNames(i))
        tdf.Connect = ";DATABASE=" & backendPath
        changedCount = i
        tdf.RefreshLink
    Next i
    MsgBox AT_U("062A,0645,0020,062A,062D,062F,064A,062B,0020,0631,0648,0627,0628,0637,0020,0642,0627,0639,062F,0629,0020,0627,0644,0628,064A,0627,0646,0627,062A,0020,0628,0646,062C,0627,062D,002E,0020,0639,062F,062F,0020,0627,0644,062C,062F,0627,0648,0644,003A,0020") & CStr(linkCount), vbInformation, "Archon Tech"
    Exit Sub

EH:
    savedErrNumber = Err.Number: eDesc = Err.Description: savedErrSource = Err.Source
    On Error Resume Next
    If Not dbBE Is Nothing Then dbBE.Close
    rollbackErrors = ""
    For i = changedCount To 1 Step -1
        Err.Clear
        Set tdf = db.TableDefs(tableNames(i))
        tdf.Connect = oldConnects(i)
        If Len(sourceNames(i)) > 0 Then tdf.SourceTableName = sourceNames(i)
        tdf.RefreshLink
        If Err.Number <> 0 Then rollbackErrors = rollbackErrors & tableNames(i) & ": " & Err.Description & vbCrLf
    Next i
    On Error GoTo 0
    If Len(rollbackErrors) = 0 Then
        Err.Raise savedErrNumber, "AT_RelinkBackend", eDesc & vbCrLf & AT_U("062A,0645,0020,0627,0633,062A,0639,0627,062F,0629,0020,0627,0644,0631,0648,0627,0628,0637,0020,0627,0644,0633,0627,0628,0642,0629,002E")
    Else
        Err.Raise savedErrNumber, "AT_RelinkBackend", eDesc & vbCrLf & AT_U("062A,0639,0630,0631,0020,0627,0633,062A,0639,0627,062F,0629,0020,0628,0639,0636,0020,0627,0644,0631,0648,0627,0628,0637,002E") & vbCrLf & rollbackErrors
    End If
End Sub

Private Function AT_TableExistsInDatabase(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    Dim tdf As DAO.TableDef
    On Error GoTo NotFound
    Set tdf = db.TableDefs(tableName)
    AT_TableExistsInDatabase = True
    Exit Function
NotFound:
    AT_TableExistsInDatabase = False
End Function
