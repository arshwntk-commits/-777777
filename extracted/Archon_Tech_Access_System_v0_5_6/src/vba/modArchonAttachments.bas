Attribute VB_Name = "modArchonAttachments"
Option Compare Database
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As LongPtr, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As LongPtr, ByVal Algid As Long, ByVal hKey As LongPtr, ByVal dwFlags As Long, ByRef phHash As LongPtr) As Long
Private Declare PtrSafe Function CryptHashData Lib "advapi32.dll" (ByVal hHash As LongPtr, ByRef pbData As Byte, ByVal dataLen As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As LongPtr, ByVal dwParam As Long, ByRef pbData As Byte, ByRef dataLen As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As LongPtr) As Long
Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As LongPtr, ByVal dwFlags As Long) As Long
#End If

Private Const AT_PROV_RSA_AES As Long = 24
Private Const AT_CALG_SHA_256 As Long = &H800C
Private Const AT_CRYPT_VERIFYCONTEXT As Long = -268435456
Private Const AT_HP_HASHVAL As Long = 2

Public Function AT_SHA256File(ByVal filePath As String) As String
    On Error GoTo EH
    Dim provider As LongPtr, hashHandle As LongPtr
    Dim fileNo As Integer, fileOpen As Boolean
    Dim remaining As Long, chunkSize As Long, digestLength As Long, i As Long
    Dim chunk() As Byte, digest(0 To 31) As Byte, result As String
    Dim savedErrNumber As Long, eDesc As String

    If Len(Dir$(filePath)) = 0 Then Err.Raise vbObjectError + 1610, "AT_SHA256File", "Attachment file does not exist."
    If CryptAcquireContext(provider, vbNullString, vbNullString, AT_PROV_RSA_AES, AT_CRYPT_VERIFYCONTEXT) = 0 Then _
        Err.Raise vbObjectError + 1611, "AT_SHA256File", "Windows CryptoAPI could not create a SHA-256 provider."
    If CryptCreateHash(provider, AT_CALG_SHA_256, 0, 0, hashHandle) = 0 Then _
        Err.Raise vbObjectError + 1612, "AT_SHA256File", "Windows CryptoAPI could not create a SHA-256 hash."

    fileNo = FreeFile
    Open filePath For Binary Access Read As #fileNo
    fileOpen = True
    remaining = LOF(fileNo)
    Do While remaining > 0
        chunkSize = remaining
        If chunkSize > 1048576 Then chunkSize = 1048576
        ReDim chunk(0 To chunkSize - 1)
        Get #fileNo, , chunk
        If CryptHashData(hashHandle, chunk(0), chunkSize, 0) = 0 Then _
            Err.Raise vbObjectError + 1613, "AT_SHA256File", "Windows CryptoAPI failed while hashing the attachment."
        remaining = remaining - chunkSize
    Loop
    Close #fileNo: fileOpen = False
    digestLength = 32
    If CryptGetHashParam(hashHandle, AT_HP_HASHVAL, digest(0), digestLength, 0) = 0 Or digestLength <> 32 Then _
        Err.Raise vbObjectError + 1614, "AT_SHA256File", "Windows CryptoAPI could not read the SHA-256 digest."
    For i = 0 To 31
        result = result & Right$("0" & Hex$(digest(i)), 2)
    Next i
    AT_SHA256File = LCase$(result)
    GoTo CleanExit
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next
    If fileOpen Then Close #fileNo
    If hashHandle <> 0 Then CryptDestroyHash hashHandle
    If provider <> 0 Then CryptReleaseContext provider, 0
    On Error GoTo 0
    Err.Raise savedErrNumber, "AT_SHA256File", eDesc
CleanExit:
    On Error Resume Next
    If hashHandle <> 0 Then CryptDestroyHash hashHandle
    If provider <> 0 Then CryptReleaseContext provider, 0
End Function


' Copies supporting files into the managed Attachments folder and links them to a document.
' The document's financial state is never changed by adding an attachment.

Public Function AT_AddAttachment(ByVal documentID As Long, ByVal sourcePath As String) As Long
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "ATTACHMENT_ADD"
    If documentID <= 0 Or DCount("*", "tblFinancialDocuments", "DocumentID=" & documentID) = 0 Then Err.Raise vbObjectError + 1600, , AT_U("0627,0644,0645,0633,062A,0646,062F,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,061B,0020,0644,0627,0020,064A,0645,0643,0646,0020,0631,0628,0637,0020,0627,0644,0645,0631,0641,0642,002E")
    If Len(Dir$(sourcePath)) = 0 Then Err.Raise vbObjectError + 1601, , AT_U("0645,0644,0641,0020,0627,0644,0645,0631,0641,0642,0020,0627,0644,0645,062D,062F,062F,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,002E")

    Dim backendPath As String, basePath As String, docFolder As String, originalName As String
    Dim targetPath As String, relativePath As String, fileBytes As Double, ext As String
    Dim rs As DAO.Recordset, db As DAO.Database, attachmentID As Long, dotPos As Long
    Dim documentType As String, linkedRecordID As Long, transactionStarted As Boolean, committed As Boolean
    backendPath = CStr(Nz(DLookup("SettingValue", "tblSystemSettings", "SettingKey='BACKEND_PATH'"), ""))
    If Len(backendPath) = 0 Then Err.Raise vbObjectError + 1602, , AT_U("0645,0633,0627,0631,0020,0642,0627,0639,062F,0629,0020,0627,0644,0628,064A,0627,0646,0627,062A,0020,0627,0644,062E,0644,0641,064A,0629,0020,063A,064A,0631,0020,0645,0636,0628,0648,0637,061B,0020,0623,0639,062F,0020,0631,0628,0637,0020,0627,0644,0642,0627,0639,062F,0629,0020,0623,0648,0020,0631,0627,062C,0639,0020,0625,0639,062F,0627,062F,0627,062A,0020,0627,0644,062A,062B,0628,064A,062A,002E")
    basePath = Left$(backendPath, InStrRev(backendPath, "\") - 1) & "\Attachments"
    If Len(Dir$(basePath, vbDirectory)) = 0 Then MkDir basePath
    docFolder = basePath & "\" & CStr(documentID)
    If Len(Dir$(docFolder, vbDirectory)) = 0 Then MkDir docFolder
    originalName = Dir$(sourcePath)
    dotPos = InStrRev(originalName, "."): If dotPos > 0 Then ext = Mid$(originalName, dotPos)
    targetPath = docFolder & "\" & originalName
    If Len(Dir$(targetPath)) > 0 Then targetPath = docFolder & "\" & Format$(Now(), "yyyymmdd_hhnnss_") & originalName
    FileCopy sourcePath, targetPath
    fileBytes = FileLen(targetPath)
    relativePath = CStr(documentID) & "\" & Mid$(targetPath, InStrRev(targetPath, "\") + 1)

    Set db = CurrentDb
    documentType = UCase$(CStr(Nz(DLookup("DocumentType", "tblFinancialDocuments", "DocumentID=" & documentID), "")))
    DBEngine(0)(0).BeginTrans: transactionStarted = True
    Set rs = db.OpenRecordset("tblAttachments", dbOpenDynaset)
    rs.AddNew: rs!DocumentID = documentID: rs!RelativePath = relativePath: rs!OriginalFileName = originalName
    rs!ContentType = AT_GuessContentType(ext): rs!FileSizeBytes = fileBytes: rs!SHA256 = AT_SHA256File(targetPath)
    rs!AddedAt = Now(): rs!AddedByUserName = Environ$("Username"): rs.Update: rs.Bookmark = rs.LastModified
    attachmentID = CLng(rs!AttachmentID): rs.Close
    If documentType = "RECEIPT" Then
        linkedRecordID = CLng(Nz(DLookup("ReceiptID", "tblReceipts", "DocumentID=" & documentID), 0))
        If linkedRecordID > 0 Then db.Execute "INSERT INTO tblReceiptAttachments (ReceiptID,AttachmentID) VALUES (" & linkedRecordID & "," & attachmentID & ")", dbFailOnError
    ElseIf documentType = "PAYMENT" Then
        linkedRecordID = CLng(Nz(DLookup("ExpenseID", "tblPayments", "DocumentID=" & documentID), 0))
        If linkedRecordID > 0 Then db.Execute "INSERT INTO tblExpenseAttachments (ExpenseID,AttachmentID) VALUES (" & linkedRecordID & "," & attachmentID & ")", dbFailOnError
    End If
    DBEngine(0)(0).CommitTrans: transactionStarted = False: committed = True
    AT_AddAttachment = attachmentID
    On Error Resume Next: AT_Audit "ATTACHMENT_ADD", "tblAttachments", attachmentID, "SUCCESS", relativePath
    On Error GoTo 0
    Exit Function
EH:

    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next
    If transactionStarted Then DBEngine(0)(0).Rollback
    If Not committed Then
        If Len(targetPath) > 0 Then
            If Len(Dir$(targetPath)) > 0 Then Kill targetPath
        End If
    End If
    AT_Audit "ATTACHMENT_ADD", "tblAttachments", 0, "FAILED", eDesc
    Err.Raise savedErrNumber, "AT_AddAttachment", eDesc
End Function

Private Function AT_GuessContentType(ByVal extensionText As String) As String
    Select Case LCase$(extensionText)
        Case ".pdf": AT_GuessContentType = "application/pdf"
        Case ".jpg", ".jpeg": AT_GuessContentType = "image/jpeg"
        Case ".png": AT_GuessContentType = "image/png"
        Case ".doc": AT_GuessContentType = "application/msword"
        Case ".docx": AT_GuessContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        Case ".xls": AT_GuessContentType = "application/vnd.ms-excel"
        Case ".xlsx": AT_GuessContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        Case Else: AT_GuessContentType = "application/octet-stream"
    End Select
End Function

Public Function AT_SelectAndAttachFile(ByVal documentID As Long) As Long
    On Error GoTo EH
    Dim dlg As Object
    Set dlg = Application.FileDialog(3)
    dlg.AllowMultiSelect = False
    dlg.Title = AT_U("0627,062E,062A,0631,0020,0645,0644,0641,064B,0627,0020,0644,0625,0631,0641,0627,0642,0647,0020,0628,0627,0644,0645,0633,062A,0646,062F")
    If dlg.Show <> -1 Then Exit Function
    AT_SelectAndAttachFile = AT_AddAttachment(documentID, CStr(dlg.SelectedItems(1)))
    Exit Function
EH:
    Err.Raise Err.Number, "AT_SelectAndAttachFile", Err.Description
End Function
