Attribute VB_Name = "modArchonCore"
Option Compare Database
Option Explicit

' Unicode-safe UI text helper. Arabic labels are encoded as code points in source files.
Public Function AT_U(ByVal codePoints As String) As String
    Dim parts As Variant, i As Long, value As Long, result As String
    parts = Split(codePoints, ",")
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(CStr(parts(i)))) > 0 Then
            value = CLng("&H" & Trim$(CStr(parts(i))))
            result = result & ChrW$(value)
        End If
    Next i
    AT_U = result
End Function

Public Function AT_TableExistsInCurrent(ByVal tableName As String) As Boolean
    Dim tdf As DAO.TableDef
    On Error Resume Next
    Set tdf = CurrentDb.TableDefs(tableName)
    AT_TableExistsInCurrent = (Err.Number = 0 And Not tdf Is Nothing)
    Err.Clear
End Function

Public Function AT_EnsureDirectory(ByVal pathText As String) As Boolean
    On Error GoTo EH
    If Len(Dir$(pathText, vbDirectory)) = 0 Then MkDir pathText
    AT_EnsureDirectory = True
    Exit Function
EH:
    Err.Raise Err.Number, "AT_EnsureDirectory", Err.Description
End Function
