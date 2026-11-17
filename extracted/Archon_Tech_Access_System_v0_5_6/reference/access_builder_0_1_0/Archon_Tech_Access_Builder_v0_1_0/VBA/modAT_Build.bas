Attribute VB_Name = "modAT_Build"
Option Compare Database
Option Explicit

' Main entry point: run BuildArchonSystem from the Immediate Window (Ctrl+G)
' or attach it to a macro/button after all modules are imported.

Public Sub BuildArchonSystem()
    Dim answer As VbMsgBoxResult
    Dim failedChecks As Long
    On Error GoTo BuildFail

    answer = MsgBox(AT_U("0633,064A,0628,062F,0623") & " " & AT_U("0628,0646,0627,0621") & " " & AT_U("0646,0638,0627,0645") & " Archon Tech " & AT_U("062F,0627,062E,0644") & " " & AT_U("0642,0627,0639,062F,0629") & " " & AT_U("0627,0644,0628,064A,0627,0646,0627,062A") & " " & AT_U("0627,0644,062D,0627,0644,064A,0629") & "." & vbCrLf & _
        AT_U("0644,0646") & " " & AT_U("064A,062D,0630,0641") & " " & AT_U("0627,0644,0625,062C,0631,0627,0621") & " " & AT_U("0627,0644,062C,062F,0627,0648,0644") & " " & AT_U("0623,0648") & " " & AT_U("0627,0644,0628,064A,0627,0646,0627,062A") & " " & AT_U("0627,0644,0645,0648,062C,0648,062F,0629,060C") & " " & AT_U("0644,0643,0646,0647") & " " & AT_U("0642,062F") & " " & AT_U("064A,0636,064A,0641") & " " & AT_U("0643,0627,0626,0646,0627,062A") & " " & AT_U("0648,062D,0642,0648,0644,064B,0627") & " " & AT_U("0645,0641,0642,0648,062F,0629") & "." & vbCrLf & _
        AT_U("0623,0646,0634,0626") & " " & AT_U("0646,0633,062E,0629") & " " & AT_U("0627,062D,062A,064A,0627,0637,064A,0629") & " " & AT_U("0625,0630,0627") & " " & AT_U("0643,0627,0646,062A") & " " & AT_U("0627,0644,0642,0627,0639,062F,0629") & " " & AT_U("062A,062D,062A,0648,064A") & " " & AT_U("0628,064A,0627,0646,0627,062A") & " " & AT_U("0645,0647,0645,0629") & "." & vbCrLf & vbCrLf & _
        AT_U("0647,0644") & " " & AT_U("062A,0631,064A,062F") & " " & AT_U("0627,0644,0645,062A,0627,0628,0639,0629,061F"), vbQuestion + vbYesNo, "Archon Tech - " & AT_U("0628,062F,0621") & " " & AT_U("0627,0644,0628,0646,0627,0621"))
    If answer <> vbYes Then Exit Sub

    AT_LogBuild "Build", "Started", "Builder version " & AT_VERSION
    DoCmd.SetWarnings False
    AT_BuildSchema
    AT_BuildQueries
    AT_BuildForms
    AT_BuildReports
    AT_SetStartupForm "frmAT_Home"
    DoCmd.SetWarnings True

    failedChecks = AT_RunSelfTests(False)
    AT_LogBuild "Build", IIf(failedChecks = 0, "Success", "ReviewRequired"), "Failed structural checks=" & failedChecks
    If failedChecks = 0 Then
        MsgBox AT_U("0627,0643,062A,0645,0644") & " " & AT_U("0625,0646,0634,0627,0621") & " " & AT_U("0627,0644,0628,0646,064A,0629") & " " & AT_U("0627,0644,0623,0648,0644,064A,0629") & " " & AT_U("0644,0644,0646,0638,0627,0645") & "." & vbCrLf & _
            AT_U("062A,0645") & " " & AT_U("0625,0646,0634,0627,0621") & " " & AT_U("0627,0644,062C,062F,0627,0648,0644") & " " & AT_U("0648,0627,0644,0627,0633,062A,0639,0644,0627,0645,0627,062A") & " " & AT_U("0648,0627,0644,0646,0645,0627,0630,062C") & " " & AT_U("0648,0627,0644,062A,0642,0627,0631,064A,0631") & " " & AT_U("0627,0644,0623,0633,0627,0633,064A,0629") & "." & vbCrLf & _
            AT_U("0627,0644,062E,0637,0648,0629") & " " & AT_U("0627,0644,062A,0627,0644,064A,0629") & ": " & AT_U("0627,0641,062A,062D") & " frmAT_Home " & AT_U("0648,0627,0628,062F,0623") & " " & AT_U("0627,0644,0627,062E,062A,0628,0627,0631,0627,062A") & " " & AT_U("0627,0644,062A,062C,0631,064A,0628,064A,0629") & " " & AT_U("0642,0628,0644") & " " & AT_U("0625,062F,062E,0627,0644") & " " & AT_U("0628,064A,0627,0646,0627,062A") & " " & AT_U("062D,0642,064A,0642,064A,0629") & ".", vbInformation, "Archon Tech"
        DoCmd.OpenForm "frmAT_Home"
    Else
        MsgBox AT_U("0627,0646,062A,0647,0649") & " " & AT_U("0625,062C,0631,0627,0621") & " " & AT_U("0627,0644,0628,0646,0627,0621") & " " & AT_U("0645,0639") & " " & AT_U("0648,062C,0648,062F") & " " & failedChecks & " " & AT_U("0639,0646,0627,0635,0631") & " " & AT_U("0641,0627,0634,0644,0629") & " " & AT_U("0641,064A") & " " & AT_U("0627,0644,0641,062D,0635") & " " & AT_U("0627,0644,0628,0646,064A,0648,064A") & ". " & AT_U("0631,0627,062C,0639") & " tblAT_BuildLog.", vbExclamation, "Archon Tech"
    End If
    Exit Sub

BuildFail:
    DoCmd.SetWarnings True
    AT_LogBuild "Build", "Failed", Err.Number & " - " & Err.Description
    MsgBox AT_U("062A,0648,0642,0641") & " " & AT_U("0627,0644,0628,0646,0627,0621") & " " & AT_U("0628,0633,0628,0628") & " " & AT_U("062E,0637,0623") & ": " & Err.Number & " - " & Err.Description & vbCrLf & _
        AT_U("0644,0627") & " " & AT_U("062A,062D,0630,0641") & " " & AT_U("0627,0644,062C,062F,0627,0648,0644") & ". " & AT_U("0623,0631,0633,0644") & " " & AT_U("0631,0633,0627,0644,0629") & " " & AT_U("0627,0644,062E,0637,0623") & " " & AT_U("0648,0633,062C,0644") & " tblAT_BuildLog " & AT_U("0644,0644,0645,0631,0627,062C,0639,0629") & ".", vbCritical, "Archon Tech"
End Sub

Public Sub OpenArchonHome()
    DoCmd.OpenForm "frmAT_Home"
End Sub

Public Sub RunArchonStructuralTests()
    Call AT_RunSelfTests(True)
End Sub
