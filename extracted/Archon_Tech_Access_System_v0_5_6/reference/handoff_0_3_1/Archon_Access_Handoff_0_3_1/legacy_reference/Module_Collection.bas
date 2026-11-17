Attribute VB_Name = "Module_Collection"
Option Explicit

Private Const OWNER_FIRST As Long = 5
Private Const OWNER_LAST As Long = 204
Private Const LEDGER_FIRST As Long = 5
Private Const LEDGER_LAST As Long = 504
Private Const AUDIT_LAST As Long = 504
Private Const SHEET_PASSWORD As String = "Sewer2026"

Public Function FindOwnerRow(ByVal ownerID As String) As Long
    Dim ws As Worksheet, foundCell As Range
    If Len(Trim$(ownerID)) = 0 Then Exit Function
    Set ws = ThisWorkbook.Worksheets("”Ã· «·„·«ﬂ")
    Set foundCell = ws.Range("A" & OWNER_FIRST & ":A" & OWNER_LAST).Find( _
        What:=Trim$(ownerID), LookIn:=xlValues, LookAt:=xlWhole, _
        SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If Not foundCell Is Nothing Then FindOwnerRow = foundCell.Row
End Function

Public Function FindMovementRow(ByVal movementID As String) As Long
    Dim ws As Worksheet, foundCell As Range
    If Len(Trim$(movementID)) = 0 Then Exit Function
    Set ws = ThisWorkbook.Worksheets("”Ã· «· Õ’Ì·")
    Set foundCell = ws.Range("A" & LEDGER_FIRST & ":A" & LEDGER_LAST).Find( _
        What:=Trim$(movementID), LookIn:=xlValues, LookAt:=xlWhole, _
        SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If Not foundCell Is Nothing Then FindMovementRow = foundCell.Row
End Function

Public Function IssueNextOwnerID() As String
    Dim wsO As Worksheet, wsS As Worksheet, c As Range
    Dim seq As Long, foundSeq As Long, s As String, oldEvents As Boolean
    On Error GoTo Failed
    oldEvents = Application.EnableEvents
    Set wsO = ThisWorkbook.Worksheets("”Ã· «·„·«ﬂ")
    Set wsS = ThisWorkbook.Worksheets("«·≈⁄œ«œ«  Êœ·Ì· «·«” Œœ«„")
    seq = CLng(Val(wsS.Range("B10").Value2))
    For Each c In wsO.Range("A" & OWNER_FIRST & ":A" & OWNER_LAST).Cells
        s = CStr(c.Value2)
        If Left$(s, 3) = "OW-" And IsNumeric(Mid$(s, 4)) Then
            foundSeq = CLng(Mid$(s, 4))
            If foundSeq > seq Then seq = foundSeq
        End If
    Next c
    seq = seq + 1
    Application.EnableEvents = False
    wsS.Range("B10").Value2 = seq
    Application.EnableEvents = oldEvents
    IssueNextOwnerID = "OW-" & Format$(seq, "0000")
    Exit Function
Failed:
    Application.EnableEvents = oldEvents
    Err.Raise Err.Number, "IssueNextOwnerID", Err.Description
End Function

Public Function IssueNextReceiptID() As String
    Dim wsL As Worksheet, wsS As Worksheet, c As Range
    Dim seq As Long, foundSeq As Long, s As String, oldEvents As Boolean
    On Error GoTo Failed
    oldEvents = Application.EnableEvents
    Set wsL = ThisWorkbook.Worksheets("”Ã· «· Õ’Ì·")
    Set wsS = ThisWorkbook.Worksheets("«·≈⁄œ«œ«  Êœ·Ì· «·«” Œœ«„")
    seq = CLng(Val(wsS.Range("B11").Value2))
    For Each c In wsL.Range("A" & LEDGER_FIRST & ":A" & LEDGER_LAST).Cells
        s = CStr(c.Value2)
        If Left$(s, 3) = "RC-" And IsNumeric(Mid$(s, 4)) Then
            foundSeq = CLng(Mid$(s, 4))
            If foundSeq > seq Then seq = foundSeq
        End If
    Next c
    seq = seq + 1
    Application.EnableEvents = False
    wsS.Range("B11").Value2 = seq
    Application.EnableEvents = oldEvents
    IssueNextReceiptID = "RC-" & Format$(seq, "000000")
    Exit Function
Failed:
    Application.EnableEvents = oldEvents
    Err.Raise Err.Number, "IssueNextReceiptID", Err.Description
End Function

Private Function NextLedgerRow() As Long
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("”Ã· «· Õ’Ì·")
    For r = LEDGER_FIRST To LEDGER_LAST
        If Len(Trim$(CStr(ws.Cells(r, 1).Value2))) = 0 Then
            NextLedgerRow = r
            Exit Function
        End If
    Next r
End Function

Private Sub ClearLedgerEntry(ByVal r As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("”Ã· «· Õ’Ì·")
    ws.Range("A" & r & ":D" & r).ClearContents
    ws.Range("G" & r & ":M" & r).ClearContents
    ws.Range("N" & r & ":W" & r).ClearContents
End Sub

Public Sub AppendAudit(ByVal operationType As String, _
    Optional ByVal ownerID As String = "", Optional ByVal movementID As String = "", _
    Optional ByVal previousValue As String = "", Optional ByVal newValue As String = "", _
    Optional ByVal reason As String = "")
    Dim ws As Worksheet, tbl As ListObject, r As Long, i As Long
    Dim oldEvents As Boolean, added As ListRow, errorText As String

    On Error GoTo Failed
    oldEvents = Application.EnableEvents
    Set ws = ThisWorkbook.Worksheets("”Ã· «· œﬁÌﬁ")
    Application.EnableEvents = False

    r = 0
    For i = 5 To AUDIT_LAST
        If Len(Trim$(CStr(ws.Cells(i, 1).Value2))) = 0 Then
            r = i
            Exit For
        End If
    Next i

    If r = 0 Then
        Set tbl = ws.ListObjects("AuditTable")
        Set added = tbl.ListRows.Add
        r = added.Range.Row
    End If

    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = operationType
    ws.Cells(r, 3).Value = ownerID
    ws.Cells(r, 4).Value = movementID
    ws.Cells(r, 5).Value = Application.UserName
    ws.Cells(r, 6).Value = previousValue
    ws.Cells(r, 7).Value = newValue
    ws.Cells(r, 8).Value = reason

    Application.EnableEvents = oldEvents
    Exit Sub

Failed:
    errorText = Err.Description
    On Error Resume Next
    Application.EnableEvents = oldEvents
    On Error GoTo 0
    MsgBox " Õ–Ì—:  ⁄–—  ”ÃÌ· «·⁄„·Ì… ›Ì ´”Ã· «· œﬁÌﬁª: " & errorText & vbCrLf & _
        " Õﬁﬁ „‰ ”Ã· «· œﬁÌﬁ ﬁ»· „ «»⁄… √Ì ≈Ã—«¡ „«·Ì.", vbCritical
End Sub

Public Sub SavePaymentAndIssueReceipt()
    Dim wsF As Worksheet, wsO As Worksheet, wsL As Worksheet
    Dim ownerID As String, paymentMethod As String, payer As String
    Dim collector As String, noteText As String, receiptID As String
    Dim ownerRow As Long, targetRow As Long, amount As Double
    Dim dueAmount As Double, paidBefore As Double, paidAfter As Double
    Dim remainingBefore As Double, overageAfter As Double
    Dim oldEvents As Boolean, validationText As String, errorText As String
    On Error GoTo Failed
    oldEvents = Application.EnableEvents
    Set wsF = ThisWorkbook.Worksheets(" ”ÃÌ· «·œ›⁄… Ê«·≈Ì’«·")
    Set wsO = ThisWorkbook.Worksheets("”Ã· «·„·«ﬂ")
    Set wsL = ThisWorkbook.Worksheets("”Ã· «· Õ’Ì·")

    ownerID = Trim$(CStr(wsF.Range("B4").Value2))
    If ownerID = "" Then
        MsgBox "√œŒ· „⁄—› «·„«·ﬂ √Ê·«.", vbExclamation
        Exit Sub
    End If
    ownerRow = FindOwnerRow(ownerID)
    If ownerRow = 0 Then
        MsgBox "„⁄—› «·„«·ﬂ €Ì— „ÊÃÊœ ›Ì ”Ã· «·„·«ﬂ. ·„  ıÕ›Ÿ «·œ›⁄….", vbCritical
        Exit Sub
    End If
    If CStr(wsO.Cells(ownerRow, 13).Value2) = "„ÊﬁÊ›" Then
        MsgBox "Â–« «·”Ã· „ÊﬁÊ›. ·« Ì„ﬂ‰  ”ÃÌ· ﬁ»÷ ÃœÌœ ·Â ﬁ»· „—«Ã⁄… «·„”ƒÊ·.", vbExclamation
        Exit Sub
    End If
    If Len(Trim$(CStr(wsO.Cells(ownerRow, 2).Value2))) = 0 _
        Or Val(wsO.Cells(ownerRow, 4).Value2) <= 0 _
        Or Val(wsO.Cells(ownerRow, 5).Value2) <= 0 _
        Or Val(wsO.Cells(ownerRow, 6).Value2) <= 0 Then
        MsgBox "»Ì«‰«  «·„«·ﬂ ‰«ﬁ’… √Ê «·„” Õﬁ €Ì— ’«·Õ. —«Ã⁄ ”Ã· «·„·«ﬂ ﬁ»· «· Õ’Ì·.", vbCritical
        Exit Sub
    End If
    If Not IsDate(wsF.Range("B12").Value) Then
        MsgBox " «—ÌŒ «·œ›⁄… €Ì— ’«·Õ.", vbExclamation
        Exit Sub
    End If
    If Not IsNumeric(wsF.Range("B13").Value2) Then
        MsgBox "√œŒ· „»·€« —ﬁ„Ì« „ÊÃ»«.", vbExclamation
        Exit Sub
    End If
    amount = CDbl(wsF.Range("B13").Value2)
    If amount <= 0 Then
        MsgBox "·« Ìı”„Õ »„»·€ ’›—Ì √Ê ”«·» ›Ì Õ—ﬂ… ﬁ»÷.", vbExclamation
        Exit Sub
    End If
    paymentMethod = Trim$(CStr(wsF.Range("B14").Value2))
    payer = Trim$(CStr(wsF.Range("B15").Value2))
    collector = Trim$(CStr(wsF.Range("B16").Value2))
    noteText = Trim$(CStr(wsF.Range("B17").Value2))
    If paymentMethod = "" Or payer = "" Or collector = "" Then
        MsgBox "√ﬂ„· ÿ—Ìﬁ… «·œ›⁄ Ê«”„ «·œ«›⁄ Ê«”„ «·„Õ’·.", vbExclamation
        Exit Sub
    End If
    If Application.WorksheetFunction.CountIf( _
       ThisWorkbook.Worksheets("«·≈⁄œ«œ«  Êœ·Ì· «·«” Œœ«„").Range("D4:D7"), paymentMethod) = 0 Then
        MsgBox "ÿ—Ìﬁ… «·œ›⁄ €Ì— „ÊÃÊœ… ›Ì ﬁ«∆„… «·≈⁄œ«œ« .", vbExclamation
        Exit Sub
    End If

    dueAmount = CDbl(wsO.Cells(ownerRow, 6).Value2)
    paidBefore = CDbl(Val(wsO.Cells(ownerRow, 7).Value2))
    remainingBefore = Application.WorksheetFunction.Max(0, dueAmount - paidBefore)
    paidAfter = paidBefore + amount
    overageAfter = Application.WorksheetFunction.Max(0, paidAfter - dueAmount)
    If amount > remainingBefore Then
        If MsgBox("Â–Â «·œ›⁄… √ﬂ»— „‰ «·„ »ﬁÌ «·Õ«·Ì." & vbCrLf & _
            "«·“Ì«œ… «·„ Êﬁ⁄…: " & Format$(overageAfter, "#,##0") & " —Ì«·." & vbCrLf & _
            "Â·  —Ìœ «·„ «»⁄… Ê ”ÃÌ· «·“Ì«œ… ’—«Õ…ø", _
            vbYesNo + vbExclamation, " √ﬂÌœ “Ì«œ… ⁄·Ï «·„” Õﬁ") <> vbYes Then Exit Sub
    End If

    targetRow = NextLedgerRow()
    If targetRow = 0 Then
        MsgBox "Ê’· ”Ã· «· Õ’Ì· ≈·Ï 500 Õ—ﬂ…. ·«  ”Ã· ÌœÊÌ«∫ Ê”¯⁄ «·ﬁ«·» Ê«·’Ì€ Ê«·ﬂÊœ ﬁ»· «·„ «»⁄….", vbCritical
        Exit Sub
    End If
    receiptID = IssueNextReceiptID()
    Application.EnableEvents = False
    With wsL
        .Cells(targetRow, 1).Value2 = receiptID
        .Cells(targetRow, 2).Value = CDate(wsF.Range("B12").Value)
        .Cells(targetRow, 3).Value2 = "ﬁ»÷"
        .Cells(targetRow, 4).Value2 = ownerID
        .Cells(targetRow, 7).Value2 = amount
        .Cells(targetRow, 8).Value2 = paymentMethod
        .Cells(targetRow, 9).Value2 = payer
        .Cells(targetRow, 10).Value2 = collector
        .Cells(targetRow, 11).Value2 = "„⁄ „œ"
        .Cells(targetRow, 12).ClearContents
        .Cells(targetRow, 13).Value2 = noteText
        .Cells(targetRow, 14).Value = Now
        .Range("O" & targetRow & ":W" & targetRow).Value = Array( _
            dueAmount, paidBefore, paidAfter, _
            Application.WorksheetFunction.Max(0, dueAmount - paidAfter), _
            overageAfter, CLng(wsO.Cells(ownerRow, 4).Value2), _
            CLng(wsO.Cells(ownerRow, 5).Value2), _
            CStr(wsO.Cells(ownerRow, 2).Value2), _
            CStr(wsO.Cells(ownerRow, 3).Text))
    End With
    Application.CalculateFull
    validationText = CStr(wsL.Cells(targetRow, 24).Value2)
    If validationText <> "”·Ì„" Then
        ClearLedgerEntry targetRow
        Application.CalculateFull
        Application.EnableEvents = oldEvents
        MsgBox "·„  ⁄ „œ «·Õ—ﬂ… ·√‰ ›Õ’ «·»Ì«‰«  ·„ Ì‰ÃÕ: " & validationText & vbCrLf & _
            "·„ Ìı’œ— ≈Ì’«· ‰Ã«Õ Ê·„  ı„”Õ «·Õ—ﬂ«  «·”«»ﬁ….", vbCritical
        Exit Sub
    End If

    AppendAudit " ”ÃÌ· ﬁ»÷", ownerID, receiptID, "", Format$(amount, "#,##0"), noteText
    wsF.Range("F4").Value2 = receiptID
    wsF.Range("B13").ClearContents
    Application.EnableEvents = oldEvents
    If MsgBox(" „ Õ›Ÿ «·œ›⁄… »‰Ã«Õ." & vbCrLf & "—ﬁ„ «·≈Ì’«·: " & receiptID & vbCrLf & _
        "Â·  —Ìœ „⁄«Ì‰… «·≈Ì’«· ··ÿ»«⁄…ø", vbYesNo + vbInformation, " „ «·Õ›Ÿ") = vbYes Then
        PrintCurrentReceipt
    End If
    Exit Sub

Failed:
    errorText = Err.Description
    On Error Resume Next
    If targetRow >= LEDGER_FIRST And targetRow <= LEDGER_LAST Then
        If CStr(wsL.Cells(targetRow, 24).Value2) <> "”·Ì„" Then ClearLedgerEntry targetRow
    End If
    Application.EnableEvents = oldEvents
    MsgBox " ⁄–— ≈ﬂ„«· «·Õ›Ÿ: " & errorText & vbCrLf & _
        " Õﬁﬁ „‰ ”Ã· «· Õ’Ì· ﬁ»· ≈⁄«œ… «·„Õ«Ê·…∫ ·«  ⁄ˆœ ≈œŒ«· «·„»·€ ≈–« ŸÂ— —ﬁ„ RC „Õ›ÊŸ.", vbCritical
End Sub

Public Sub PostReversalOrAdjustment()
    Dim wsL As Worksheet, wsO As Worksheet, wsS As Worksheet
    Dim sourceID As String, movementType As String, choice As String
    Dim reasonText As String, dateText As String, methodText As String, recipientText As String, amountText As String
    Dim sourceRow As Long, ownerRow As Long, targetRow As Long
    Dim amount As Double, signedAmount As Double, sourceAmount As Double
    Dim reversedBefore As Double, dueAmount As Double, paidBefore As Double, paidAfter As Double
    Dim receiptID As String, oldEvents As Boolean, errorText As String
    On Error GoTo Failed
    oldEvents = Application.EnableEvents
    Set wsL = ThisWorkbook.Worksheets("”Ã· «· Õ’Ì·")
    Set wsO = ThisWorkbook.Worksheets("”Ã· «·„·«ﬂ")
    Set wsS = ThisWorkbook.Worksheets("«·≈⁄œ«œ«  Êœ·Ì· «·«” Œœ«„")

    sourceID = Trim$(InputBox("√œŒ· —ﬁ„ «·Õ—ﬂ… «·√’·Ì… «·„⁄ „œ… (RC-xxxxxx):", "⁄ﬂ” /  ’ÕÌÕ Õ—ﬂ…"))
    If sourceID = "" Then Exit Sub
    sourceRow = FindMovementRow(sourceID)
    If sourceRow = 0 Then
        MsgBox "—ﬁ„ «·Õ—ﬂ… «·√’·Ì… €Ì— „ÊÃÊœ.", vbExclamation
        Exit Sub
    End If
    If CStr(wsL.Cells(sourceRow, 11).Value2) <> "„⁄ „œ" Or CStr(wsL.Cells(sourceRow, 24).Value2) <> "”·Ì„" Then
        MsgBox "·« Ì„ﬂ‰ «· ’ÕÌÕ »«·—ÃÊ⁄ ≈·Ï Õ—ﬂ… €Ì— „⁄ „œ… √Ê €Ì— ”·Ì„….", vbCritical
        Exit Sub
    End If

    choice = Trim$(InputBox("«Œ — ‰Ê⁄ «·⁄„·Ì…:" & vbCrLf & _
        "1 = ⁄ﬂ” ﬁ»÷ / «” —œ«œ" & vbCrLf & _
        "2 =  ”ÊÌ…  ’ÕÌÕÌ… „ÊÃ»…" & vbCrLf & _
        "3 =  ”ÊÌ…  ’ÕÌÕÌ… ”«·»…", "‰Ê⁄ «· ’ÕÌÕ", "1"))
    Select Case choice
        Case "1"
            If CStr(wsL.Cells(sourceRow, 3).Value2) <> "ﬁ»÷" Then
                MsgBox "⁄ﬂ” «·ﬁ»÷ ÌÃ» √‰ Ì‘Ì— ≈·Ï Õ—ﬂ… ﬁ»÷ √’·Ì….", vbExclamation
                Exit Sub
            End If
            movementType = "⁄ﬂ” ﬁ»÷ / «” —œ«œ"
        Case "2": movementType = " ”ÊÌ…  ’ÕÌÕÌ… „ÊÃ»…"
        Case "3": movementType = " ”ÊÌ…  ’ÕÌÕÌ… ”«·»…"
        Case Else: Exit Sub
    End Select

    amountText = Trim$(InputBox("√œŒ· „»·€ «· ’ÕÌÕ „ÊÃ»«:", "„»·€ «· ’ÕÌÕ", "0"))
    If amountText = "" Then Exit Sub
    If Not IsNumeric(amountText) Then
        MsgBox "«·„»·€ €Ì— ’«·Õ.", vbExclamation
        Exit Sub
    End If
    amount = CDbl(amountText)
    If amount <= 0 Then
        MsgBox "«·„»·€ ÌÃ» √‰ ÌﬂÊ‰ √ﬂ»— „‰ ’›—.", vbExclamation
        Exit Sub
    End If
    reasonText = Trim$(InputBox("«ﬂ » ”»» «·⁄ﬂ” √Ê «· ”ÊÌ… (≈·“«„Ì):", "”»» «· ’ÕÌÕ"))
    If reasonText = "" Then
        MsgBox "ÌÃ»  ÊÀÌﬁ «·”»».", vbExclamation
        Exit Sub
    End If
    dateText = InputBox(" «—ÌŒ «·Õ—ﬂ… („À«· 2026-10-16):", " «—ÌŒ «· ’ÕÌÕ", Format$(Date, "yyyy-mm-dd"))
    If Not IsDate(dateText) Then
        MsgBox "«· «—ÌŒ €Ì— ’«·Õ.", vbExclamation
        Exit Sub
    End If

    ownerRow = FindOwnerRow(CStr(wsL.Cells(sourceRow, 4).Value2))
    If ownerRow = 0 Then
        MsgBox "„«·ﬂ «·Õ—ﬂ… «·√’·Ì… €Ì— „ÊÃÊœ∫ √Êﬁ› «·⁄„·Ì… Ê—«Ã⁄ «·”Ã·« .", vbCritical
        Exit Sub
    End If
    dueAmount = CDbl(Val(wsO.Cells(ownerRow, 6).Value2))
    paidBefore = CDbl(Val(wsO.Cells(ownerRow, 7).Value2))
    sourceAmount = CDbl(Val(wsL.Cells(sourceRow, 7).Value2))

    If movementType = "⁄ﬂ” ﬁ»÷ / «” —œ«œ" Then
        reversedBefore = Application.WorksheetFunction.SumIfs( _
            wsL.Range("G5:G504"), wsL.Range("L5:L504"), sourceID, _
            wsL.Range("C5:C504"), "⁄ﬂ” ﬁ»÷ / «” —œ«œ", _
            wsL.Range("K5:K504"), "„⁄ „œ", wsL.Range("X5:X504"), "”·Ì„")
        If amount > sourceAmount - reversedBefore Or amount > paidBefore Then
            MsgBox "„»·€ «·«” —œ«œ Ì Ã«Ê“ «·Ã“¡ «·ﬁ«»· ··⁄ﬂ” √Ê ’«›Ì «·„”œœ «·Õ«·Ì.", vbCritical
            Exit Sub
        End If
        methodText = Trim$(InputBox("ÿ—Ìﬁ… —œ «·„»·€ (‰ﬁœ« /  ÕÊÌ· »‰ﬂÌ / „Õ›Ÿ… ≈·ﬂ —Ê‰Ì… / √Œ—Ï):", _
            "ÿ—Ìﬁ… «·«” —œ«œ", "‰ﬁœ«"))
        If Application.WorksheetFunction.CountIf(wsS.Range("D4:D7"), methodText) = 0 Then
            MsgBox "ÿ—Ìﬁ… «·«” —œ«œ €Ì— „ÊÃÊœ… ›Ì ﬁ«∆„… «·≈⁄œ«œ« .", vbExclamation
            Exit Sub
        End If
        signedAmount = -amount
    ElseIf movementType = " ”ÊÌ…  ’ÕÌÕÌ… ”«·»…" Then
        If amount > paidBefore Then
            MsgBox "«· ”ÊÌ… «·”«·»… ” Ã⁄· ’«›Ì «·„”œœ ”«·»«∫ √Êﬁ› «·⁄„·Ì… Ê—«Ã⁄ «·„”ƒÊ·.", vbCritical
            Exit Sub
        End If
        methodText = "√Œ—Ï"
        signedAmount = -amount
    Else
        methodText = "√Œ—Ï"
        signedAmount = amount
        If paidBefore + amount > dueAmount Then
            If MsgBox("” ƒœÌ «· ”ÊÌ… «·„ÊÃ»… ≈·Ï “Ì«œ… ﬁœ—Â« " & _
                Format$(Application.WorksheetFunction.Max(0, paidBefore + amount - dueAmount), "#,##0") & _
                " —Ì«· ›Êﬁ «·„” Õﬁ. Â·  —Ìœ «·«” „—«—ø", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
        End If
    End If

    recipientText = Trim$(InputBox("«”„ «·‘Œ’ / «·ÃÂ… «·„— »ÿ… »«· ’ÕÌÕ:", "»Ì«‰«  «· ’ÕÌÕ", _
        CStr(wsO.Cells(ownerRow, 2).Value2)))
    If recipientText = "" Then recipientText = CStr(wsO.Cells(ownerRow, 2).Value2)
    paidAfter = paidBefore + signedAmount
    If paidAfter < 0 Then
        MsgBox "·« Ì„ﬂ‰ «⁄ „«œ ’«›Ì „”œœ ”«·».", vbCritical
        Exit Sub
    End If
    targetRow = NextLedgerRow()
    If targetRow = 0 Then
        MsgBox "Ê’· ”Ã· «· Õ’Ì· ≈·Ï 500 Õ—ﬂ…. Ê”¯⁄ «·ﬁ«·» Ê«·’Ì€ Ê«·ﬂÊœ ﬁ»· «·„ «»⁄….", vbCritical
        Exit Sub
    End If
    receiptID = IssueNextReceiptID()
    Application.EnableEvents = False
    With wsL
        .Cells(targetRow, 1).Value2 = receiptID
        .Cells(targetRow, 2).Value = CDate(dateText)
        .Cells(targetRow, 3).Value2 = movementType
        .Cells(targetRow, 4).Value2 = CStr(wsL.Cells(sourceRow, 4).Value2)
        .Cells(targetRow, 7).Value2 = amount
        .Cells(targetRow, 8).Value2 = methodText
        .Cells(targetRow, 9).Value2 = recipientText
        .Cells(targetRow, 10).Value2 = Application.UserName
        .Cells(targetRow, 11).Value2 = "„⁄ „œ"
        .Cells(targetRow, 12).Value2 = sourceID
        .Cells(targetRow, 13).Value2 = reasonText
        .Cells(targetRow, 14).Value = Now
        .Range("O" & targetRow & ":W" & targetRow).Value = Array( _
            dueAmount, paidBefore, paidAfter, _
            Application.WorksheetFunction.Max(0, dueAmount - paidAfter), _
            Application.WorksheetFunction.Max(0, paidAfter - dueAmount), _
            CLng(wsO.Cells(ownerRow, 4).Value2), CLng(wsO.Cells(ownerRow, 5).Value2), _
            CStr(wsO.Cells(ownerRow, 2).Value2), CStr(wsO.Cells(ownerRow, 3).Text))
    End With
    Application.CalculateFull
    If CStr(wsL.Cells(targetRow, 24).Value2) <> "”·Ì„" Then
        ClearLedgerEntry targetRow
        Application.CalculateFull
        Application.EnableEvents = oldEvents
        MsgBox "·„ Ì „ «⁄ „«œ «· ’ÕÌÕ ·√‰ ›Õ’ «·Õ—ﬂ… ·„ Ì‰ÃÕ.", vbCritical
        Exit Sub
    End If
    AppendAudit "⁄ﬂ” /  ”ÊÌ… „«·Ì…", CStr(wsL.Cells(targetRow, 4).Value2), receiptID, sourceID, _
        Format$(signedAmount, "#,##0"), reasonText
    Application.EnableEvents = oldEvents
    MsgBox " „  ”ÃÌ· «·Õ—ﬂ… «·„— »ÿ…: " & receiptID & vbCrLf & _
        "«·„—Ã⁄ «·√’·Ì: " & sourceID & vbCrLf & "«·√À— ⁄·Ï «·Õ”«»: " & Format$(signedAmount, "#,##0") & " —Ì«·.", _
        vbInformation
    Exit Sub

Failed:
    errorText = Err.Description
    On Error Resume Next
    If targetRow >= LEDGER_FIRST And targetRow <= LEDGER_LAST Then
        If CStr(wsL.Cells(targetRow, 24).Value2) <> "”·Ì„" Then ClearLedgerEntry targetRow
    End If
    Application.EnableEvents = oldEvents
    MsgBox " ⁄–—  ‰›Ì– «· ’ÕÌÕ: " & errorText & ". —«Ã⁄ ”Ã· «· Õ’Ì· ﬁ»· «·„Õ«Ê·… „—… √Œ—Ï.", vbCritical
End Sub

Public Sub CancelPendingMovement()
    Dim wsL As Worksheet, movementID As String, reasonText As String
    Dim r As Long, oldEvents As Boolean, oldNotes As String
    On Error GoTo Failed
    oldEvents = Application.EnableEvents
    Set wsL = ThisWorkbook.Worksheets("”Ã· «· Õ’Ì·")
    movementID = Trim$(InputBox("√œŒ· —ﬁ„ «·Õ—ﬂ… €Ì— «·„⁄ „œ… «·„—«œ ≈·€«ƒÂ«:", "≈·€«¡ Õ—ﬂ… „⁄·ﬁ…"))
    If movementID = "" Then Exit Sub
    r = FindMovementRow(movementID)
    If r = 0 Then
        MsgBox "«·Õ—ﬂ… €Ì— „ÊÃÊœ….", vbExclamation
        Exit Sub
    End If
    If CStr(wsL.Cells(r, 11).Value2) <> "ﬁÌœ «·„—«Ã⁄…" Then
        MsgBox "Ì„ﬂ‰ ≈·€«¡ «·Õ—ﬂ… «· Ì Õ«· Â« ´ﬁÌœ «·„—«Ã⁄…ª ›ﬁÿ. ·« Ì„ﬂ‰ ≈·€«¡ Õ—ﬂ… „⁄ „œ… √Ê ”»ﬁ ≈·€«ƒÂ«.", vbCritical
        Exit Sub
    End If
    reasonText = Trim$(InputBox("”»» «·≈·€«¡ (≈·“«„Ì):", "”»» «·≈·€«¡"))
    If reasonText = "" Then Exit Sub
    oldNotes = CStr(wsL.Cells(r, 13).Value2)
    Application.EnableEvents = False
    wsL.Cells(r, 11).Value2 = "„·€«…"
    If oldNotes <> "" Then oldNotes = oldNotes & " | "
    wsL.Cells(r, 13).Value2 = oldNotes & "”»» «·≈·€«¡: " & reasonText
    Application.CalculateFull
    AppendAudit "≈·€«¡ Õ—ﬂ… €Ì— „⁄ „œ…", CStr(wsL.Cells(r, 4).Value2), movementID, oldNotes, _
        CStr(wsL.Cells(r, 11).Value2), reasonText
    Application.EnableEvents = oldEvents
    MsgBox " „  ⁄·Ì„ «·Õ—ﬂ… ﬂ„·€«… „⁄ «·«Õ ›«Ÿ »”Ã·Â«.", vbInformation
    Exit Sub
Failed:
    Application.EnableEvents = oldEvents
    MsgBox " ⁄–— ≈·€«¡ «·Õ—ﬂ…: " & Err.Description, vbCritical
End Sub

Public Sub PrintCurrentReceipt()
    Dim wsF As Worksheet, wsL As Worksheet
    Dim receiptID As String, r As Long
    Dim errorText As String

    On Error GoTo Failed
    Set wsF = ThisWorkbook.Worksheets(" ”ÃÌ· «·œ›⁄… Ê«·≈Ì’«·")
    Set wsL = ThisWorkbook.Worksheets("”Ã· «· Õ’Ì·")

    receiptID = Trim$(CStr(wsF.Range("F4").Value2))
    If receiptID = "" Then
        MsgBox "√œŒ· —ﬁ„ «·≈Ì’«· RC-xxxxxx ›Ì «·Œ«‰… F4 √Ê·«.", vbExclamation
        Exit Sub
    End If

    r = FindMovementRow(receiptID)
    If r = 0 Then
        MsgBox "—ﬁ„ «·≈Ì’«· €Ì— „ÊÃÊœ. ·„ Ì „ ≈‰‘«¡ Õ—ﬂ… ÃœÌœ….", vbExclamation
        Exit Sub
    End If

    If CStr(wsL.Cells(r, 11).Value2) <> "„⁄ „œ" _
        Or CStr(wsL.Cells(r, 24).Value2) <> "”·Ì„" Then
        MsgBox "·« Ì„ﬂ‰ „⁄«Ì‰… ≈Ì’«· ·Õ—ﬂ… €Ì— „⁄ „œ… √Ê €Ì— ”·Ì„….", vbCritical
        Exit Sub
    End If

    wsF.Activate
    wsF.PrintPreview
    Exit Sub

Failed:
    errorText = Err.Description
    MsgBox " ⁄–— › Õ „⁄«Ì‰… «·≈Ì’«·: " & errorText, vbCritical
End Sub

Public Sub ConfigureCollectionProtection()
    Dim ws As Worksheet

    ' UserInterfaceOnly must be applied on every workbook open.
    ' Unprotect every worksheet before changing locked cells/protection options.
    For Each ws In ThisWorkbook.Worksheets
        If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then
            ws.Unprotect Password:=SHEET_PASSWORD
        End If
    Next ws

    With ThisWorkbook.Worksheets("”Ã· «·„·«ﬂ")
        .Cells.Locked = True
        .Range("B5:E204,L5:M204").Locked = False
        .Protect Password:=SHEET_PASSWORD, UserInterfaceOnly:=True, _
            AllowFiltering:=True, AllowSorting:=True
    End With

    With ThisWorkbook.Worksheets("”Ã· «· Õ’Ì·")
        .Cells.Locked = True
        .Range("B5:D504,G5:J504,L5:M504").Locked = False
        .Protect Password:=SHEET_PASSWORD, UserInterfaceOnly:=True, _
            AllowFiltering:=True, AllowSorting:=True
    End With

    With ThisWorkbook.Worksheets(" ”ÃÌ· «·œ›⁄… Ê«·≈Ì’«·")
        .Cells.Locked = True
        .Range("B4,B12:B17,F4").Locked = False
        .Protect Password:=SHEET_PASSWORD, UserInterfaceOnly:=True
    End With

    With ThisWorkbook.Worksheets("«·≈⁄œ«œ«  Êœ·Ì· «·«” Œœ«„")
        .Cells.Locked = True
        .Range("B4:B5,B7:B8").Locked = False
        .Protect Password:=SHEET_PASSWORD, UserInterfaceOnly:=True
    End With

    With ThisWorkbook.Worksheets("«· ﬁ«—Ì— Ê«·≈ﬁ›«·")
        .Cells.Locked = True
        .Range("B3:B4,B34,B37:B39,B41,B43").Locked = False
        .Protect Password:=SHEET_PASSWORD, UserInterfaceOnly:=True, _
            AllowFiltering:=True, AllowSorting:=True
    End With

    With ThisWorkbook.Worksheets("·ÊÕ… «· Õﬂ„")
        .Cells.Locked = True
        .Protect Password:=SHEET_PASSWORD, UserInterfaceOnly:=True
    End With

    With ThisWorkbook.Worksheets("”Ã· «· œﬁÌﬁ")
        .Cells.Locked = True
        .Protect Password:=SHEET_PASSWORD, UserInterfaceOnly:=True, _
            AllowFiltering:=True, AllowSorting:=True
    End With
End Sub

Private Sub CreateCollectionButtonsCore()
    Dim wsF As Worksheet, wsL As Worksheet, i As Long
    Dim btn As Button

    Set wsF = ThisWorkbook.Worksheets(" ”ÃÌ· «·œ›⁄… Ê«·≈Ì’«·")
    Set wsL = ThisWorkbook.Worksheets("”Ã· «· Õ’Ì·")

    ' This macro may be run manually while the sheets are protected.
    wsF.Unprotect Password:=SHEET_PASSWORD
    wsL.Unprotect Password:=SHEET_PASSWORD

    ' Remove only this project's existing buttons, avoiding duplicates.
    For i = wsF.Buttons.Count To 1 Step -1
        If wsF.Buttons(i).Caption = "Õ›Ÿ «·œ›⁄… Ê≈’œ«— «·≈Ì’«·" _
            Or wsF.Buttons(i).Caption = "„⁄«Ì‰… / ÿ»«⁄… «·≈Ì’«·" Then
            wsF.Buttons(i).Delete
        End If
    Next i

    For i = wsL.Buttons.Count To 1 Step -1
        If wsL.Buttons(i).Caption = "⁄ﬂ” /  ”ÊÌ… „ÊÀﬁ…" _
            Or wsL.Buttons(i).Caption = "≈·€«¡ Õ—ﬂ… €Ì— „⁄ „œ…" Then
            wsL.Buttons(i).Delete
        End If
    Next i

    ' Change row heights only after removing protection.
    wsF.Rows(20).RowHeight = 30
    wsL.Rows(2).RowHeight = 32

    Set btn = wsF.Buttons.Add(wsF.Range("A20").Left, wsF.Range("A20").Top, 220, 28)
    btn.Caption = "Õ›Ÿ «·œ›⁄… Ê≈’œ«— «·≈Ì’«·"
    btn.OnAction = "SavePaymentAndIssueReceipt"

    Set btn = wsF.Buttons.Add(wsF.Range("A21").Left, wsF.Range("A21").Top, 220, 28)
    btn.Caption = "„⁄«Ì‰… / ÿ»«⁄… «·≈Ì’«·"
    btn.OnAction = "PrintCurrentReceipt"

    Set btn = wsL.Buttons.Add(wsL.Range("A2").Left, wsL.Range("A2").Top, 205, 28)
    btn.Caption = "⁄ﬂ” /  ”ÊÌ… „ÊÀﬁ…"
    btn.OnAction = "PostReversalOrAdjustment"

    Set btn = wsL.Buttons.Add(wsL.Range("F2").Left, wsL.Range("F2").Top, 205, 28)
    btn.Caption = "≈·€«¡ Õ—ﬂ… €Ì— „⁄ „œ…"
    btn.OnAction = "CancelPendingMovement"

    With wsF.PageSetup
        .PrintArea = "$E$1:$H$23"
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .PrintGridlines = False
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.3)
        .BottomMargin = Application.InchesToPoints(0.3)
        On Error Resume Next
        .PaperSize = xlPaperA5
        On Error GoTo 0
    End With
End Sub

Public Sub InstallCollectionButtons()
    Dim errorText As String
    On Error GoTo Failed

    CreateCollectionButtonsCore
    ConfigureCollectionProtection
    Exit Sub

Failed:
    errorText = Err.Description
    On Error Resume Next
    ConfigureCollectionProtection
    On Error GoTo 0
    MsgBox " ⁄–—  À»Ì  √“—«— «· Õ’Ì·: " & errorText & vbCrLf & _
        "√ı⁄Ìœ  „Õ«Ê·… Õ„«Ì… «·√Ê—«ﬁ. ·«  ‘€¯·  ”ÃÌ· «·œ›⁄«  ﬁ»· „—«Ã⁄… «·‰ ÌÃ….", vbCritical
End Sub

Public Sub ValidateCollectionSetup()
    Dim requiredSheets As Variant, requiredTables As Variant
    Dim i As Long, ws As Worksheet, tbl As ListObject
    Dim missingItems As String
    On Error GoTo Failed

    requiredSheets = Array( _
        "·ÊÕ… «· Õﬂ„", "”Ã· «·„·«ﬂ", "”Ã· «· Õ’Ì·", _
        " ”ÃÌ· «·œ›⁄… Ê«·≈Ì’«·", "«· ﬁ«—Ì— Ê«·≈ﬁ›«·", _
        "«·≈⁄œ«œ«  Êœ·Ì· «·«” Œœ«„", "”Ã· «· œﬁÌﬁ")

    For i = LBound(requiredSheets) To UBound(requiredSheets)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(requiredSheets(i)))
        On Error GoTo Failed
        If ws Is Nothing Then missingItems = missingItems & vbCrLf & "Ê—ﬁ… „›ﬁÊœ…: " & CStr(requiredSheets(i))
    Next i

    requiredTables = Array("OwnersTable", "CollectionsTable", "AuditTable")
    For i = LBound(requiredTables) To UBound(requiredTables)
        Set tbl = Nothing
        For Each ws In ThisWorkbook.Worksheets
            On Error Resume Next
            Set tbl = ws.ListObjects(CStr(requiredTables(i)))
            On Error GoTo Failed
            If Not tbl Is Nothing Then Exit For
        Next ws
        If tbl Is Nothing Then missingItems = missingItems & vbCrLf & "ÃœÊ· „›ﬁÊœ: " & CStr(requiredTables(i))
    Next i

    If Len(missingItems) > 0 Then
        MsgBox "›‘· ›Õ’ «— »«ÿ «·ﬂÊœ »«·ﬁ«·»:" & missingItems, vbCritical
    Else
        MsgBox "«Ã «“ ›Õ’ √”„«¡ «·√Ê—«ﬁ Ê«·Ãœ«Ê· «·√”«”Ì…. Â–« «·›Õ’ ·« Ì€‰Ì ⁄‰ «Œ »«— «·√“—«— Ê«·œ›⁄«  «· Ã—Ì»Ì….", vbInformation
    End If
    Exit Sub

Failed:
    MsgBox " ⁄–— ≈ﬂ„«· ›Õ’ «·„’‰›: " & Err.Description, vbCritical
End Sub

Public Sub SetWorkbookRTLAndButtons()
    Dim ws As Worksheet
    Dim errorText As String
    On Error GoTo Failed

    ' Remove protection before changing view direction.
    For Each ws In ThisWorkbook.Worksheets
        If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then
            ws.Unprotect Password:=SHEET_PASSWORD
        End If
    Next ws

    For Each ws In ThisWorkbook.Worksheets
        ws.DisplayRightToLeft = True
    Next ws

    CreateCollectionButtonsCore
    ConfigureCollectionProtection
    Application.CalculateFull
    Exit Sub

Failed:
    errorText = Err.Description
    On Error Resume Next
    ConfigureCollectionProtection
    On Error GoTo 0
    MsgBox " ⁄–— ≈ﬂ„«·  ÂÌ∆… «·„’‰›: " & errorText & vbCrLf & _
        "·«  ⁄ „œ ⁄·Ï «·√“—«— ﬁ»· «· Õﬁﬁ „‰ Õ„«Ì… «·√Ê—«ﬁ ÊŸÂÊ— «·√“—«—.", vbCritical
End Sub
