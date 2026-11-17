Attribute VB_Name = "modArchonProperty"
Option Compare Database
Option Explicit

' Property, building, unit, ownership-history and campaign services.
' Every financial charge stores unit and owner snapshots; ownership changes do not rewrite history.

Public Sub AT_RefreshOwnerSummaryCounts(ByVal ownerID As Long)
    Dim buildingCount As Long, unitCount As Long
    buildingCount = CLng(Nz(DCount("*", "tblBuildingOwnership", "OwnerID=" & ownerID & " AND StartDate<=Date() AND (EndDate Is Null OR EndDate>=Date())"), 0))
    unitCount = CLng(Nz(DLookup("UnitCount", "qryAT_OwnerCurrentUnitCounts", "OwnerID=" & ownerID), 0))
    CurrentDb.Execute "UPDATE tblOwners SET HouseCount=" & buildingCount & ", ApartmentCount=" & unitCount & " WHERE OwnerID=" & ownerID, dbFailOnError
End Sub

Public Function AT_CreateBuilding(ByVal buildingCode As String, ByVal buildingName As String, ByVal addressText As String, _
                                  ByVal initialOwnerID As Long, ByVal effectiveDate As Date, _
                                  Optional ByVal transferDocumentNo As String = "", Optional ByVal notes As String = "") As Long
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "BUILDING_EDIT"
    buildingCode = Trim$(buildingCode): buildingName = Trim$(buildingName)
    If Len(buildingCode) = 0 Or Len(buildingName) = 0 Then Err.Raise vbObjectError + 810, , AT_U("0631,0645,0632,0020,0627,0644,0645,0628,0646,0649,0020,0648,0627,0633,0645,0647,0020,0645,0637,0644,0648,0628,0627,0646,002E")
    If initialOwnerID <= 0 Or DCount("*", "tblOwners", "OwnerID=" & initialOwnerID & " AND AdminStatus='ACTIVE'") = 0 Then Err.Raise vbObjectError + 811, , AT_U("0627,062E,062A,0631,0020,0645,0627,0644,0643,064B,0627,0020,0645,0639,062A,0645,062F,064B,0627,0020,0644,0644,0645,0628,0646,0649,002E")
    If DCount("*", "tblBuildings", "BuildingCode='" & Replace(buildingCode, "'", "''") & "'") > 0 Then Err.Raise vbObjectError + 812, , AT_U("0631,0645,0632,0020,0627,0644,0645,0628,0646,0649,0020,0645,0633,062A,062E,062F,0645,0020,0645,0633,0628,0642,064B,0627,002E")
    If effectiveDate = 0 Then effectiveDate = Date

    Dim db As DAO.Database, rs As DAO.Recordset, buildingID As Long
    Set db = CurrentDb: DBEngine(0)(0).BeginTrans
    Set rs = db.OpenRecordset("tblBuildings", dbOpenDynaset)
    rs.AddNew: rs!BuildingCode = buildingCode: rs!BuildingName = buildingName
    rs!Address = addressText: rs!IsActive = True: rs!CreatedAt = Now(): rs!UpdatedAt = Now(): rs!Notes = notes
    rs.Update: rs.Bookmark = rs.LastModified: buildingID = CLng(rs!BuildingID): rs.Close

    Set rs = db.OpenRecordset("tblBuildingOwnership", dbOpenDynaset)
    rs.AddNew: rs!BuildingID = buildingID: rs!OwnerID = initialOwnerID: rs!StartDate = effectiveDate
    rs!TransferDocumentNo = transferDocumentNo: rs!Reason = AT_U("0625,0633,0646,0627,062F,0020,0627,0644,0645,0644,0643,064A,0629,0020,0639,0646,062F,0020,062A,0633,062C,064A,0644,0020,0627,0644,0645,0628,0646,0649")
    rs!CreatedAt = Now(): rs!CreatedByUserName = Environ$("Username"): rs.Update: rs.Close
    AT_RefreshOwnerSummaryCounts initialOwnerID
    DBEngine(0)(0).CommitTrans
    AT_Audit "BUILDING_CREATE", "tblBuildings", buildingID, "SUCCESS", "Building=" & buildingCode
    AT_CreateBuilding = buildingID
    Exit Function
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "BUILDING_CREATE", "tblBuildings", 0, "FAILED", eDesc
    Err.Raise savedErrNumber, "AT_CreateBuilding", eDesc
End Function

Public Function AT_AddUnit(ByVal buildingID As Long, ByVal unitCode As String, ByVal floorNo As String, _
                           Optional ByVal contributionEligible As Boolean = True, Optional ByVal notes As String = "") As Long
    Dim unitErr As Long, unitDesc As String
    On Error GoTo EH
    AT_RequirePermission "UNIT_EDIT"
    unitCode = Trim$(unitCode)
    If buildingID <= 0 Or DCount("*", "tblBuildings", "BuildingID=" & buildingID & " AND IsActive=True") = 0 Then Err.Raise vbObjectError + 820, , AT_U("0627,0644,0645,0628,0646,0649,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0020,0623,0648,0020,063A,064A,0631,0020,0646,0634,0637,002E")
    If Len(unitCode) = 0 Then Err.Raise vbObjectError + 821, , AT_U("0631,0645,0632,0020,0627,0644,0634,0642,0629,0020,0645,0637,0644,0648,0628,002E")
    If DCount("*", "tblUnits", "BuildingID=" & buildingID & " AND UnitCode='" & Replace(unitCode, "'", "''") & "'") > 0 Then Err.Raise vbObjectError + 822, , AT_U("0631,0645,0632,0020,0627,0644,0634,0642,0629,0020,0645,0633,062A,062E,062F,0645,0020,062F,0627,062E,0644,0020,0647,0630,0627,0020,0627,0644,0645,0628,0646,0649,002E")

    Dim db As DAO.Database, rs As DAO.Recordset, unitID As Long, ownerID As Long
    Set db = CurrentDb: DBEngine(0)(0).BeginTrans
    Set rs = db.OpenRecordset("tblUnits", dbOpenDynaset)
    rs.AddNew: rs!BuildingID = buildingID: rs!UnitCode = unitCode: rs!FloorNo = floorNo
    rs!IsActive = True: rs!IsContributionEligible = contributionEligible: rs!Notes = notes
    rs!CreatedAt = Now(): rs!UpdatedAt = Now(): rs.Update: rs.Bookmark = rs.LastModified
    unitID = CLng(rs!UnitID): rs.Close
    ownerID = CLng(Nz(DLookup("OwnerID", "tblBuildingOwnership", "BuildingID=" & buildingID & " AND StartDate<=Date() AND (EndDate Is Null OR EndDate>=Date())"), 0))
    If ownerID > 0 Then AT_RefreshOwnerSummaryCounts ownerID
    DBEngine(0)(0).CommitTrans
    AT_Audit "UNIT_CREATE", "tblUnits", unitID, "SUCCESS", "BuildingID=" & buildingID & "; UnitCode=" & unitCode
    AT_AddUnit = unitID
    Exit Function
EH:
    unitErr = Err.Number: unitDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "UNIT_CREATE", "tblUnits", 0, "FAILED", unitDesc
    Err.Raise unitErr, "AT_AddUnit", unitDesc
End Function

Public Sub AT_TransferBuildingOwnership(ByVal buildingID As Long, ByVal newOwnerID As Long, ByVal effectiveDate As Date, _
                                        ByVal transferDocumentNo As String, ByVal reason As String)
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "BUILDING_EDIT"
    If buildingID <= 0 Or newOwnerID <= 0 Then Err.Raise vbObjectError + 830, , AT_U("0627,0644,0645,0628,0646,0649,0020,0648,0627,0644,0645,0627,0644,0643,0020,0627,0644,062C,062F,064A,062F,0020,0645,0637,0644,0648,0628,0627,0646,002E")
    If Len(Trim$(reason)) = 0 Then Err.Raise vbObjectError + 831, , AT_U("0633,0628,0628,0020,0646,0642,0644,0020,0627,0644,0645,0644,0643,064A,0629,0020,0625,0644,0632,0627,0645,064A,002E")
    If DCount("*", "tblOwners", "OwnerID=" & newOwnerID & " AND AdminStatus='ACTIVE'") = 0 Then Err.Raise vbObjectError + 832, , AT_U("0627,0644,0645,0627,0644,0643,0020,0627,0644,062C,062F,064A,062F,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0020,0623,0648,0020,063A,064A,0631,0020,0645,0639,062A,0645,062F,002E")
    If effectiveDate = 0 Then effectiveDate = Date

    Dim db As DAO.Database, rs As DAO.Recordset, oldOwnerID As Long, oldStart As Date
    Set db = CurrentDb: DBEngine(0)(0).BeginTrans
    Set rs = db.OpenRecordset("SELECT OwnershipID,OwnerID,StartDate FROM tblBuildingOwnership WHERE BuildingID=" & buildingID & " AND StartDate<=#" & Format$(DateValue(effectiveDate), "mm\/dd\/yyyy") & "# AND (EndDate Is Null OR EndDate>=#" & Format$(DateValue(effectiveDate), "mm\/dd\/yyyy") & "#)", dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise vbObjectError + 833, , AT_U("0644,0627,0020,062A,0648,062C,062F,0020,0645,0644,0643,064A,0629,0020,062D,0627,0644,064A,0629,0020,0645,0633,062C,0644,0629,0020,0644,0647,0630,0627,0020,0627,0644,0645,0628,0646,0649,002E")
    oldOwnerID = CLng(rs!OwnerID): oldStart = CDate(rs!StartDate): rs.Close
    If oldOwnerID = newOwnerID Then Err.Raise vbObjectError + 834, , AT_U("0627,0644,0645,0627,0644,0643,0020,0627,0644,0645,062E,062A,0627,0631,0020,0647,0648,0020,0627,0644,0645,0627,0644,0643,0020,0627,0644,062D,0627,0644,064A,0020,0628,0627,0644,0641,0639,0644,002E")
    If DateValue(effectiveDate) <= DateValue(oldStart) Then Err.Raise vbObjectError + 835, , AT_U("062A,0627,0631,064A,062E,0020,0627,0644,0646,0642,0644,0020,064A,062C,0628,0020,0623,0646,0020,064A,0643,0648,0646,0020,0628,0639,062F,0020,0628,062F,0627,064A,0629,0020,0627,0644,0645,0644,0643,064A,0629,0020,0627,0644,062D,0627,0644,064A,0629,002E")
    If DCount("*", "tblBuildingOwnership", "BuildingID=" & buildingID & " AND StartDate>#" & Format$(DateValue(effectiveDate), "mm\/dd\/yyyy") & "# AND (EndDate Is Null OR EndDate>=#" & Format$(DateValue(effectiveDate), "mm\/dd\/yyyy") & "#)") > 0 Then Err.Raise vbObjectError + 836, , AT_U("064A,0648,062C,062F,0020,0633,062C,0644,0020,0645,0644,0643,064A,0629,0020,0645,0633,062A,0642,0628,0644,064A,0020,064A,062A,0639,0627,0631,0636,0020,0645,0639,0020,062A,0627,0631,064A,062E,0020,0627,0644,0646,0642,0644,061B,0020,0631,0627,062C,0639,0020,0627,0644,062A,0633,0644,0633,0644,0020,0627,0644,0632,0645,0646,064A,0020,0623,0648,0644,064B,0627,002E")

    db.Execute "UPDATE tblBuildingOwnership SET EndDate=#" & Format$(DateAdd("d", -1, DateValue(effectiveDate)), "mm\/dd\/yyyy") & "# WHERE BuildingID=" & buildingID & " AND (EndDate Is Null OR EndDate>=#" & Format$(DateValue(effectiveDate), "mm\/dd\/yyyy") & "#)", dbFailOnError
    Set rs = db.OpenRecordset("tblBuildingOwnership", dbOpenDynaset)
    rs.AddNew: rs!BuildingID = buildingID: rs!OwnerID = newOwnerID: rs!StartDate = DateValue(effectiveDate)
    rs!TransferDocumentNo = transferDocumentNo: rs!Reason = reason: rs!CreatedAt = Now()
    rs!CreatedByUserName = Environ$("Username"): rs.Update: rs.Close
    AT_RefreshOwnerSummaryCounts oldOwnerID: AT_RefreshOwnerSummaryCounts newOwnerID
    DBEngine(0)(0).CommitTrans
    AT_Audit "OWNERSHIP_TRANSFER", "tblBuildings", buildingID, "SUCCESS", "OldOwner=" & oldOwnerID & "; NewOwner=" & newOwnerID, reason
    Exit Sub
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "OWNERSHIP_TRANSFER", "tblBuildings", buildingID, "FAILED", eDesc, reason
    Err.Raise savedErrNumber, "AT_TransferBuildingOwnership", eDesc
End Sub

Public Function AT_AddUnitOccupancy(ByVal unitID As Long, ByVal familyName As String, ByVal householdCount As Long, _
                                    ByVal startDate As Date, Optional ByVal notes As String = "") As Long
    On Error GoTo EH
    AT_RequirePermission "UNIT_EDIT"
    If unitID <= 0 Or DCount("*", "tblUnits", "UnitID=" & unitID & " AND IsActive=True") = 0 Then Err.Raise vbObjectError + 840, , AT_U("0627,0644,0634,0642,0629,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0629,0020,0623,0648,0020,063A,064A,0631,0020,0646,0634,0637,0629,002E")
    If Len(Trim$(familyName)) = 0 Or householdCount < 1 Then Err.Raise vbObjectError + 841, , AT_U("0627,0633,0645,0020,0627,0644,0623,0633,0631,0629,0020,0648,0639,062F,062F,0020,0623,0641,0631,0627,062F,0647,0627,0020,0645,0637,0644,0648,0628,0627,0646,002E")
    If startDate = 0 Then startDate = Date
    Dim rs As DAO.Recordset, occupancyID As Long, db As DAO.Database
    Set db = CurrentDb: DBEngine(0)(0).BeginTrans
    Set rs = db.OpenRecordset("SELECT OccupancyID,StartDate FROM tblUnitOccupancy WHERE UnitID=" & unitID & " AND (EndDate Is Null OR EndDate>=#" & Format$(DateValue(startDate), "mm\/dd\/yyyy") & "# OR StartDate>=#" & Format$(DateValue(startDate), "mm\/dd\/yyyy") & "#)", dbOpenSnapshot)
    Do While Not rs.EOF
        If DateValue(rs!StartDate) >= DateValue(startDate) Then Err.Raise vbObjectError + 842, , AT_U("062A,0627,0631,064A,062E,0020,0627,0644,0625,0634,063A,0627,0644,0020,0627,0644,062C,062F,064A,062F,0020,064A,062A,062F,0627,062E,0644,0020,0645,0639,0020,0633,062C,0644,0020,0633,0627,0628,0642,0020,0623,0648,0020,064A,0633,0628,0642,0647,002E")
        db.Execute "UPDATE tblUnitOccupancy SET EndDate=#" & Format$(DateAdd("d", -1, DateValue(startDate)), "mm\/dd\/yyyy") & "# WHERE OccupancyID=" & CLng(rs!OccupancyID), dbFailOnError
        rs.MoveNext
    Loop
    rs.Close
    Set rs = db.OpenRecordset("tblUnitOccupancy", dbOpenDynaset)
    rs.AddNew: rs!UnitID = unitID: rs!FamilyName = Trim$(familyName): rs!HouseholdCount = householdCount
    rs!StartDate = DateValue(startDate): rs!Notes = notes: rs!CreatedAt = Now(): rs.Update: rs.Bookmark = rs.LastModified
    occupancyID = CLng(rs!OccupancyID): rs.Close
    DBEngine(0)(0).CommitTrans
    AT_Audit "UNIT_OCCUPANCY_CREATE", "tblUnitOccupancy", occupancyID, "SUCCESS", "UnitID=" & unitID
    AT_AddUnitOccupancy = occupancyID
    Exit Function
EH:
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "UNIT_OCCUPANCY_CREATE", "tblUnitOccupancy", 0, "FAILED", Err.Description
    Err.Raise Err.Number, "AT_AddUnitOccupancy", Err.Description
End Function

Public Function AT_RequestUnitExemption(ByVal unitID As Long, ByVal effectiveFrom As Date, ByVal effectiveTo As Variant, _
                                        ByVal reason As String, ByVal evidenceReference As String) As Long
    On Error GoTo EH
    AT_RequirePermission "UNIT_EXEMPTION_REQUEST"
    If unitID <= 0 Or DCount("*", "tblUnits", "UnitID=" & unitID & " AND IsActive=True") = 0 Then Err.Raise vbObjectError + 850, , AT_U("0627,0644,0634,0642,0629,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0629,0020,0623,0648,0020,063A,064A,0631,0020,0646,0634,0637,0629,002E")
    If Len(Trim$(reason)) = 0 Then Err.Raise vbObjectError + 851, , AT_U("0633,0628,0628,0020,0637,0644,0628,0020,0627,0644,0625,0639,0641,0627,0621,0020,0645,0637,0644,0648,0628,002E")
    If effectiveFrom = 0 Then effectiveFrom = Date
    If Not IsNull(effectiveTo) Then If CDate(effectiveTo) < effectiveFrom Then Err.Raise vbObjectError + 852, , AT_U("062A,0627,0631,064A,062E,0020,0646,0647,0627,064A,0629,0020,0627,0644,0625,0639,0641,0627,0621,0020,064A,0633,0628,0642,0020,062A,0627,0631,064A,062E,0020,0628,062F,0627,064A,062A,0647,002E")
    Dim rs As DAO.Recordset, requestID As Long
    Set rs = CurrentDb.OpenRecordset("tblUnitExemptionRequests", dbOpenDynaset)
    rs.AddNew: rs!UnitID = unitID: rs!RequestDate = Date: rs!EffectiveFrom = effectiveFrom
    If Not IsNull(effectiveTo) Then rs!EffectiveTo = CDate(effectiveTo)
    rs!Reason = reason: rs!EvidenceReference = evidenceReference: rs!RequestedByUserName = Environ$("Username")
    rs!StatusCode = "PENDING": rs!CreatedAt = Now(): rs.Update: rs.Bookmark = rs.LastModified
    requestID = CLng(rs!ExemptionRequestID): rs.Close
    AT_Audit "UNIT_EXEMPTION_REQUEST", "tblUnitExemptionRequests", requestID, "SUCCESS", "UnitID=" & unitID, reason
    AT_RequestUnitExemption = requestID
    Exit Function
EH:
    AT_Audit "UNIT_EXEMPTION_REQUEST", "tblUnitExemptionRequests", 0, "FAILED", Err.Description, reason
    Err.Raise Err.Number, "AT_RequestUnitExemption", Err.Description
End Function

Public Sub AT_DecideUnitExemption(ByVal requestID As Long, ByVal approveRequest As Boolean, ByVal decisionReason As String)
    Dim decisionErr As Long, decisionDesc As String
    On Error GoTo EH
    AT_RequirePermission "UNIT_EXEMPTION_APPROVE"
    If requestID <= 0 Or DCount("*", "tblUnitExemptionRequests", "ExemptionRequestID=" & requestID & " AND StatusCode='PENDING'") = 0 Then Err.Raise vbObjectError + 860, , AT_U("0637,0644,0628,0020,0627,0644,0625,0639,0641,0627,0621,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0020,0623,0648,0020,0633,0628,0642,0020,0627,0644,0628,062A,0020,0641,064A,0647,002E")
    If Len(Trim$(decisionReason)) = 0 Then Err.Raise vbObjectError + 861, , AT_U("0633,0628,0628,0020,0627,0644,0642,0631,0627,0631,0020,0625,0644,0632,0627,0645,064A,002E")
    Dim statusText As String, unitID As Long, ownerID As Long, rs As DAO.Recordset
    statusText = IIf(approveRequest, "APPROVED", "REJECTED")
    unitID = CLng(Nz(DLookup("UnitID", "tblUnitExemptionRequests", "ExemptionRequestID=" & requestID), 0))
    DBEngine(0)(0).BeginTrans
    CurrentDb.Execute "UPDATE tblUnitExemptionRequests SET StatusCode='" & statusText & "', DecidedByUserName='" & Replace(Environ$("Username"), "'", "''") & "', DecidedAt=Now(), DecisionReason='" & Replace(decisionReason, "'", "''") & "' WHERE ExemptionRequestID=" & requestID & " AND StatusCode='PENDING'", dbFailOnError
    Set rs = CurrentDb.OpenRecordset("SELECT BO.OwnerID FROM tblBuildingOwnership AS BO INNER JOIN tblUnits AS U ON BO.BuildingID=U.BuildingID WHERE U.UnitID=" & unitID & " AND BO.EndDate Is Null", dbOpenSnapshot)
    If Not rs.EOF Then ownerID = CLng(rs!OwnerID)
    rs.Close
    If ownerID > 0 Then AT_RefreshOwnerSummaryCounts ownerID
    DBEngine(0)(0).CommitTrans
    AT_Audit "UNIT_EXEMPTION_DECISION", "tblUnitExemptionRequests", requestID, "SUCCESS", statusText, decisionReason
    Exit Sub
EH:
    decisionErr = Err.Number: decisionDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "UNIT_EXEMPTION_DECISION", "tblUnitExemptionRequests", requestID, "FAILED", decisionDesc, decisionReason
    Err.Raise decisionErr, "AT_DecideUnitExemption", decisionDesc
End Sub

Public Function AT_CreateCampaign(ByVal campaignCode As String, ByVal campaignName As String, ByVal contributionTypeID As Long, _
                                  ByVal targetScope As String, ByVal amountPerUnit As Currency, ByVal rateID As Long, _
                                  ByVal periodStart As Date, ByVal periodEnd As Date, ByVal dueDate As Date, ByVal reason As String) As Long
    On Error GoTo EH
    AT_RequirePermission "CAMPAIGN_DRAFT"
    campaignCode = Trim$(campaignCode): campaignName = Trim$(campaignName): targetScope = UCase$(Trim$(targetScope))
    If Len(campaignCode) = 0 Or Len(campaignName) = 0 Then Err.Raise vbObjectError + 870, , AT_U("0631,0645,0632,0020,0627,0644,062D,0645,0644,0629,0020,0648,0627,0633,0645,0647,0627,0020,0645,0637,0644,0648,0628,0627,0646,002E")
    If targetScope <> "ALL" And targetScope <> "BUILDINGS" And targetScope <> "UNITS" Then Err.Raise vbObjectError + 871, , AT_U("0646,0637,0627,0642,0020,0627,0644,062D,0645,0644,0629,0020,063A,064A,0631,0020,0635,0627,0644,062D,002E")
    If amountPerUnit <= 0 Or amountPerUnit <> Fix(amountPerUnit) Then Err.Raise vbObjectError + 872, , AT_U("0642,064A,0645,0629,0020,0627,0644,0645,0633,0627,0647,0645,0629,0020,0644,0643,0644,0020,0634,0642,0629,0020,064A,062C,0628,0020,0623,0646,0020,062A,0643,0648,0646,0020,0639,062F,062F,064B,0627,0020,0635,062D,064A,062D,064B,0627,0020,0645,0648,062C,0628,064B,0627,0020,0628,0627,0644,0631,064A,0627,0644,0020,0627,0644,064A,0645,0646,064A,002E")
    If periodEnd <> 0 And periodStart <> 0 Then If periodEnd < periodStart Then Err.Raise vbObjectError + 873, , AT_U("0646,0647,0627,064A,0629,0020,0627,0644,0641,062A,0631,0629,0020,062A,0633,0628,0642,0020,0628,062F,0627,064A,062A,0647,0627,002E")
    If DCount("*", "tblContributionTypes", "ContributionTypeID=" & contributionTypeID & " AND IsActive=True") = 0 Then Err.Raise vbObjectError + 874, , AT_U("0646,0648,0639,0020,0627,0644,0645,0633,0627,0647,0645,0629,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0020,0623,0648,0020,063A,064A,0631,0020,0646,0634,0637,002E")
    If DCount("*", "tblContributionCampaigns", "CampaignCode='" & Replace(campaignCode, "'", "''") & "'") > 0 Then Err.Raise vbObjectError + 875, , AT_U("0631,0645,0632,0020,0627,0644,062D,0645,0644,0629,0020,0645,0633,062A,062E,062F,0645,0020,0645,0633,0628,0642,064B,0627,002E")
    If rateID > 0 Then If DCount("*", "tblContributionRates", "RateID=" & rateID & " AND IsActive=True") = 0 Then Err.Raise vbObjectError + 876, , AT_U("0645,0639,062F,0644,0020,0627,0644,0645,0633,0627,0647,0645,0629,0020,063A,064A,0631,0020,0641,0639,0627,0644,0020,0623,0648,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,002E")
    Dim rs As DAO.Recordset, campaignID As Long
    Set rs = CurrentDb.OpenRecordset("tblContributionCampaigns", dbOpenDynaset)
    rs.AddNew: rs!CampaignCode = campaignCode: rs!CampaignName = campaignName: rs!ContributionTypeID = contributionTypeID
    rs!TargetScope = targetScope: If rateID > 0 Then rs!RateID = rateID
    rs!AmountPerUnit = amountPerUnit: If periodStart <> 0 Then rs!PeriodStart = periodStart
    If periodEnd <> 0 Then rs!PeriodEnd = periodEnd
    If dueDate <> 0 Then rs!DueDate = dueDate
    rs!Reason = reason: rs!StatusCode = "DRAFT": rs!CreatedAt = Now(): rs!CreatedByUserName = Environ$("Username")
    rs.Update: rs.Bookmark = rs.LastModified: campaignID = CLng(rs!CampaignID): rs.Close
    AT_Audit "CAMPAIGN_CREATE", "tblContributionCampaigns", campaignID, "SUCCESS", "Code=" & campaignCode, reason
    AT_CreateCampaign = campaignID
    Exit Function
EH:
    AT_Audit "CAMPAIGN_CREATE", "tblContributionCampaigns", 0, "FAILED", Err.Description, reason
    Err.Raise Err.Number, "AT_CreateCampaign", Err.Description
End Function

Public Sub AT_AddCampaignBuildingTarget(ByVal campaignID As Long, ByVal buildingID As Long)
    AT_RequirePermission "CAMPAIGN_DRAFT"
    If DCount("*", "tblContributionCampaigns", "CampaignID=" & campaignID & " AND StatusCode='DRAFT' AND TargetScope='BUILDINGS'") = 0 Then Err.Raise vbObjectError + 880, , AT_U("0627,0644,062D,0645,0644,0629,0020,0644,064A,0633,062A,0020,0645,0633,0648,062F,0629,0020,0628,0646,0637,0627,0642,0020,0627,0644,0645,0628,0627,0646,064A,002E")
    If DCount("*", "tblBuildings", "BuildingID=" & buildingID & " AND IsActive=True") = 0 Then Err.Raise vbObjectError + 881, , AT_U("0627,0644,0645,0628,0646,0649,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0020,0623,0648,0020,063A,064A,0631,0020,0646,0634,0637,002E")
    If DCount("*", "tblCampaignBuildingTargets", "CampaignID=" & campaignID & " AND BuildingID=" & buildingID) = 0 Then CurrentDb.Execute "INSERT INTO tblCampaignBuildingTargets (CampaignID,BuildingID) VALUES (" & campaignID & "," & buildingID & ")", dbFailOnError
End Sub

Public Sub AT_AddCampaignUnitTarget(ByVal campaignID As Long, ByVal unitID As Long)
    AT_RequirePermission "CAMPAIGN_DRAFT"
    If DCount("*", "tblContributionCampaigns", "CampaignID=" & campaignID & " AND StatusCode='DRAFT' AND TargetScope='UNITS'") = 0 Then Err.Raise vbObjectError + 882, , AT_U("0627,0644,062D,0645,0644,0629,0020,0644,064A,0633,062A,0020,0645,0633,0648,062F,0629,0020,0628,0646,0637,0627,0642,0020,0627,0644,0648,062D,062F,0627,062A,002E")
    If DCount("*", "tblUnits", "UnitID=" & unitID & " AND IsActive=True") = 0 Then Err.Raise vbObjectError + 883, , AT_U("0627,0644,0634,0642,0629,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0629,0020,0623,0648,0020,063A,064A,0631,0020,0646,0634,0637,0629,002E")
    If DCount("*", "tblCampaignUnitTargets", "CampaignID=" & campaignID & " AND UnitID=" & unitID) = 0 Then CurrentDb.Execute "INSERT INTO tblCampaignUnitTargets (CampaignID,UnitID) VALUES (" & campaignID & "," & unitID & ")", dbFailOnError
End Sub

Public Function AT_GenerateCampaignCharges(ByVal campaignID As Long, ByVal chargeDate As Date) As Long
    Dim savedErrNumber As Long, eDesc As String
    On Error GoTo EH
    AT_RequirePermission "CAMPAIGN_DRAFT"
    Dim scopeText As String, amount As Currency, campaignName As String, campaignReason As String
    Dim rsCampaign As DAO.Recordset, rsUnits As DAO.Recordset, rsOwner As DAO.Recordset, rsCharge As DAO.Recordset, rsLine As DAO.Recordset, rsDoc As DAO.Recordset
    Dim groups As Object, bucket As Collection, ownerKey As Variant, unitRec As Variant, item As Variant
    Dim sql As String, ownerID As Long, chargeID As Long, docID As Long, lineNo As Long, ownerCode As String, ownerName As String
    If chargeDate = 0 Then chargeDate = Date
    Set rsCampaign = CurrentDb.OpenRecordset("SELECT C.*,T.TypeCode,T.CalculationMethod FROM tblContributionCampaigns AS C INNER JOIN tblContributionTypes AS T ON C.ContributionTypeID=T.ContributionTypeID WHERE C.CampaignID=" & campaignID, dbOpenSnapshot)
    If rsCampaign.EOF Then rsCampaign.Close: Err.Raise vbObjectError + 890, , AT_U("0627,0644,062D,0645,0644,0629,0020,063A,064A,0631,0020,0645,0648,062C,0648,062F,0629,002E")
    If rsCampaign!StatusCode <> "DRAFT" Then rsCampaign.Close: Err.Raise vbObjectError + 891, , AT_U("064A,0645,0643,0646,0020,0625,0646,0634,0627,0621,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,0627,062A,0020,0645,0646,0020,062D,0645,0644,0629,0020,0645,0633,0648,062F,0629,0020,0641,0642,0637,002E")
    scopeText = CStr(rsCampaign!TargetScope): amount = CCur(rsCampaign!AmountPerUnit): campaignName = CStr(rsCampaign!CampaignName): campaignReason = Nz(rsCampaign!Reason, "")
    If rsCampaign!CalculationMethod <> "PER_UNIT" Then rsCampaign.Close: Err.Raise vbObjectError + 892, , AT_U("0637,0631,064A,0642,0629,0020,0627,062D,062A,0633,0627,0628,0020,0627,0644,062D,0645,0644,0629,0020,063A,064A,0631,0020,0645,062F,0639,0648,0645,0629,0020,0641,064A,0020,0647,0630,0647,0020,0627,0644,0646,0633,062E,0629,061B,0020,0627,0644,0645,062A,0627,062D,0020,0647,0648,0020,0645,0628,0644,063A,0020,0644,0643,0644,0020,0648,062D,062F,0629,002E")
    rsCampaign.Close
    If scopeText = "BUILDINGS" And DCount("*", "tblCampaignBuildingTargets", "CampaignID=" & campaignID) = 0 Then Err.Raise vbObjectError + 893, , AT_U("0623,0636,0641,0020,0645,0628,0646,0649,0020,0648,0627,062D,062F,064B,0627,0020,0639,0644,0649,0020,0627,0644,0623,0642,0644,0020,0625,0644,0649,0020,0646,0637,0627,0642,0020,0627,0644,062D,0645,0644,0629,002E")
    If scopeText = "UNITS" And DCount("*", "tblCampaignUnitTargets", "CampaignID=" & campaignID) = 0 Then Err.Raise vbObjectError + 894, , AT_U("0623,0636,0641,0020,0634,0642,0629,0020,0648,0627,062D,062F,0629,0020,0639,0644,0649,0020,0627,0644,0623,0642,0644,0020,0625,0644,0649,0020,0646,0637,0627,0642,0020,0627,0644,062D,0645,0644,0629,002E")
    If DCount("*", "tblChargeLines", "CampaignID=" & campaignID) > 0 Then Err.Raise vbObjectError + 895, , AT_U("0633,0628,0642,0020,0625,0646,0634,0627,0621,0020,062A,0641,0627,0635,064A,0644,0020,0627,0633,062A,062D,0642,0627,0642,0627,062A,0020,0644,0647,0630,0647,0020,0627,0644,062D,0645,0644,0629,061B,0020,0645,0646,0639,064B,0627,0020,0644,0644,062A,0643,0631,0627,0631,0020,0644,0645,0020,064A,064F,0646,0634,0623,0020,0634,064A,0621,0020,062C,062F,064A,062F,002E")

    sql = "SELECT U.UnitID,U.UnitCode,U.BuildingID,B.BuildingCode,BO.OwnerID FROM ((tblUnits AS U INNER JOIN tblBuildings AS B ON U.BuildingID=B.BuildingID) INNER JOIN tblBuildingOwnership AS BO ON B.BuildingID=BO.BuildingID) " & _
          "WHERE U.IsActive=True AND U.IsContributionEligible=True AND B.IsActive=True AND BO.StartDate<=#" & Format$(chargeDate, "mm\/dd\/yyyy") & "# AND (BO.EndDate Is Null OR BO.EndDate>=#" & Format$(chargeDate, "mm\/dd\/yyyy") & "#) " & _
          "AND NOT EXISTS (SELECT X.ExemptionRequestID FROM tblUnitExemptionRequests AS X WHERE X.UnitID=U.UnitID AND X.StatusCode='APPROVED' AND X.EffectiveFrom<=#" & Format$(chargeDate, "mm\/dd\/yyyy") & "# AND (X.EffectiveTo Is Null OR X.EffectiveTo>=#" & Format$(chargeDate, "mm\/dd\/yyyy") & "#)) "
    If scopeText = "BUILDINGS" Then sql = sql & "AND EXISTS (SELECT T.CampaignBuildingTargetID FROM tblCampaignBuildingTargets AS T WHERE T.CampaignID=" & campaignID & " AND T.BuildingID=U.BuildingID) "
    If scopeText = "UNITS" Then sql = sql & "AND EXISTS (SELECT T.CampaignUnitTargetID FROM tblCampaignUnitTargets AS T WHERE T.CampaignID=" & campaignID & " AND T.UnitID=U.UnitID) "
    sql = sql & "ORDER BY BO.OwnerID,B.BuildingCode,U.UnitCode;"
    Set rsUnits = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    Set groups = CreateObject("Scripting.Dictionary")
    Do While Not rsUnits.EOF
        ownerKey = CStr(CLng(rsUnits!OwnerID))
        If Not groups.Exists(ownerKey) Then Set bucket = New Collection: groups.Add ownerKey, bucket
        groups(ownerKey).Add Array(CLng(rsUnits!UnitID), CStr(rsUnits!UnitCode), CLng(rsUnits!BuildingID), CStr(rsUnits!BuildingCode))
        rsUnits.MoveNext
    Loop
    rsUnits.Close
    If groups.Count = 0 Then Err.Raise vbObjectError + 896, , AT_U("0644,0627,0020,062A,0648,062C,062F,0020,0648,062D,062F,0627,062A,0020,0645,0624,0647,0644,0629,0020,0636,0645,0646,0020,0646,0637,0627,0642,0020,0627,0644,062D,0645,0644,0629,0020,0628,062A,0627,0631,064A,062E,0020,0627,0644,0627,0633,062A,062D,0642,0627,0642,002E")

    DBEngine(0)(0).BeginTrans
    CurrentDb.Execute "UPDATE tblContributionCampaigns SET StatusCode='GENERATING' WHERE CampaignID=" & campaignID & " AND StatusCode='DRAFT'", dbFailOnError
    If CurrentDb.RecordsAffected <> 1 Then Err.Raise vbObjectError + 897, , AT_U("062A,063A,064A,0631,062A,0020,062D,0627,0644,0629,0020,0627,0644,062D,0645,0644,0629,0020,0642,0628,0644,0020,0628,062F,0621,0020,0627,0644,0625,0646,0634,0627,0621,002E")
    For Each ownerKey In groups.Keys
        ownerID = CLng(ownerKey)
        Set rsOwner = CurrentDb.OpenRecordset("SELECT OwnerCode,FullName FROM tblOwners WHERE OwnerID=" & ownerID & " AND AdminStatus='ACTIVE'", dbOpenSnapshot)
        If rsOwner.EOF Then rsOwner.Close: Err.Raise vbObjectError + 898, , AT_U("0625,062D,062F,0649,0020,0627,0644,0648,062D,062F,0627,062A,0020,0644,0627,0020,062A,0631,062A,0628,0637,0020,0628,0645,0627,0644,0643,0020,0646,0634,0637,002E")
        ownerCode = Nz(rsOwner!OwnerCode, ""): ownerName = Nz(rsOwner!FullName, ""): rsOwner.Close
        Set rsDoc = CurrentDb.OpenRecordset("tblFinancialDocuments", dbOpenDynaset)
        rsDoc.AddNew: rsDoc!DocumentType = "CHARGE": rsDoc!TransactionDate = chargeDate: rsDoc!StatusCode = "DRAFT"
        rsDoc!CreatedAt = Now(): rsDoc!CreatedByUserName = Environ$("Username"): rsDoc!Reason = campaignReason
        rsDoc.Update: rsDoc.Bookmark = rsDoc.LastModified: docID = CLng(rsDoc!DocumentID): rsDoc.Close
        Set rsCharge = CurrentDb.OpenRecordset("tblOwnerCharges", dbOpenDynaset)
        rsCharge.AddNew: rsCharge!DocumentID = docID: rsCharge!OwnerID = ownerID: rsCharge!CampaignID = campaignID
        rsCharge!ChargeDate = chargeDate: rsCharge!ApartmentCountAtCharge = groups(ownerKey).Count
        rsCharge!RateAmount = amount: rsCharge!GrossAmount = CCur(groups(ownerKey).Count) * amount
        rsCharge!DiscountAmount = 0: rsCharge!NetAmount = rsCharge!GrossAmount
        rsCharge!OwnerCodeAtCharge = ownerCode: rsCharge!OwnerNameAtCharge = ownerName
        rsCharge!ChargeStatus = "DRAFT": rsCharge!CreatedAt = Now(): rsCharge.Update
        rsCharge.Bookmark = rsCharge.LastModified: chargeID = CLng(rsCharge!ChargeID): rsCharge.Close
        lineNo = 0
        For Each unitRec In groups(ownerKey)
            Set rsLine = CurrentDb.OpenRecordset("tblChargeLines", dbOpenDynaset)
            rsLine.AddNew: rsLine!ChargeID = chargeID: rsLine!CampaignID = campaignID
            rsLine!UnitID = CLng(unitRec(0)): rsLine!UnitCodeSnapshot = CStr(unitRec(1))
            rsLine!BuildingID = CLng(unitRec(2)): rsLine!BuildingCodeSnapshot = CStr(unitRec(3))
            rsLine!OwnerIDSnapshot = ownerID: rsLine!RateAmount = amount: rsLine!GrossAmount = amount
            rsLine!DiscountAmount = 0: rsLine!NetAmount = amount: rsLine!CreatedAt = Now(): rsLine.Update: rsLine.Close
            lineNo = lineNo + 1
        Next unitRec
        AT_GenerateCampaignCharges = AT_GenerateCampaignCharges + 1
    Next ownerKey
    CurrentDb.Execute "UPDATE tblContributionCampaigns SET StatusCode='GENERATED' WHERE CampaignID=" & campaignID, dbFailOnError
    DBEngine(0)(0).CommitTrans
    AT_Audit "CAMPAIGN_GENERATE", "tblContributionCampaigns", campaignID, "SUCCESS", "Owner charges generated=" & AT_GenerateCampaignCharges, campaignName
    Exit Function
EH:
    savedErrNumber = Err.Number: eDesc = Err.Description
    On Error Resume Next: DBEngine(0)(0).Rollback
    AT_Audit "CAMPAIGN_GENERATE", "tblContributionCampaigns", campaignID, "FAILED", eDesc
    Err.Raise savedErrNumber, "AT_GenerateCampaignCharges", eDesc
End Function

Public Function AT_ApproveCampaignCharges(ByVal campaignID As Long) As Long
    On Error GoTo EH
    AT_RequirePermission "CAMPAIGN_APPROVE"
    Dim rs As DAO.Recordset, ids As Collection, item As Variant, pendingCount As Long
    Set ids = New Collection
    Set rs = CurrentDb.OpenRecordset("SELECT C.ChargeID FROM tblOwnerCharges AS C INNER JOIN tblFinancialDocuments AS D ON C.DocumentID=D.DocumentID WHERE C.CampaignID=" & campaignID & " AND C.ChargeStatus='DRAFT' AND D.StatusCode='DRAFT' ORDER BY C.ChargeID", dbOpenSnapshot)
    Do While Not rs.EOF: ids.Add CLng(rs!ChargeID): rs.MoveNext: Loop: rs.Close
    If ids.Count = 0 Then Err.Raise vbObjectError + 899, , AT_U("0644,0627,0020,062A,0648,062C,062F,0020,0627,0633,062A,062D,0642,0627,0642,0627,062A,0020,0645,0633,0648,062F,0629,0020,0644,0627,0639,062A,0645,0627,062F,0647,0627,002E")
    For Each item In ids
        AT_ApproveCharge CLng(item), Environ$("Username")
        AT_ApproveCampaignCharges = AT_ApproveCampaignCharges + 1
    Next item
    If DCount("*", "tblOwnerCharges", "CampaignID=" & campaignID & " AND ChargeStatus<>'ACTIVE'") = 0 Then CurrentDb.Execute "UPDATE tblContributionCampaigns SET StatusCode='APPROVED', ApprovedAt=Now(), ApprovedByUserName='" & Replace(Environ$("Username"), "'", "''") & "' WHERE CampaignID=" & campaignID, dbFailOnError
    AT_Audit "CAMPAIGN_APPROVE", "tblContributionCampaigns", campaignID, "SUCCESS", "Approved owner charges=" & AT_ApproveCampaignCharges
    Exit Function
EH:
    AT_Audit "CAMPAIGN_APPROVE", "tblContributionCampaigns", campaignID, "FAILED", Err.Description
    Err.Raise Err.Number, "AT_ApproveCampaignCharges", Err.Description
End Function

Public Function AT_UnitIsExempt(ByVal unitID As Long, ByVal asOfDate As Date) As Boolean
    AT_UnitIsExempt = (DCount("*", "tblUnitExemptionRequests", "UnitID=" & unitID & " AND StatusCode='APPROVED' AND EffectiveFrom<=#" & Format$(asOfDate, "mm\/dd\/yyyy") & "# AND (EffectiveTo Is Null OR EffectiveTo>=#" & Format$(asOfDate, "mm\/dd\/yyyy") & "#)") > 0)
End Function
