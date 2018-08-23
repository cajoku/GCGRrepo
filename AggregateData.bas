Attribute VB_Name = "AggregateData"
'
Private projectNAME As Variant
Private projectLOCATION As Variant
Private opsjobNUM As Variant
Private pretaskNUM As Variant
Private docDATE As Variant
Private projectCOST As Variant
Private vacation As Variant
Private preconSTART As Variant
Private preconEND As Variant
Private conStart As Variant
Private conEnd As Variant
Private preconLIT As Variant
Private conLIT As Variant
Private preconRAISE As Variant
Private conRAISE As Variant
Private IT As Variant
Private gas As Variant
Private phone As Variant
Private units As Variant
Private area As Variant
Private preTOTAL As Variant
Private conTOTAL As Variant
Private gcTOTAL As Variant
Private grTOTAL As Variant
Private gcARR() As Variant
Private grARR() As Variant
Private staffARR() As Variant

Private jsonBuilder As JSON
Private paJSON As JSON

Sub uploadData()


    MsgBox "Success!, your data has been uploaded to the AWS server"



End Sub

Sub mctest()

    Dim j1 As JSON
    Dim jArr As JSON
    Dim jObj As JSON
    
    
    Set j1 = New JSON
    
    With j1
        .Key = "mc obj"
        .addLeaf "leaf", "value"
        .addLeaf "leaf", "value"
        .addLeaf "leaf", "value"
        Set jArr = j1.addArray("jArr")
        Set jObj = j1.addObj
    End With
    
    Dim i As Integer
    Dim j4 As JSON
    
    For i = 1 To 10
        Set j4 = New JSON
        j4.addLeaf str(i), i
        jArr.appendArray j4
    Next
    
    Debug.Print j1.JSONstr

End Sub


Sub mikecolbytry()

    Dim j1 As JSON, settingsJ As JSON, staffJ As JSON
    Dim j2 As JSON, j3 As JSON
    
    Set j1 = New JSON
    
    With j1
        .Key = "master obj"
        Set settingsJ = .addObj
        Set staffJ = .addArray("Staff Items")
        
        'for each employee in doc.employees
            Set j2 = New JSON
            With j2
                .addLeaf "Position", ""
                .addLeaf "Position", ""
            End With
            
            'for each month in employee.months
                Set j3 = j2.addObj
                With j3
                    .addLeaf "month int": i
                End With
                
            'next
            staffJ.appendArray j2
        'next


End Sub

Sub testjson()

    Dim WB As Workbook
    Dim WS As Worksheet
    Dim sdWS As Worksheet
    
    Set WB = masterWB
    Set WS = masterOBJ.WS
    Set sdWS = masterOBJ.sdWS
    
    Set jsonBuilder = New JSON
    With jsonBuilder
        .Key = "GCGR Project"
        .addLeaf "UserName", Environ("USERNAME")
        .addLeaf "DateSubmitted", Format(Now(), "Short Date")
    End With
        
    'For project atts
    ScanProjAtts WS
    
    'For staff dtails
    ScanStaffDetail sdWS
    
    'For GC items
    ScanLineItems masterOBJ.gcdWS
    
    'For GR items
    ScanLineItems masterOBJ.grdWS
    
    'Debug.Print jsonBuilder.JSONstr
    
    jsonBuilder.createJSONfile MasterDir, "jsonTest2"

End Sub

Sub ScanProjAtts(WS As Worksheet)
    
    Set paJSON = jsonBuilder.addObj
    paJSON.Key = "Project Attributes"
    
    
    projectATTS WS.[\proj]
    projectATTS WS.[\loc]
    projectATTS WS.[\ops]
    projectATTS WS.[\task]
    projectATTS WS.[\date]
    projectATTS WS.[\pcost]
    projectATTS WS.[\vaca]
    projectATTS WS.[\pstart]
    projectATTS WS.[\pend]
    DaysFromStart WS.[\pend], "PreconEnd Delta"
    projectATTS WS.[\cstart]
    DaysFromStart WS.[\cstart], "ConStart Delta"
    projectATTS WS.[\cend]
    DaysFromStart WS.[\cend], "ConEnd Delta"
    projectATTS WS.[\preLIT]
    projectATTS WS.[\conLIT]
    projectATTS WS.[\preRaise]
    projectATTS WS.[\conRaise]
    projectATTS WS.[\IT]
    projectATTS WS.[\gas]
    projectATTS WS.[\iphone]
    projectATTS WS.[\unit]
    projectATTS WS.[\area]
    projectATTS WS.[\pretotal]
    projectATTS WS.[\contotal]
    projectATTS WS.[\gcontotal]
    projectATTS WS.[\greqtotal]
    
    
End Sub

Sub ScanStaffDetail(WS As Worksheet)

    Dim positionRan As Range
    Dim m2mRAN As Range
    Dim arr() As Variant, tempARR As Variant
    Dim pNameRan As Range, salRan As Range, autoRan As Range, sMonthRan As Range, durRan As Range, sDateRan As Range, eDateRan As Range, pWorkRan As Range
    Dim pNameStr As String, salVal As Double, autoVal As Double, sMonth As Integer, dur As Integer, sDate As Date, eDate As Date, pWork As Double
    Dim staffCount As Integer, i As Integer, j As Integer, p As Integer, q As Integer
    Dim keySTR As String, valOBJ As Variant
    Dim staffJSON As JSON, topJSON As JSON, tempJSON As JSON
    
    Set positionRan = boxRANGE(WS, "\c_Position", "\r_start", "\r_end")
    trimRANGE positionRan, dsupdown
    
    Set m2mRAN = boxRANGE(WS, "\r_start", "\c_negStart", "\c_durEND")
    trimRANGE m2mRAN, dssides
    
    Set pNameRan = WS.[\c_posName]
    Set salRan = WS.[\salary]
    Set autoRan = WS.[\auto]
    Set sMonthRan = WS.[\c_jobStart]
    Set durRan = WS.[\c_jobDur]
    Set sDateRan = WS.[\c_posStart]
    Set eDateRan = WS.[\c_posEnd]
    Set pWorkRan = WS.[\c_perTIME]
    
    Set staffJSON = jsonBuilder.addArray("Staff Items")
    
    
    For Each var In positionRan
        If Intersect(var.EntireRow, sDateRan.EntireColumn).Value <> 0 And Intersect(var.EntireRow, durRan.EntireColumn).Value <> 0 Then
            staffCount = staffCount + 1
        End If
    Next
    
    ReDim arr(staffCount - 1, 10 + m2mRAN.Cells.Count)
    Dim keyARR() As Variant
    keyARR() = Array("Position", "Name", "Salary", "Auto", "StartDate", "StartDate Delta", "EndDate", "EndDate Delta", "Rate", "GrandTotal")
    
    For Each var In positionRan
        If Intersect(var.EntireRow, sDateRan.EntireColumn).Value <> 0 And Intersect(var.EntireRow, durRan.EntireColumn).Value <> 0 Then
            
            arr(i, 0) = var.Value ': keyARR(i) = "Position"
            arr(i, 1) = Intersect(var.EntireRow, pNameRan.EntireColumn).Value ': keyARR(i) = "Name"
            arr(i, 2) = Intersect(var.EntireRow, salRan.EntireColumn).Value ': keyARR(i) = "Salary"
            arr(i, 3) = Intersect(var.EntireRow, autoRan.EntireColumn).Value ': keyARR(i) = "Auto"
            arr(i, 4) = Intersect(var.EntireRow, sDateRan.EntireColumn).Value ': keyARR(i) = "StartDate"
            arr(i, 5) = DateDiff("d", WS.[\pstart].Value, Intersect(var.EntireRow, sDateRan.EntireColumn).Value)
            arr(i, 6) = Intersect(var.EntireRow, eDateRan.EntireColumn).Value ': keyARR(i) = "EndDate"
            arr(i, 7) = DateDiff("d", WS.[\pstart].Value, Intersect(var.EntireRow, eDateRan.EntireColumn).Value)
            arr(i, 8) = Round(Intersect(var.EntireRow, WS.[\c_durEND].Offset(0, 1).EntireColumn).Value, 2) ': keyARR(i) = "Rate"
            arr(i, 9) = Round(Intersect(var.EntireRow, WS.[\c_durEND].Offset(0, 2).EntireColumn).Value, 2) ': keyARR(i) = "GrandTotal"
            'arr(i, 8) = "PercentWork:" & Intersect(var.EntireRow, pWorkRan.EntireColumn).Value
            p = 1
            For j = 10 To m2mRAN.Cells.Count + 9
                arr(i, j) = m2mRAN.Cells(1, p).Value & ":" & Intersect(var.EntireRow, m2mRAN.Cells(1, p).EntireColumn).Value
                p = p + 1
            Next
            i = i + 1
        End If
    Next
            
    For i = 0 To UBound(arr, 1)
        Set topJSON = New JSON
        For p = 0 To 9 'UBound(arr, 2) - 1
            'tempARR = Split(arr(i, p), ":")
            'keySTR = tempARR(0): valOBJ = tempARR(1)
            keySTR = keyARR(p): valOBJ = arr(i, p)
            topJSON.addLeaf keySTR, valOBJ
        Next
        Set tempJSON = topJSON.addObj
        tempJSON.Key = "Percent Work Allocation"
        For j = 10 To m2mRAN.Cells.Count + 9
            tempARR = Split(arr(i, j), ":")
            keySTR = tempARR(0): If tempARR(1) <> "" Then valOBJ = CDbl(tempARR(1)) Else valOBJ = Null
            tempJSON.addLeaf keySTR, valOBJ
        Next
        staffJSON.appendArray topJSON
    Next




End Sub

Sub ScanLineItems(WS As Worksheet)

    Dim RAN As Range, cell As Range
    Dim groupRAN As Range, valRAN As Range, qtyRAN As Range, uomRAN As Range
    Dim arr() As Variant, tempARR As Variant
    Dim p As Integer, i As Integer, printSTR As String, iCount As Integer
    Dim lineJSON As JSON, topJSON As JSON, tempJSON As JSON
    Dim keySTR As String, valOBJ As Variant
    
    Set RAN = boxRANGE(WS, "\r_start", "\r_end", "\c_desc")
    Set groupRAN = WS.[\c_group]
    Set valRAN = WS.[\c_val].Offset(0, -1)
    Set qtyRAN = WS.[\c_qt]
    Set uomRAN = WS.[\c_qt].Offset(0, 1)
    
    Set lineJSON = jsonBuilder.addArray(WS.Name & " Items")
    
    For Each cell In RAN
        If Intersect(cell.EntireRow, groupRAN.EntireColumn).Value = 1 Then
            iCount = iCount + 1
        End If
    Next
    
    ReDim arr(iCount - 1, 4)
    Dim keyARR() As Variant
    keyARR() = Array("Description", "Value", "Quantity", "UnitOfMeasure", "CostCode")
    
    For Each cell In RAN
        If Intersect(cell.EntireRow, groupRAN.EntireColumn).Value = 1 Then
            arr(i, 0) = cell.Value
            arr(i, 1) = Intersect(cell.EntireRow, valRAN.EntireColumn).Value
            arr(i, 2) = Intersect(cell.EntireRow, qtyRAN.EntireColumn).Value
            arr(i, 3) = Intersect(cell.EntireRow, uomRAN.EntireColumn).Value
            arr(i, 4) = CStr(Intersect(cell.EntireRow, groupRAN.Offset(0, 4).EntireColumn).Value)
            i = i + 1
        End If
    Next
    
    For i = 0 To UBound(arr, 1)
        Set topJSON = New JSON
        For p = 0 To UBound(arr, 2)
            'tempARR = Split(arr(i, p), ":")
            keySTR = keyARR(p): valOBJ = arr(i, p)
            topJSON.addLeaf keySTR, valOBJ
        Next
        lineJSON.appendArray topJSON
    Next
    
    'Debug.Print printSTR
    
End Sub

Sub projectATTS(attRAN As Range)

    Dim attNAME As String, attVAL As Variant

    attNAME = attRAN.Offset(0, -1).Value
    If attNAME = "Date" Then attVAL = Format(attRAN.Value, "Short Date") Else attVAL = attRAN.Value
    
    paJSON.addLeaf attNAME, attVAL
    

End Sub

Sub DaysFromStart(RAN As Range, KeyName As String)

    Dim settingWS As Worksheet
    Dim WS As Worksheet
    Dim OGstart As Date
    Dim delta As Variant
    
    Set settingWS = masterOBJ.WS
    Set WS = RAN.Parent
    OGstart = settingWS.[\pstart].Value
    
    delta = DateDiff("d", OGstart, RAN.Value)
    
    paJSON.addLeaf KeyName, delta
    

End Sub

Sub CostLabCSV()


    Dim csvWB As Workbook
    Dim beginRAN As Range
    Dim i As Integer, j As Double
    Dim pfrm As progressFRM

    j = 0.3
    Set pfrm = New progressFRM
    pfrm.progressON "CostLab Item Import", "Aggregating Line Items"
    
    grARR = detailITEMS(masterOBJ.grdWS)
    preTOTAL = masterOBJ.WS.[\prelabor]
    conTOTAL = masterOBJ.WS.[\conlabor]
    gcTOTAL = masterOBJ.WS.[\gctotal]
    
    pfrm.progressUPDATE "Creating CSV", 0
    
    Set csvWB = Application.Workbooks.Open(fileName:=itemImportFile, ReadOnly:=True)
    csvWB.Windows(1).Visible = False
    Set beginRAN = csvWB.Worksheets(1).[A1]
    
    Do Until beginRAN.Offset(0, 1).Value = ""
        Set beginRAN = beginRAN.Offset(1, 0)
    Loop
    
    For i = 0 To UBound(grARR, 1)
        beginRAN.Value = grARR(i, 0)
        beginRAN.Offset(0, 2).Value = grARR(i, 1)
        beginRAN.Offset(0, 3).Value = grARR(i, 2)
        beginRAN.Offset(0, 4).Value = grARR(i, 3)
        beginRAN.Offset(0, 12).Value = grARR(i, 4)
        Set beginRAN = beginRAN.Offset(1, 0)
        pfrm.progressUPDATE "Creating CSV", ((i + 1) / UBound(grARR, 1))
    Next
    
    With beginRAN
        .Value = "General Conditions"
        .Offset(0, 2).Value = 1
        .Offset(0, 3).Value = "lsum"
        .Offset(0, 4).Value = gcTOTAL
        .Offset(0, 12).Value = "98 00 00"
    End With
    Set beginRAN = beginRAN.Offset(1, 0)
    
    With beginRAN
        .Value = "Preconstruction Staffing"
        .Offset(0, 2).Value = 1
        .Offset(0, 3).Value = "lsum"
        .Offset(0, 4).Value = preTOTAL
        .Offset(0, 12).Value = "98 11 00"
    End With
    Set beginRAN = beginRAN.Offset(1, 0)
    
    With beginRAN
        .Value = "Construction Staffing"
        .Offset(0, 2).Value = 1
        .Offset(0, 3).Value = "lsum"
        .Offset(0, 4).Value = conTOTAL
        .Offset(0, 12).Value = "98 21 00"
    End With

    pfrm.turnOFF
    
    csvWB.Windows(1).Visible = True
    
End Sub

Function detailITEMS(WS As Worksheet) As Variant

    
    Dim RAN As Range, cell As Range
    Dim groupRAN As Range, valRAN As Range, qtyRAN As Range, uomRAN As Range
    Dim itemCOLL As Collection, valCOLL As Collection, qtyCOLL As Collection, uomCOLL As Collection, ccCOLL As Collection
    Dim arr() As Variant
    Dim iCount As Integer, i As Integer
    
    Set RAN = boxRANGE(WS, "\r_start", "\r_end", "\c_desc")
    Set groupRAN = WS.[\c_group]
    Set valRAN = WS.[\c_val].Offset(0, -1)
    Set qtyRAN = WS.[\c_qt]
    Set uomRAN = WS.[\c_qt].Offset(0, 1)
    Set itemCOLL = New Collection
    Set valCOLL = New Collection
    Set qtyCOLL = New Collection
    Set uomCOLL = New Collection
    Set ccCOLL = New Collection
    
    For Each cell In RAN
        If Intersect(cell.EntireRow, groupRAN.EntireColumn).Value = 1 Then
            itemCOLL.Add cell.Value
            valCOLL.Add Intersect(cell.EntireRow, valRAN.EntireColumn).Value
            qtyCOLL.Add Intersect(cell.EntireRow, qtyRAN.EntireColumn).Value
            uomCOLL.Add Intersect(cell.EntireRow, uomRAN.EntireColumn).Value
            ccCOLL.Add Intersect(cell.EntireRow, groupRAN.Offset(0, 4).EntireColumn).Value
        End If
    Next

    iCount = itemCOLL.Count
    
    ReDim arr(iCount - 1, 4)
    
    For i = 0 To iCount - 1
        arr(i, 0) = itemCOLL(i + 1)
        arr(i, 1) = qtyCOLL(i + 1)
        arr(i, 2) = uomCOLL(i + 1)
        arr(i, 3) = valCOLL(i + 1)
        arr(i, 4) = ccCOLL(i + 1)
    Next
    
    detailITEMS = arr
    
End Function
