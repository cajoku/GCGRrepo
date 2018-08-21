Attribute VB_Name = "StaffDetailMod"
Option Explicit
Private staffCOLL As Collection
Private countCOLL As Collection
Private percentCOLL As Collection

Sub InsertStaff(sdWSran As Range, codeRAN As Variant)
On Error GoTo e1

    Dim ranColl As Collection
    Dim cell As Range, insertRAN As Range, tempRAN As Range, percentRan As Range, staffStart As Range, staffEnd As Range, sMonth As Range, durRan As Range
    Dim topRAN As Range, botRAN As Range, sortRAN As Range
    Dim WS As Worksheet, gcoWS As Worksheet
    Dim shp As Shape
    
    Set WS = sdWSran.Parent
    Set gcoWS = masterWB.Worksheets("GCs Owner")
    Set insertRAN = sdWSran.Item(1) '.Offset(1, 0)
    
    EnS 0
    
    WS.Unprotect
    'gcoWS.Unprotect
    
    Set ranColl = dependentCOLL(insertRAN)
            
    WS.[\r_tempCON].EntireRow.Hidden = False
    WS.[\r_tempPRECON].EntireRow.Hidden = False
    WS.[\r_tempPhase].EntireRow.Hidden = False
    WS.[\r_tempEST].EntireRow.Hidden = False
    
        'For Each cell In ranColl
            If sdWSran.Value = "Precon" Then
                If codeRAN.Offset(0, -1).Value = "\est" Then WS.[\r_tempEST].EntireRow.Copy Else WS.[\r_tempPRECON].EntireRow.Copy
            ElseIf sdWSran.Value = "Construction" Then
                WS.[\r_tempCON].EntireRow.Copy
            Else
                WS.[\r_tempPhase].EntireRow.Copy
                insertRAN.Offset(1, 0).EntireRow.Insert
                Set tempRAN = insertRAN.Offset(1, 0)
                Set percentRan = Intersect(tempRAN.EntireRow, WS.[\c_perTIME].EntireColumn)
                Set staffStart = Intersect(tempRAN.EntireRow, WS.[\c_posStart].EntireColumn)
                Set staffEnd = Intersect(tempRAN.EntireRow, WS.[\c_posEnd].EntireColumn)
                Set sMonth = Intersect(tempRAN.EntireRow, WS.[\c_jobStart].EntireColumn)
                Set durRan = Intersect(tempRAN.EntireRow, WS.[\c_jobDur].EntireColumn)
                percentRan.Value = 0
                staffStart.Formula = "=IFERROR(" & Intersect(sdWSran.EntireRow, WS.[\c_posStart].EntireColumn).Address & ","""")"
                staffEnd.Formula = "=IFERROR(" & Intersect(sdWSran.EntireRow, WS.[\c_posEnd].EntireColumn).Address & ","""")"
                If IsDate(staffStart.Value) And IsDate(staffEnd.Value) Then
                    durRan.Value = cDateDiff("m", staffStart.Value, staffEnd.Value)
                    If Intersect(sdWSran.EntireRow, WS.[\c_posStart].EntireColumn).Value < WS.[\cstart].Value Then
                        sMonth.Value = cDateDiff("m", WS.[\cstart].Value, Intersect(sdWSran.EntireRow, WS.[\c_posStart].EntireColumn).Value)
                    Else
                        sMonth.Value = cDateDiff("m", WS.[\cstart].Value, Intersect(sdWSran.EntireRow, WS.[\c_posStart].EntireColumn).Value) + 1
                    End If
                End If
                GoTo movingon
            End If
                insertRAN.Offset(1, 0).EntireRow.Insert
                Set tempRAN = insertRAN.Offset(1, 0)
                Set percentRan = Intersect(tempRAN.EntireRow, WS.[\c_perTIME].EntireColumn)
                Set staffStart = Intersect(tempRAN.EntireRow, WS.[\c_posStart].EntireColumn)
                Set staffEnd = Intersect(tempRAN.EntireRow, WS.[\c_posEnd].EntireColumn)
                Set sMonth = Intersect(tempRAN.EntireRow, WS.[\c_jobStart].EntireColumn)
                Set durRan = Intersect(tempRAN.EntireRow, WS.[\c_jobDur].EntireColumn)
                percentRan.Value = 1
'            ElseIf sdWSran.Value = "Construction" Then
'                WS.[\r_tempCON].EntireRow.Copy
'                insertRAN.Offset(1, 0).EntireRow.Insert
'                Set tempRAN = insertRAN.Offset(1, 0)
'                Set percentRan = Intersect(tempRAN.EntireRow, WS.[\c_perTIME].EntireColumn)
'                Set staffStart = Intersect(tempRAN.EntireRow, WS.[\c_posStart].EntireColumn)
'                Set staffEnd = Intersect(tempRAN.EntireRow, WS.[\c_posEnd].EntireColumn)
'                Set sMonth = Intersect(tempRAN.EntireRow, WS.[\c_jobStart].EntireColumn)
'                Set durRan = Intersect(tempRAN.EntireRow, WS.[\c_jobDur].EntireColumn)
'                percentRan.Value = 1
'            Else
'                WS.[\r_tempPhase].EntireRow.Copy
'                insertRAN.Offset(1, 0).EntireRow.Insert
'                Set tempRAN = insertRAN.Offset(1, 0)
'                Set percentRan = Intersect(tempRAN.EntireRow, WS.[\c_perTIME].EntireColumn)
'                Set staffStart = Intersect(tempRAN.EntireRow, WS.[\c_posStart].EntireColumn)
'                Set staffEnd = Intersect(tempRAN.EntireRow, WS.[\c_posEnd].EntireColumn)
'                Set sMonth = Intersect(tempRAN.EntireRow, WS.[\c_jobStart].EntireColumn)
'                Set durRan = Intersect(tempRAN.EntireRow, WS.[\c_jobDur].EntireColumn)
'                percentRan.Value = 1
'                staffStart.Formula = "=" & Intersect(sdWSran.EntireRow, WS.[\c_posStart].EntireColumn).Address
'                staffEnd.Formula = "=" & Intersect(sdWSran.EntireRow, WS.[\c_posEnd].EntireColumn).Address
'                If IsDate(staffStart.Value) And IsDate(staffEnd.Value) Then
'                    durRan.Value = cDateDiff("m", staffStart.Value, staffEnd.Value)
'                    If Intersect(sdWSran.EntireRow, WS.[\c_posStart].EntireColumn).Value < WS.[\cstart].Value Then
'                        sMonth.Value = cDateDiff("m", WS.[\cstart].Value, Intersect(sdWSran.EntireRow, WS.[\c_posStart].EntireColumn).Value)
'                    Else
'                        sMonth.Value = cDateDiff("m", WS.[\cstart].Value, Intersect(sdWSran.EntireRow, WS.[\c_posStart].EntireColumn).Value) + 1
'                    End If
'                End If
'            End If
movingon:
            AddStaffProperties tempRAN, codeRAN
            
        For Each cell In ranColl
            With cell
                'If .Parent.Name = masterOBJ.gantWS.Name And sdWSran.Value = "Construction" Then
                    '.Parent.Unprotect
                    .EntireRow.Copy
                    .EntireRow.Offset(1, 0).Insert
                    .Parent.Range("\r_lineitem").EntireRow.Hidden = False
                    .Parent.Range("\r_lineitem").EntireRow.Copy
                    .Offset(1, 0).EntireRow.PasteSpecial xlPasteFormats
                    .Parent.Range("\r_lineitem").EntireRow.Hidden = True
                    If .Parent.Name = masterOBJ.gantWS.Name And sdWSran.Offset(1, 1).Value <> "Precon" Then
'                        For Each shp In .Parent.Shapes
'                            If shp.TopLeftCell.EntireRow.Cells(1, 1).Address = .Offset(1, 0).EntireRow.Cells(1, 1).Address Then shp.Delete
'                        Next
                        '.Offset(1, -1).ClearContents
                        newGant.createBar .Offset(1, 0)
                    End If
                    '.Parent.Range("\r_lineitem").EntireRow.Hidden = True
                    'basicPROTECT .Parent, True
'                ElseIf .Parent.Name = masterOBJ.gantWS.Name And sdWSran.Offset(1, 1).Value <> "Precon" Then
'                    .EntireRow.Copy
'                    .EntireRow.Offset(1, 0).Insert
'                    .Parent.Range("\r_lineitem").EntireRow.Hidden = False
'                    .Parent.Range("\r_lineitem").EntireRow.Copy
'                    .Offset(1, 0).EntireRow.PasteSpecial xlPasteFormats
'                    .Parent.Range("\r_lineitem").EntireRow.Hidden = True
'                    newGant.createBar cell.Offset(1, 0)
'                Else
'                    .EntireRow.Copy
'                    .EntireRow.Offset(1, 0).Insert
'                    .Parent.Range("\r_lineitem").EntireRow.Hidden = False
'                    .Parent.Range("\r_lineitem").EntireRow.Copy
'                    .Offset(1, 0).EntireRow.PasteSpecial xlPasteFormats
'                    .Parent.Range("\r_lineitem").EntireRow.Hidden = True
'                End If
                    
                
            End With
        Next
        
        
    Application.CutCopyMode = False
    
    WS.[\r_tempCON].EntireRow.Hidden = True
    WS.[\r_tempPRECON].EntireRow.Hidden = True
    WS.[\r_tempPhase].EntireRow.Hidden = True
    WS.[\r_tempEST].EntireRow.Hidden = True
    
    UpdateFormulas percentRan, WS, False
'''''TO Determine range for sort
    Set topRAN = sdWSran
    Set sortRAN = sdWSran.Offset(1, 0)
    Do Until sortRAN.Interior.ColorIndex <> -4142
        Set sortRAN = sortRAN.Offset(1, 0)
    Loop
    Set botRAN = sortRAN
    Set sortRAN = Range(topRAN, botRAN)
    SortJobs sortRAN
''''''''''''''''''''''''''''''''

    EnS 1
    basicPROTECT WS, True
    'basicPROTECT gcoWS, True
Exit Sub
e1:
    LogError "ToolKit", "InsertStaff", Err.Description, Err
    basicPROTECT WS, True
    'basicPROTECT gcoWS, True
End Sub

Sub insertPreset(safetySTRING As String)
On Error GoTo e1

    Dim sdWS As Worksheet, codeWS As Worksheet, gcoWS As Worksheet
    Dim preRan As Range, conRan As Range, cell As Range, cell2 As Range, codeRAN As Range, percentRan As Range, safeRAN As Range, sCell As Range, staffCell As Range, sortRAN As Range
    Dim position As String, counter As Integer, percent As Double, newCount As Integer
    Dim conCOLL As Collection, checkCOLL As Collection
    Dim i As Integer, j As Integer, k As Integer
    
    Set sdWS = masterWB.Worksheets("Staff Detail")
    Set codeWS = masterWB.Worksheets("Code")
    Set gcoWS = masterWB.Worksheets("GCs Owner")
    
    
    EnS 0
    
    sdWS.Unprotect
    'gcoWS.Unprotect
    
    Set conRan = boxRANGE(sdWS, "\r_constr", "\c_Position")
    Set safeRAN = boxRANGE(sdWS, "\r_constr", "\r_almostEnd", "\c_Position")
    trimRANGE safeRAN, dsupdown
    
    Set conCOLL = dependentCOLL(conRan)
    
    formatSTR safetySTRING
    
    Set checkCOLL = New Collection
    
    If Not safeRAN Is Nothing Then
        Set safeRAN = safeRAN.Offset(0, -1)
        For k = 1 To staffCOLL.Count
            newCount = 0
            For Each sCell In safeRAN
                If sCell.Value = "s" Then
                    If sCell.Offset(0, 1).Value = codeWS.Range(staffCOLL(k)).Offset(0, 1).Value Then
                        newCount = newCount + 1
                        Intersect(sCell.EntireRow, sdWS.[\c_perTIME].EntireColumn).Value = percentCOLL(k)
                    End If
                End If
            Next
            checkCOLL.Add newCount
        Next
    End If
    
    If staffCOLL.Count <> 0 Then
        For j = 1 To staffCOLL.Count
            sdWS.[\r_tempCON].EntireRow.Hidden = False
            position = staffCOLL(j)
            If checkCOLL.Count = countCOLL.Count Then
                counter = countCOLL(j) - checkCOLL(j)
            Else
                counter = countCOLL(j)
            End If
            percent = percentCOLL(j)
            Set codeRAN = codeWS.Range(position)
            
            For i = 1 To counter
                    sdWS.[\r_tempCON].EntireRow.Copy
                    conRan.Offset(1, 0).EntireRow.Insert
                    conRan.Offset(1, -1).Value = "s"
                    Application.CutCopyMode = False
                    AddStaffProperties conRan.Offset(1, 0), codeRAN
                    Set percentRan = Intersect(conRan.Offset(1, 0).EntireRow, sdWS.[\c_perTIME].EntireColumn)
                    If percent <> 0 Then
                        percentRan.Value = percent
                    End If
                For Each cell2 In conCOLL
                    cell2.Parent.Unprotect
                    cell2.EntireRow.Copy
                    cell2.Offset(1, 0).EntireRow.Insert
                    cell2.Parent.Range("\r_lineitem").EntireRow.Hidden = False
                    cell2.Parent.Range("\r_lineitem").EntireRow.Copy
                    cell2.Offset(1, 0).EntireRow.PasteSpecial xlPasteFormats
                    cell2.Parent.Range("\r_lineitem").EntireRow.Hidden = True
                    Application.CutCopyMode = False
                    If cell2.Parent.Name = masterOBJ.gantWS.Name Then newGant.createBar cell2.Offset(1, 0)
                    'basicPROTECT cell2.Parent, True
                Next
    
            Next
            sdWS.[\r_tempCON].EntireRow.Hidden = True
            If Not percentRan Is Nothing Then UpdateFormulas percentRan, sdWS, False
        Next
    End If
      
    Set safeRAN = conRan.Offset(1, 0)
    Do Until safeRAN.Interior.ColorIndex <> -4142
        Set safeRAN = safeRAN.Offset(1, 0)
    Loop
    Set safeRAN = Range(conRan, safeRAN)
    SortJobs safeRAN
        
    EnS 1
    
    basicPROTECT sdWS, True
    'basicPROTECT gcoWS, True
    
Exit Sub
e1:
    LogError "ToolKit", "InsertPreset", Err.Description, Err
    basicPROTECT sdWS, True
    'basicPROTECT gcoWS, True
    
End Sub

Sub formatSTR(safetySTRING As String)
On Error GoTo e1

    Dim stringARR() As String
    Dim pARR() As String, pARR2() As String
    Dim i As Integer, j As Integer
    
    Set staffCOLL = New Collection
    Set countCOLL = New Collection
    Set percentCOLL = New Collection
    
    If InStr(safetySTRING, ";") Then
        stringARR = Split(safetySTRING, ";")
        For i = 0 To UBound(stringARR)
            pARR = Split(stringARR(i), ",")
            staffCOLL.Add pARR(0)
            countCOLL.Add pARR(1)
            percentCOLL.Add pARR(2)
        Next
    Else
        If safetySTRING <> "" Then
            pARR = Split(safetySTRING, ",")
            staffCOLL.Add pARR(0)
            countCOLL.Add pARR(1)
            percentCOLL.Add pARR(2)
        End If
    End If
    
Exit Sub
e1:
    LogError "ToolKit", "formatSTR", Err.Description, Err

End Sub

Function safetySTR(tblRAN As Range, valueRAN As Range, region As String) As String
On Error GoTo e1

    Dim WS As Worksheet
    Dim hRan As Range, tran As Range
    Dim pcost As Double
    Dim cell As Range
    Dim i As Integer
    Dim imax As Integer
    
    
    Set hRan = Range(tblRAN.Cells(1, 2), tblRAN.Cells(1, 2).Offset(0, tblRAN.Columns.Count - 2))
    Set tran = hRan.Find(What:=region)

    pcost = valueRAN.Value
    Set cell = tblRAN.Cells(2, 1)
    For i = 0 To tblRAN.Rows.Count - 2
        If cell.Offset(i, 0).Value > pcost Then
            safetySTR = Intersect(cell.Offset(i - 1, 0).EntireRow, tran.EntireColumn).Value
            Exit For
        End If
        If i = tblRAN.Rows.Count - 2 Then
            safetySTR = Intersect(cell.Offset(i, 0).EntireRow, tran.EntireColumn).Value
        End If
    Next
    
Exit Function
e1:
   LogError "ToolKit", "safetySTR", Err.Description, Err
End Function

Sub AddStaffProperties(RAN As Range, codeRAN As Variant)
On Error GoTo e1

    Dim sdWS As Worksheet, cWS As Worksheet
    Dim posRan As Range, salRan As Range, carRan As Range, rankRan As Range, orderRan As Range
    Dim posRan2 As Range, salRan2 As Range, carRan2 As Range, rankRan2 As Range, orderRan2 As Range
    
    Set sdWS = RAN.Parent
    Set cWS = codeRAN.Parent
    
    Set posRan = sdWS.Cells(RAN.row, sdWS.[\c_Position].Column)
    Set salRan = sdWS.Cells(RAN.row, sdWS.[\salary].Column)
    Set carRan = sdWS.Cells(RAN.row, sdWS.[\auto].Column)

    Set orderRan = sdWS.Cells(RAN.row, sdWS.[\c_order].Column)
    
    Set posRan2 = cWS.Cells(codeRAN.row, cWS.[\position].Column)
    Set salRan2 = cWS.Cells(codeRAN.row, cWS.[\salary].Column)
    Set carRan2 = cWS.Cells(codeRAN.row, cWS.[\auto].Column)
    Set orderRan2 = cWS.Cells(codeRAN.row, cWS.[\order].Column)
    
    posRan.Value = posRan2.Value
    salRan.Value = salRan2.Value
    carRan.Value = carRan2.Value
    orderRan.Value = orderRan2.Value

Exit Sub
e1:
    LogError "ToolKit", "AddStaffProperties", Err.Description, Err
End Sub

Sub SortJobs(sdRan As Range)
'On Error GoTo e1

    Dim sdWS As Worksheet
    Dim sortRAN As Range, sortRan2 As Range, rankRan As Range, rankRan2 As Range
    
    Set sdWS = sdRan.Parent
    
    trimRANGE sdRan, dsupdown
    
    Set rankRan = Intersect(sdRan.EntireRow, sdWS.[\c_order].EntireColumn)
    
'    Set sortRAN = boxRANGE(sdWS, "\r_precon", "\r_constr", "\c_Position", "\c_rateEnd")
'    trimRANGE sortRAN, dsUPDOWN
'
'    Set rankRan = boxRANGE(sdWS, "\r_precon", "\r_constr", "\c_order")
'    trimRANGE rankRan, dsUPDOWN
'
'    Set sortRan2 = boxRANGE(sdWS, "\r_constr", "\r_end", "\c_Position", "\c_rateEnd")
'    trimRANGE sortRan2, dsUPDOWN
'
'    Set rankRan2 = boxRANGE(sdWS, "\r_constr", "\r_end", "\c_order")
'    trimRANGE rankRan2, dsUPDOWN
    
    On Error Resume Next
'    sortRAN.Sort rankRan, xlAscending
'    sortRan2.Sort rankRan2, xlAscending
    sdRan.EntireRow.Sort rankRan, xlAscending
Exit Sub
e1:
    LogError "ToolKit", "SortJobs", Err.Description, Err
End Sub

