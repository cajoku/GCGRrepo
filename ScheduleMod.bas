Attribute VB_Name = "ScheduleMod"
Option Explicit

Sub addTemp_Click()
On Error GoTo e1

    Dim sht As Worksheet, destWS As Worksheet
    Dim colTemp As Range, colEnd As Range
    Dim dCOLL As Collection, cell As Range

    Set sht = ActiveSheet
    Set destWS = masterOBJ.grdWS 'masterWB.Worksheets("GRs Detail")
    Set colTemp = sht.[\c_schedtemp]
    Set colEnd = sht.[\c_schedend]

    EnS 0
    sht.Unprotect
    colTemp.Cells.MergeArea.EntireColumn.Hidden = False
    colTemp.Cells.MergeArea.EntireColumn.Copy
    colEnd.EntireColumn.Insert
    colTemp.Cells.MergeArea.EntireColumn.Hidden = True
    
    Application.CutCopyMode = False
    

            With destWS
                .Unprotect
                .[\r_scheditem].EntireRow.Hidden = False
                Set dCOLL = dependentCOLL(Intersect(.[\r_insertsched].Offset(-1, 0).EntireRow, .[\c_desc].EntireColumn))
                
                     .[\r_insertsched].EntireRow.Insert
                     .[\r_scheditem].EntireRow.Copy
                     .[\r_insertsched].Offset(-1, 0).EntireRow.PasteSpecial xlPasteFormats
                     .[\r_insertsched].Offset(-1, 0).EntireRow.PasteSpecial xlPasteFormulas
                     .[\r_insertsched].Offset(-1, 0).EntireRow.Hidden = .[\r_insertsched].Offset(-2, 0).EntireRow.Hidden
                     .[\r_scheditem].EntireRow.Hidden = True
                     Application.CutCopyMode = False
                     
                    sht.Activate
                    add2GRs
                For Each cell In dCOLL
                     cell.EntireRow.Copy
                     cell.Offset(1, 0).EntireRow.Insert
                     'cell.Parent.[\r_lineitem].EntireRow.Hidden = False
                     'cell.Parent.[\r_lineitem].EntireRow.Copy
                     'cell.Offset(-1, 0).EntireRow.PasteSpecial xlPasteFormats
                     'cell.Parent.[\r_lineitem].EntireRow.Hidden = True
                     Application.CutCopyMode = False
                 Next
            End With
        


    
    basicPROTECT destWS, True
    basicPROTECT sht, True
    EnS 1
            
Exit Sub
e1:
    LogError "ScheduleMod", "addTemp_Click", Err.Description, Err
    EnS 1, , True
    basicPROTECT destWS, True
    basicPROTECT sht, True
    
End Sub

Sub add2GRs()
On Error GoTo e1

    Dim destWS As Worksheet, WS As Worksheet
    Dim cell As Range
    Dim headers As Range
    Dim schedRAN As Range, carryRAN As Range
    Dim headFORM As String, carryFORM As String
    Dim headCOLL As Collection
    Dim carryCOLL As Collection
    Dim i As Integer
    
    
    Set WS = masterOBJ.schedWS 'masterWB.Worksheets("Labor")
    Set destWS = masterOBJ.grdWS 'masterWB.Worksheets("GRs Detail")
    Set headCOLL = New Collection
    Set carryCOLL = New Collection
    
    Set headers = boxRANGE(WS, "\r_header", "\c_schedEnd", "\c_schedStart")
    Set headers = headers.Offset(-1, 0)
    trimRANGE headers, dsSIDES
    trimRANGE headers, dsLEFT
    
    Set schedRAN = boxRANGE(destWS, "\r_schedStart", "\c_desc", "\r_insertSched")
    trimRANGE schedRAN, dsupdown
    
    'EnS 0
    For Each cell In headers
        If cell.Value <> "" Then
            headFORM = "='" & headers.Parent.Name & "'!" & cell.Address
            headCOLL.Add headFORM
            Set carryRAN = Intersect(cell.EntireColumn, WS.[\r_carry].EntireRow).Offset(0, 3)
            carryFORM = "='" & carryRAN.Parent.Name & "'!" & carryRAN.Address(False, False)
            carryCOLL.Add carryFORM
        End If
    Next
    
    i = schedRAN.Count
    schedRAN(i).Formula = headCOLL(i)
    schedRAN(i).Offset(0, 5).Formula = carryCOLL(i)

    'EnS 1
    
Exit Sub
e1:
    LogError "ScheduleMod", "add2GRs", Err.Description, Err
    'EnS 1, , True
    
End Sub

Sub addITEM_Click()
On Error GoTo e1

    Dim sht As Worksheet, gsht As Worksheet
    Dim inStart As Range, inTemp As Range, inEnd As Range, colTemp As Range, colEnd As Range, colStart As Range, headRAN As Range
    Dim printRAN As Range, cell As Range, endRAN As Range, tempRAN As Range
    Dim itemStr As String, durNum As Double, newdate As Date
    Dim schedItem As Range, durCell As Range
    Dim i As Integer, startweek As Integer, findRAN As Range, bottomRAN As Range
 
    
    Set sht = ActiveSheet
    Set inStart = sht.[\r_start]
    Set inTemp = sht.[\r_schedtemp]
    Set inEnd = sht.[\r_end].EntireRow
    Set colTemp = sht.[\c_schedtemp]
    Set colEnd = sht.[\c_schedend]
    Set colStart = sht.[\c_schedStart]
    Set headRAN = sht.[\r_header]
    Set schedItem = sht.[\scheditem]
    Set durCell = sht.[\scheddur]
    
    Set tempRAN = Intersect(inStart.EntireRow, colStart.EntireColumn)
    
    itemStr = schedItem.Value
    durNum = durCell.Value
    
    If durNum = 0 Then Exit Sub
    
    sht.Unprotect
    EnS 0
    
        '''''''''PlaceHolder for some process to push to GantSHT''''''
    
    If masterOBJ Is Nothing Then masterOBJ = New ClassMaster
    Set gsht = masterOBJ.gantWS
    
    On Error Resume Next
    startweek = Intersect(inStart.Offset(-1, 0).EntireRow, colStart.EntireColumn).Value
    If startweek = 0 Then startweek = 1 Else startweek = startweek + 1
    Set findRAN = gsht.[\r_wheader].EntireRow.Find(What:=startweek, LookIn:=xlValues, LookAt:=xlWhole)
    If Not findRAN Is Nothing Then
        Set bottomRAN = Intersect(gsht.[\r_gbottom].EntireRow, findRAN.EntireColumn)
        With findRAN
            '.EntireColumn.Interior.ColorIndex = 40
            .Offset(-1, 0).Value = itemStr
            .Offset(-1, 0).Orientation = xlUpward
            .Offset(-1, 0).Interior.ColorIndex = 40
            .Offset(-1, 0).Font.ColorIndex = 49
            .Offset(-1, 0).Font.Bold = True
            .Font.ColorIndex = 49
            .Font.Bold = True
'            .AddComment itemStr & " Start"
'            '.Comment.Shape.Width = .Width * 4
'            .Comment.Shape.Height = .Height
'            .Comment.Shape.Line.DashStyle = msoLineDashDot
'            .Comment.Shape.Fill.ForeColor.RGB = RGB(0, 0, 102)
'            .Comment.Shape.TextFrame.Characters.Font.ColorIndex = 2
'            .Comment.Shape.TextFrame.Characters.Font.Bold = True
'            .Comment.Shape.TextFrame.AutoSize = True
        End With
        Set findRAN = Range(findRAN, bottomRAN)
        trimRANGE findRAN, dsbottom
        findRAN.Interior.ColorIndex = 40
        Intersect(findRAN.EntireColumn, gsht.[\r_lineitem].EntireRow).Interior.ColorIndex = 40
        Intersect(findRAN.EntireColumn, gsht.[\r_phaseitem].EntireRow).Interior.ColorIndex = 40
    End If
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error GoTo e1
    
    inTemp.EntireRow.Hidden = False
    For i = 1 To durNum
        inStart.EntireRow.Insert
    Next
    
    Set printRAN = Range(tempRAN.Offset(-1, 0), tempRAN.Offset(-durNum, 0))
    inTemp.EntireRow.Copy
    printRAN.EntireRow.PasteSpecial xlPasteFormulasAndNumberFormats
    inTemp.EntireRow.Copy
    printRAN.EntireRow.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    inTemp.EntireRow.Hidden = True
    
    With printRAN.Offset(0, -1)
        .Merge
        .Value = itemStr
        .Orientation = xlUpward
        .VerticalAlignment = xlVAlignCenter
        .HorizontalAlignment = xlCenter
    End With
    
    schedItem.Cells.MergeArea.ClearContents
    durCell.Cells.MergeArea.ClearContents
    
    
    EnS 1
    basicPROTECT sht, True
    
Exit Sub
e1:
    LogError "ScheduleMod", "addITEM_Click", Err.Description, Err
    EnS 1, , True
    basicPROTECT sht, True
    
End Sub


Sub deleteGRs(deleteSTR As String)
On Error GoTo e1

    Dim destWS As Worksheet, WS As Worksheet
    Dim cell As Range
    Dim headers As Range
    Dim schedRAN As Range, deleteRAN As Range
    Dim headFORM As String
    Dim headCOLL As Collection
    Dim i As Integer
    Dim deleCOL As Collection
    
    Set destWS = masterOBJ.grdWS 'masterWB.Worksheets("GRs Detail")
    destWS.Unprotect
    
    Set schedRAN = boxRANGE(destWS, "\r_schedStart", "\c_desc", "\r_insertSched")
    trimRANGE schedRAN, dsupdown
    
    Set deleteRAN = schedRAN.Find(What:=deleteSTR, LookIn:=xlFormulas)
    
    Set deleCOL = dependentCOLL(deleteRAN)
    For Each cell In deleCOL
        cell.EntireRow.Delete
    Next
    deleteRAN.EntireRow.Delete
    basicPROTECT destWS, True
    
Exit Sub
e1:
    LogError "ScheduleMod", "deleteGRs", Err.Description, Err
    basicPROTECT destWS, True
    
End Sub


Sub deleteSchedItem()
On Error GoTo e1

    Dim sht As Worksheet, gsht As Worksheet
    Dim deleteRAN As Range, findRAN As Range, bottomRAN As Range
    Dim startweek As Integer
    
    Set sht = ActiveSheet
    Set gsht = masterOBJ.gantWS
    Set deleteRAN = Selection
    
    sht.Unprotect
    'gsht.Unprotect
    EnS 0
    
    startweek = Intersect(sht.[\c_schedStart].EntireColumn, deleteRAN.Cells(1, 1).EntireRow).Value
    Set findRAN = gsht.[\r_wheader].EntireRow.Find(What:=startweek, LookIn:=xlValues, LookAt:=xlWhole)
    If Not findRAN Is Nothing Then
        Set bottomRAN = Intersect(gsht.[\r_gbottom].EntireRow, findRAN.EntireColumn)
        With findRAN
            .Interior.ColorIndex = 2
            .Offset(-1, 0).Interior.ColorIndex = 2
            .Offset(-1, 0).ClearContents
            .Font.ColorIndex = 1
            .Font.Bold = False
        End With
        Set findRAN = Range(findRAN.Offset(1, 0), bottomRAN)
        trimRANGE bottomRAN, dsbottom
        findRAN.Interior.ColorIndex = -4142
        Intersect(findRAN.EntireColumn, gsht.[\r_lineitem].EntireRow).Interior.ColorIndex = -4142
        Intersect(findRAN.EntireColumn, gsht.[\r_phaseitem].EntireRow).Interior.ColorIndex = -4142
    End If
    deleteRAN.EntireRow.Delete xlShiftUp
    
    
    EnS 1
    'basicPROTECT gsht, True
    basicPROTECT sht, True
    
Exit Sub
e1:
    LogError "ScheduleMod", "deleteSchedItem", Err.Description, Err
    Debug.Print "For Colby: error on deleting schedule labor item"
    Set buttonCls = Nothing
    EnS 1, , True
    basicPROTECT sht, True
    'basicPROTECT gsht, True
    
End Sub

Sub deleteTemp()
On Error GoTo e1

    Dim sht As Worksheet
    Dim deleteRAN As Range
    Dim deleteSTR As String
    Dim deleCOL As Collection
    Dim cell As Range
    
    Set sht = ActiveSheet
    Set deleteRAN = Selection
    deleteSTR = deleteRAN(1).Address
    
    sht.Unprotect
    EnS 0
    
    deleteGRs deleteSTR
    deleteRAN.EntireColumn.Delete 'xlLeft
    sht.Activate
    basicPROTECT sht, True
    
    EnS 1
    
Exit Sub
e1:
    LogError "ScheduleMod", "deleteTemp", Err.Description, Err
    Debug.Print "For Colby: error on deleting template column"
    Set buttonCls = Nothing
    EnS 1, , True
    basicPROTECT sht, True
    
End Sub


Sub NumberSchedRows(weekAmount As Double)
On Error GoTo e1

    Dim sht As Worksheet
    Dim inStart As Range, inTemp As Range
    Dim existRan As Range, tempRAN As Range
    Dim i As Integer, currentRows As Integer, delta As Integer
    
    Set sht = masterOBJ.schedWS
    Set inStart = sht.[\r_start]
    Set inTemp = sht.[\r_schedtemp]
    Set existRan = boxRANGE(sht, "\r_header", "\r_start", "\c_schedStart")
    trimRANGE existRan, dsupdown
    
    EnS 0
        
    If Not existRan Is Nothing Then currentRows = existRan.Rows.Count
    
    delta = weekAmount - currentRows
    
    If delta > 0 Then
'        For i = 1 To delta
'            inStart.EntireRow.Insert
            sht.Range(inStart, inStart.Offset(delta - 1, 0)).EntireRow.Insert

            Set tempRAN = boxRANGE(sht, "\r_header", "\r_start", "\c_schedStart")
            trimRANGE tempRAN, dsupdown

            inTemp.EntireRow.Hidden = False
            inTemp.EntireRow.Copy
            tempRAN.EntireRow.PasteSpecial xlPasteFormats
            inTemp.EntireRow.Copy
            tempRAN.EntireRow.PasteSpecial xlPasteFormulas

            tempRAN.Offset(0, -1).Merge

            Application.CutCopyMode = False

            inTemp.EntireRow.Hidden = True
        'Next
    Else
        'For i = 1 To Abs(delta)
            'inStart.Offset(-1, 0).EntireRow.Delete
            sht.Range(inStart.Offset(-1, 0), inStart.Offset(-1 + delta + 1, 0)).EntireRow.Delete
        'Next
    End If
    
    EnS 1
    
    Exit Sub
e1:
    LogError "ScheduleMod", "NumberSchedRows", Err.Description, Err
    
End Sub


Sub ToggleSchedule(region As String)
On Error GoTo e1

    Dim codeWS As Worksheet, settingsWS As Worksheet, sWS As Worksheet
    Dim headRAN As Range, codeRAN As Range, regRAN As Range, findRAN As Range
    Dim cell As Range
    Dim i As Integer
    
    'If schedBOOL = False Then
    'Set settingsWS = masterOBJ.WS
    Set codeWS = masterWB.Worksheets("Code")
    Set sWS = masterOBJ.schedWS
    Set headRAN = boxRANGE(sWS, "\r_header", "\c_schedEnd", "\c_schedStart")
    Set headRAN = headRAN.Offset(-1, 0)
    trimRANGE headRAN, dsSIDES
    trimRANGE headRAN, dsLEFT

    Set codeRAN = codeWS.[\headertbl]
    'Set regRAN = settingsWS.[\reg]
    If region = "" Then Exit Sub
    Set findRAN = codeRAN.Find(region)
    
    i = 1
    For Each cell In headRAN
        If cell.Value <> "" Then
            cell.Value = findRAN.Offset(0, i).Value
            i = i + 1
        End If
    Next
    'schedBOOL = True
    'End If

    Exit Sub
e1:
    LogError "ScheduleMod", "ToggleSchedule", Err.Description, Err

End Sub

Sub AddLaborPhase(phaseName As String)

    Dim sht As Worksheet
    Dim inStart As Range, inTemp As Range, inEnd As Range, colStart As Range, colEnd As Range
    Dim findRAN As Range, tempRAN As Range, mergeRAN As Range, printRAN As Range
    Dim i As Integer, currentRows As Integer, delta As Integer
    
    Set sht = masterOBJ.schedWS
    Set inStart = sht.[\r_start]
    Set inTemp = sht.[\r_schedtemp]
    Set inEnd = sht.[\r_end].EntireRow
    Set colEnd = sht.[\c_schedend]
    Set colStart = sht.[\c_schedStart]
    'Set printRAN = Intersect(sht.[\r_header].Offset(1, 0).EntireRow, sht.[\c_merge].Offset(0, -1).EntireColumn)
    Set printRAN = Intersect(sht.[\r_header].Offset(1, 0).EntireRow, sht.[\c_print].EntireColumn)
    Set findRAN = boxRANGE(sht, "\r_header", "\r_start", "\c_schedStart")
    trimRANGE findRAN, dsupdown
    Set mergeRAN = findRAN.Offset(0, -1)
    
    If mergeRAN.MergeCells Then mergeRAN.MergeArea.UnMerge
            
    Do Until printRAN.Value = ""
        Set printRAN = printRAN.Offset(1, 0)
    Loop
    
    printRAN.Value = phaseName
    
    


End Sub


Sub CreateLaborSection(phaseName As String, startweek As Integer, Optional durNum As Integer)
On Error GoTo e1

    Dim sht As Worksheet
    Dim findRAN As Range, tempRAN As Range, mergeRAN As Range, printRAN As Range, bottomRAN As Range, weekRAN As Range, findMergeRAN As Range, merge2RAN As Range, cell As Range
    Dim topCell As Range, botCell As Range, tempItem As String
    Dim oldWeek As Integer, phi As Integer
    
    
    Application.DisplayAlerts = False
    
    Set sht = masterOBJ.schedWS
    Set printRAN = Intersect(sht.[\r_header].Offset(1, 0).EntireRow, sht.[\c_print].EntireColumn)
    Set findRAN = boxRANGE(sht, "\r_header", "\r_start", "\c_schedStart")
    trimRANGE findRAN, dsbottom
    

    Set weekRAN = findRAN.Find(What:=startweek, LookIn:=xlValues)
    If weekRAN Is Nothing Then Exit Sub
    
    Set mergeRAN = findRAN.Offset(0, -1)
    'trimRANGE mergeRAN, dsTOP
    Set weekRAN = weekRAN.Offset(0, -1)
    
    Set findMergeRAN = mergeRAN.Find(What:=phaseName)
    If Not findMergeRAN Is Nothing Then
        oldWeek = findMergeRAN.Offset(0, 1).Value
        If startweek > oldWeek Then
            If findMergeRAN.MergeCells = True Then Intersect(findMergeRAN.MergeArea.EntireRow, printRAN.EntireColumn).ClearContents Else Intersect(findMergeRAN.EntireRow, printRAN.EntireColumn).ClearContents
            If findMergeRAN.MergeCells = True Then findMergeRAN.MergeArea.ClearContents Else findMergeRAN.ClearContents
            phi = findMergeRAN.MergeArea.Rows.Count
            findMergeRAN.MergeArea.UnMerge
            Set findMergeRAN = Range(weekRAN, findMergeRAN.Cells(phi, 1))
            With findMergeRAN
                .Merge
                .Value = phaseName
                .Orientation = xlUpward
                .VerticalAlignment = xlVAlignCenter
                .HorizontalAlignment = xlCenter
                With .Cells(1, 1).Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .Color = 7810048
                End With
                Intersect(.EntireRow, printRAN.EntireColumn).Value2 = phaseName
            End With
            GoTo RemergeCells
        ElseIf startweek < oldWeek Then
            If weekRAN.MergeCells = True Then weekRAN.MergeArea.ClearContents: weekRAN.MergeArea.UnMerge
            If findMergeRAN.MergeCells Then findMergeRAN.MergeArea.ClearContents Else findMergeRAN.ClearContents
            Set findMergeRAN = Range(findMergeRAN, weekRAN)
            With findMergeRAN
                .Merge
                .Value = phaseName
                .Orientation = xlUpward
                .VerticalAlignment = xlVAlignCenter
                .HorizontalAlignment = xlCenter
                With .Cells(1, 1).Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .Color = 7810048
                End With
                Intersect(.EntireRow, printRAN.EntireColumn).Value2 = phaseName
            End With
            GoTo RemergeCells
        End If
    End If
    
    If weekRAN.MergeCells = True Then weekRAN.MergeArea.ClearContents: weekRAN.MergeArea.UnMerge
    
    On Error GoTo e1
    If durNum = 0 Then
        Set bottomRAN = weekRAN
        Do Until bottomRAN.Value <> "" Or bottomRAN.Offset(1, -2).Address = sht.[\r_start].Address Or bottomRAN.Offset(1, 0).MergeCells = True
            Set bottomRAN = bottomRAN.Offset(1, 0)
        Loop
    Else
        Set bottomRAN = weekRAN.Offset(durNum - 1, 0)
    End If
    
    Set merge2RAN = sht.Range(weekRAN, bottomRAN)
    Set tempRAN = Intersect(merge2RAN.EntireRow, printRAN.EntireColumn)
    tempRAN.Value2 = phaseName
    
    With merge2RAN
        With .Cells(1, 1).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = 7810048
        End With
        .Merge
        .Value = phaseName
        .Orientation = xlUpward
        .VerticalAlignment = xlVAlignCenter
        .HorizontalAlignment = xlCenter
    End With
    
RemergeCells:
    
    BalanceMergeSection
            
    Application.DisplayAlerts = True

    Exit Sub
e1:
   LogError "ScheduleMod", "CreateLaborSection", Err.Description, Err
   Application.DisplayAlerts = True
   
End Sub

Sub RemoveMergeSection(phaseName As String)
On Error GoTo e1

    Dim WS As Worksheet
    Dim printRAN As Range, findRAN As Range, mergeRAN As Range, deleteRAN As Range
    
    Set WS = masterOBJ.schedWS
    Set printRAN = Intersect(WS.[\r_header].Offset(1, 0).EntireRow, WS.[\c_print].EntireColumn)
    Set mergeRAN = boxRANGE(WS, "\r_header", "\r_start", "\c_merge")
    trimRANGE mergeRAN, dsupdown

    Set deleteRAN = mergeRAN.Find(What:=phaseName)
    
    If Not deleteRAN Is Nothing Then
        With deleteRAN
            Intersect(.MergeArea.EntireRow, printRAN.EntireColumn).ClearContents
            If .MergeCells = True Then .MergeArea.ClearContents: .MergeArea.UnMerge
        End With
    BalanceMergeSection
    End If
    
    Exit Sub
e1:
    LogError "ScheduleMod", "RemoveMergeSection", Err.Description, Err
End Sub


Private Sub BalanceMergeSection()
On Error GoTo e1

    Dim sht As Worksheet
    Dim topCell As Range, botCell As Range, cell As Range
    Dim tempItem As String
    Dim mergeRAN As Range, printRAN As Range
    
    
    Set sht = masterOBJ.schedWS
    Set printRAN = Intersect(sht.[\r_header].Offset(1, 0).EntireRow, sht.[\c_print].EntireColumn)
    Set mergeRAN = boxRANGE(sht, "\r_header", "\r_start", "\c_merge")
    trimRANGE mergeRAN, dsupdown
    
    For Each cell In Intersect(mergeRAN.EntireRow, printRAN.EntireColumn)
        If cell.Value = "" Then
            Set topCell = cell
            Set botCell = cell
            Do Until topCell.Value <> "" Or topCell.Offset(-1, 0).row = sht.[\r_header].row
                Set topCell = topCell.Offset(-1, 0)
            Loop
            tempItem = topCell.Value
            Do Until botCell.Offset(1, 0).Value <> "" Or botCell.Offset(1, 0).row = sht.[\r_start].row
                Set botCell = botCell.Offset(1, 0)
            Loop
            sht.Range(topCell, botCell).Value2 = tempItem
            Intersect(sht.Range(topCell, botCell).EntireRow, mergeRAN.EntireColumn).Merge
            Intersect(sht.Range(topCell, botCell).EntireRow, mergeRAN.EntireColumn).Value = tempItem
        ElseIf cell.Value <> "" And Intersect(cell.EntireRow, mergeRAN).MergeCells = False Then
            Set topCell = cell
            Set botCell = cell
            tempItem = cell.Value
            Do
                If botCell.Offset(1, 0).Value = "" Then botCell.Offset(1, 0).Value = tempItem: Set botCell = botCell.Offset(1, 0)
                If botCell.Offset(1, 0).Value <> tempItem Then Exit Do Else Set botCell = botCell.Offset(1, 0)
                
            Loop
            sht.Range(topCell, botCell).Value2 = tempItem
            Intersect(sht.Range(topCell, botCell).EntireRow, mergeRAN.EntireColumn).Merge
            Intersect(sht.Range(topCell, botCell).EntireRow, mergeRAN.EntireColumn).Value = tempItem
        End If
    Next
    
    Exit Sub
e1:
    LogError "ScheduleMod", "BalanceMergeSection", Err.Description, Err
End Sub
