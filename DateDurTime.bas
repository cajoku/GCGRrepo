Attribute VB_Name = "DateDurTime"
Option Explicit

Sub UpdateFormulas(cell As Range, WS As Worksheet, monthly As Boolean)
On Error GoTo ehandle

    Dim monthRan As Range, pcell As Range, formRAN As Range, cell2 As Range
    Dim endCell As Range, startcell As Range, durCell As Range
    Dim endcol As Integer
    Dim cellArr As Variant
    
    endcol = WS.Range("\c_durEND").Offset(0, -1).Column
    Set endCell = Cells(cell.row, endcol)
    Set startcell = Intersect(cell.EntireRow, WS.[\c_negStart].Offset(0, 1).EntireColumn)
    Set pcell = Intersect(cell.EntireRow, WS.[\c_perTIME].EntireColumn)
    Set durCell = Cells(cell.row, WS.[\c_jobDur].Column)
    
    Set monthRan = WS.Range(startcell, endCell)
    
    If durCell.Value = 0 Then Exit Sub
    Set formRAN = GetFormulaRange(cell, WS)
    If formRAN.Count < 1 Then Exit Sub
    
    cellArr = formRAN.Value2
            
    If monthly = False Then
        monthRan.ClearContents
        formRAN.Formula = "=" + cell.Address(False, True)
    ElseIf monthly = True Then
        monthRan.ClearContents
        formRAN.Value2 = cellArr
        For Each cell2 In formRAN
            If cell2.Value = "" Then cell2.Value = cell.Value
        Next
        pcell.Formula = "=AVERAGE(" + formRAN.Address(False, False) + ")"
    End If

Exit Sub

ehandle:
    LogError "DateDurTime", "UpdateFormulas", Err.Description, Err
End Sub

Function GetFormulaRange(cell As Range, WS As Worksheet) As Range
On Error GoTo ehandle

    Dim durCell As Range, perCell As Range, startcell As Range, mRan As Range
    Dim fstart As Integer, fend As Integer
    Dim startDate As Date, endDate As Date
    
    Set perCell = Cells(cell.row, Range("\c_perTIME").Column)
    Set mRan = Cells(cell.row, Range("\c_durSTART").Column)
    Set durCell = Cells(cell.row, Range("\c_jobDur").Column)
    Set startcell = Cells(cell.row, Range("\c_jobStart").Column)
    
    startDate = Intersect(cell.EntireRow, WS.[\c_posStart].EntireColumn).Value
    endDate = Intersect(cell.EntireRow, WS.[\c_posEnd].EntireColumn).Value
    
    'added if statement for startcell.value to account for negative start month. Will need to added negative start month functionality at some point 2/19
    If startcell.Value < 0 Then
        fstart = startcell.Value
    ElseIf startcell.Value > 0 And WS.[\negMin].Value = 0 Then
        fstart = startcell.Value - 1
    Else
        fstart = startcell.Value - 1 'if there are neg columns subtract one bc we skip a "0" col
    End If
    
    fend = fstart + durCell.Value
    
    If durCell.Value = "" Then
        Set GetFormulaRange = mRan.Offset(0, fstart)
    Else
        Set GetFormulaRange = WS.Range(mRan.Offset(0, fstart), mRan.Offset(0, fend - 1))
    End If

Exit Function

ehandle:
    LogError "DateDurTime", "GetFormulaRange", Err.Description, Err
End Function

Sub CreateStart(cell As Range, WS As Worksheet, Optional keep As Boolean = False)
On Error GoTo ehandle

    Dim sMonth As Double
    Dim cStart As Date, cEnd As Date, jStart As Date, jEnd As Date, preStart As Date, preEnd As Date
    Dim perRan As Range
    Dim dRan As Range, sRan As Range
    Dim ans As Variant
    
    
    Set perRan = Cells(cell.row, Range("\c_perTIME").Column)
    
    Set dRan = Cells(cell.row, WS.Range("\c_jobDur").Column)
    Set sRan = Cells(cell.row, WS.Range("\c_posStart").Column)
    
    sMonth = cell.Value '+ 1 'to offset 0 for date adds and date diffs
    
    cStart = WS.Range("\cstart").Value
    cEnd = WS.Range("\cend").Value
    preStart = WS.[\pstart].Value
    preEnd = WS.[\pend].Value
    jEnd = Intersect(cell.EntireRow, WS.[\c_posEnd].EntireColumn).Value
    
    If sMonth = 1 Or sMonth = WS.[\pDur].Value * -1 Then GoTo quickEND
    If cell.Value > 0 Then
        jStart = DateAdd("m", sMonth - 1, cStart)
        jStart = DateSerial(DatePart("yyyy", jStart), DatePart("m", jStart), 1)
    ElseIf cell.Value < 0 Then
        jStart = DateAdd("m", sMonth, preEnd)
        jStart = DateSerial(DatePart("yyyy", jStart), DatePart("m", jStart), 1)
    End If
    
    If jStart < preStart Or preStart = 0 Then
        sRan.Value = jStart
        NegColCount DateDiff("m", jStart, preEnd), WS
    End If
    
quickEND:
    If sMonth = 1 Then
        sRan.Formula = "=\cstart"
    ElseIf sMonth = WS.[\pDur].Value * -1 Then
        sRan.Formula = "=\pstart"
    Else
        sRan.Value = jStart
    End If
    If dRan.Value = "" Then
        If keep = True Then
            UpdateFormulas perRan, WS, True
        Else
            UpdateFormulas perRan, WS, False
        End If
    Else
        'If jEnd <> 0 Then dRan.Value = DateDiff("m", sRan.Value, jEnd)  7/23 changed this so that start month changed offset the project duration instead of getting a datedif between new start date & old end date
        If keep = True Then
            CreateFinish dRan, WS, True
        Else
            CreateFinish dRan, WS, False
        End If
    End If
    
Exit Sub
    
ehandle:
    LogError "DateDurTime", "CreateStart", Err.Description, Err
End Sub

Sub CreateFinish(cell As Range, WS As Worksheet, Optional keep As Boolean = False)
On Error GoTo ehandle

    Dim sMonth As Double, duration As Double
    Dim cStart As Date, cEnd As Date, jStart As Date, jEnd As Date, preStart As Date
    Dim perRan As Range
    Dim fRan As Range
    Dim ans As Variant

    
    Set perRan = Cells(cell.row, Range("\c_perTIME").Column)
    
    Set fRan = Cells(cell.row, Range("\c_posEnd").Column)
    
    cEnd = WS.Range("\cend").Value
    jStart = Cells(cell.row, WS.Range("\c_posStart").Column).Value
    sMonth = Cells(cell.row, WS.Range("\c_jobStart").Column).Value
    duration = cell.Value
    jEnd = DateAdd("m", duration, jStart)
    preStart = WS.[\pstart].Value
    
    If jEnd <= cEnd Then
        If sMonth = 1 And duration = WS.[\duration].Value Then
            fRan.Formula = "=\cend"
        ElseIf sMonth = WS.[\pDur].Value * -1 And duration = WS.[\pDur].Value Then
            fRan.Formula = "=\pend"
        Else
            fRan.Value = jEnd
        End If
    ElseIf jEnd > cEnd Then
'        ans = MsgBox("Would you like to change the end date of this entire project?", vbYesNo, "Finish Date Exceeds Project End Date")
'        If ans = vbYes Then
'            fRan.Value = jEnd
'            EnS 1
'            WS.Range("\cend").Value = jEnd
'            EnS 0
'        ElseIf ans = vbNo Then
'            Exit Sub
'        End If
        fRan.Value = jEnd
        DurColumnCount DateDiff("m", WS.[\cstart].Value, jEnd), WS
    End If
    
   
    If keep = True Then
        UpdateFormulas perRan, WS, True
    Else
        UpdateFormulas perRan, WS, False
    End If
    
Exit Sub
    
ehandle:
    LogError "DateDurTime", "CreateFinish", Err.Description, Err
End Sub

Sub monthtester(cell As Range, WS As Worksheet)
On Error GoTo ehandle

    Dim monthRan As Range, formRAN As Range
    Dim endCell As Range, startcell As Range, durCell As Range, smonthCell As Range, originRan As Range, cell2 As Range
    Dim endcol As Integer, temp As Integer
    Dim cellArr As Variant
    
    endcol = WS.Range("\c_durEND").Offset(0, -1).Column
    Set endCell = Cells(cell.row, endcol)
    Set startcell = Cells(cell.row, Range("\c_negStart").Offset(0, 1).Column)
    Set durCell = Cells(cell.row, WS.[\c_jobDur].Column)
    Set smonthCell = Cells(cell.row, WS.[\c_jobStart].Column)
    
    Set monthRan = WS.Range(startcell, endCell)
    Set originRan = GetFormulaRange(cell, WS)
   
    Do Until startcell.Value <> ""
        Set startcell = startcell.Offset(0, 1)
        startcell.Select
    Loop
    
    Do Until endCell.Value <> ""
        Set endCell = endCell.Offset(0, -1)
        endCell.Select
    Loop
    
    Set formRAN = WS.Range(startcell, endCell)
    
    If formRAN.Count > durCell.Value Then
        temp = formRAN.Count - durCell.Value
        If InRange(originRan.Offset(0, -1), formRAN) Then
            durCell.Value = durCell.Value + temp
            smonthCell.Value = Intersect(WS.[\r_start].EntireRow, cell.EntireColumn).Value
                CreateStart smonthCell, WS, True
        ElseIf InRange(originRan.Offset(0, 1), formRAN) Then
            durCell.Value = durCell.Value + temp
                CreateFinish durCell, WS, True
        End If
    ElseIf formRAN.Count < durCell.Value Then
        temp = durCell.Value - formRAN.Count
        If Union(formRAN, formRAN.Offset(0, -1)).Address = originRan.Address Then
            durCell.Value = durCell.Value - temp
            smonthCell.Value = Intersect(WS.[\r_start].EntireRow, formRAN(1).EntireColumn).Value
            CreateStart smonthCell, WS, True
        ElseIf Union(formRAN, formRAN.Offset(0, 1)).Address = originRan.Address Then
            durCell.Value = durCell.Value - temp
            CreateFinish durCell, WS, True
        End If
    End If
    
    For Each cell2 In formRAN
        If cell2.Value = "" Then cell2.Value = 0
    Next
        
Exit Sub
    
ehandle:
    LogError "DateDurTime", "monthtester", Err.Description, Err
End Sub



