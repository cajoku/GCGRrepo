Attribute VB_Name = "ColumnEditors"
Option Explicit
Sub DurColumnCount(RAN As Double, sht As Worksheet)
On Error GoTo ehandle

    Dim durColl As Collection
    Dim headRAN As Range, cell As Range
    
    EnS 0
    sht.Unprotect
    
    Set durColl = New Collection
    
    Set headRAN = boxRANGE(sht, "\c_durSTART", "\r_start", "\c_durEND", "\r_start")
    trimRANGE headRAN, dsRIGHT
    
    For Each cell In headRAN
        durColl.Add cell
    Next
    
    CreateColumns durColl, RAN, sht
    
    basicPROTECT sht, True

Exit Sub

ehandle:
    LogError "ColumnEditors", "DurColumnCount", Err.Description, Err
    basicPROTECT sht, True
End Sub

Private Sub CreateColumns(dCOLL As Collection, dRan As Double, sht As Worksheet)
On Error GoTo ehandle

    Dim startCol As Range, endcol As Range, newCol As Range
    Dim diff As Integer, i As Integer
    
    Set startCol = sht.Range("\c_durSTART").EntireColumn
    Set endcol = sht.Range("\c_durEND").EntireColumn
    
    diff = Abs(dRan - dCOLL.Count)
    
    If dRan > 500 Then Exit Sub
    
    If dRan = 0 Or dRan = 1 Then
        If dRan < dCOLL.Count Then
            For i = 2 To diff ' i = 2 to delete all but 1 month column when duration is cleared or equal to one month
                startCol.Offset(0, 1).Delete
            Next
            Set newCol = sht.[\c_monthDETAIL].EntireColumn
            If newCol.Hidden = True Then
                If Not sht.Shapes("\\moreMONTHdetail").Visible = msoTrue Then
                    sht.Shapes("\\moreMONTHdetail").Visible = msoTrue
                    sht.Shapes("\\lessMONTHdetail").Visible = msoFalse
                End If
            Else
                If Not sht.Shapes("\\moreMONTHdetail").Visible = msoFalse Then
                    sht.Shapes("\\moreMONTHdetail").Visible = msoFalse
                    sht.Shapes("\\lessMONTHdetail").Visible = msoTrue
                End If
            End If
        End If
    ElseIf dRan > dCOLL.Count Then
        Set newCol = sht.[\c_monthDETAIL].EntireColumn
        If newCol.Hidden = False Then sht.[\c_posTemp].EntireColumn.Hidden = False
            sht.[\c_posTemp].EntireColumn.Copy
            sht.Range(endcol, endcol.Offset(0, diff - 1)).EntireColumn.Insert
        sht.[\c_posTemp].EntireColumn.Hidden = True
        If newCol.Hidden = True Then
            If Not sht.Shapes("\\moreMONTHdetail").Visible = msoTrue Then
                sht.Shapes("\\moreMONTHdetail").Visible = msoTrue
                sht.Shapes("\\lessMONTHdetail").Visible = msoFalse
            End If
        Else
            If Not sht.Shapes("\\moreMONTHdetail").Visible = msoFalse Then
                sht.Shapes("\\moreMONTHdetail").Visible = msoFalse
                sht.Shapes("\\lessMONTHdetail").Visible = msoTrue
            End If
        End If
        UpdateHeaders sht
    ElseIf dRan < dCOLL.Count Then
            sht.Range(endcol.Offset(0, -diff), endcol.Offset(0, -1)).EntireColumn.Delete
        Set newCol = sht.[\c_monthDETAIL].EntireColumn
        If newCol.Hidden = True Then
            If Not sht.Shapes("\\moreMONTHdetail").Visible = msoTrue Then
                sht.Shapes("\\moreMONTHdetail").Visible = msoTrue
                sht.Shapes("\\lessMONTHdetail").Visible = msoFalse
            End If
        Else
            If Not sht.Shapes("\\moreMONTHdetail").Visible = msoFalse Then
                sht.Shapes("\\moreMONTHdetail").Visible = msoFalse
                sht.Shapes("\\lessMONTHdetail").Visible = msoTrue
            End If
        End If
        UpdateHeaders sht
    Else
        Exit Sub
    End If
    
Exit Sub
    
ehandle:
    LogError "ColumnEditors", "CreateColumns", Err.Description, Err
End Sub

Sub UpdateHeaders(sht As Worksheet)
On Error GoTo ehandle

    Dim i As Integer
    Dim headRAN As Range, cell As Range
    
    Set headRAN = boxRANGE(sht, "\c_durSTART", "\r_start", "\c_durEND", "\r_start")
    trimRANGE headRAN, dsRIGHT
    
    i = 1
    For Each cell In headRAN
        cell.Value = CStr(i)
        i = i + 1
    Next
    
Exit Sub
    
ehandle:
    LogError "ColumnEditors", "UpdateHeaders", Err.Description, Err
End Sub
Sub NegColCount(RAN As Double, sht As Worksheet)
On Error GoTo ehandle

    Dim headRAN As Range, cell As Range
    Dim negRan As Range, dsRan As Range
    Dim negStr As String, dsStr As String, nStr As String
    Dim negs As Collection
    
    EnS 0
    sht.Unprotect
    
    Set negRan = sht.[\c_negStart]
    negStr = negRan.Name.Name
    Set dsRan = sht.[\c_durSTART]
    dsStr = dsRan.Name.Name
    
    Set negs = New Collection
    

    If negRan.Offset(0, 1).Address <> dsRan.Address Then
        Set headRAN = boxRANGE(sht, "\c_negStart", "\r_start", "\c_durSTART", "\r_start")
        trimRANGE headRAN, dssides
    Else
        Set headRAN = sht.[\c_negStart]
    End If

    For Each cell In headRAN
        negs.Add cell
    Next
    
    If negs.Count = 1 And negs(1).Address = sht.[\c_negStart].Address Then
        CreateNegatives RAN, negs, sht
    ElseIf negs.Count > 1 Then
        CreateNegatives RAN, negs, sht, 1
    Else
        CreateNegatives RAN, negs, sht, 1
    End If

    basicPROTECT sht, True
    
Exit Sub
    
ehandle:
    LogError "ColumnEditors", "NegColCount", Err.Description, Err
    basicPROTECT sht, True
End Sub
Private Sub CreateNegatives(RAN As Double, negs As Collection, sht As Worksheet, Optional counter As Integer = 0)
On Error GoTo ehandle

    Dim negCol As Range, perCol As Range, durStart As Range, jobCOL As Range, jobCOL2 As Range, negTemp As Range
    Dim sMonth As Double
    Dim i As Integer, diff As Integer, j As Integer

    Set negCol = sht.Range("\c_negStart").EntireColumn
    Set negTemp = sht.Range("\c_negTemp").EntireColumn
    Set perCol = sht.Range("\c_perTIME").EntireColumn
    Set durStart = sht.Range("\c_durSTART").EntireColumn
    
    If RAN > 500 Then Exit Sub 'put in as safegaurd for entering cstart first on Setting page & high dur on precon
    If RAN = 0 And negs.Count = 1 Then Exit Sub
    sMonth = RAN
    
    If counter = 0 Then
        diff = Abs(sMonth)
    Else
        diff = Abs(Abs(sMonth) - negs.Count)
    End If
    

    If Abs(sMonth) >= negs.Count Then
        If diff = 0 Then Exit Sub
        negTemp.EntireColumn.Hidden = False
            With negCol
                negTemp.Copy
                sht.Range(.Offset(0, 1), .Offset(0, diff)).Insert
                Application.CutCopyMode = False
                sht.Range(.Offset(0, 1), .Offset(0, diff)).EntireColumn.Hidden = sht.[\c_durSTART].EntireColumn.Hidden
            End With
        negTemp.EntireColumn.Hidden = True
        UpdateNegHeaders sht
    ElseIf Abs(sMonth) < negs.Count Then
        j = Abs(negs.Count - sMonth)
            sht.Range(negCol.Offset(0, 1), negCol.Offset(0, j)).Delete
        UpdateNegHeaders sht
    End If
    
Exit Sub

ehandle:
    LogError "ColumnEditors", "CreateNegatives", Err.Description, Err
End Sub

Sub UpdateNegHeaders(sht As Worksheet)
On Error GoTo ehandle

    Dim headRAN As Range, cell As Range
    Dim i As Integer, j As Integer
    
    Set headRAN = boxRANGE(sht, "\c_negStart", "\r_start", "\c_durSTART", "\r_start")
    trimRANGE headRAN, dssides
    
    If Not headRAN Is Nothing Then
        j = headRAN.Count
        i = -1 * j
        For Each cell In headRAN
            cell.Value = CStr(i)
            i = i + 1
        Next
    End If
    
Exit Sub

ehandle:
    LogError "ColumnEditors", "UpdateNegHeaders", Err.Description, Err
End Sub

Sub CheckNegs(RAN As Range, sht As Worksheet)
On Error GoTo ehandle

    Dim headRAN As Range, cell As Range, negCol As Range ', negEnd As Range
    Dim i As Integer, j As Integer
    Dim sMonth As Double
    Dim ans As Variant
    Dim lastRAN As Range, lastFORM As String

    Set headRAN = boxRANGE(sht, "\c_negStart", "\r_start", "\c_durSTART", "\r_start")
    trimRANGE headRAN, dssides
    Set negCol = sht.Range("\c_negStart").EntireColumn
    'Set negEnd = sht.Range("\c_negEnd").EntireColumn
    
On Error GoTo recover
    If sht.[\c_negStart].Offset(0, 1).Name.Name = sht.[\c_durSTART].Name.Name Then
        j = 0
    Else
recover:
        j = Abs(headRAN.Count - Abs(sht.[\negMin].Value))
    End If
    
On Error GoTo ehandle
    sMonth = RAN.Value
    
    If sMonth > 0 And sMonth < Abs(sht.[\negMin].Value) Then ' <> 0 Then
        ans = MsgBox("Would you like to reduce the preconstruction duration of the entire project?", vbYesNo, "PreCon Start Is Later Than The Earliest Staff Start")
        If ans = vbYes Then
            For i = 1 To Abs(sht.[\negMin].Value) - sMonth
                negCol.Offset(0, 1).Delete
            Next
        ElseIf ans = vbNo Then
            'ran.Formula = lastForm
            Exit Sub
        End If
        
    
'        For i = 1 To j
'            negCol.Offset(0, 1).Delete
'        Next
'        UpdateNegHeaders sht
'    ElseIf sMonth > 0 And sht.[\negMin].Value = 0 Then
'        Exit Sub
'    ElseIf sMonth > 0 And headRan.count > 0 Then
'        For i = 1 To j
'            negCol.Offset(0, 1).Delete
'        Next
    Else
        For i = 1 To j
            negCol.Offset(0, 1).Delete
        Next
        UpdateHeaders sht
    End If
    
    Exit Sub
    
ehandle:
    LogError "ColumnEditors", "CheckNegs", Err.Description, Err
    
End Sub

