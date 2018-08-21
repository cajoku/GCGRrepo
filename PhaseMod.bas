Attribute VB_Name = "PhaseMod"
Option Explicit

Sub addPhase_Click()
On Error GoTo e1

    Dim WS As Worksheet
    Dim phaseRAN As Range
    Dim insertRAN As Range
    Dim cell As Range
    Dim shp As Shape
    Dim i As Single
    Dim stuffCOLL As Collection
    
    
    EnS 0
    
    Set WS = ActiveSheet
    WS.Unprotect
    'Set phaseRAN = Union(WS.[\r_phase], WS.[\r_phase].Offset(1, 0))
    Set phaseRAN = WS.[\r_phase]
    Set insertRAN = WS.[\r_end].Offset(-1, 0).EntireRow
    
    
    Set stuffCOLL = dependentCOLL(insertRAN.Cells(1, 2).Offset(-1, 0))
    
    'For Each cell In stuffCOLL
        With phaseRAN
            .EntireRow.Hidden = False
            WS.Shapes("\phasestaff").Visible = msoTrue
            .EntireRow.Copy
        End With
        'Set otherWS = cell.Parent
        insertRAN.Insert
    For Each cell In stuffCOLL
        'If cell.Parent.Name <> masterOBJ.gantWS.Name Then
'            With cell
'                .EntireRow.Copy
'                .EntireRow.Offset(1, 0).Insert
'                .Parent.Range("\r_phaseitem").EntireRow.Hidden = False
'                .Parent.Range("\r_phaseitem").EntireRow.Copy
'                .Offset(1, 0).EntireRow.PasteSpecial xlPasteFormats
'                .Parent.Range("\r_phaseitem").EntireRow.Hidden = True
'            End With
'        Else
            With cell
                '.Parent.Unprotect
                .EntireRow.Copy
                .EntireRow.Offset(1, 0).Insert
                .Parent.Range("\r_phaseitem").EntireRow.Hidden = False
                .Parent.Range("\r_phaseitem").EntireRow.Copy
                .Offset(1, 0).EntireRow.PasteSpecial xlPasteFormats
                .Parent.Range("\r_phaseitem").EntireRow.Hidden = True
                If .Parent.Name = masterOBJ.gantWS.Name Then
                    For Each shp In .Parent.Shapes
                        If Left(shp.Name, 2) = "\b" Then
                            If shp.TopLeftCell.EntireRow.Cells(1, 1).Address = cell.Offset(1, 0).EntireRow.Cells(1, 1).Address Then shp.Delete
                        End If
                    Next
'                    .Parent.Range("\r_phaseitem").Copy
'                    .Offset(1, -1).PasteSpecial xlPasteFormulas
'                    newGant.createBar .Offset(1, 0)
                End If
                '.Parent.Range("\r_phaseitem").EntireRow.Hidden = True
                'basicPROTECT .Parent, True
            End With
        'End If
    Next
    
    'insertGLOBAL insertRAN.Cells(1, 2).Offset(-1, 0), False, "\r_phase"
    phaseRAN.EntireRow.Hidden = True
    WS.Shapes("\phasestaff").Visible = msoFalse
    Application.CutCopyMode = False
    
    For Each shp In WS.Shapes
        i = Rnd
        If shp.Name = "\phasestaff" Then
            If shp.Visible = msoTrue Then
                shp.Name = shp.Name & i
            End If
        End If
    Next
        
    EnS 1
    basicPROTECT WS, True
Exit Sub
e1:
    LogError "PhaseMod", "addPhase_Click", Err.Description, Err
    EnS 1, , True
    basicPROTECT WS, True

End Sub
Sub deletePHASE()
On Error GoTo e1

    Dim WS As Worksheet
    Dim RAN As Range, cell As Range
    Dim deleCOLL As Collection
    Dim shp As Shape
    
    
    
    Set WS = ActiveSheet
    WS.Unprotect
    Set RAN = Intersect(activecell.EntireRow, WS.[\c_Position].EntireColumn)
    
    EnS 0
    
    Set buttonCls = Nothing
    Set deleCOLL = dependentCOLL(RAN)
    
    For Each shp In WS.Shapes
        If Left(shp.Name, 3) = "\ph" Then
            If shp.TopLeftCell.Address = RAN.Address Then
                shp.Delete
            End If
        End If
    Next
    
    RAN.EntireRow.Delete
    
    For Each cell In deleCOLL
        'cell.Parent.Unprotect
        cell.EntireRow.Delete
        'basicPROTECT cell.Parent, True
    Next
    
    basicPROTECT WS, True
    EnS 1
    
Exit Sub
e1:
    LogError "PhaseMod", "deletePHASE", Err.Description, Err
    EnS 1, , True
End Sub

Sub registerPHASE(RAN As Range)

    Dim WS As Worksheet
    Dim staffRAN As Range
    Dim topRAN As Range
    Dim botRAN As Range
    Dim tempRAN As Range
    Dim cell As Range
    
    Set WS = RAN.Parent
'    Set tempRAN = WS.[\r_tempPhase].EntireRow
    
'    If RAN.Column = WS.[\c_posStart].Column Then
'        Intersect(RAN.EntireColumn, tempRAN).Formula = "=" & RAN.Address
'    ElseIf RAN.Column = WS.[\c_posEnd].Column Then
'        Intersect(RAN.EntireColumn, tempRAN).Formula = "=" & RAN.Address
'    End If
    
    Set topRAN = RAN
    Set botRAN = RAN.Offset(1, 0)
    
    Do Until botRAN.EntireRow.Cells(1, 1).Value = "dp" Or botRAN.row = WS.[\r_almostEnd].row
        Set botRAN = botRAN.Offset(1, 0)
    Loop
    
    Set staffRAN = Range(topRAN, botRAN)
    trimRANGE staffRAN, dsupdown
    
    
    
    If Not staffRAN Is Nothing Then
        For Each cell In staffRAN
'            If cell.Formula <> Intersect(RAN.EntireColumn, tempRAN).Formula Then
'                cell.Formula = Intersect(RAN.EntireColumn, tempRAN).Formula
'            End If
            If cell.Formula = "=IFERROR(" & RAN.Address(True, True) & ","""")" Then
                EnS 1
                cell.Value = RAN.Value
                EnS 0
                cell.Formula = "=IFERROR(" & RAN.Address(True, True) & ","""")"
            End If
        Next
    End If
    

    
End Sub
