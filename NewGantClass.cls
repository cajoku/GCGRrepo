VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NewGantClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_barH As Double
Private p_barColor As Long
Private p_borderW As Double
Private p_borderColor As Long
'
Public anchorRow As Range
Public barShp As Shape
'
Public t0Col As Range
Public durCol As Range
Public startCol As Range
Public pworkCol As Range
Private WithEvents WS As Worksheet
Attribute WS.VB_VarHelpID = -1
Private WithEvents staffWS As Worksheet
Attribute staffWS.VB_VarHelpID = -1

'Private barColl As Collection
Private barNum As Integer
'
Private tZeroDefined As Boolean
Private wsRegistered As Boolean
'Private isMasterGantt As Boolean



Private Sub sOff()
    Application.ScreenUpdating = False
End Sub


Private Sub sOn()
    Application.ScreenUpdating = True
End Sub


Private Sub updateDesign()
On Error GoTo e1

    Dim var
    Dim shp As Shape
    sOff
    For Each var In barColl
        With var.barShp
            .Fill.ForeColor.RGB = p_barColor
            .Line.Weight = p_borderW
            .Line.ForeColor.RGB = p_borderColor
            .Height = p_barH
        End With
        var.centerOnRow
    Next
    sOn

    Exit Sub
e1:
    LogError "NewGantClass", "updateDesign", Err.Description, Err
End Sub

Sub centerOnRow()
On Error GoTo e1

    sOff
    With barShp
        .Top = anchorRow.Top + (anchorRow.Height / 2 - .Height / 2)
    End With
    sOn
    
    Exit Sub
e1:
    LogError "NewGantClass", "centerOnRow", Err.Description, Err
End Sub

'==============================================================================
'                   PROPERTIES
'==============================================================================

'SETS

Property Let myFillColorNumber(colorNum As Long)

    p_barColor = colorNum
    updateDesign

End Property


Property Let myFillColorRGB(R As Integer, G As Integer, B As Integer)

    p_barColor = RGB(R, G, B)
    updateDesign

End Property

Property Let myBorderColorNumber(colorNum As Long)

    p_borderColor = colorNum
    updateDesign

End Property

Property Let myBorderColorRGB(R As Integer, G As Integer, B As Integer)

    p_borderColor = RGB(R, G, B)
    updateDesign

End Property


Property Let myBorderWeight(Weight As Double)

    p_borderW = Weight
    updateDesign

End Property


Property Let myBarHeight(Height As Double)

    p_barH = Height
    updateDesign

End Property

Property Set myWS(sht As Worksheet)
    Set WS = sht
End Property

'GETS

Property Get TzeroColumn() As Range

    Set TzeroColumn = t0Col
    
End Property

'==============================================================================
'                   CONSTRUCTORS
'==============================================================================

Sub createBar(anyRangeInRow As Range)
On Error GoTo e1

    Dim ganttBar As NewGantClass
    Dim shp As Shape, shpL As Double, shpW As Double
    Dim cell As Range, cell2 As Range, pcell As Range
    Dim aRan As Range
    Dim fillcolor As Long
    
    'If Not wsRegistered Then Exit Sub
    Application.Calculate
    sOff
    
    fillcolor = 7434613
    Set ganttBar = New NewGantClass
    Set aRan = anyRangeInRow.Cells(1, 1).EntireRow
    Set ganttBar.anchorRow = aRan
    
    On Error Resume Next
    Set pcell = Intersect(aRan, pworkCol)
    
    Set cell = Intersect(aRan, t0Col).Offset(Intersect(aRan, startCol).Value - 1)
    If cell Is Nothing Then Set cell = Intersect(aRan, t0Col)

    shpL = cell.Left
    
    On Error Resume Next
    Set cell2 = cell.Offset(0, Intersect(aRan, durCol).Value - 1) 'TODO Validate only integers in durCol & startCol)
    If cell2 Is Nothing Then Set cell2 = Intersect(aRan, t0Col)
    
    On Error GoTo e1
    shpW = cell2.Left + cell2.Width - cell.Left
    If shpW = 0 Then shpW = 1
    Set shp = WS.Shapes.AddShape(msoShapeRectangle, shpL, 0, shpW, p_barH)
    inheritParentProps shp
    barNum = barNum + 1
    shp.Name = "\bar_" & barNum 'aRan.row
    
    With ganttBar
        Set .barShp = shp
        Set .myWS = WS
        Set .durCol = durCol
        Set .startCol = startCol
        Set .t0Col = t0Col
        Set .pworkCol = pworkCol
        If pcell.Value <> 1 And aRan.Cells(1, 1).Value <> "dp" Then .barShp.Fill.ForeColor.RGB = fillcolor: .barShp.Line.ForeColor.RGB = fillcolor
        If aRan.Cells(1, 1).Value = "dp" Then .barShp.Fill.ForeColor.RGB = SCCIred: .barShp.Line.ForeColor.RGB = SCCIred
    End With
    ganttBar.centerOnRow
    
    barColl.Add ganttBar
    sOn
    
    Exit Sub
e1:
    LogError "NewGantClass", "createBar", Err.Description, Err
    Exit Sub

e2:
    If cell Is Nothing Then Set cell = Intersect(aRan, t0Col)
    Resume Next
    
End Sub

Sub registerBars()
On Error GoTo e1

    Dim ganttBar As NewGantClass
    Dim shp As Shape, shpL As Double, shpW As Double
    Dim cell As Range, cell2 As Range
    Dim aRan As Range
    
    sOff
    
        For Each shp In WS.Shapes
            If Left(shp.Name, 4) = "\bar" Then
                Set ganttBar = New NewGantClass
                Set aRan = shp.TopLeftCell.Cells(1, 1).EntireRow
                Set ganttBar.anchorRow = aRan
                barNum = barNum + 1
                shp.Name = "\bar_" & barNum
                With ganttBar
                    Set .barShp = shp
                    Set .myWS = WS
                    Set .durCol = durCol
                    Set .startCol = startCol
                    Set .t0Col = t0Col
                    Set .pworkCol = pworkCol
                End With
                barColl.Add ganttBar
            End If
        Next

    sOn
    
    Exit Sub
e1:
    LogError "NewGantClass", "registerBars", Err.Description, Err

End Sub

Private Sub inheritParentProps(bar As Shape)
On Error GoTo e1

    With bar
        .Height = p_barH
        .Line.Weight = p_borderW
        .Line.ForeColor.RGB = p_borderColor
        .Fill.ForeColor.RGB = p_barColor
    End With
    
    Exit Sub
e1:
    LogError "NewGantClass", "inheritParentProps", Err.Description, Err
End Sub


Sub registerGanttWS(ganttWS As Worksheet, AnyRangeInStartCol As Range, Optional AnyRangeInDurCol As Range, Optional AnyRangeInTzeroColumn As Range, Optional AnyRangeInPWorkCol As Range)
On Error GoTo e1

    Set WS = ganttWS
    
    Set startCol = AnyRangeInStartCol.Cells(1, 1).EntireColumn
    
    If AnyRangeInDurCol Is Nothing Then
        Set durCol = startCol.Offset(0, 1)
    Else
        Set durCol = AnyRangeInDurCol.Cells(1, 1).EntireColumn
    End If
    
    If AnyRangeInPWorkCol Is Nothing Then
        Set pworkCol = startCol.Offset(0, 2)
    Else
        Set pworkCol = AnyRangeInPWorkCol.Cells(1, 1).EntireColumn
    End If
    
    If AnyRangeInTzeroColumn Is Nothing Then
        Set t0Col = startCol.Offset(0, 4)
    Else
        Set t0Col = AnyRangeInTzeroColumn.Cells(1, 1).EntireColumn
    End If
    

    If Not isMasterGantt And WS.Shapes.Count > 0 Then
        Set barColl = New Collection
        isMasterGantt = True
        registerBars
    Else
        Set barColl = New Collection
        isMasterGantt = True
    End If
    
    Exit Sub
e1:
    LogError "NewGantClass", "registerGanttWS", Err.Description, Err

End Sub

Sub updateBar()
On Error GoTo e1

    Dim shpL As Double, shpW As Double
    Dim cell As Range, cell2 As Range, pcell As Range
    Dim fillcolor As Long
    
    If Not isMasterGantt Then Exit Sub
    
    
    fillcolor = 7434613
    'sOff
    
    If Not anchorRow Is Nothing Then
    
        On Error Resume Next
        Set pcell = Intersect(anchorRow, pworkCol)
        
        Set cell = Intersect(anchorRow, t0Col).Offset(0, Intersect(anchorRow, startCol).Value - 1) 'TODO Validate only integers in durCol & startCol)
        If cell Is Nothing Then Set cell = Intersect(anchorRow, t0Col)
        
        shpL = cell.Left
        
        On Error Resume Next
        Set cell2 = cell.Offset(0, Intersect(anchorRow, durCol).Value - 1) 'TODO Validate only integers in durCol & startCol)
        If cell2 Is Nothing Then Set cell2 = Intersect(anchorRow, t0Col)
        
        On Error GoTo e1
        shpW = cell2.Left + cell2.Width - cell.Left
        If shpW = 0 Then shpW = 1
        
        With barShp
            .Left = shpL
            .Width = shpW
            .Top = anchorRow.Top + (anchorRow.Height / 2 - .Height / 2)
            If pcell.Value <> 1 And anchorRow.Cells(1, 1).Value <> "dp" Then
                .Fill.ForeColor.RGB = fillcolor: .Line.ForeColor.RGB = fillcolor
            ElseIf anchorRow.Cells(1, 1).Value = "dp" Then
                .barShp.Fill.ForeColor.RGB = SCCIred: .barShp.Line.ForeColor.RGB = SCCIred
            Else
                inheritParentProps barShp
            End If
        End With
    'sOn
    End If
    
    Exit Sub
e1:
    'LogError "NewGantClass", "updateBar", Err.Description, Err
    Exit Sub
e2:
    If cell Is Nothing Then Set cell = Intersect(anchorRow, t0Col)
    Resume Next
    
End Sub

Sub NumberColumns(weeks As Integer)
On Error GoTo e1

    Dim i As Integer
    Dim countRAN As Range
    Dim lastCol As Range
    Dim tempRAN As Range
    Dim newRAN As Range
    Dim deleteRAN As Range
    Dim delta As Integer
    Dim firstCol As Range
    
    Set firstCol = WS.[\colStart]
    Set lastCol = WS.[\c_endcol]
    Set countRAN = Range(firstCol, lastCol)
    trimRANGE countRAN, dsRIGHT
    Set tempRAN = WS.[\c_gtemp]
    Set deleteRAN = lastCol.Offset(0, -1)
    EnS 0
    
    
    
    'firstCol.EntireRow.ClearContents
    'Range(firstCol, firstCol.End(xlToRight)).ClearContents
    delta = (weeks) - countRAN.Columns.Count
    
    'For i = 1 To Abs(delta)
        If delta >= 1 Then
            'lastCol.EntireColumn.Insert
            WS.Range(lastCol, lastCol.Offset(0, delta - 1)).EntireColumn.Insert
        ElseIf delta < 1 Then
            WS.Range(deleteRAN.Offset(0, delta + 1), deleteRAN).EntireColumn.Delete
            'lastCol.Offset(0, -1).EntireColumn.Delete
            'If i = Abs(delta) Then Set deleteRAN = Range(deleteRAN, lastCol.Offset(0, -1)): deleteRAN.EntireColumn.Delete
        End If
    'Next

    Set newRAN = Range(firstCol.Offset(0, countRAN.Columns.Count), lastCol)
    trimRANGE newRAN, dsRIGHT
    
    If delta >= 1 Then
    With tempRAN
        .EntireColumn.Hidden = False
        .EntireColumn.Copy
        newRAN.EntireColumn.PasteSpecial xlPasteFormats
        .EntireColumn.Copy
        newRAN.EntireColumn.PasteSpecial xlPasteFormulas
        .EntireColumn.Hidden = True
    End With
    End If
    Application.CutCopyMode = False

    EnS 1
    
    Exit Sub
e1:
    LogError "NewGantClass", "NumberColumns", Err.Description, Err
    EnS 1, , True
End Sub


'==============================================================================
'                   INITIALIZE/TERMINATE
'==============================================================================

Private Sub Class_Initialize()
On Error GoTo e1

    'default values
    p_barColor = RGB(0, 44, 119)
    p_borderColor = RGB(0, 54, 144)
    p_borderW = 1
    p_barH = 8

    If Not isMasterGantt Then registerGanttWS masterOBJ.gantWS, masterOBJ.gantWS.[\c_gstart], masterOBJ.gantWS.[\c_gdur], masterOBJ.gantWS.[\c_t0]
    'Set barColl = New Collection
    Set staffWS = masterOBJ.sdWS
    'If masterOBJ.gantWS.Visible <> xlVisible Then If staffWS.[\duration].Value <> 0 Then masterOBJ.gantWS.Visible = xlSheetVisible

    Exit Sub
e1:
    LogError "NewGantClass", "Class_Initialize", Err.Description, Err
End Sub


Private Sub Class_Terminate()
    'isMasterGantt = False
End Sub


Private Sub staffWS_Change(ByVal Target As Range)
On Error GoTo e1

    Dim conDUR As Range
    Dim conDUR2 As Range
    Dim conStart As Range
    Dim conEnd As Range
    
    Set conDUR = staffWS.[\wdur]
    Set conDUR2 = staffWS.[\duration]
    Set conStart = staffWS.[\cstart]
    Set conEnd = staffWS.[\cend]
    
    If Union(Target, conDUR).Address = conDUR.Address Then
        If conDUR.Value <> 0 Then NumberColumns conDUR.Value
    ElseIf Union(Target, conDUR2).Address = conDUR2.Address Then
        If conDUR2.Value <> 0 Then Application.Calculate: NumberColumns conDUR.Value
    End If
    
    Exit Sub
e1:
    LogError "NewGantClass", "staffWS_Change", Err.Description, Err
End Sub


Private Sub WS_Activate()
On Error GoTo e1
    

    If Not isMasterGantt Then Exit Sub
    

    If barColl.Count <= 0 Then Exit Sub
    

    sOff
    
    Me.updateBar
    
    sOn
    
    Exit Sub
e1:

End Sub

'Private Sub WS_Calculate()
'On Error GoTo e1
'
'    Dim cell As Range
'    Dim gcls As Variant
'
'
'    If Not isMasterGantt Then Exit Sub
'
'    On Error GoTo e2
'    If barColl.Count <= 0 Then Exit Sub
'
'    On Error GoTo e1
'    sOff
'
'    'For Each gcls In barColl
'        'gcls.updateBar
'    'Next
'    Me.updateBar
'
'    sOn
'
'
'    Exit Sub
'e1:
'    'Debug.Print "error from ws_calc"
'    Exit Sub
'e2:
'    registerBars
'    Resume Next
    
'End Sub


Private Sub WS_Change(ByVal Target As Range)
On Error GoTo e1

    Dim cell As Range

    
    sOff
    For Each cell In Target.Cells
        If Not anchorRow Is Nothing Then
            If cell.Address = Intersect(anchorRow, durCol).Address Or cell.Address = Intersect(anchorRow, startCol).Address Or cell.Address = Intersect(anchorRow, pworkCol).Address Then
                If IsNumeric(cell.Value) And cell.Value > 0 Then
                    If cell.Value / 1 <> cell.Value Then
                        Application.EnableEvents = False
                        cell.Value = Int(cell.Value)
                        Application.EnableEvents = False
                    End If
                    updateBar
                End If
                Exit For
            End If
            'updateBar
        End If
    Next
    sOn


Exit Sub
e1:
    'LogError "NewGantClass", "WS_Change", Err.Description, Err
End Sub




