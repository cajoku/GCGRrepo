Attribute VB_Name = "GantChart"
Option Explicit

Public gantCls As GantChart_cls
Public gant As Chart
Public x As Shape
Public plus As Shape

Sub MakeChart()
On Error GoTo e1

    Dim WS As Worksheet
    Dim startRAN As Range, durRan As Range, chartRan As Range, nameRAN As Range, tempRAN As Range, cell As Range, formRAN As Range, cell2 As Range
    Dim durArr() As Double, gradientloc As Double, formRanArr() As Variant, pt As Point, gc As GradientStop
    Dim i As Integer, j As Integer, k As Integer, p As Integer, gradientCount As Integer
    Dim L As Long, T As Long, W As Long, H As Long

    If masterOBJ Is Nothing Then onOPEN
    If gantCls Is Nothing Then
        Set gantCls = New GantChart_cls
    ElseIf Not gantCls Is Nothing Then
        Exit Sub
    Else
        Exit Sub
    End If
    
    Application.EnableEvents = False
    
    Set WS = masterWB.Worksheets("Staff Detail")
    WS.Unprotect
    
    L = Application.Left * 0.5
    T = Application.Top * 0.5
    W = Application.UsableWidth * (1 / 3)
    H = Application.UsableHeight * (1 / 3)
    
    Set chartRan = WS.[\chartran]
    Set gant = WS.ChartObjects.Add(L, T, W, H).Chart
    Set startRAN = GetRangeVals(sDate)
    Set durRan = GetRangeVals(dur)
    
    i = 0
    For Each cell In durRan
        ReDim Preserve durArr(i)
        durArr(i) = cell.Value * 30.4167  'convert duration to days. Necessary for gant chart
        i = i + 1
    Next
    
    Set nameRAN = Intersect(startRAN.EntireRow, WS.[\c_posName].EntireColumn)

    With gant
        .HasTitle = True
        .ChartTitle.Text = "Staff Duration Chart"
        .ChartType = xlBarStacked
        .Axes(xlCategory).ReversePlotOrder = True
        .Axes(xlCategory).MajorTickMark = xlNone
        .Axes(xlValue).MinimumScale = WS.[\pstart].Value
        .Axes(xlValue).MaximumScale = WS.[\cend].Value
        '.Axes(xlValue).MajorUnit = 150
        .Axes(xlValue).TickLabels.NumberFormat = "[$-en-US]mmm-yy;@"
        .SeriesCollection.NewSeries
        With .FullSeriesCollection(1)
            .Name = "StartDate"
            .Values = startRAN
            .XValues = nameRAN
            .Format.Fill.Visible = msoFalse
        End With
        .ChartGroups(1).GapWidth = 10
        .SeriesCollection.NewSeries
        With .FullSeriesCollection(2)
            .Name = "Duration"
            .Values = durArr
            .Format.Fill.ForeColor.RGB = SCCIblue
        End With

        .Legend.Delete
        .Parent.Name = "\chart"
    End With

    MakeExitButton gant.Parent.Left, gant.Parent.Top, gant.Parent.Width, gant.Parent.Height, WS
    Application.EnableEvents = True
    basicPROTECT WS, True
    
Exit Sub
    
e1:
    LogError "GantChart", "MakeChart", Err.Description, Err
    Set gantCls = Nothing
    Application.EnableEvents = True
    basicPROTECT WS, True
    
End Sub

Function GetRangeVals(rtype As dsRangeType) As Range
On Error GoTo e1

    Dim WS As Worksheet
    Dim RAN As Range, ran2 As Range, tempRAN As Range, finalran As Range, cell As Range, tempcell As Range
    Dim rColl As Collection
    Dim i As Integer
    Set WS = masterWB.Worksheets("Staff Detail")
    Set rColl = New Collection
    
    Select Case rtype
        Case 1
            Set RAN = boxRANGE(WS, "\r_precon", "\c_posStart", "\r_constr")
            trimRANGE RAN, dsupdown
            Set ran2 = boxRANGE(WS, "\r_constr", "\c_posStart", "\r_end")
            trimRANGE ran2, dsupdown
            If Not RAN Is Nothing And Not ran2 Is Nothing Then
                Set tempRAN = Union(RAN, ran2)
            ElseIf RAN Is Nothing And Not ran2 Is Nothing Then
                Set tempRAN = ran2
            ElseIf Not RAN Is Nothing And ran2 Is Nothing Then
                Set tempRAN = RAN
            End If
        Case 2
            Set RAN = boxRANGE(WS, "\r_precon", "\c_posEnd", "\r_constr")
            trimRANGE RAN, dsupdown
            Set ran2 = boxRANGE(WS, "\r_constr", "\c_posEnd", "\r_end")
            trimRANGE ran2, dsupdown
            If Not RAN Is Nothing And Not ran2 Is Nothing Then
                Set tempRAN = Union(RAN, ran2)
            ElseIf RAN Is Nothing And Not ran2 Is Nothing Then
                Set tempRAN = ran2
            ElseIf Not RAN Is Nothing And ran2 Is Nothing Then
                Set tempRAN = RAN
            End If
        Case 3
            Set RAN = boxRANGE(WS, "\r_precon", "\c_jobDur", "\r_constr")
            trimRANGE RAN, dsupdown
            Set ran2 = boxRANGE(WS, "\r_constr", "\c_jobDur", "\r_end")
            trimRANGE ran2, dsupdown
            If Not RAN Is Nothing And Not ran2 Is Nothing Then
                Set tempRAN = Union(RAN, ran2)
            ElseIf RAN Is Nothing And Not ran2 Is Nothing Then
                Set tempRAN = ran2
            ElseIf Not RAN Is Nothing And ran2 Is Nothing Then
                Set tempRAN = RAN
            End If
    End Select
    
    For Each cell In tempRAN
        If cell.Value <> "" Then
            rColl.Add cell
        End If
    Next
    
    For i = 1 To rColl.Count
        If Not finalran Is Nothing Then
            Set finalran = Union(finalran, rColl(i))
        Else
            Set finalran = rColl(i)
        End If
    Next
    
    Set GetRangeVals = finalran
        
Exit Function
    
e1:
    LogError "GantChart", "GetRangeVals", Err.Description, Err
End Function

Sub MakeExitButton(L As Long, T As Long, W As Long, H As Long, WS As Worksheet)
On Error GoTo ehandle

    Dim temp As Shape

    Set x = WS.Shapes.AddShape(msoShapeRoundedRectangle, L + (W - 20), T + 5, 12, 12)
    Set plus = WS.Shapes.AddShape(msoShapeRoundedRectangle, L + (W - 20), T + 20, 12, 12)
    
    With x
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextRange.Text = "X"
            .TextRange.Font.Fill.ForeColor.RGB = vbWhite
            .TextRange.Font.Size = 10
            .TextRange.Font.Name = "Arial Black"
            .MarginBottom = 0
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
        End With
        .Fill.ForeColor.RGB = SCCIred
        .Line.Visible = True
        .Line.ForeColor.RGB = SCCIred
        .Line.Weight = 1
        .Name = "\x"
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "onXAction"
    End With
    
    With plus
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextRange.Text = "+"
            .TextRange.Font.Fill.ForeColor.RGB = vbWhite
            .TextRange.Font.Size = 13
            .TextRange.Font.Name = "Arial Black"
            .MarginBottom = 0
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
        End With
        .Fill.ForeColor.RGB = SCCIblue
        .Line.Visible = True
        .Line.ForeColor.RGB = SCCIblue
        .Line.Weight = 1
        .Name = "\cPlus"
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "onPlusAction"
    End With
    
    Set temp = WS.Shapes.Range(Array("\x", "\chart", "\cPlus")).Group

Exit Sub

ehandle:
    LogError "GantChart", "MakeExitButton", Err.Description, Err
    
End Sub

Sub onXAction()
On Error GoTo ehandle
    
    Set gantCls = Nothing
    
Exit Sub
    
ehandle:
    LogError "GantChart", "OnXAction", Err.Description, Err
    Set gantCls = Nothing
End Sub

Sub onPlusAction()
On Error GoTo ehandle

    Dim WS As Worksheet
    Dim chrt As ChartObject
    Dim startRAN As Range, durRan As Range, cell As Range, nameRAN As Range
    Dim durArr() As Double, i As Integer
    Dim sr As Series
    
    Set WS = masterWB.Worksheets("Staff Detail")
    
    For Each chrt In WS.ChartObjects
        Charts.Add After:=Worksheets(Worksheets.Count)
        ActiveSheet.Move After:=Worksheets(Worksheets.Count)
    Next

    Set startRAN = GetRangeVals(sDate)
    Set durRan = GetRangeVals(dur)
    
    i = 0
    For Each cell In durRan
        ReDim Preserve durArr(i)
        durArr(i) = cell.Value * 30.4167  'convert duration to days. Necessary for gant chart
        i = i + 1
    Next
    
    Set nameRAN = Intersect(startRAN.EntireRow, WS.[\c_posName].EntireColumn)

    With ActiveSheet
        For Each sr In .SeriesCollection
            sr.Delete
        Next
        .HasTitle = True
        .ChartTitle.Text = "Staff Duration Chart"
        .ChartType = xlBarStacked
        .Axes(xlCategory).ReversePlotOrder = True
        .Axes(xlCategory).MajorTickMark = xlNone
        .Axes(xlValue).MinimumScale = WS.[\pstart].Value
        .Axes(xlValue).MaximumScale = WS.[\cend].Value
        .Axes(xlValue).TickLabels.NumberFormat = "[$-en-US]mmm-yy;@"
        .SeriesCollection.NewSeries
        .FullSeriesCollection(1).Name = "StartDate"
        .FullSeriesCollection(1).Values = startRAN
        .FullSeriesCollection(1).XValues = nameRAN
        .ChartGroups(1).GapWidth = 10
        .FullSeriesCollection(1).Format.Fill.Visible = msoFalse
        .SeriesCollection.NewSeries
        .FullSeriesCollection(2).Name = "Duration"
        .FullSeriesCollection(2).Values = durArr
        .FullSeriesCollection(2).Format.Fill.ForeColor.RGB = SCCIblue
        .Legend.Delete
    End With

    Set gantCls = Nothing
        
Exit Sub
    
ehandle:
    LogError "GantChart", "OnPlusAction", Err.Description, Err
    Set gantCls = Nothing
End Sub
