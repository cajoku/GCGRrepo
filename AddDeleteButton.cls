VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AddDeleteButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
Private anchorRAN As Range
Private WithEvents WS As Worksheet
Attribute WS.VB_VarHelpID = -1
Private WithEvents wb As Workbook
Attribute wb.VB_VarHelpID = -1
Private addSHP As Shape
Private deleSHP As Shape
Private schedDele As Shape
Private trailerDele As Shape
Private phaseDele As Shape
Private ganttDele As Shape
'
Const rightOFFSET As Long = 1
Const topOFFSET As Long = 1
Const pFIX As String = "\_\"
Const Wmain As Long = 12
Const Hmain As Long = 12
Const Wmenu As Long = 135
Const Hitem As Long = 15
Const Vspace As Long = 0

Private myL As Long
Private myT As Long

'
Public aRan As Range
Public aSHT As Worksheet

Private Sub Class_Terminate()
On Error GoTo ehandle

    clearALL

    Exit Sub
    
ehandle:
    LogError "AddDeleteButton", "Class_Terminate", Err.Description, Err
    
End Sub

Sub setANCHOR(RAN As Range)
On Error GoTo ehandle
    'navarro|  flatten ran into a single cell, otherwise the buttons will appear in funny places on a multi cell selection
    
    Set anchorRAN = RAN.Cells(1, 1)
    
    Set WS = ActiveSheet
    Set wb = ActiveWorkbook
    
    With anchorRAN
        myL = .Left + .Width - Wmain
        myT = .Top + .Height / 2 - Hmain / 2
    End With
    
    clearALL
    makeMAIN

    Exit Sub

ehandle:
    LogError "AddDeleteButton", "setAnchor", Err.Description, Err
    
End Sub

Private Sub makeMAIN()
On Error GoTo ehandle

    Set addSHP = WS.Shapes.AddShape(msoShapeOval, myL, myT, Wmain, Hmain)
    Set deleSHP = WS.Shapes.AddShape(msoShapeOval, myL - 20, myT, Wmain, Hmain)
    
    With addSHP
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextRange.Text = "+"
            .TextRange.Font.Fill.ForeColor.RGB = vbWhite
            .TextRange.Font.Size = 15
            .TextRange.Font.Name = "Arial Black"
            '.TextRange.Font.RGB = vbWhite
           ' .MarginBottom = 1
           ' .MarginLeft = 1
            '.MarginRight = 1
            '.MarginTop = 1
            '.Orientation = msoTextOrientationUpward
        End With
        .Fill.ForeColor.RGB = SCCIblue
        .Line.Visible = True
        .Line.ForeColor.RGB = SCCIblue
        .Line.Weight = 1
        .Name = pFIX & "add"
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "insertCOPYperson"
    End With
    
    With deleSHP
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextRange.Text = "-"
            .TextRange.Font.Fill.ForeColor.RGB = vbWhite
            .TextRange.Font.Size = 18
            .TextRange.Font.Name = "Arial Black"
            '.TextRange.Font.Color = vbWhite
            .MarginBottom = 4
            .MarginLeft = 1
            .MarginRight = 1
            .MarginTop = 1
            '.Orientation = msoTextOrientationUpward
        End With
        .Fill.ForeColor.RGB = SCCIred
        .Line.Visible = True
        .Line.ForeColor.RGB = SCCIred
        .Line.Weight = 1
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "deletePERSON"
        .Name = pFIX & "minus"
    End With
    
    Exit Sub

ehandle:
    LogError "AddDeleteButton", "makeMAIN", Err.Description, Err
End Sub

Sub clearALL()
On Error GoTo ehandle

    Dim shp As Shape
    
    For Each shp In WS.Shapes
        If Left(shp.Name, 3) = pFIX Then
            shp.Delete
        End If
    Next
    
    Exit Sub

ehandle:
    LogError "AddDeleteButton", "clearALL", Err.Description, Err
End Sub
Sub setGanttAnchor(RAN As Range)
On Error GoTo e1

    Set anchorRAN = RAN
    Set WS = ActiveSheet
    Set wb = ActiveWorkbook
    

        With anchorRAN
            myL = .Offset(0, -1).Left
            myT = .Top
        End With

    
    clearALL
    makeGanttMain
    
Exit Sub
e1:
    LogError "AddDeleteButton", "setSchedAnchor", Err.Description, Err
End Sub

Sub makeGanttMain()
On Error GoTo e1

    Set ganttDele = WS.Shapes.AddShape(msoShapeRectangle, myL, myT, anchorRAN.Offset(0, -1).Width, anchorRAN.Height)
    
    With ganttDele
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextRange.Text = "x"
            .TextRange.Font.Fill.ForeColor.RGB = vbWhite
            .TextRange.Font.Size = 18
            .TextRange.Font.Name = "Arial"
            .MarginBottom = 3
            .MarginLeft = 1
            .MarginRight = 1
            .MarginTop = 1
        End With
        .Fill.ForeColor.RGB = SCCIred
        .Line.Visible = True
        .Line.ForeColor.RGB = SCCIred
        .Line.Weight = 1
        .Name = pFIX & "ganttDele"
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "deleteGantt"
    End With

Exit Sub
e1:
    LogError "AddDeleteButton", "makeSchedMain", Err.Description, Err
End Sub

Sub setSchedAnchor(RAN As Range, Optional col As Boolean = False)
On Error GoTo e1

    Set anchorRAN = RAN
    Set WS = ActiveSheet
    Set wb = ActiveWorkbook
    
    If col = False Then
        With anchorRAN
            myL = .Left - (Wmain + 25)
            myT = .Top + .Height / 2 - (Hmain + 20) / 2
        End With
    ElseIf col = True Then
        With anchorRAN
            myL = .Left + (.Width / 2) - Wmain
            myT = .Top - (.Height + 20)
        End With
    End If
    
    clearALL
    makeSchedMain col
    
Exit Sub
e1:
    LogError "AddDeleteButton", "setSchedAnchor", Err.Description, Err
End Sub

Sub makeSchedMain(col As Boolean)
On Error GoTo e1

    Set schedDele = WS.Shapes.AddShape(msoShapeRectangle, myL, myT, Wmain + 13, Hmain + 13)
    
    With schedDele
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextRange.Text = "X"
            .TextRange.Font.Fill.ForeColor.RGB = vbWhite
            .TextRange.Font.Size = 24
            .TextRange.Font.Name = "Comic Sans MS"
            '.TextRange.Font.Color = vbWhite
            .MarginBottom = 1
            .MarginLeft = 1
            .MarginRight = 1
            .MarginTop = 1
            '.Orientation = msoTextOrientationUpward
        End With
        .Fill.ForeColor.RGB = SCCIred
        .Line.Visible = True
        .Line.ForeColor.RGB = SCCIred
        .Line.Weight = 1
        .Name = pFIX & "schedDele"
        If col = False Then
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "deleteSchedItem"
        ElseIf col = True Then
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "deleteTemp"
        End If
    End With

Exit Sub
e1:
    LogError "AddDeleteButton", "makeSchedMain", Err.Description, Err
End Sub

Sub trailerANCHOR(RAN As Range)
On Error GoTo ehandle
    
    Set WS = ActiveSheet
    Set wb = ActiveWorkbook
    
    Set anchorRAN = Intersect(WS.[\c_group].EntireColumn, RAN.EntireRow)
    
    Do Until anchorRAN.Value = ""
        Set anchorRAN = anchorRAN.Offset(-1, 0)
    Loop
    
    Set anchorRAN = Intersect(WS.[\c_qt].EntireColumn, anchorRAN.EntireRow).Offset(0, -1)
    
    With anchorRAN
        myL = .Left + .Width / 2
        myT = .Top + .Height / 6 '- Hmain / 2
    End With
    
    clearALL
    makeTRAILER

    Exit Sub

ehandle:
    LogError "AddDeleteButton", "trailerANCHOR", Err.Description, Err
    
End Sub

Sub makeTRAILER()
On Error GoTo e1

    Set trailerDele = WS.Shapes.AddShape(msoShapeOval, myL, myT, Hmain + 1, Hmain)
    
    With trailerDele
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextRange.Text = "-"
            .TextRange.Font.Fill.ForeColor.RGB = vbWhite
            .TextRange.Font.Size = 26
            .TextRange.Font.Name = "Arial"
            .MarginBottom = 6
            .MarginLeft = 2
            .MarginRight = 1
            .MarginTop = 1
        End With
        .Fill.ForeColor.RGB = SCCIred
        .Line.Visible = True
        .Line.ForeColor.RGB = SCCIred
        .Line.Weight = 1
        .Name = pFIX & "trailerDele"
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "deleteTRAILER"
    End With

Exit Sub
e1:
    LogError "AddDeleteButton", "makeTRAILER", Err.Description, Err
End Sub


Sub setDETAILANCHOR(RAN As Range)
On Error GoTo ehandle
    'navarro|  flatten ran into a single cell, otherwise the buttons will appear in funny places on a multi cell selection
    
    Set anchorRAN = RAN.Cells(1, 1)
    
    Set WS = ActiveSheet
    Set wb = ActiveWorkbook
    
    With anchorRAN
        myL = .Left + .Width - Wmain
        myT = .Top + .Height / 2 - Hmain / 2
    End With
    
    clearALL
    makeMAIN

    Exit Sub

ehandle:
    LogError "AddDeleteButton", "setAnchor", Err.Description, Err
    
End Sub

Sub setPHASE(RAN As Range)

    Set anchorRAN = RAN.Cells(1, 1)
    
    Set WS = ActiveSheet
    Set wb = ActiveWorkbook
    
    With anchorRAN
        myL = .Left
        myT = .Top
    End With
    
    clearALL
    makePHASE

End Sub


Sub makePHASE()

'On Error GoTo e1

    Set phaseDele = WS.Shapes.AddShape(msoShapeRoundedRectangle, myL, myT, anchorRAN.Width, anchorRAN.Height)
    
    With phaseDele
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextRange.Text = "x"
            .TextRange.Font.Fill.ForeColor.RGB = vbWhite
            .TextRange.Font.Size = 18
            .TextRange.Font.Name = "Arial"
            .MarginBottom = 3
            .MarginLeft = 1
            .MarginRight = 1
            .MarginTop = 1
        End With
        .Fill.ForeColor.RGB = SCCIred
        .Line.Visible = True
        .Line.ForeColor.RGB = SCCIred
        .Line.Weight = 1
        .Name = pFIX & "phaseDele"
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "deletePHASE"
    End With

Exit Sub
e1:
    LogError "AddDeleteButton", "makePHASE", Err.Description, Err




End Sub
