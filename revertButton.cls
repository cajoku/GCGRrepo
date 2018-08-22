VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "revertButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
Private anchorRAN As Range
Private WithEvents WS As Worksheet
Attribute WS.VB_VarHelpID = -1
Private WithEvents WB As Workbook
Attribute WB.VB_VarHelpID = -1
Private arrowSHP As Shape

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
    LogError "revertButton", "Class_Terminate", Err.Description, Err
End Sub

Sub setANCHOR(RAN As Range)
On Error GoTo ehandle
    'navarro|  flatten ran into a single cell, otherwise the buttons will appear in funny places on a multi cell selection
    
    Set anchorRAN = RAN.Cells(1, 1)
    
    Set WS = ActiveSheet
    Set WB = ActiveWorkbook
    
    With anchorRAN
        myL = .Left + .Width + 5 '- Wmain
        myT = .Top + .Height / 2 - Hmain / 2
    End With
    
    clearALL
    makeMAIN

Exit Sub
    
ehandle:
    LogError "revertButton", "setANCHOR", Err.Description, Err
End Sub
Private Sub makeMAIN()
On Error GoTo ehandle

    Set arrowSHP = WS.Shapes.AddShape(msoShapeCurvedRightArrow, myL, myT, Wmain, Hmain)
    
    With arrowSHP
        .flip msoFlipVertical
        .flip msoFlipHorizontal
        .Fill.ForeColor.RGB = SCCIblue
        .Line.Visible = True
        .Line.ForeColor.RGB = SCCIblue
        .Line.Weight = 1
        .Name = pFIX & "arrow"
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "RevertValue"
    End With
    
Exit Sub

ehandle:
    LogError "revertButton", "makeMAIN", Err.Description, Err
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
    LogError "revertButton", "clearALL", Err.Description, Err
End Sub

Sub setANCHOR2(RAN As Range)
On Error GoTo ehandle
    'navarro|  flatten ran into a single cell, otherwise the buttons will appear in funny places on a multi cell selection
    
    Set anchorRAN = RAN.Cells(1, 1)
    
    Set WS = ActiveSheet
    Set WB = ActiveWorkbook
    
    With anchorRAN
        myL = .Left + .Width + 5 '- Wmain
        myT = .Top + .Height / 2 - Hmain / 2
    End With
    
    clearALL
    makeMAIN2

Exit Sub
    
ehandle:
    LogError "revertButton", "setANCHOR2", Err.Description, Err
End Sub

Private Sub makeMAIN2()
On Error GoTo ehandle

    Set arrowSHP = WS.Shapes.AddShape(msoShapeCurvedRightArrow, myL, myT, Wmain, Hmain)
    
    With arrowSHP
        .flip msoFlipVertical
        .flip msoFlipHorizontal
        .Fill.ForeColor.RGB = SCCIblue
        .Line.Visible = True
        .Line.ForeColor.RGB = SCCIblue
        .Line.Weight = 1
        .Name = pFIX & "arrow"
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "RevertValue2"
    End With
    
Exit Sub

ehandle:
    LogError "revertButton", "makeMAIN2", Err.Description, Err
End Sub
