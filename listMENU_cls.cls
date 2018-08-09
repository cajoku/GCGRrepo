VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "listMENU_cls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private WithEvents uiWS As Worksheet
Attribute uiWS.VB_VarHelpID = -1
Private codeWS As Worksheet
Private anchorRAN As Range
Private listSOURCEran As Range
Private listLINKran As Range
'
Private myL As Long
Private myT As Long
'
Private mainSHP As Shape
Private xSHP As Shape
Private addSHP As Shape
Private lBOX As ListBox
'
Const pFIX As String = "\_\"
Const Wmain As Long = 15
Const Hmain As Long = 15
Const Wmin As Long = 135
Const Hdefault As Long = 200
'
Public resultCOLL As Collection
Private OnAccept As String


'=========================================
'BEHAVIOURS
'=========================================

Sub hideME()
    On Error GoTo e1

    Dim shp As Shape

    For Each shp In uiWS.Shapes
        If Left(shp.Name, 3) = pFIX Then shp.Visible = msoFalse
    Next
    Exit Sub
e1:
    LogError "menuCLS", "hideME", Err.Description, Err
    
End Sub

Sub minimizeME()
    On Error GoTo e1

    Dim shp As Shape
    
    For Each shp In uiWS.Shapes
        If Left(shp.Name, 4) = pFIX & "\" Then shp.Delete
    Next
   ' mainSHP.Visible = msoTrue
    
   
    Exit Sub
e1:
    LogError "menuCLS", "minimizeME", Err.Description, Err
End Sub

Sub clearALL()
On Error GoTo e1

    Dim shp As Shape

    For Each shp In uiWS.Shapes
        If Left(shp.Name, 3) = pFIX Then
            shp.Delete
        End If
    Next

Exit Sub
e1:
    LogError "menuCLS", "clearALL", Err.Description, Err
End Sub


Sub openLIST()
On Error GoTo e1
    
    'mainSHP.Visible = False
    makeLIST
    makeX
    makeCHK
    
Exit Sub
e1:
    LogError "menuCLS", "openMENU", Err.Description, Err
End Sub

Sub Accept()
On Error GoTo e1

    Dim str As String
    Dim i As Integer
    
    Set resultCOLL = New Collection
    
    For i = 1 To lBOX.ListCount
        If lBOX.Selected(i) Then
            resultCOLL.Add listSOURCEran.Cells(i, 1)
        End If
    Next

    Application.Run OnAccept, resultCOLL, anchorRAN
    
Exit Sub
    
e1:
    LogError "Accept", "listMENU_cls", Err.Description, Err
End Sub

'=========================================
'INITIALIZE/TERMINATE
'=========================================


Sub setANCHOR(Target As Range, listRAN As Range, OnAction As String)
On Error GoTo e1

    Set anchorRAN = Target
    Set uiWS = Target.Parent
    Set listSOURCEran = listRAN
    Set codeWS = listRAN.Parent
    
    OnAccept = OnAction
    
    With anchorRAN
        myL = .Left + .Width - Wmain
        myT = .Top + .Height / 2 - Hmain / 2
    End With
    
    clearALL
    'makeMAIN

    
Exit Sub
e1:
    LogError "menuCLS", "setANCHOR", Err.Description, Err
End Sub



Private Sub Class_Terminate()
On Error GoTo e1

    clearALL
    
Exit Sub
e1:
    LogError "menuCLS", "Class_Terminate", Err.Description, Err
End Sub


'=========================================
'MAKER
'=========================================

'
'Private Sub makeMAIN()
'    On Error GoTo e1
'
'    Set mainSHP = uiWS.Shapes.AddShape(msoShapeRoundedRectangle, myL, myT, Wmain, Hmain)
'
'    With mainSHP
''        .Adjustments.Item(1) = -0.5
''        .Adjustments.Item(2) = 0.7
'        With .TextFrame2
'            .VerticalAnchor = msoAnchorMiddle
'            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
'            .TextRange.Text = "|||"
'            .TextRange.Font.Fill.ForeColor.RGB = SCCIblue
'            .TextRange.Font.Size = 10
'            .TextRange.Font.Name = "Arial Black"
'            .MarginBottom = 0
'            .MarginLeft = 0
'            .MarginRight = 1.2
'            .MarginTop = 0
'            .Orientation = msoTextOrientationUpward
'        End With
'        .Fill.ForeColor.RGB = vbWhite
'        .Line.Visible = True
'        .Line.ForeColor.RGB = SCCIblue
'        .Line.Weight = 1
'        .Name = pFIX & "main"
'        .OnAction = "'" & masterWB.Name & "'!" & "listMENU_CLICK"
'    End With
'
'
'    Exit Sub
'e1:
'    logerror "menuCLS", "makeMAIN",err.description, Err
'End Sub

Private Sub makeLIST()
On Error GoTo e1

    Dim setW As Long
    Dim setH As Long
    
    setH = listSOURCEran.Cells.Count * 10 + 10
    If setH > Hdefault Then setH = Hdefault
    
    With anchorRAN
        If .Width > Wmin Then setW = .Width Else setW = Wmin
        Set lBOX = uiWS.ListBoxes.Add(.Left, .Top + .Height, setW, setH)
    End With
    
    With lBOX
        .Name = pFIX & "\list"
        .ListFillRange = "'" & codeWS.Name & "'!" & listSOURCEran.Address
        .MultiSelect = -4154
    End With


Exit Sub
e1:
    LogError "menuCLS", "makeTITLE", Err.Description, Err
End Sub

Private Sub makeX()
On Error GoTo e1
    
    Set xSHP = uiWS.Shapes.AddShape(msoShapeRound1Rectangle, lBOX.Left + lBOX.Width, lBOX.Top, Hmain, Hmain)
    
    With xSHP
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            With .TextRange
                .ParagraphFormat.Alignment = msoAlignCenter
                .Text = "X"
                .Font.Fill.ForeColor.RGB = vbWhite
                .Font.Name = "Arial Black"
                .Font.Size = 7
            End With
        End With
        .Fill.ForeColor.RGB = SCCIred
        .Line.Visible = False
        .Name = pFIX & "\X"
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "listMENU_EXIT"
    End With


Exit Sub
e1:
    LogError "menuCLS", "makeX", Err.Description, Err
End Sub

Private Sub makeCHK()
On Error GoTo e1
    
    Set xSHP = uiWS.Shapes.AddShape(msoShapeRoundedRectangle, lBOX.Left + lBOX.Width, lBOX.Top + xSHP.Height, Hmain, Hmain)
    
    With xSHP
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            With .TextRange
                .ParagraphFormat.Alignment = msoAlignCenter
                .Text = "+"
                .Font.Fill.ForeColor.RGB = vbWhite
                .Font.Name = "Arial Black"
                .Font.Size = 13
            End With
        End With
        .Fill.ForeColor.RGB = SCCIblue
        .Line.Visible = False
        .Name = pFIX & "\CHK"
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "listMENU_ACCEPT"
    End With


Exit Sub
e1:
    LogError "menuCLS", "makeX", Err.Description, Err
End Sub

