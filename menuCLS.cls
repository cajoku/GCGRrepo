VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "menuCLS"
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
Private mainSHP As Shape
Private xSHP As Shape
Private titleSHP As Shape
Private cmdSHP As Shape ', cmdSHP2 As Shape, cmdSHP3 As Shape, cmdSHP4 As Shape
Private cmdCOLL As Collection
Private myARR() As String
'
Const rightOFFSET As Long = 1
Const topOFFSET As Long = 1
Const pFIX As String = "\_\"

'<navarro
'<4-18-16
'Const Wmain As Long = 15
'Const Hmain As Long = 15
Private Wmain As Long
Private Hmain As Long
'4-18-16>
'navarro>

Const Wmenu As Long = 135
Const Hitem As Long = 15
Const Vspace As Long = 0
'
Public aRan As Range
Public aSHT As Worksheet

Sub setANCHOR(RAN As Range, menuARR() As String)
On Error GoTo e1

    Set anchorRAN = RAN
    Set aRan = RAN
    Set aSHT = RAN.Parent
    
    'position
    With anchorRAN
        If .Top <= Hmain Then
            mainSHP.Top = .Top
            mainSHP.Left = .Left - mainSHP.Width
        Else
            mainSHP.Top = .Top '- Hmain + topOFFSET
            mainSHP.Left = .Left - mainSHP.Width '- rightOFFSET
        End If
    End With

    myARR = menuARR

    'handle error with
    
    
Exit Sub
e1:
    LogError "menuCLS", "setANCHOR", Err.Description, Err
End Sub


Sub clearALL()
On Error GoTo e1

    Dim shp As Shape

    For Each shp In WS.Shapes
        If Left(shp.Name, 3) = pFIX Then
            shp.Delete
        End If
    Next

Exit Sub
e1:
    LogError "menuCLS", "clearALL", Err.Description, Err
End Sub

Private Sub makeMAIN()
On Error GoTo e1
    
    '<navarro
    '<4-18-18
    
    Set mainSHP = WS.Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, Wmain, Hmain)
    
    '4-18-18>
    'navarro>
    
    With mainSHP
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextRange.Text = "="
            .TextRange.Font.Fill.ForeColor.RGB = vbWhite
            .TextRange.Font.Size = 12
            .TextRange.Font.Name = "Arial Black"
            .MarginBottom = 0
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
        End With
        .Fill.ForeColor.RGB = buyGREEN
        .Line.Visible = True
        .Line.ForeColor.RGB = buyGREEN
        .Line.Weight = 1
        .Name = pFIX & "main"
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "menuACCEPT"
    End With


Exit Sub
e1:
    LogError "menuCLS", "makeMAIN", Err.Description, Err
End Sub

Private Sub makeTITLE()
On Error GoTo e1
    
    Set titleSHP = WS.Shapes.AddShape(msoShapeRectangle, mainSHP.Left, mainSHP.Top, Wmenu, Hmain)
    
    With titleSHP
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            With .TextRange
                .ParagraphFormat.Alignment = msoAlignRight
                .Text = "dataSMART"
                .Font.Fill.ForeColor.RGB = vbWhite
                .Font.Size = 8
                .Font.Name = "Verdana"
                .Font.Bold = msoCTrue
            End With
        End With
        .Fill.ForeColor.RGB = SCCIgray
        .Line.Visible = False
        .Name = pFIX & "\title"
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "dummy"
    End With

Exit Sub
e1:
    LogError "menuCLS", "makeTITLE", Err.Description, Err
End Sub


Private Sub makeX()
On Error GoTo e1
    
    Set xSHP = WS.Shapes.AddShape(msoShapeRoundedRectangle, mainSHP.Left, mainSHP.Top, Hmain, Hmain)
    
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
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "menuEXIT"
    End With


Exit Sub
e1:
    LogError "menuCLS", "makeX", Err.Description, Err
End Sub

Sub makeITEMs()
On Error GoTo e1
    
    Dim i As Integer
    Dim incrementT As Long
    Dim uI As Integer, lI As Integer
    'Dim tLBL As Label
    uI = UBound(myARR, 1)
    lI = LBound(myARR, 1)
    
'    Set cmdCOLL = New Collection
'
'    Do Until lI = uI
'    If incrementT = 0 Then incrementT = mainSHP.Top + mainSHP.Height + Vspace _
'    Else incrementT = incrementT + Hitem + Vspace
'    cmdSHP1 = WS.Shapes.AddShape(msoShapeRectangle, mainSHP.Left, incrementT, Wmenu, Hitem)
'    With cmdCOLL
'        .Add cmdSHP1
'        .Add cmdSHP2
'        .Add cmdSHP3
'        .Add cmdSHP4
'    End With
'    For i = lI To uI
'
'        If incrementT = 0 Then incrementT = mainSHP.Top + mainSHP.Height + Vspace _
'            Else incrementT = incrementT + Hitem + Vspace

    
    'array test here:
    
    For i = lI To uI
            
        If incrementT = 0 Then incrementT = mainSHP.Top + mainSHP.Height + Vspace _
            Else incrementT = incrementT + Hitem + Vspace

        Set cmdSHP = WS.Shapes.AddShape(msoShapeRectangle, mainSHP.Left, incrementT, Wmenu, Hitem)
        
        With cmdSHP
            With .TextFrame2
                .VerticalAnchor = msoAnchorMiddle
                With .TextRange
                    .ParagraphFormat.Alignment = msoAlignLeft
                    .Text = myARR(i, 1)
                    .Font.Fill.ForeColor.RGB = SCCIblue
                    .Font.Name = "Verdana"
                    .Font.Size = 8
                End With
            End With
            .Fill.ForeColor.RGB = vbWhite
            .Line.ForeColor.RGB = RGB(220, 220, 220)
            .Line.Weight = 0.5
            .Line.Visible = True
            .Name = pFIX & "\menu_" & i
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "listACCEPT" '& myARR(i, 2) ' & " " & myARR(i, 2) ' & "(" & myARR(i, 2) & ")"
        End With
        
    Next
    
    
Exit Sub
e1:
    LogError "menuCLS", "makeITEMs", Err.Description, Err
End Sub

Sub openMENU()
On Error GoTo e1
    
    mainSHP.Visible = False
    'makeTITLE
    makeX
    makeITEMs
    
Exit Sub
e1:
    LogError "menuCLS", "openMENU", Err.Description, Err
End Sub

Function createARRAY() As Variant
On Error GoTo e1

    Dim tARR(1 To 4, 1 To 2) As String
    Dim sht As Worksheet
    Dim mDur As Range, wDur As Range, pcost As Range, getrange As Range
    Dim staffCount As Range

    Set sht = masterOBJ.WS
    Set mDur = sht.[\duration]
    Set wDur = sht.[\wdur]
    Set pcost = sht.[\pcost]
    Set staffCount = sht.[staff]
    'staffCount.Value = countSTAFF
    'Set getrange = GetRangeVals(sDate)
    'staffCount.Value = getrange.Count
    
    tARR(1, 1) = "Construction Duration (mos)"
    tARR(2, 1) = "Construction Duration (wks)"
    tARR(3, 1) = "Project Cost"
    tARR(4, 1) = "Staff Count"
    
    tARR(1, 2) = mDur.Name
    tARR(2, 2) = wDur.Name
    tARR(3, 2) = pcost.Name
    tARR(4, 2) = staffCount.Name
    
    'ReDim createARRAY(1 To 3, 1 To 2)
    
    createARRAY = tARR
    
Exit Function
e1:
    LogError "menuCLS", "createARRAY", Err.Description, Err
End Function

Sub listSELECT()
On Error GoTo e1

    Dim i As Integer
    
    i = CInt(Right(Application.Caller, 1))
    
    ActiveCell.Formula = myARR(i, 2)

Exit Sub
e1:
    LogError "menuCLS", "listSELECT", Err.Description, Err
    
End Sub


Private Sub Class_Initialize()
On Error GoTo e1

    Set WS = ActiveSheet
    Set wb = ActiveWorkbook
    
    'WS.Unprotect
    '<navarro
    '<4-18-18
    Hmain = ActiveCell.RowHeight
    Wmain = Hmain
    'navarro>
    '4-18-18>
    
    clearALL
    makeMAIN
    
    'basicPROTECT WS, True
Exit Sub
e1:
    LogError "menuCLS", "Class_Initialize", Err.Description, Err
    'basicPROTECT WS, True
    
End Sub

Private Sub Class_Terminate()
On Error GoTo e1

    clearALL
    
Exit Sub
e1:
    LogError "menuCLS", "Class_Terminate", Err.Description, Err
End Sub
