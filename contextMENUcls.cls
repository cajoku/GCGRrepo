VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "contextMENUcls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents WS As Worksheet
Attribute WS.VB_VarHelpID = -1
Private masterSHP As Shape
Private lineSHP As Shape
Private Const pFIX As String = "\c\"
Private activeROW As Range
Private xSHP As Shape
Private Const btnW As Double = 21
Private Const btnH As Double = 21
Private Const btnW2 As Double = 120
Private Const lH As Double = 2
Private Const lMargin As Double = 10
Private Const btnIndent As Double = 4
Private myWS As Worksheet
Private WithEvents moveWS As Worksheet
Attribute moveWS.VB_VarHelpID = -1
Private moveSHP As Shape
Private lastRAN As Range
Private lastROWh As Double
Private lastCOLw As Double
Private lastFORM As String
Private shpCOLL As Collection
Private insertRAN As Range
Private modeSTR As String
Private WithEvents deleteWS As Worksheet
Attribute deleteWS.VB_VarHelpID = -1
Private tempTOP As Double
Private windowCOLL As Collection

'============================================================
'               INITIALIZE/TERMINATE
'============================================================

Sub register(sht As Worksheet)
On Error GoTo e1

    Set WS = sht
    Set myWS = sht
    Set masterSHP = myWS.Shapes("\c_main")
    Set xSHP = myWS.Shapes("\c_x")
    Set lineSHP = myWS.Shapes("\c_line")
    Set activeMenu = Me
    killTempShapes

Exit Sub
e1:
    LogError "contextMENUcls", "register", Err.Description, Err
End Sub

Sub anchor(RAN As Range)
On Error GoTo e1

    masterSHP.Top = RAN.Top - (masterSHP.Height - RAN.Height)
    masterSHP.Left = lMargin
    align_Burger2Line
    Set activeMenu = Me
    Set activeROW = ActiveCell.EntireRow
    
Exit Sub
e1:
    LogError "contextMENUcls", "anchor", Err.Description, Err
End Sub

Sub assembleMENU()
On Error GoTo e1

    Dim shp As Shape, shp2 As Shape, shp3 As Shape
    Dim masterL As Double
    Dim lastTOP As Double

    Application.ScreenUpdating = False
    
    masterL = masterSHP.Left + btnIndent
    lastTOP = masterSHP.Top + masterSHP.Height + 1
    
    For Each shp In myWS.Shapes
        If Left(shp.Name, Len(pFIX)) = pFIX And Right(shp.Name, 1) <> "\" Then
            
            With shp
                .Width = btnW
                .Height = btnH
                .Left = masterL
                .Top = lastTOP
                lastTOP = lastTOP + .Height
                .Visible = True
            End With
            Set shp2 = myWS.Shapes(shp.Name & "\")
            With shp2
                .Height = shp.Height
                .Width = btnW2
                .Left = shp.Left + shp.Width
                .Top = shp.Top
                .Visible = True
            End With
            Set shp3 = myWS.Shapes(shp.Name & "_\")
            With shp3
                .Height = shp.Height
                .Width = shp.Width + shp2.Width
                .Left = shp.Left
                .Top = shp.Top
                .Line.ForeColor.RGB = RGB(240, 240, 240)
                .Line.Weight = 0.5
                .OnAction = "'" & ThisWorkbook.Name & "'!" & "context_click"
                .Visible = True
            End With
        End If
    Next
    
    xSHP.Visible = True
    
    Application.ScreenUpdating = True
    
Exit Sub
e1:
    LogError "contextMENUcls", "assembleMENU", Err.Description, Err
End Sub

Sub align_Burger2Line()
On Error GoTo e1

    Dim lW As Double
    
    With myWS.Range("\body")
'        lW = .Columns(.Columns.Count).Offset(0, 1).Left
        lW = .Columns(2).Left
        lW = lW - masterSHP.Left - 4
    End With
    
    With lineSHP
        .Height = lH
        .Top = masterSHP.Top + masterSHP.Height - lineSHP.Height
        .Left = masterSHP.Left + 4
        .Width = lW
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "dummy"
        .Visible = True
    End With
    
    With xSHP
        .Top = masterSHP.Top
        .Left = masterSHP.Left
        .Width = masterSHP.Height
        .Height = .Width
        .Visible = False
    End With
    
    masterSHP.Visible = True
        
Exit Sub
e1:
    LogError "contextMENUcls", "align_Burger2Line", Err.Description, Err
End Sub

Private Sub Class_Terminate()
On Error GoTo e1

    dockMAIN

Exit Sub
e1:
    LogError "contextMENUcls", "Class_Terminate", Err.Description, Err
End Sub

'============================================================
'                  MODES
'============================================================

Sub MODE_cancel()
On Error GoTo e1

    dockMAIN
    Set moveWS = Nothing
    'Set deleteWS = Nothing
    Set WS = myWS
    modeSTR = ""
    If Not lastRAN Is Nothing Then
        If lastROWh <> 0 Then
            lastRAN.EntireRow.RowHeight = lastROWh
        End If
    End If
    clearCACHE
    'other stuff
    
Exit Sub
e1:
    LogError "contextMENUcls", "MODE_cancel", Err.Description, Err
End Sub

Sub MODE_delete()
On Error GoTo e1

    Dim docRAN As Range
    Dim deleteRAN As Range
    Dim shp As Shape
    
    'prep
    Set docRAN = myWS.[\body]
    Set deleteRAN = Intersect(activeROW, docRAN)
    modeSTR = "delete"
    
    'promote move btn
    promote "\c\minus_\"
    
    'custom stuff
    With deleteRAN
        Set shp = myWS.Shapes.AddShape(msoShapeRectangle, .Left, .Top, .Width, .Height)
    End With
    With shp
        .Line.Visible = True
        .Line.Weight = 2
        .Line.ForeColor.RGB = SCCIred
        .Fill.ForeColor.RGB = SCCIred
        .Fill.Transparency = 0.75
        .Name = "\c\deleterange~"
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "dummy"
    End With
    shpCOLL.Add shp
    
    Set shp = myWS.Shapes("\c\caption~")
    With shp
        .TextFrame.Characters.Font.Color = SCCIred
    End With
    
    EnS 0, , True
    myWS.[A1].Select
    EnS 1, , False

Exit Sub
e1:
    LogError "contextMENUcls", "MODE_delete", Err.Description, Err
    EnS 1, , False
End Sub

Sub MODE_move()
On Error GoTo e1

    Dim docRAN As Range
    Dim moveRAN As Range
    Dim shp As Shape

    'prep
    Set docRAN = myWS.[\body]
    Set docRAN = Union(docRAN, myWS.[\xfer])
    Set moveRAN = Intersect(activeROW, docRAN)
    Set WS = Nothing
    Set moveWS = myWS
    modeSTR = "move"
    
    'promote move btn
    promote "\c\updown_\"
    
    'custom stuff
    moveRAN.CopyPicture xlScreen, xlPicture
    myWS.Paste moveRAN
    Selection.Name = "\c\moverange~"
    EnS 0, , True
    moveRAN.Select
    
    Set shp = myWS.Shapes("\c\moverange~")
    With shp
        .Line.Visible = True
        .Line.Weight = 2
        .Line.ForeColor.RGB = RGB(240, 240, 240)
        .OnAction = .OnAction = "'" & ThisWorkbook.Name & "'!" & "dummy"
    End With
    Set moveSHP = shp
    shpCOLL.Add moveSHP
    
    EnS 1, , False
    
Exit Sub
e1:
    LogError "contextMENUcls", "MODE_move", Err.Description, Err
    EnS 1, , False
End Sub



'============================================================
'                  EXECUTE
'============================================================

Private Sub executeMOVE()
On Error GoTo e1

    EnS 0
    dockMAIN
    activeROW.Cut
    insertRAN.Insert
    Application.CutCopyMode = False
    hideENDS myWS
    EnS 1
    
Exit Sub
e1:
    LogError "contextMENUcls", "executeMOVE", Err.Description, Err
    EnS 1, , True
End Sub

Sub addTEMPLATE(tempSTR As String)
On Error GoTo e1

    Dim copyRAN As Range
    
    EnS 0, , True
    dockMAIN
    Set copyRAN = myWS.Range(tempSTR).EntireRow
    copyRAN.Hidden = False
    copyRAN.Copy
    activeROW.Insert
    copyRAN.Hidden = True
    Application.CutCopyMode = False
    Intersect(activeROW.Offset(-1, 0).EntireRow, myWS.[\c_desc].EntireColumn).Select
    hideENDS myWS
    EnS 1
    
Exit Sub
e1:
    LogError "contextMENUcls", "addTEMPLATE", Err.Description, Err
    EnS 1, , True
End Sub

Private Sub executeDELETE()
On Error GoTo e1

    'Dim insertRAN As Range
    Dim formSTR As String
    Dim trashRAN As Range
    
    Set insertRAN = myWS.[\r_bin]
    Set insertRAN = insertRAN.Rows(insertRAN.Rows.Count)
    
    
    EnS 0
    dockMAIN
    formSTR = "=" & activeROW.Cells(1, 1).Address(False, False)
    activeROW.Cut
    insertRAN.Insert
    insertRAN.Offset(-1, 0).Cells(1, 1).Formula = formSTR
    Application.CutCopyMode = False
    Set trashRAN = myWS.[\r_bin]
    'maintain max
    With trashRAN
        If .Rows.Count > 12 Then
            .Rows(2).EntireRow.Delete
        End If
    End With
    hideENDS myWS
    EnS 1
    
Exit Sub
e1:
    LogError "contextMENUcls", "executeDELETE", Err.Description, Err
    EnS 1, , True
End Sub

Private Sub executeXFER()
On Error GoTo e1

    Dim toWS As Worksheet
    
    If myWS.Name = masterOBJ.groWS.Name Then
        Set toWS = masterOBJ.gcoWS
    ElseIf myWS.Name = masterOBJ.gcoWS.Name Then
        Set toWS = masterOBJ.groWS
    End If
    
    Set insertRAN = toWS.[\r_xfer].EntireRow
    
    toWS.Unprotect
    EnS 0
    dockMAIN
    'activeROW.Copy
    insertRAN.Offset(1, 0).EntireRow.Insert
    activeROW.Copy
    insertRAN.Offset(1, 0).EntireRow.PasteSpecial xlPasteFormats
    activeROW.Copy
    insertRAN.Offset(1, 0).EntireRow.PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    activeROW.Delete
    Set activeROW = Nothing
    hideENDS myWS
    EnS 1
    basicPROTECT toWS, True
    
Exit Sub
e1:
    LogError "contextMENUcls", "executeXFER", Err.Description, Err
    EnS 1, , True
    basicPROTECT toWS, True
    
End Sub


'============================================================
'                  CLEANUP
'============================================================

Private Sub killTempShapes()
On Error GoTo e1

    Dim shp As Shape
    
    For Each shp In myWS.Shapes
        If Right(shp.Name, 1) = "~" Then shp.Delete
    Next

Exit Sub
e1:
    LogError "contextMENUcls", "killTempShapes", Err.Description, Err
End Sub

Sub clearCACHE()
On Error GoTo e1

    lastROWh = 0
    lastCOLw = 0
    Set lastRAN = Nothing
    Set shpCOLL = Nothing

Exit Sub
e1:
    LogError "contextMENUcls", "clearCACHE", Err.Description, Err
End Sub

'============================================================
'                  METHODS
'============================================================


Private Sub promote(btnSTR As String)
On Error GoTo e1

    Dim myW As Double
    Dim yesSHP As Shape
    Dim noSHP As Shape
    Dim captionSHP As Shape
    Dim myT As Double
    Dim captionSTR As String
    
    
    'hide all
    dockMAIN
    
    'create accept & Declines
    myW = activeROW.Height
    captionSTR = myWS.Shapes(Replace(btnSTR, "_", "")).TextFrame.Characters.Text
    Set shpCOLL = New Collection
    
    'this is a clear box over the x icon
    Set noSHP = myWS.Shapes.AddShape(msoShapeRectangle, lMargin + btnIndent, activeROW.Top, myW, myW)
    With noSHP
        .Line.Visible = False
        .Fill.Visible = False
        .Name = "\c\cancel~"
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "context_click"
    End With
    shpCOLL.Add noSHP
    
    'this is the X icon
    With myWS.Shapes("\c\cancel\")
        .Left = noSHP.Left
        .Top = noSHP.Top
        .Width = noSHP.Width
        .Height = noSHP.Height
        .Visible = True
    End With
    shpCOLL.Add myWS.Shapes("\c\cancel\")
    
    'this is the check icon
    With myWS.Shapes("\c\ok\")
        .Left = noSHP.Left + noSHP.Width
        .Top = noSHP.Top
        .Width = noSHP.Width
        .Height = noSHP.Height
        .Visible = True
    End With
    shpCOLL.Add myWS.Shapes("\c\ok\")
    
    'this is the caption box
    myW = myWS.[\body].Cells(1, 1).Left - (noSHP.Left + noSHP.Width)
    Set captionSHP = myWS.Shapes.AddShape(msoShapeRectangle, myWS.Shapes("\c\ok\").Left + myWS.Shapes("\c\ok\").Width, activeROW.Top, myW - myWS.Shapes("\c\ok\").Width, noSHP.Height)
    With captionSHP
        .Line.Visible = False
        .Placement = xlMove
        With .TextFrame.Characters
            .Font.Color = SCCIblue
            .Font.Name = "Futura Bk BT"
            .Font.Size = 11
            .Font.Bold = True
            .Text = captionSTR
        End With
        .Fill.ForeColor.RGB = vbWhite
        .Name = "\c\caption~"
    End With
    shpCOLL.Add captionSHP
    
    'this is a clear box over the check icon and caption
    Set yesSHP = myWS.Shapes.AddShape(msoShapeRectangle, noSHP.Left + noSHP.Width, activeROW.Top, myW, noSHP.Height)
    With yesSHP
        .Line.Visible = False
        .Fill.Visible = False
        .Name = "\c\ok~"
        .Placement = xlMove
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "context_click"
    End With
    shpCOLL.Add yesSHP
    
   
Exit Sub
e1:
    LogError "contextMENUcls", "promote", Err.Description, Err
End Sub

Sub parseCMD(cmdSTR As String)
On Error GoTo e1

    Debug.Print myWS.Name & "==>" & cmdSTR
    
    myWS.Unprotect
    Select Case cmdSTR
        Case "\c_main"
            assembleMENU
        Case "\c_x"
            collapse
        Case "\c\updown_\"
            MODE_move
        Case "\c\minus_\"
            MODE_delete
        Case "\c\H_\"
            addTEMPLATE "\r_heading"
        Case "\c\a_\"
            addTEMPLATE "\r_lineitem"
        Case "\c\arrow_\"
            executeXFER
        Case "\c\plus_\"
            addTEMPLATE "\r_blank"
        Case "\c\cancel~"
            MODE_cancel
        Case "\c\ok~"
            Select Case modeSTR
                Case "move"
                    MODE_cancel
                    executeMOVE
                Case "delete"
                    MODE_cancel
                    executeDELETE
                Case Else
                    MODE_cancel
            End Select
    End Select
    basicPROTECT myWS, True
    
Exit Sub
e1:
    LogError "contextMENUcls", "parseCMD", Err.Description, Err
    basicPROTECT myWS, True
    
End Sub

'============================================================
'                  BEHAVIOURS
'============================================================

Sub collapse()
On Error GoTo e1

    Application.ScreenUpdating = False
    
    dockBTNS
    xSHP.Visible = False

    'Set activeMenu = Nothing
    Application.ScreenUpdating = True

Exit Sub
e1:
    LogError "contextMENUcls", "collapse", Err.Description, Err
End Sub

Sub dockBTNS()
On Error GoTo e1

    Dim shp As Shape
    
    For Each shp In myWS.Shapes
        If Left(shp.Name, Len(pFIX)) = pFIX Then
            shp.Left = myWS.[\shapeSTORE].Left
            shp.Top = myWS.[\shapeSTORE].Top
            shp.Visible = False
        End If
    Next
    killTempShapes
    
Exit Sub
e1:
    LogError "contextMENUcls", "dockBTNS", Err.Description, Err
End Sub

Sub dockMAIN()
On Error GoTo e1

    Dim dockRAN As Range
    
    Set dockRAN = myWS.[\shapeSTORE]
    
    With masterSHP
        .Top = dockRAN.Top
        .Left = dockRAN.Left
        .Visible = False
    End With
    With lineSHP
        .Top = dockRAN.Top
        .Left = dockRAN.Left
        .Visible = False
    End With
    With xSHP
        .Top = dockRAN.Top
        .Left = dockRAN.Left
        .Visible = False
    End With
    dockBTNS

Exit Sub
e1:
    LogError "contextMENUcls", "dockMAIN", Err.Description, Err
End Sub

'============================================================
'                  EVENTS
'============================================================

Private Sub moveWS_SelectionChange(ByVal Target As Range)
On Error GoTo e1

    Dim docRAN As Range
    Dim tempROW As Range
    Dim vSHIFT As Double
    Dim lastTOP As Double
    
    'moveWS.Unprotect
    
    Set docRAN = myWS.[\body]
    'reset last
    If Not lastRAN Is Nothing Then
        lastRAN.EntireRow.RowHeight = lastROWh
    End If
        
    If Union(ActiveCell, docRAN).Address = docRAN.Address Then
        Set tempROW = ActiveCell.EntireRow
        Set insertRAN = tempROW
        'record for reset
        lastROWh = tempROW.RowHeight
        Set lastRAN = tempROW
        
        
        tempROW.RowHeight = lastROWh + moveSHP.Height
        vSHIFT = tempROW.Top - moveSHP.Top
        collSHIFT shpCOLL, vSHIFT
    End If
    
    'basicPROTECT moveWS, True
    
Exit Sub
e1:
    LogError "contextMENUcls", "moveWS_SelectionChange", Err.Description, Err
End Sub

Private Sub WS_Activate()
On Error GoTo e1

    'ActiveWindow.DisplayHeadings = False
    Set activeMenu = Me 'added 4/23 colby
    
Exit Sub
e1:
    LogError "contextMENUcls", "WS_Activate", Err.Description, Err
End Sub

Private Sub WS_Deactivate()
On Error GoTo e1

    'ActiveWindow.DisplayHeadings = True
    dockMAIN

Exit Sub
e1:
    LogError "contextMENUcls", "WS_Deactivate", Err.Description, Err
End Sub

Private Sub WS_SelectionChange(ByVal Target As Range)
On Error GoTo e1

    Dim documentRAN As Range
    Dim xferRAN As Range
    
    If myWS.[\c_settings].EntireColumn.Hidden Then Exit Sub
    
    Set documentRAN = WS.[\body]
    Set xferRAN = WS.[\xfer]
    trimRANGE xferRAN, dsupdown
    If xferRAN Is Nothing Then Set xferRAN = WS.[xx1000] 'not nothing
    
    If Union(ActiveCell, documentRAN.EntireRow).Address = documentRAN.EntireRow.Address Then
        dockBTNS
        anchor ActiveCell
    ElseIf Union(ActiveCell, xferRAN.EntireRow).Address = xferRAN.EntireRow.Address Then
        dockBTNS
        anchor ActiveCell
    Else
        dockMAIN
    End If
    
Exit Sub
e1:
    LogError "contextMENUcls", "WS_SelectionChange", Err.Description, Err
End Sub

