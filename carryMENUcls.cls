VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "carryMENUcls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<navarro>
Option Compare Text 'added for createREPORT method : ajoku 4/20
Private WithEvents myWS As Worksheet
Attribute myWS.VB_VarHelpID = -1
Private carrySHP As Shape
Private burgerSHP As Shape
Private carryCOL As Range
Private groupCOL As Range
Private contextCOL As Range
Private activeROW As Range
Private Const pFIXc As String = "\c\"
Private Const pFIXm As String = "\m\"
Private xSHP As Shape
Private Const btnW As Double = 21
Private Const btnH As Double = 21
Private Const btnW2 As Double = 80
Private Const lH As Double = 2
Private Const lMargin As Double = 10
Private Const btnIndent As Double = 4


'============================================================
'                  INITIALIZE
'============================================================

Sub register(sht As Worksheet)
On Error GoTo e1

    Set myWS = sht
    Set carrySHP = myWS.Shapes("\m\carry")
    Set burgerSHP = myWS.Shapes("\m\burger")
    Set xSHP = myWS.Shapes("\m\x")
    Set contextCOL = myWS.[\c_context].EntireColumn
    Set carryCOL = myWS.[\c_carry].EntireColumn
    Set groupCOL = myWS.[\c_group].EntireColumn
    Set activeMenu = Me
    'killTempShapes
    carrySHP.Visible = False
    burgerSHP.Visible = False
    dockBTNS
    
Exit Sub
e1:
    LogError "carryMENUcls", "register", Err.Description, Err
End Sub

Sub anchor(RAN As Range)
On Error GoTo e1

    Dim tran As Range
    Set tran = Intersect(RAN.EntireRow, contextCOL)
    
    EnS 0
    
    dockBTNS
    
    carrySHP.Top = RAN.Top - (carrySHP.Height - RAN.Height) / 2
    carrySHP.Left = myWS.[\c_desc].Left - carrySHP.Width
    
    With burgerSHP
        .Height = tran.RowHeight
        .Width = .Height
        .Top = tran.Top
        .Left = tran.Left - .Width
        xSHP.Height = .Height
        xSHP.Width = .Width
        xSHP.Left = .Left
        xSHP.Top = .Top
        xSHP.Visible = False
    End With
    
    
    Set activeMenu = Me
    Set activeROW = activecell.EntireRow
    
    If isCARRYrow(RAN) Then
        burgerSHP.Visible = True
        carrySHP.Visible = False
    ElseIf Intersect(RAN.EntireRow, myWS.[\c_val].EntireColumn).Value > 0 Then
        carrySHP.Visible = True
        burgerSHP.Visible = False
    Else
        carrySHP.Visible = False
        burgerSHP.Visible = False
    End If
    
    EnS 1
    
Exit Sub
e1:
    LogError "carryMENUcls", "anchor", Err.Description, Err
    EnS 1, , True
End Sub

'============================================================
'                  FUNCTIONS
'============================================================

Private Function isCARRYrow(RAN As Range) As Boolean
On Error GoTo e1

    isCARRYrow = Intersect(RAN.EntireRow, carryCOL).Value = 1

Exit Function
e1:
    LogError "carryMENUcls", "isCARRYrow", Err.Description, Err
End Function

Private Function isHEADERtotal(RAN As Range) As Boolean
On Error GoTo e1

    isHEADERtotal = Intersect(RAN.EntireRow, myWS.[\c_group].EntireColumn).Value = "["

Exit Function
e1:
    LogError "carryMENUcls", "isHEADERtotal", Err.Description, Err
End Function

Private Function myTOTALrange_Address(RAN As Range) As String
On Error GoTo e1

    'returns address in the correct format
    Dim tran As Range
    
    Set tran = Intersect(RAN.EntireRow, myWS.[\c_group].EntireColumn)
    
    Do Until tran.Value = "["
        Set tran = tran.Offset(-1, 0)
        If tran.row = 1 Then GoTo e1
    Loop
    
    Set tran = Intersect(tran.EntireRow, myWS.[\c_val].EntireColumn)
    myTOTALrange_Address = tran.Address
    
Exit Function
e1:
    Debug.Print "myTOTALrange failed on cell " & RAN.Parent.Name & "!" & RAN.Address
    myTOTALrange = ""
    LogError "carryMENUcls", "myTOTALrange_Address", Err.Description, Err
End Function

'============================================================
'                  METHODS
'============================================================

Sub dockBTNS()
On Error GoTo e1

    Dim shp As Shape
    
    For Each shp In myWS.Shapes
        If Left(shp.Name, Len(pFIXc)) = pFIXc Or Left(shp.Name, Len(pFIXm)) = pFIXm Then
            With myWS.[\shapeSTORE]
                shp.Visible = False
                shp.Left = .Left
                shp.Top = .Top
            End With
        End If
    Next

Exit Sub
e1:
    LogError "carryMENUcls", "dockBTNS", Err.Description, Err
End Sub

Sub parseCMD(cmdSTR As String)
On Error GoTo e1

    Debug.Print myWS.Name & "==>" & cmdSTR
    
    myWS.Unprotect
    Select Case cmdSTR
        Case "\m\carry"
            activateCARRY
            SYNCrequired True
        Case "\m\burger"
            assembleMENU
        Case "\m\x"
            anchor activecell
        Case "\c\minus_\"
            removeCARRY
            SYNCrequired True
        Case "\s\sync"
            createREPORT
            'SYNCrequired False 'TO DO - need procedure here to sync
    End Select
    basicPROTECT myWS, True
    
Exit Sub
e1:
    LogError "carryMENUcls", "parseCMD", Err.Description, Err
    basicPROTECT myWS, True
    
End Sub

Sub assembleMENU()
On Error GoTo e1

    Dim shp As Shape, shp2 As Shape, shp3 As Shape
    Dim masterL As Double
    Dim lastTOP As Double

    Application.ScreenUpdating = False
    
    masterL = burgerSHP.Left + btnIndent
    lastTOP = burgerSHP.Top + burgerSHP.Height + 1
    
    For Each shp In myWS.Shapes
        If Left(shp.Name, Len(pFIXc)) = pFIXc And Right(shp.Name, 1) <> "\" Then
            
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
    LogError "carryMENUcls", "assembleMENU", Err.Description, Err
End Sub

Sub SYNCrequired(boo As Boolean)
On Error GoTo e1

    Dim RAN As Range
    
    Set RAN = myWS.[\sync]
    
    With myWS.Shapes("\s\sync")
        .Left = RAN.Left
        .Top = RAN.Top
        .Visible = boo
    End With

    With myWS.Shapes("\s\ready")
        .Left = RAN.Left
        .Top = RAN.Top
        .Visible = Not boo
    End With

    If boo Then
        myWS.[\sync].Value = ""
    Else
        myWS.[\sync].Value = 1
    End If
    
Exit Sub
e1:
    LogError "carryMENUcls", "SYNCrequired", Err.Description, Err
End Sub



'============================================================
'                  HANDLE CARRY
'============================================================

Private Sub removeCARRY()
On Error GoTo e1

    Dim RAN As Range
    Dim formatRAN As Range
    Dim allCARRYS As Range
    Dim headerRAN As Range
    
    Set RAN = Intersect(activeROW, myWS.[\c_context].EntireColumn)
    Set formatRAN = myWS.[\defaultFORMAT]
    
    EnS 0
    RAN.ClearContents
    formatRAN.Copy
    RAN.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    
    Intersect(RAN.EntireRow, myWS.[\c_carry].EntireColumn).Value = ""
    Intersect(RAN.EntireRow, myWS.[\c_carry].EntireColumn).Offset(0, 4).Value = ""
    
    Set allCARRYS = RAN
    Transform2AllCarrys allCARRYS, headerRAN
    
    If Not allCARRYS Is Nothing Then
        ProcessFormulasWith allCARRYS.Cells(1, 1)
    ElseIf Not headerRAN Is Nothing Then
        LinkMe2myTotal headerRAN
    End If
    dockBTNS
    EnS 1

Exit Sub
e1:
    LogError "carryMENUcls", "removeCARRY", Err.Description, Err
    EnS 1, , True
End Sub

Private Sub activateCARRY()
On Error GoTo e1

    Dim RAN As Range
    Dim formatRAN As Range
    
    Set RAN = Intersect(activeROW, myWS.[\c_context].EntireColumn)
    Set formatRAN = myWS.[\carryFORMAT]
    
    EnS 0
    
    formatRAN.Copy
    RAN.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    
    Intersect(RAN.EntireRow, myWS.[\c_carry].EntireColumn).Value = 1
    
    If isHEADERtotal(RAN) Then
        handleHEADER RAN
    Else
        ProcessFormulasWith RAN
    End If
    
    anchor RAN
    EnS 1
    
Exit Sub
e1:
    LogError "carryMENUcls", "activateCARRY", Err.Description, Err
    EnS 1, , True
End Sub

Private Sub ProcessFormulasWith(ByRef RAN As Range)
On Error GoTo e1

    Dim myTOTAL As Range
    Dim allCARRYS As Range
    Dim headerADDRESS As String
    Dim headerRAN As Range
    Dim cell As Range
    
    Set myTOTAL = Intersect(RAN.EntireRow, myWS.[\c_val].EntireColumn)
    Set allCARRYS = RAN
    Transform2AllCarrys allCARRYS, headerRAN, True
    
    If headerRAN Is Nothing Then
        'header is not one of the carrys, so prorate all carrys
        headerADDRESS = myTOTALrange_Address(RAN)
        For Each cell In allCARRYS.Cells
            Intersect(cell.EntireRow, myWS.[\c_context].EntireColumn).Formula = _
                "=" & cell.Address(False, True) & _
                "/sum(" & allCARRYS.Address & ")" & _
                "*" & headerADDRESS
        Next
    Else
        'header is part of the carrys, so make it the catch all and link all other carrys to there total
        For Each cell In allCARRYS.Cells
            LinkMe2myTotal cell
        Next
        Set allCARRYS = Intersect(allCARRYS.EntireRow, myWS.[\c_context].EntireColumn)
        Intersect(headerRAN.EntireRow, myWS.[\c_context].EntireColumn).Formula = _
            "=" & headerRAN.Address & "-sum(" & allCARRYS.Address & ")"
    End If

Exit Sub
e1:
    LogError "carryMENUcls", "ProcessFormulasWith", Err.Description, Err
End Sub

Private Sub handleHEADER(ByRef headerRAN As Range)
On Error GoTo e1

    Dim allCARRYS As Range
    Dim cell As Range
    
    Set allCARRYS = headerRAN
    Transform2AllCarrys allCARRYS
    
    If allCARRYS Is Nothing Then
        'only the header is a carry, link to its total
        LinkMe2myTotal headerRAN
    Else
        'other carrys exist, link them to their total, make the header a catch
        For Each cell In allCARRYS.Cells
            LinkMe2myTotal cell
        Next
        Set allCARRYS = Intersect(allCARRYS.EntireRow, myWS.[\c_context].EntireColumn)
        headerRAN.Formula = "=" & Intersect(headerRAN.EntireRow, myWS.[\c_val].EntireColumn).Address & _
                      "-sum(" & allCARRYS.Address & ")"
    End If

Exit Sub
e1:
    LogError "carryMENUcls", "handleHEADER", Err.Description, Err
End Sub

Private Sub makeMEaCatch(ByRef RAN As Range)


End Sub

Private Sub LinkMe2myTotal(ByRef RAN As Range)
On Error GoTo e1

    Dim tran As Range
    Dim t2ran As Range
    
    Set tran = Intersect(RAN.EntireRow, myWS.[\c_val].EntireColumn)
    Set t2ran = Intersect(RAN.EntireRow, myWS.[\c_context].EntireColumn)
    t2ran.Formula = "=" & tran.Address(False, False)

Exit Sub
e1:
    LogError "carryMENUcls", "LinkMe2myTotal", Err.Description, Err
End Sub

Private Sub Transform2AllCarrys(RAN As Range, Optional ByRef headerRAN As Range, Optional includeME As Boolean)
On Error GoTo e1

    'returns range
    Dim groupRAN As Range
    Dim carryRAN As Range
    Dim WS As Worksheet
    
    Set groupRAN = Intersect(RAN.EntireRow, myWS.[\c_group].EntireColumn)
    Set WS = groupRAN.Parent
    
    Do Until groupRAN.Value = "["
        Set groupRAN = groupRAN.Offset(-1, 0)
        If groupRAN.row = 1 Then GoTo e1
    Loop
    
    Do Until groupRAN.Value = "]"
        If Intersect(groupRAN.EntireRow, myWS.[\c_carry].EntireColumn) = 1 Then
            If includeME Then
                If groupRAN.Value = "[" Then
                    Set headerRAN = Intersect(groupRAN.EntireRow, myWS.[\c_val].EntireColumn)
                Else
                    If carryRAN Is Nothing Then Set carryRAN = groupRAN Else Set carryRAN = Union(carryRAN, groupRAN)
                End If
            ElseIf RAN.row <> groupRAN.row Then
                If groupRAN.Value = "[" Then
                    Set headerRAN = Intersect(groupRAN.EntireRow, myWS.[\c_val].EntireColumn)
                Else
                    If carryRAN Is Nothing Then Set carryRAN = groupRAN Else Set carryRAN = Union(carryRAN, groupRAN)
                End If
            End If
        End If
        
        Set groupRAN = groupRAN.Offset(1, 0)
        If groupRAN.Value = "[" Or groupRAN.row >= WS.UsedRange.Rows.Count Then GoTo e1
        
    Loop
    
    If Not carryRAN Is Nothing Then
        Set RAN = Intersect(carryRAN.EntireRow, myWS.[\c_val].EntireColumn)
    Else
        Set RAN = Nothing
    End If
    
Exit Sub
e1:
    Debug.Print "Transform2AllCarrys failed on cell " & RAN.Parent.Name & "!" & RAN.Address
    LogError "carryMENUcls", "Transform2AllCarrys", Err.Description, Err
End Sub

Private Sub createREPORT()
On Error GoTo e1

    Dim destWS As Worksheet
    Dim headerRAN As Range
    Dim RAN As Range
    Dim cell As Range
    Dim tran As Range
    Dim groupRAN As Range
    Dim carryRAN As Range, formRAN As Range
    Dim carryFORM As String
    Dim insertRAN As Range
    Dim insertSTR As String
    Dim destRAN As Range
    Dim carryCOLL As Collection, syncCOLL As Collection, headerCOLL As Collection, valCOLL As Collection
    Dim i As Integer
    
    If myWS.Name = masterOBJ.gcdWS.Name Then
        Set destWS = masterOBJ.gcoWS
    ElseIf myWS.Name = masterOBJ.grdWS.Name Then
        Set destWS = masterOBJ.groWS
    End If
    
    Set destRAN = Intersect(destWS.[\c_desc].EntireColumn, destWS.[\body])
    Set carryCOLL = New Collection
    Set syncCOLL = New Collection
    Set headerCOLL = New Collection
    Set valCOLL = New Collection
    Set carryRAN = boxRANGE(myWS, "\r_start", "\c_carry", "\r_end")
    trimRANGE carryRAN, dsupdown
    
    For Each cell In carryRAN
        Set tran = Intersect(cell.EntireRow, myWS.[\c_desc].EntireColumn)
        If cell.Value = 1 And cell.Offset(0, 4).Value <> 1 Then
            carryCOLL.Add tran
            cell.Offset(0, 4).Value = 1
        End If
    Next
    
    If carryCOLL.Count = 0 Then
        MsgBox "There Are No Carry Items To Sync", vbOKOnly, "Sync Failed"
        Exit Sub
    Else
        SYNCrequired False
    End If
    
    For i = 1 To carryCOLL.Count
        Set groupRAN = Intersect(carryCOLL(i).EntireRow, myWS.[\c_group].EntireColumn)
        Do Until groupRAN.Value = "["
            Set groupRAN = groupRAN.Offset(-1, 0)
        Loop
        insertSTR = Intersect(groupRAN.EntireRow, myWS.[\c_desc].EntireColumn).Value
        Set formRAN = Intersect(carryCOLL(i).EntireRow, myWS.[\c_context].EntireColumn)
        carryFORM = "='" & formRAN.Parent.Name & "'!" & formRAN.Address
        syncCOLL.Add carryCOLL(i)
        headerCOLL.Add insertSTR
        valCOLL.Add carryFORM
    Next
    
    EnS 0
    destWS.Unprotect
    destWS.[\r_lineitem].EntireRow.Hidden = False
    
    For i = 1 To syncCOLL.Count
        Set insertRAN = destRAN.Find(What:=syncCOLL(i).Value)
        If Not insertRAN Is Nothing Then Set insertRAN = insertRAN.Offset(1, 0)
        If insertRAN Is Nothing Then
            Set insertRAN = destRAN.Find(What:=headerCOLL(i))
            If Not insertRAN Is Nothing Then Set insertRAN = insertRAN.Offset(1, 0)
        End If
        If insertRAN Is Nothing Then Set insertRAN = destWS.[\r_xfer].Offset(1, 0)
        insertRAN.EntireRow.Insert
        With insertRAN.Offset(-1, 0)
            Intersect(.EntireRow, destWS.[\c_desc].EntireColumn).Value = StrConv(syncCOLL(i).Value, vbProperCase)
            Intersect(.EntireRow, destWS.[\c_total].EntireColumn).Formula = valCOLL(i)
            destWS.[\r_lineitem].EntireRow.Copy
            .EntireRow.PasteSpecial xlPasteFormats
        End With
    Next
    
    destWS.[\r_lineitem].EntireRow.Hidden = True
    Application.CutCopyMode = False
    basicPROTECT destWS, True
    EnS 1

Exit Sub
e1:
    LogError "carryMENUcls", "createREPORT", Err.Description, Err
    EnS 1, , True
    basicPROTECT destWS, True
    
End Sub

'============================================================
'                  EVENTS
'============================================================

Private Sub myWS_Activate()
On Error GoTo e1

    Set activeMenu = Me
    
Exit Sub
e1:
    LogError "carryMENUcls", "myWS_Activate", Err.Description, Err
End Sub

Private Sub myWS_SelectionChange(ByVal Target As Range)
On Error GoTo e1

    Dim RAN As Range
    Dim showRAN As Range
    Dim boo1 As Boolean, boo2 As Boolean
    
    Set RAN = Intersect(activecell.EntireRow, contextCOL)
    
    Set contextCOL = myWS.[\c_context].EntireColumn
    Set carryCOL = myWS.[\c_carry].EntireColumn
    Set groupCOL = myWS.[\c_group].EntireColumn
    
    If Union(RAN, contextCOL).Address = contextCOL.Address And Not myWS.[\c_settings].EntireColumn.Hidden Then
        If Intersect(RAN.EntireRow, groupCOL).Value = 1 Then
            anchor RAN
        ElseIf Intersect(RAN.EntireRow, groupCOL).Value = "[" Then
            anchor RAN
        Else
            dockBTNS
        End If
    Else
        dockBTNS
    End If
    
Exit Sub
e1:
    LogError "carryMENUcls", "myWS_SelectionChange", Err.Description, Err
End Sub





