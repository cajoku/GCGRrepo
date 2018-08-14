Attribute VB_Name = "Navarro"
Option Explicit
'Public cMenuCOLL As Collection
'Public cMenu As contextMENUcls
Public activeMenu As Variant
Public cmenu As contextMENUcls

Sub death2Links_names()
    Dim nm As Name
    Dim arr() As String
    Dim str As String
    
    For Each nm In ActiveWorkbook.Names
        If InStr(nm.RefersTo, "]") > 0 Then
            arr = Split(nm.RefersTo, "]")
            str = Replace(arr(0), "=", "") & "]"
            nm.RefersTo = Replace(nm.RefersTo, str, "'")
        End If
    Next


End Sub

Sub fixFORM()
    Dim sht As Worksheet
    Dim arr() As String
    Dim str As String
    Dim cell As Range
    Dim i As Integer, j As Integer
    Dim nm As Name
    
    Set sht = ActiveSheet
    
    
    For Each cell In sht.UsedRange.Cells
        If InStr(cell.Formula, "]") > 0 Then
          '  i = instr(cell.Formula,"'"
            
            
            
            arr = Split(cell.Formula, "]")
            str = Replace(arr(0), "=", "") & "]"
            nm.RefersTo = Replace(nm.RefersTo, str, "'")
        End If
    Next
    
    For Each nm In ActiveWorkbook.Names
        If InStr(nm.RefersTo, "]") > 0 Then
            arr = Split(nm.RefersTo, "]")
            str = Replace(arr(0), "=", "") & "]"
            nm.RefersTo = Replace(nm.RefersTo, str, "'")
        End If
    Next


End Sub

Sub hkjl()

Dim shp As Shape
For Each shp In ActiveSheet.Shapes
    shp.LockAspectRatio = False
    
    
Next

End Sub

Sub callCONTEXT()
    
    Set cmenu = New contextMENUcls
    cmenu.register ActiveSheet
    'Set cMenuCOLL = New Collection
   ' cMenuCOLL.Add cmenu
    
End Sub

Sub context_click()
On Error GoTo e1
        
    activeMenu.parseCMD (Application.Caller)
    
'    Select Case Application.Caller
'        Case "\c_main"
'            cMenu.assembleMENU
'        Case "\c_x"
'            cMenu.collapse
'        Case Else
'            MsgBox Application.Caller
'    End Select
Exit Sub
e1:
    LogError "Navarro", "context_click", Err.Description, Err
End Sub


'====================================================
'   ADMIN
'====================================================

Sub navALIGN()
    
    Dim WS As Worksheet
    Dim shp As Shape
    Dim i As Integer
    Dim myW As Double
    Dim lastL As Double
    
    Set WS = ActiveSheet
    myW = 144
    lastL = 39
    For i = 1 To 10
        On Error GoTo e1
        Set shp = WS.Shapes("\N\" & i)
        With shp
            .Top = 0
            .Width = myW
            .Height = shp.TopLeftCell.RowHeight
            .Left = lastL
            lastL = lastL + myW
        End With
    Next
    
e1:

End Sub

Sub shpOnAction_realignALL()
    Dim shp As Shape
    Dim WS As Worksheet
    Dim cmdSTR As String
    Dim arr() As String
    
    Set WS = ActiveSheet
    
    Debug.Print "==Realign Shape OnAction for: " & WS.Name
    
    For Each shp In WS.Shapes
        cmdSTR = shp.OnAction
        
        Debug.Print "~~" & shp.Name & ".OnAction= " & shp.OnAction
        
        If cmdSTR <> "" Then
            arr = Split(cmdSTR, "!")
            If UBound(arr) > 0 Then
                shp.OnAction = arr(1)
            End If
        End If
        
        Debug.Print " >" & shp.Name & ".OnAction= " & shp.OnAction
        
    Next



End Sub

Sub toggleADMINmode()
On Error GoTo e1

    Dim WS As Worksheet
    
    If masterOBJ Is Nothing Then onOPEN
    
    Set WS = ActiveSheet
    On Error Resume Next
    WS.[\r_admin].EntireRow.Hidden = Not WS.[\c_admin].EntireColumn.Hidden
    WS.[\c_admin].EntireColumn.Hidden = Not WS.[\c_admin].EntireColumn.Hidden
    
Exit Sub
e1:
    LogError "Navarro", "toggleADMINmode", Err.Description, Err
End Sub

Sub formatGROUPbtns()
On Error GoTo e1

    Dim WS As Worksheet
    Dim shp As Shape
    Dim boo As Boolean
    
    Set WS = ActiveSheet
    
    For Each shp In WS.Shapes
        boo = False
        If InStr(shp.Name, "less") > 0 Then
            shp.OnAction = "'" & ThisWorkbook.Name & "'!" & "less_click"
            boo = True
        ElseIf InStr(shp.Name, "more") > 0 Then
            shp.OnAction = "'" & ThisWorkbook.Name & "'!" & "more_click"
            boo = True
        End If
        
        If boo Then
            With shp.TopLeftCell
                shp.Top = .Top + .Height / 2 - shp.Height / 2
                shp.Left = .Left + .Width - shp.Width - (.Height - shp.Height) / 2
            End With
        End If
            
    Next

Exit Sub
e1:
    LogError "Navarro", "formatGROUPbtns", Err.Description, Err
End Sub



'====================================================
'   PROCEDURES
'====================================================

Sub toggleNAV(WS As Worksheet, Optional IsVisible As Boolean)
On Error GoTo e1

    Dim shp As Shape
    
    For Each shp In WS.Shapes
        If Left(shp.Name, 3) = "\N\" Then shp.Visible = IsVisible
    Next

Exit Sub
e1:
    LogError "Navarro", "toggleNAV", Err.Description, Err
End Sub

Sub hideENDS(WS As Worksheet)
On Error GoTo e1

    Dim hideROW As Range
    Dim hideCOL As Range
    Dim RAN As Range
    'Dim WS As Worksheet
    
    On Error Resume Next
    Set hideROW = WS.[\r_adminEND].EntireRow
    Set hideCOL = WS.[\c_adminEND].EntireColumn
    With WS
        Set RAN = .Rows(.Rows.Count).EntireRow
        Range(hideROW, RAN).EntireRow.Hidden = True
        Set RAN = .Columns(.Columns.Count).EntireColumn
        Range(hideCOL, RAN).EntireColumn.Hidden = True
    End With
    On Error GoTo 0
Exit Sub
e1:
    LogError "Navarro", "hideENDS", Err.Description, Err
End Sub

Sub toggleGC_GR(RAN As Range)
    
    Dim onBOO As Boolean
    Dim i As Integer
    
    onBOO = (RAN.Style = "btnON")
    EnS 0
      
    If RAN.Value = "GC" Then
        i = 1
    ElseIf RAN.Value = "GR" Then
        i = -1
    Else
        Exit Sub
    End If
    
    If RAN.Style = "btnON" Then
        RAN.Style = "btnOFF"
    ElseIf RAN.Style = "btnOFF" Then
        RAN.Style = "btnON"
        RAN.Offset(0, i).Style = "btnOFF"
    End If

    RAN.Parent.[\r_settings].Cells(1, 1).Select
    
    'sync set to needed
    With RAN.Parent.[\sync]
        .Style = "syncNEED"
        .Value = "q"
        .Offset(0, 1).Style = "adminRED"
    End With

    EnS 1

End Sub

Sub collSHIFT(COLL As Collection, Optional vMOVE As Double, Optional hMOVE As Double)
On Error GoTo e1

    Dim var
    
    For Each var In COLL
        var.Left = var.Left + hMOVE
        var.Top = var.Top + vMOVE
    Next

Exit Sub
e1:
    LogError "Navarro", "collSHIFT", Err.Description, Err
End Sub
'====================================================
'   BUTTON CLICKS
'====================================================

Sub dummy()

End Sub

Sub clearContext()
On Error GoTo e1

Dim var As contextMENUcls

On Error Resume Next
For Each var In masterOBJ.cMenuCOLL
    var.dockMAIN
    'clear temp shapes
    'reset modes
Next
On Error GoTo 0

Exit Sub
e1:
    LogError "Navarro", "clearContext", Err.Description, Err
    
End Sub

Sub plusCLICK()
On Error GoTo ehandle

    Dim WS As Worksheet
    
    'navarro|  users tend to click buttons when it appears the programming is 'off' so
    '           so button clicks are a great opportunity to 'turn on' the program if an
    '           unhandled error or break in the code caused all the objects to terminate
    If masterOBJ Is Nothing Then onOPEN
    
    EnS 0
    clearContext
    
    Set WS = ActiveSheet
    WS.Shapes("\\plus").Visible = False
    WS.Shapes("\\minus").Visible = True
    WS.[\r_settings].EntireRow.Hidden = False
    On Error Resume Next
    WS.[\c_settings].EntireColumn.Hidden = False
    
    '<navarro 4-19-18
'    If WS.[\sync].Value = 1 Then
'        WS.Shapes("\s\ready").Visible = True
'    Else
'        WS.Shapes("\s\sync").Visible = True
'    End If
    'navarro 4-19-18>
    On Error GoTo 0
    
    hideENDS WS
    toggleNAV WS, True
    EnS 1
    
    
Exit Sub
     
ehandle:
    LogError "Navarro", "plusCLICK", Err.Description, Err
    EnS 1, , True
End Sub

Sub minusCLICK()
On Error GoTo ehandle

    Dim WS As Worksheet
    
    'navarro|  users tend to click buttons when it appears the programming is 'off' so
    '           so button clicks are a great opportunity to 'turn on' the program if an
    '           unhandled error or break in the code caused all the objects to terminate
    If masterOBJ Is Nothing Then onOPEN
    
    EnS 0
    clearContext
    activeMenu.dockBTNS
    
    Set WS = ActiveSheet
    WS.Shapes("\\plus").Visible = True
    WS.Shapes("\\minus").Visible = False
    WS.[\r_settings].EntireRow.Hidden = True
    On Error Resume Next
    WS.[\c_settings].EntireColumn.Hidden = True
    WS.Shapes("\s\sync").Visible = False '<navarro 4-19-18>
    WS.Shapes("\s\ready").Visible = False '<navarro 4-19-18>
    On Error GoTo 0
    hideENDS WS
    toggleNAV WS, False
    EnS 1
    
Exit Sub
    
    
    
ehandle:
    LogError "Navarro", "minusCLICK", Err.Description, Err
    EnS 1, , True
End Sub

Sub moreMONTHdetailCLICK()
On Error GoTo ehandle

    Dim WS As Worksheet
    
    Set WS = ActiveSheet
    WS.Shapes("\\moreMONTHdetail").Visible = False
    WS.Shapes("\\lessMONTHdetail").Visible = True
    WS.[\c_monthDETAIL].EntireColumn.Hidden = False
    
Exit Sub

ehandle:
    LogError "Navarro", "moreMONTHdetailCLICK", Err.Description, Err
End Sub

Sub lessMONTHdetailCLICK()
On Error GoTo ehandle

    Dim WS As Worksheet
    
    Set WS = ActiveSheet
    WS.Shapes("\\moreMONTHdetail").Visible = True
    WS.Shapes("\\lessMONTHdetail").Visible = False
    WS.[\c_monthDETAIL].EntireColumn.Hidden = True

Exit Sub

ehandle:
    LogError "Navarro", "lessMONTHdetailCLICK", Err.Description, Err
End Sub

Sub moreRATEdetailCLICK()
On Error GoTo ehandle

    Dim WS As Worksheet
    
    Set WS = ActiveSheet
    WS.Shapes("\\moreRATEdetail").Visible = False
    WS.Shapes("\\lessRATEdetail").Visible = True
    WS.[\c_rateDETAIL].EntireColumn.Hidden = False

Exit Sub

ehandle:
    LogError "Navarro", "moreRATEdetailCLICK", Err.Description, Err
End Sub

Sub lessRATEdetailCLICK()
On Error GoTo ehandle

    Dim WS As Worksheet
    
    Set WS = ActiveSheet
    WS.Shapes("\\moreRATEdetail").Visible = True
    WS.Shapes("\\lessRATEdetail").Visible = False
    WS.[\c_rateDETAIL].EntireColumn.Hidden = True

Exit Sub

ehandle:
    LogError "Navarro", "lessRATEdetailCLICK", Err.Description, Err
End Sub

Sub more_click()
On Error GoTo e1

    Dim targetSHP As Shape
    Dim shpSTR As String
    Dim WS As Worksheet
    Dim RAN As Range
    
    If masterOBJ Is Nothing Then onOPEN
    
    shpSTR = Application.Caller
    Set WS = ActiveSheet
    Set targetSHP = WS.Shapes(shpSTR)
    shpSTR = Replace(shpSTR, "more", "less")
    
    'hide me, show my evil twin brother "less"
    targetSHP.Visible = False
    WS.Shapes(shpSTR).Visible = True
    
    Set RAN = groupRAN(targetSHP.TopLeftCell, True)
    RAN.Hidden = False
    
Exit Sub
e1:
    LogError "Navarro", "more_click", Err.Description, Err
End Sub

Sub less_click()
On Error GoTo e1

    Dim targetSHP As Shape
    Dim shpSTR As String
    Dim WS As Worksheet
    Dim RAN As Range
    
    If masterOBJ Is Nothing Then onOPEN
    
    shpSTR = Application.Caller
    Set WS = ActiveSheet
    Set targetSHP = WS.Shapes(shpSTR)
    shpSTR = Replace(shpSTR, "less", "more")
    
    'hide me, show my evil twin brother "more"
    targetSHP.Visible = False
    WS.Shapes(shpSTR).Visible = True
    
    Set RAN = groupRAN(targetSHP.TopLeftCell)
    RAN.Hidden = True
Exit Sub
e1:
    LogError "Navarro", "less_click", Err.Description, Err
End Sub

Sub restoreDETAIL()
On Error GoTo e1

    Dim scanWS As Worksheet, destWS As Worksheet
    Dim descCOL As Range, uomCOL As Range, valCOL As Range, qtyCOL As Range, cutRAN As Range
    Dim groupCOL As Range
    Dim cell As Range, tempRAN As Range, descRAN As Range, totalRAN As Range
    Dim dCOLL As Collection, uCOLL As Collection, vCOLL As Collection, qCOLL As Collection
    ''''''
    Dim beginRAN As Range, valRAN2 As Range, qtyRAN2 As Range, uomRAN2 As Range
    Dim i As Integer
    Dim ans As Variant
    Dim pfrm As progressFRM
    Dim pcount As Integer


    
    ans = MsgBox("Are You Sure You Would Like to Revert Back to Defaults?", vbYesNo, "Restore Defaults")
    If ans = vbYes Then

        EnS 0
        
        Set pfrm = New progressFRM
        pfrm.progressON "Restoring Defaults", "Linking Detail Page"
        
        clearContext
        activeMenu.dockBTNS
    
        If ActiveSheet.CodeName = "Sheet4" Then
            Set scanWS = masterOBJ.grdWS
            Set destWS = masterOBJ.groWS
        ElseIf ActiveSheet.CodeName = "Sheet6" Then
            Set scanWS = masterOBJ.gcdWS
            Set destWS = masterOBJ.gcoWS
        End If
    
        Set descCOL = boxRANGE(scanWS, "\r_start", "\r_end", "\c_desc")
        Set groupCOL = scanWS.[\c_group].EntireColumn
        Set uomCOL = scanWS.[\c_qt].Offset(0, 1).EntireColumn
        Set valCOL = scanWS.[\c_val].Offset(0, -1).EntireColumn
        Set qtyCOL = scanWS.[\c_qt].EntireColumn
    
    
        Set cutRAN = Range(destWS.[\r_dstart].EntireRow, destWS.[\r_dend].EntireRow)
        trimRANGE cutRAN, dsupdown
    
        Set beginRAN = Intersect(destWS.[\r_dstart].EntireRow, destWS.[\c_desc].EntireColumn)
        'Set descRAN = destWS.[\c_desc].EntireColumn
        Set valRAN2 = destWS.[\c_desc].Offset(0, 4).EntireColumn
        Set qtyRAN2 = destWS.[\c_desc].Offset(0, 5).EntireColumn
        Set uomRAN2 = destWS.[\c_desc].Offset(0, 6).EntireColumn
        Set totalRAN = destWS.[\c_total].EntireColumn
    
        'scanWS.Unprotect
        destWS.Unprotect
    
        'scanWS.Rows.EntireRow.Hidden = False
        destWS.Rows.EntireRow.Hidden = False
        
'        destWS.[\r_heading].EntireRow.Hidden = False
'        destWS.[\r_lineitem].EntireRow.Hidden = False
'        destWS.[\r_blank].EntireRow.Hidden = False
        
        cutRAN.Delete
        
        pcount = 1
        For Each cell In descCOL
            
            If Intersect(cell.EntireRow, groupCOL).Value = "[" Then
                Set dCOLL = New Collection
                Set uCOLL = New Collection
                Set vCOLL = New Collection
                Set qCOLL = New Collection
    
                dCOLL.Add cell
                uCOLL.Add Intersect(cell.EntireRow, uomCOL)
                vCOLL.Add Intersect(cell.EntireRow, valCOL)
                qCOLL.Add Intersect(cell.EntireRow, qtyCOL)
    
                Set tempRAN = cell
                Do Until Intersect(tempRAN.EntireRow, groupCOL).Value = "]"
                    Set tempRAN = tempRAN.Offset(1, 0)
                    dCOLL.Add tempRAN
                    uCOLL.Add Intersect(tempRAN.EntireRow, uomCOL)
                    vCOLL.Add Intersect(tempRAN.EntireRow, valCOL)
                    qCOLL.Add Intersect(tempRAN.EntireRow, qtyCOL)
                Loop
             'End If
                'With beginRAN '.Offset(1, 0)
                    For i = 1 To dCOLL.Count
                        beginRAN.Offset(1, 0).EntireRow.Insert
                        beginRAN.Offset(1, 0).Formula = "='" & dCOLL(i).Parent.Name & "'!" & dCOLL(i).Address(False, True)
                        Intersect(beginRAN.Offset(1, 0).EntireRow, valRAN2).Formula = "='" & vCOLL(i).Parent.Name & "'!" & vCOLL(i).Address(False, True)
                        Intersect(beginRAN.Offset(1, 0).EntireRow, qtyRAN2).Formula = "='" & qCOLL(i).Parent.Name & "'!" & qCOLL(i).Address(False, True)
                        Intersect(beginRAN.Offset(1, 0).EntireRow, uomRAN2).Formula = "='" & uCOLL(i).Parent.Name & "'!" & uCOLL(i).Address(False, True)
                        Intersect(beginRAN.Offset(1, 0).EntireRow, totalRAN).Formula = "=" & Intersect(beginRAN.Offset(1, 0).EntireRow, valRAN2).Address(False, True) & "*" & Intersect(beginRAN.Offset(1, 0).EntireRow, qtyRAN2).Address(False, True)
                        If i = 1 Then
                            destWS.[\r_heading].EntireRow.Copy
                            beginRAN.Offset(1, 0).EntireRow.PasteSpecial xlPasteFormats
                        ElseIf i = dCOLL.Count Then
                            destWS.[\r_blank].EntireRow.Copy
                            beginRAN.Offset(1, 0).EntireRow.PasteSpecial xlPasteFormats
                        Else
                            If beginRAN.Offset(1, 0).Value = "0" Then
                                destWS.[\r_blank].EntireRow.Copy
                                beginRAN.Offset(1, 0).EntireRow.PasteSpecial xlPasteFormats
                                beginRAN.Offset(1, 0).EntireRow.RowHeight = 4.5
                            Else
                                destWS.[\r_lineitem].EntireRow.Copy
                                beginRAN.Offset(1, 0).EntireRow.PasteSpecial xlPasteFormats
                            End If
                        End If
                        Set beginRAN = beginRAN.Offset(1, 0)
                    Next
                    pfrm.progressUPDATE "", pcount / descCOL.Cells.Count
                    Application.CutCopyMode = False
                'End With
            'ElseIf Intersect(cell.EntireRow, groupCOL).Value = 1 Then
    
            End If
            pcount = pcount + 1
        Next
    
        minusCLICK
        pfrm.turnOFF
        EnS 1
        'basicPROTECT scanWS, True
        basicPROTECT destWS, True
    Else
        Exit Sub
    End If
    
    
    
Exit Sub
e1:
    LogError "Navarro", "restoreDETAIL", Err.Description, Err
    EnS 1, , True
    basicPROTECT scanWS, True
    basicPROTECT destWS, True
    
End Sub

Sub N1_Click()
On Error GoTo e1

    Dim WS As Worksheet
    Dim defaultRAN As Range, deleteRAN As Range, startRAN As Range, endRAN As Range, insertRAN As Range
    Dim ans As Variant

    Set WS = ActiveSheet
    Set deleteRAN = Range(WS.[\r_dstart].EntireRow, WS.[\r_dend].EntireRow)
    trimRANGE deleteRAN, dsupdown
    Set startRAN = WS.[\r_defaultstart].EntireRow
    Set endRAN = WS.[\r_defaultend].EntireRow
    Set defaultRAN = Range(startRAN, endRAN)
    trimRANGE defaultRAN, dsupdown
    Set insertRAN = WS.[\r_dstart]
    
    ans = MsgBox("Are You Sure You Would Like to Revert Back to Defaults?", vbYesNo, "Restore Defaults")
    If ans = vbYes Then
        activeMenu.dockMAIN
        WS.Unprotect
        EnS 0
        defaultRAN.EntireRow.Hidden = False
        defaultRAN.EntireRow.Copy
        insertRAN.Offset(1, 0).EntireRow.Insert
        deleteRAN.EntireRow.Delete
        defaultRAN.EntireRow.Hidden = True
        EnS 1
    End If
    'basicPROTECT WS, True
    
Exit Sub
e1:
    LogError "Navarro", "N1_Click", Err.Description, Err
    'basicPROTECT WS, True

End Sub

Sub N2_Click()
On Error GoTo e1

    Dim WS As Worksheet
    Dim nameRAN As Range
    
    Set WS = ActiveSheet
    Set nameRAN = WS.[\c_name]
    
    If nameRAN.EntireColumn.Hidden = False Then
        nameRAN.EntireColumn.Hidden = True
    ElseIf nameRAN.EntireColumn.Hidden = True Then
        nameRAN.EntireColumn.Hidden = False
    End If
    
Exit Sub
e1:
    LogError "Navarro", "N2_Click", Err.Description, Err

End Sub



'====================================================
'   FUNCTIONS
'====================================================

Function groupRAN(RAN As Range, Optional region_sectorCHK As Boolean) As Range
On Error GoTo e1

    Dim WS As Worksheet
    Dim ran1 As Range, ran2 As Range
    Dim i As Integer
    Dim sector As String
    Dim region As String
    Dim boo As Boolean
    
    Const imax As Integer = 200
    
    Set WS = RAN.Cells(1, 1).Parent
    Set ran1 = Intersect(RAN.EntireRow, WS.[\c_group].EntireColumn)
    sector = masterOBJ.WS.[\sec].Value
    region = masterOBJ.WS.[\reg].Value
    
    
    If ran1.Value = "[" Then
        Set ran2 = ran1.Offset(1, 0)
        Do Until ran2.Value = 1
            If ran2.Value = "]" Then Exit Do
            Set ran2 = ran2.Offset(1, 0)
        Loop
        Set ran1 = ran2
        Do Until ran2.Value = "]" Or i = imax
            i = i + 1
            If region_sectorCHK Then
                boo = (InStr(Intersect(WS.[\c_sector].EntireColumn, ran2.EntireRow).Value, sector) > 0 Or _
                      Intersect(WS.[\c_sector].EntireColumn, ran2.EntireRow).Value = "") And _
                      (InStr(Intersect(WS.[\c_region].EntireColumn, ran2.EntireRow).Value, region) > 0 Or _
                      Intersect(WS.[\c_region].EntireColumn, ran2.EntireRow).Value = "")
            Else
                boo = True
            End If
            
            If boo Then Set ran1 = Union(ran1, ran2)
            Set ran2 = ran2.Offset(1, 0)
        Loop
    End If
    
    If i <> imax Then
        Set groupRAN = ran1.EntireRow
    End If
    
Exit Function
e1:
    LogError "Navarro", "groupRAN", Err.Description, Err
End Function

Function tbl(tblRAN As Range, metricRAN As Range, Optional colNUM As Integer = 1) As Double
On Error GoTo e1

    Dim tran As Range
    Dim metricVAL As Double
    Dim cell As Range
    Dim i As Integer
    Dim imax As Integer
    
    'On Error GoTo quickout
    
    'Set tRAN = masterOBJ.WB.Names(tblNAME).RefersToRange
    metricVAL = metricRAN.Value
    
    Set cell = tblRAN.Cells(1, 1)
    For i = 0 To tblRAN.Rows.Count - 1
        If cell.Offset(i, 0).Value > metricVAL Then
            tbl = cell.Offset(i - 1, colNUM).Value
            Exit For
        End If
        If i = tblRAN.Rows.Count - 1 Then
            tbl = cell.Offset(i, colNUM).Value
        End If
    Next
    
Exit Function
e1:
   'LogError "Navarro", "tbl",err.description, Err
End Function


Function trailertbl(referRAN As Range, trailerRAN As Range) As Double
On Error GoTo e1

    Dim tVAL As Double
    Dim tNAME As String
    Dim tblRAN As Range
    Dim startcell As Range, headers As Range, cell As Range, selectcell As Range, attRAN As Range, attcell As Range
    Dim i As Integer
    
    tNAME = trailerRAN.Value
    Set tblRAN = Worksheets("Code").[\t_trailer]
    Set startcell = tblRAN.Cells(1, 1)
    Set headers = Union(startcell.Offset(0, 1), startcell.Offset(0, 2), startcell.Offset(0, 3))
    Set attRAN = Union(startcell.Offset(1, 0), startcell.Offset(2, 0), startcell.Offset(3, 0), startcell.Offset(4, 0), startcell.Offset(5, 0), startcell.Offset(6, 0), startcell.Offset(7, 0), startcell.Offset(8, 0))

    Set selectcell = headers.Find(What:=tNAME, MatchCase:=True)
    Set attcell = attRAN.Find(referRAN.Value)
    
    trailertbl = Intersect(selectcell.EntireColumn, attcell.EntireRow).Value
    
Exit Function
e1:
    
End Function







