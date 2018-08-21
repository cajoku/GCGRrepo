VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents thisAPP As Application
Attribute thisAPP.VB_VarHelpID = -1
Public WithEvents sdWS As Worksheet
Attribute sdWS.VB_VarHelpID = -1
Public WithEvents grdWS As Worksheet
Attribute grdWS.VB_VarHelpID = -1
Public WithEvents gcoWS As Worksheet
Attribute gcoWS.VB_VarHelpID = -1
Public WithEvents WS As Worksheet
Attribute WS.VB_VarHelpID = -1
Public WithEvents WB As Workbook
Attribute WB.VB_VarHelpID = -1
Public WithEvents gcdWS As Worksheet
Attribute gcdWS.VB_VarHelpID = -1
Public WithEvents groWS As Worksheet
Attribute groWS.VB_VarHelpID = -1
Public WithEvents schedWS As Worksheet
Attribute schedWS.VB_VarHelpID = -1
Public WithEvents gantWS As Worksheet
Attribute gantWS.VB_VarHelpID = -1
Private cStartRan As Range, cEndRan As Range, perRan As Range, durRan As Range, jstartRan As Range, jendRan As Range, pStartRan As Range, pEndRan As Range
Private durCol As Range, perCol As Range, perCOL2 As Range, dendCOL As Range, jobCOL As Range, jobCOL2 As Range, jobDUR As Range, jobDUR2 As Range
Private hROW As Range, eROW As Range
Private preCon As Range, conRan As Range, rateRan As Range, rateRan2 As Range
Private dZone As Range, mZone As Range, infoRan As Range, rateZone As Range
Private lastRAN As Range
Private lastFORM As String
'<navarro 4/11/18
Public cMenuCOLL As Collection
'navarro>

Private Sub Class_Initialize()
On Error GoTo ehandle

    Dim cmenu As contextMENUcls '<navarro 4/11/18>
    Dim carrymenu As carryMENUcls '<navarro 4-18-18>
    
    Set WB = masterWB
    Set WS = WB.Worksheets("Settings")
    Set sdWS = WB.Worksheets("Staff Detail")
    Set gcdWS = WB.Worksheets("GCs Detail")
    Set grdWS = WB.Worksheets("GRs Detail")
    Set gcoWS = WB.Worksheets("GCs Owner")
    Set groWS = WB.Worksheets("GRs Owner")
    Set schedWS = WB.Worksheets("Non-Staff Labor")
    Set gantWS = WB.Worksheets("Labor Gantt")
    
    '<navarro
    '<4-11-18
    Set cMenuCOLL = New Collection
    Set cmenu = New contextMENUcls
    cmenu.register gcoWS
    cMenuCOLL.Add cmenu
    Set cmenu = New contextMENUcls
    cmenu.register groWS
    cMenuCOLL.Add cmenu
    '4-11-18>
    
    '<4/18/18
'    Set carrymenu = New carryMENUcls
'    carrymenu.register gcdWS
'    cMenuCOLL.Add carrymenu
'
'    Set carrymenu = New carryMENUcls
'    carrymenu.register grdWS
'    cMenuCOLL.Add carrymenu
    '4-18-18>
    'navarro>

    schedBOOL = False
    
Exit Sub
    
ehandle:
    LogError "ClassMaster", "Class_Initialize", Err.Description, Err
End Sub

Private Sub Class_Terminate()
    isMasterGantt = False
End Sub



Private Sub WB_BeforeClose(Cancel As Boolean)
On Error GoTo e1

    onCLOSE Cancel
    
    Exit Sub
e1:
    LogError WB.Name, "Workbook_BeforeClose", "Error With Closing WB", Err
End Sub


Private Sub WS_Change(ByVal Target As Range)
On Error GoTo ehandle

    Dim startRAN As Range, endRAN As Range, durRan As Range, preStart As Range, preEnd As Range, preDur As Range, costRan As Range, wDur As Range, pwDur As Range
    Dim regRAN As Range, secRan As Range, sht As Worksheet
    Dim safetySTRING As String
    
    Set startRAN = WS.Range("\cstart")
    Set endRAN = WS.Range("\cend")
    Set durRan = WS.Range("\duration")
    Set preStart = WS.[\pstart]
    Set preEnd = WS.[\pend]
    Set preDur = WS.[\pDur]
    Set costRan = WS.[\pcost]
    Set wDur = WS.[\wdur]
    Set pwDur = WS.[\pwdur]
    Set regRAN = WS.[\reg]
    Set secRan = WS.[\sec]
    
    Application.DisplayAlerts = False
    
    sdWS.Shapes("\\moreratedetail").Select
    
    If Union(Target, startRAN).Address = startRAN.Address Then
        If IsDate(Target.Value) Then
            sdWS.Range("\cstart").Formula = "=" & WS.Name & "!" & startRAN.Address(False, False)
        Else
            If Target.Address = lastRAN.Address And Target.Value <> 0 Then Target.Formula = lastFORM
        End If
    ElseIf Union(Target, endRAN).Address = endRAN.Address Then
        If IsDate(Target.Value) And Target.Value > startRAN.Value Then
            sdWS.Range("\cend").Formula = "=" & WS.Name & "!" & endRAN.Address(False, False)
        Else
            If Target.Address = lastRAN.Address And Target.Value <> 0 Then Target.Formula = lastFORM
        End If
    ElseIf Union(Target, durRan).Address = durRan.Address Then 'recalcs the start & finish based on duration change
        If IsNumeric(Target.Value) Then
            If startRAN.Value <> 0 Then
                ReCalcDur Target.Value, Target, WS, True
                sdWS.[\duration].Formula = "=" & WS.Name & "!" & durRan.Address(False, False)
            Else
                If Target.Address = lastRAN.Address Then Target.Formula = lastFORM
            End If
        Else
            If Target.Address = lastRAN.Address Then Target.Formula = lastFORM
        End If
    ElseIf Union(Target, wDur).Address = wDur.Address Then
        If IsNumeric(Target.Value) Then
            If startRAN.Value <> 0 Then
                ReCalcDur Target.Value, durRan, WS, False
                sdWS.[\wdur].Formula = "=" & WS.Name & "!" & wDur.Address(False, False)
            Else
                If Target.Address = lastRAN.Address Then Target.Formula = lastFORM
            End If
        Else
            If Target.Address = lastRAN.Address Then Target.Formula = lastFORM
        End If
    End If
    
    If Union(Target, preStart).Address = preStart.Address Then
        If IsDate(Target.Value) Then
            sdWS.[\pstart].Formula = "=" & WS.Name & "!" & preStart.Address(False, False)
        Else
            If Target.Address = lastRAN.Address And Target.Value <> 0 Then Target.Formula = lastFORM
        End If
    ElseIf Union(Target, preEnd).Address = preEnd.Address Then
        If IsDate(Target.Value) And Target.Value > preStart.Value Then
            sdWS.[\pend].Formula = "=" & WS.Name & "!" & preEnd.Address(False, False)
        Else
            If Target.Address = lastRAN.Address And Target.Value <> 0 Then Target.Formula = lastFORM
        End If
    ElseIf Union(Target, preDur).Address = preDur.Address Then
        If IsNumeric(Target.Value) Then
            If preEnd.Value <> 0 Then
                ReCalcDur Target.Value, Target, WS, True, True
                sdWS.[\pDur].Formula = "=" & WS.Name & "!" & preDur.Address(False, False)
            Else
                If Target.Address = lastRAN.Address Then Target.Formula = lastFORM
            End If
        Else
            If Target.Address = lastRAN.Address Then Target.Formula = lastFORM
        End If
    ElseIf Union(Target, pwDur).Address = pwDur.Address Then
        If IsNumeric(Target.Value) Then
            If preEnd.Value <> 0 Then
                ReCalcDur Target.Value, preDur, WS, False, True
                sdWS.[\pwdur].Formula = "=" & WS.Name & "!" & pwDur.Address(False, False)
            Else
                If Target.Address = lastRAN.Address Then Target.Formula = lastFORM
            End If
        Else
            If Target.Address = lastRAN.Address Then Target.Formula = lastFORM
        End If
    End If
    
    If Union(Target, costRan).Address = costRan.Address Then
        If IsNumeric(Target.Value) And Target.Value <> "" Then
            safetySTRING = safetySTR(masterWB.Worksheets("Code").[\safetyTABLE], Target, WS.[\reg].Value)
            insertPreset safetySTRING
            WS.Activate
        End If
    End If
    
    If Union(Target, regRAN).Address = regRAN.Address Then
        ToggleSchedule Target.Value
        ToggleWorkRegion Target.Value
    End If
    
    If startRAN.Value <> 0 And endRAN.Value <> 0 Then
        For Each sht In masterWB.Worksheets
            If sht.Visible = xlSheetHidden Then
                sht.Visible = xlSheetVisible
            End If
        Next
    End If
    
    'sdWS.[A1].Select
    Application.DisplayAlerts = True
quickout:
    EnS 1, , True
Exit Sub
    
ehandle:
    LogError "ClassMaster", "WS_Change", Err.Description, Err
    Debug.Print "For Colby: Error on Settings WS Change Event"
    EnS 1, , True
End Sub
Private Sub WS_SelectionChange(ByVal Target As Range)
On Error GoTo ehandle

    Set lastRAN = activecell
    lastFORM = lastRAN.Formula
    
Exit Sub
    
ehandle:
    LogError "ClassMaster", "WS_SelectionChange", Err.Description, Err
End Sub


Private Sub sdWS_Change(ByVal Target As Range)
    
    Dim ans As Variant
    Dim cell As Range
    
On Error Resume Next

    '============ Defining Variables =================='
    Set infoRan = sdWS.Range("\info")
    
    Set cStartRan = sdWS.Range("\cstart")
    Set cEndRan = sdWS.Range("\cend")
    Set pStartRan = sdWS.[\pstart]
    Set pEndRan = sdWS.[\pend]
    
    Set durRan = sdWS.Range("\duration")
    
    Set hROW = sdWS.Range("\r_start").EntireRow
    Set eROW = sdWS.Range("\r_end").EntireRow
    
    Set durCol = sdWS.Range("\c_durSTART").EntireColumn
    Set dendCOL = sdWS.Range("\c_durEND").EntireColumn
    
    Set perCol = boxRANGE(sdWS, "\c_perTIME", "\r_constr", "\r_precon")
    trimRANGE perCol, dsupdown
    trimRANGE perCol, dsbottom
    Set perCOL2 = boxRANGE(sdWS, "\c_perTIME", "\r_constr", "\r_end")
    trimRANGE perCOL2, dsupdown
    trimRANGE perCOL2, dsbottom
    If Not perCol Is Nothing And Not perCOL2 Is Nothing Then
        Set perCol = Union(perCol, perCOL2)
    ElseIf perCol Is Nothing Then
        Set perCol = perCOL2
    End If
    
    Set jobCOL = boxRANGE(sdWS, "\c_jobStart", "\r_constr", "\r_precon")
    trimRANGE jobCOL, dsupdown
    trimRANGE jobCOL, dsbottom
    Set jobCOL2 = boxRANGE(sdWS, "\c_jobStart", "\r_constr", "\r_end")
    trimRANGE jobCOL2, dsupdown
    trimRANGE jobCOL2, dsbottom
    If Not jobCOL Is Nothing And Not jobCOL2 Is Nothing Then
        Set jobCOL = Union(jobCOL, jobCOL2)
    ElseIf jobCOL Is Nothing Then
        Set jobCOL = jobCOL2
    End If
    
    Set jobDUR = boxRANGE(sdWS, "\c_jobDur", "\r_constr", "\r_precon")
    trimRANGE jobDUR, dsupdown
    trimRANGE jobDUR, dsbottom
    Set jobDUR2 = boxRANGE(sdWS, "\c_jobDur", "\r_constr", "\r_end")
    trimRANGE jobDUR2, dsupdown
    trimRANGE jobDUR2, dsbottom
    If Not jobDUR Is Nothing And Not jobDUR2 Is Nothing Then
        Set jobDUR = Union(jobDUR, jobDUR2)
    ElseIf jobDUR Is Nothing Then
        Set jobDUR = jobDUR2
    End If
    
    Set rateRan = boxRANGE(sdWS, "\r_precon", "\r_constr", "\c_order", "\c_rateEnd")
    trimRANGE rateRan, dsupdown
    trimRANGE rateRan, dsbottom
    Set rateRan2 = boxRANGE(sdWS, "\r_constr", "\r_end", "\c_order", "\c_rateEnd")
    trimRANGE rateRan2, dsupdown
    trimRANGE rateRan2, dsbottom
    If Not rateRan Is Nothing And Not rateRan2 Is Nothing Then
        Set rateZone = Union(rateRan, rateRan2)
    ElseIf rateRan Is Nothing Then
        Set rateZone = rateRan2
    ElseIf rateRan2 Is Nothing Then
        Set rateZone = rateRan
    End If
    
    Set dZone = sdWS.Range(Range("\c_durSTART").EntireColumn, Range("\c_durEND").EntireColumn)
    Set mZone = sdWS.Range(Cells(Range("\r_start").row, Range("\c_negStart").Offset(0, 1).Column), Cells(Range("\r_end").row, Range("\c_durEND").Offset(0, -1).Column))
    
    
    Set jstartRan = boxRANGE(sdWS, "\r_precon", "\r_end", "\c_posStart")
    trimRANGE jstartRan, dsupdown
    
    Set jendRan = boxRANGE(sdWS, "\r_precon", "\r_end", "\c_posEnd")
    trimRANGE jendRan, dsupdown
      
    
On Error GoTo ehandle
    '============ Settings header change events =================='
    If Union(Target, infoRan).Address = infoRan.Address Then
        EnS 0
        If InStr(Target.Name.Name, "\cstart") Or InStr(Target.Name.Name, "\cend") Then
            If IsDate(Target.Value) Then
                WS.Range(Mid(Target.Name.Name, 16)) = Target.Value
                If cStartRan.Value <> 0 Then Application.Calculate
                Target.Formula = "=" & "Settings!" & WS.Range(Mid(Target.Name.Name, 16)).Address(False, False)
                'to catch gap btwn constart & preconend
                If Target.Address = cStartRan.Address Then ' And pEndRan.Value = 0 Then
                    WS.Range(Mid(pEndRan.Name.Name, 16)) = Target.Value
                    DurColumnCount sdWS.Range("\duration").Value, sdWS
                    Application.Calculate
                    NegColCount sdWS.[\pDur].Value, sdWS
                    If durRan.Value <> 0 Then newGant.NumberColumns durRan.Offset(1, 0).Value: NumberSchedRows durRan.Offset(1, 0).Value
                ElseIf Target.Address = cEndRan.Address And cEndRan.Value > cStartRan.Value Then
                    DurColumnCount sdWS.Range("\duration").Value, sdWS
                    Application.Calculate
                    If durRan.Value <> 0 Then newGant.NumberColumns durRan.Offset(1, 0).Value: NumberSchedRows durRan.Offset(1, 0).Value
                Else
                    If Target.Address = lastRAN.Address Then Target.Formula = lastFORM
                End If
            Else
                If Target.Address = lastRAN.Address Then
                    If Target.Value = 0 Then
                        WS.Range(Mid(Target.Name.Name, 16)) = Target.Value ''triggers an auto duration calc with high month & week totals. Fortunately the calc does not trigger events
                        Target.Formula = "=" & "Settings!" & WS.Range(Mid(Target.Name.Name, 16)).Address(False, False)
                    End If
                    If Target.Value <> 0 Then Target.Formula = lastFORM
                End If
            End If
        ElseIf InStr(Target.Name.Name, "\duration") Then
            If IsNumeric(Target.Value) Then
                If cStartRan.Value <> 0 Then
                    ReCalcDur Target.Value, WS.Range(Mid(Target.Name.Name, 16)), WS, True
                    Target.Formula = "=" & "Settings!" & WS.Range(Mid(Target.Name.Name, 16)).Address(False, False)
                    DurColumnCount Target.Value, sdWS
                    Application.Calculate: NumberSchedRows Target.Offset(1, 0).Value: newGant.NumberColumns Target.Offset(1, 0).Value
                End If
            Else
                If Target.Address = lastRAN.Address Then Target.Formula = lastFORM
            End If
        ElseIf InStr(Target.Name.Name, "\wdur") Then
            If IsNumeric(Target.Value) Then
                If cStartRan.Value <> 0 Then
                    Application.Calculate
                    ReCalcDur Target.Value, WS.Range(Mid(Target.Name.Name, 16)).Offset(-1, 0), WS, False
                    Application.Calculate
                    Target.Formula = "=" & "Settings!" & WS.Range(Mid(Target.Name.Name, 16)).Address(False, False)
                    DurColumnCount sdWS.Range("\duration").Value, sdWS
                    Application.Calculate: NumberSchedRows Target.Value: newGant.NumberColumns Target.Value
                End If
            Else
                If Target.Address = lastRAN.Address Then Target.Formula = lastFORM
            End If
        ElseIf InStr(Target.Name.Name, "\pstart") Or InStr(Target.Name.Name, "\pend") Then
            If IsDate(Target.Value) Then
                WS.Range(Mid(Target.Name.Name, 16)) = Target.Value
                If pStartRan.Value <> 0 Then Application.Calculate
                Target.Formula = "=" & "Settings!" & WS.Range(Mid(Target.Name.Name, 16)).Address(False, False)
                ' added to prevent gaps between preCon End and cStart
                If Target.Address = pEndRan.Address And pEndRan.Value > pStartRan.Value Then 'And cStartRan.Value = 0 Then
                    WS.Range(Mid(cStartRan.Name.Name, 16)) = Target.Value
                    NegColCount sdWS.[\pDur].Value, sdWS
                    Application.Calculate
                    DurColumnCount sdWS.Range("\duration").Value, sdWS
                Else
                    NegColCount sdWS.[\pDur], sdWS  'added for negcol precon feature 3/7
                End If
                
            Else
                If Target.Address = lastRAN.Address Then
                    If Target.Value = 0 Then
                        WS.Range(Mid(Target.Name.Name, 16)) = Target.Value ''triggers an auto duration calc with high month & week totals. Fortunately the calc does not trigger events
                        Target.Formula = "=" & "Settings!" & WS.Range(Mid(Target.Name.Name, 16)).Address(False, False)
                    End If
                    If Target.Value <> 0 Then Target.Formula = lastFORM
                End If
            End If
        ElseIf InStr(Target.Name.Name, "\pdur") Then
            If IsNumeric(Target.Value) Then
                If pEndRan.Value <> 0 Then
                    Application.Calculate
                    ReCalcDur Target.Value, WS.Range(Mid(Target.Name.Name, 16)), WS, True, True
                    Target.Formula = "=" & "Settings!" & WS.Range(Mid(Target.Name.Name, 16)).Address(False, False)
                    NegColCount sdWS.[\pDur], sdWS
                End If
            Else
                If Target.Address = lastRAN.Address Then Target.Formula = lastFORM
            End If
        ElseIf InStr(Target.Name.Name, "\pwdur") Then
            If IsNumeric(Target.Value) Then
                If pEndRan.Value <> 0 Then
                    Application.Calculate
                    ReCalcDur Target.Value, WS.Range(Mid(Target.Name.Name, 16)).Offset(-1, 0), WS, False, True
                    Application.Calculate
                    Target.Formula = "=" & "Settings!" & WS.Range(Mid(Target.Name.Name, 16)).Address(False, False)
                    NegColCount sdWS.[\pDur], sdWS
                End If
            Else
                If Target.Address = lastRAN.Address Then Target.Formula = lastFORM
            End If
        ElseIf InStr(Target.Name.Name, "\") Then
            If IsNumeric(Target.Value) Then
                WS.Range(Mid(Target.Name.Name, 16)) = Target.Value
                Target.Formula = "=" & "Settings!" & WS.Range(Mid(Target.Name.Name, 16)).Address(False, False)
            Else
                If Target.Address = lastRAN.Address Then Target.Formula = lastFORM
            End If
        End If
        'below is added to prevent default zero value dates to show up in blank date cells
        If cStartRan.Value = 0 Or cEndRan.Value = 0 Or pStartRan.Value = 0 Or pEndRan.Value = 0 Then
            HideZeros
        Else
            ActiveWindow.DisplayZeros = True
        End If
        Application.CalculateFull
    End If
    
    For Each cell In Target.Cells
    
        '============ % Time column change event =================='
        If Not perCol Is Nothing Then
            If Union(cell, perCol).Address = perCol.Address And cell.EntireRow.Cells(1, 1).Value <> "dp" Then
                EnS 0
                If IsNumeric(cell.Value) Then
                    UpdateFormulas cell, sdWS, False
                Else
                    If cell.Address = lastRAN.Address Then cell.Formula = lastFORM
                End If
                Application.CalculateFull
            End If
        End If
        '============ Monthly breakdown zone change event =================='
        If Not mZone Is Nothing Then
            If Union(cell, mZone).Address = mZone.Address And cell.EntireRow.Cells(1, 1).Value <> "dp" Then
                EnS 0
                If IsNumeric(cell.Value) Then
                    If Intersect(cell.EntireRow, jstartRan.EntireColumn).Value <> "" And Intersect(cell.EntireRow, jendRan.EntireColumn).Value <> "" Then
                        monthtester cell, sdWS
                        UpdateFormulas cell, sdWS, True
                    Else
                        If cell.Address = lastRAN.Address And cell.Value <> 0 Then cell.Formula = lastFORM
                    End If
                Else
                    If cell.Address = lastRAN.Address Then cell.Formula = lastFORM
                End If
                Application.CalculateFull
            End If
        End If
        '============ Job Title start month cell change event =================='
        If Not jobCOL Is Nothing Then
            If Union(cell, jobCOL).Address = jobCOL.Address And cell.EntireRow.Cells(1, 1).Value <> "dp" Then
                EnS 0
                If IsNumeric(cell.Value) And (cell.Value > 0 Or cell.Value < 0) Then
                    If cStartRan.Value <> 0 And pEndRan.Value <> 0 Then  ' pStartRan.Value <> 0 And
                        CreateStart cell, sdWS, True
                        Application.Calculate
                        If cEndRan.Value >= sdWS.[\negMax].Value Then DurColumnCount durRan.Value, sdWS
                        If pStartRan.Value <= sdWS.[\negMin].Value Then NegColCount sdWS.[\pDur].Value, sdWS
                    End If
                Else
                    If cell.Address = lastRAN.Address And cell.Value <> 0 Then cell.Formula = lastFORM
                End If
                Application.CalculateFull
            End If
        End If
        '============ Job Title duration cell change event =================='
        If Not jobDUR Is Nothing Then
            If Union(cell, jobDUR).Address = jobDUR.Address And cell.EntireRow.Cells(1, 1).Value <> "dp" Then
                EnS 0
                If IsNumeric(cell.Value) And Intersect(cell.EntireRow, jobCOL.EntireColumn).Value = "" And Intersect(cell.EntireRow, jstartRan.EntireColumn).Value = "" Then GoTo quickout
                If IsNumeric(cell.Value) And cell.Value > 0 Then
                    If cStartRan.Value <> 0 And pEndRan.Value <> 0 Then ' pStartRan.Value <> 0 And
                        CreateFinish cell, sdWS, True
                        Application.Calculate
                        If cEndRan.Value >= sdWS.[\negMax].Value Then DurColumnCount durRan.Value, sdWS
                        If pStartRan.Value <= sdWS.[\negMin].Value Then NegColCount sdWS.[\pDur].Value, sdWS
                    End If
                Else
                    If cell.Address = lastRAN.Address And cell.Value <> 0 Then cell.Formula = lastFORM
                End If
                Application.CalculateFull
            End If
        End If
        '============ Job Title start date cell and finish date cell change events =================='
        
        If Union(cell, jstartRan).Address = jstartRan.Address And cell.EntireRow.Cells(1, 1).Value <> "dp" Then
            EnS 0
            If IsDate(cell.Value) Then
                If Intersect(cell.EntireRow, jobDUR.EntireColumn).Value = "" Or Intersect(cell.EntireRow, jendRan.EntireColumn).Value = "" Then GoTo quickout
                ReCalcJobDur cell, True
                If cell.Value < pStartRan.Value Or pStartRan.Value = 0 Then
'                    ans = MsgBox("Would you like to change the preCon start of this project?", vbYesNo, "Start Date is Earlier Than PreCon Start")
'                        If ans = vbYes Then
'                            EnS 1
'                            pStartRan.Value = cell.Value
'                            EnS 0
                            NegColCount DateDiff("m", cell.Value, pEndRan.Value), sdWS
                            UpdateFormulas Intersect(cell.EntireRow, perCol.EntireColumn), sdWS, True
'                        ElseIf ans = vbNo Then
'                            If cell.Address = lastRAN.Address Then cell.Formula = lastFORM
'                            GoTo quickout
'                        End If
                Else
                    Application.Calculate
                    If cEndRan.Value >= sdWS.[\negMax].Value Then DurColumnCount durRan.Value, sdWS
                    If pStartRan.Value <= sdWS.[\negMin].Value Then NegColCount sdWS.[\pDur].Value, sdWS
                    UpdateFormulas Intersect(cell.EntireRow, perCol.EntireColumn), sdWS, True
                End If
            Else
                If cell.Address = lastRAN.Address And cell.Value <> 0 Then cell.Formula = lastFORM
            End If
            Application.CalculateFull
        ElseIf Union(cell, jstartRan).Address = jstartRan.Address And cell.EntireRow.Cells(1, 1).Value = "dp" Then
            registerPHASE cell
        ElseIf Union(cell, jendRan).Address = jendRan.Address And cell.EntireRow.Cells(1, 1).Value <> "dp" Then
            EnS 0
            If IsDate(cell.Value) Then
                If Intersect(cell.EntireRow, jstartRan.EntireColumn).Value = "" Or cStartRan.Value = 0 Then GoTo quickout
                ReCalcJobDur cell, False  'may need to put this inside of each if statement below
                If cell.Value > cEndRan.Value Then
'                    ans = MsgBox("Would you like to change the end date of this entire project?", vbYesNo, "Finish Date Exceeds Project End Date")
'                        If ans = vbYes Then
'                            EnS 1
'                            cEndRan.Value = cell.Value
'                            EnS 0
                            DurColumnCount DateDiff("m", cStartRan.Value, cell.Value), sdWS
                            UpdateFormulas Intersect(cell.EntireRow, perCol.EntireColumn), sdWS, True
                            
'                        ElseIf ans = vbNo Then
'                            If cell.Address = lastRAN.Address Then cell.Formula = lastFORM
'                            GoTo quickout
'                        End If
                Else
                    Application.Calculate
                    If cEndRan.Value >= sdWS.[\negMax].Value Then DurColumnCount durRan.Value, sdWS
                    If pStartRan.Value <= sdWS.[\negMin].Value Then NegColCount sdWS.[\pDur].Value, sdWS
                    UpdateFormulas Intersect(cell.EntireRow, perCol.EntireColumn), sdWS, True
                End If
            Else
                If cell.Address = lastRAN.Address And cell.Value <> 0 Then cell.Formula = lastFORM
            End If
            Application.CalculateFull
        ElseIf Union(cell, jendRan).Address = jendRan.Address And cell.EntireRow.Cells(1, 1).Value = "dp" Then
            registerPHASE cell
        End If
        
        
        '============ log changes to the rateDetail section ============='
        If Not rateZone Is Nothing Then
            If Union(cell, rateZone).Address = rateZone.Address And cell.EntireRow.Cells(1, 1).Value <> "dp" Then
                If lastFORM <> cell.Formula Then
                    sdWS.Unprotect
                    With cell 'start formatting comment
                        If Not .Comment Is Nothing Then .Comment.Delete
                        .ClearComments
                        .AddComment lastFORM
                        .Comment.Shape.Fill.ForeColor.RGB = RGB(0, 0, 102)
                        .Comment.Shape.TextFrame.Characters.Font.ColorIndex = 2
                        .Comment.Shape.TextFrame.Characters.Font.Bold = True
                        .Font.Bold = True
                        .Font.Italic = True
                        .Font.Color = RGB(250, 1, 1)
                    End With
                    basicPROTECT sdWS, True
                End If
                Application.CalculateFull
            End If
        End If
        
    Next
    
quickout:
    EnS 1, , True
Exit Sub
    
ehandle:
    LogError "ClassMaster", "sdWS_Change", Err.Description, Err
    Debug.Print "For Colby: Error on Staff Detail Change Event"
    EnS 1, , True
    basicPROTECT sdWS, True
End Sub

Private Sub sdWS_SelectionChange(ByVal Target As Range)
On Error Resume Next

    Dim preTitle As Range, conTitle As Range, phaseTitle As Range, rateZoneRan As Range
    
    
    Set buttonCls = Nothing
    Set gantCls = Nothing
    Set lastRAN = activecell
    lastFORM = lastRAN.Formula
    
    Set revertButtoncls = Nothing
    Set rateRan = boxRANGE(sdWS, "\r_precon", "\r_constr", "\c_order", "\c_rateEnd")
    trimRANGE rateRan, dsupdown
    Set rateRan2 = boxRANGE(sdWS, "\r_constr", "\r_end", "\c_order", "\c_rateEnd")
    trimRANGE rateRan2, dsupdown
    If Not rateRan Is Nothing And Not rateRan2 Is Nothing Then
        Set rateZoneRan = Union(rateRan, rateRan2)
    ElseIf rateRan Is Nothing Then
        Set rateZoneRan = rateRan2
    ElseIf rateRan2 Is Nothing Then
        Set rateZoneRan = rateRan
    End If
        
    Set preTitle = boxRANGE(sdWS, "\r_precon", "\r_constr", "\c_Position")
    trimRANGE preTitle, dsupdown
    trimRANGE preTitle, dsbottom
    
    Set conTitle = boxRANGE(sdWS, "\r_constr", "\r_end", "\c_Position")
    trimRANGE conTitle, dsupdown
    trimRANGE conTitle, dsbottom
    
    
On Error GoTo ehandle
    
    If Target.Cells.Count = 1 Then
        If Not preTitle Is Nothing Then
            If Union(Target, preTitle).Address = preTitle.Address Then
                Set buttonCls = New AddDeleteButton
                buttonCls.setANCHOR Target
            End If
        End If
        
        If Not conTitle Is Nothing Then
            If Union(Target, conTitle).Address = conTitle.Address Then
                If Target.Offset(0, -1).Value <> "dp" Then
                'If Target.Offset(0, -1).Value <> "dp" Then
                    Set buttonCls = New AddDeleteButton
                    buttonCls.setANCHOR Target
                'End If
                End If
            End If
        End If
        
        If Not conTitle Is Nothing Then
            If Union(Target, conTitle.EntireRow).Address = conTitle.EntireRow.Address Then
                If Target.EntireRow.Cells(1, 1).Value = "dp" Then
                    Set buttonCls = New AddDeleteButton
                    buttonCls.setPHASE Target.EntireRow.Cells(1, 1)
                End If
            End If
        End If
        
        If Not rateZoneRan Is Nothing Then
            If Union(Target, rateZoneRan).Address = rateZoneRan.Address And Target.Font.Color = RGB(250, 1, 1) And Target.EntireRow.Cells(1, 1).Value <> "dp" Then
                Set revertButtoncls = New revertButton
                revertButtoncls.setANCHOR Target
            End If
        End If
    End If
    
quickout:
Exit Sub
    
ehandle:
    LogError "ClassMaster", "sdWS_SelectionChange", Err.Description, Err
    Debug.Print "For Colby: Error on sdws selection change"
    
End Sub

Private Sub schedWS_SelectionChange(ByVal Target As Range)
On Error GoTo e1

    Dim mergeRAN As Range, catRAN As Range
    
    Set buttonCls = Nothing

    Set catRAN = boxRANGE(schedWS, "\r_header", "\c_editStart", "\c_schedEnd")
    trimRANGE catRAN, dsRIGHT
    Set catRAN = catRAN.Offset(-1, 0)
    
    If Not catRAN Is Nothing Then
        If Union(Target, catRAN).Address = catRAN.Address Then
            Set buttonCls = New AddDeleteButton
            buttonCls.setSchedAnchor Target, True
        End If
    End If

Exit Sub
e1:
    LogError "ClassMaster", "schedWS_SelectionChange", Err.Description, Err
    
End Sub
Private Sub gantWS_Change(ByVal Target As Range)
On Error GoTo e1

    Dim phaseRAN As Range, phaseStart As Range, phaseDur As Range, tempPhase As Range, descCOL As Range, cell As Range, tempRAN As Range
    Dim predRAN As Range, lagRAN As Range
    Dim phaseName As String, defaultName As String, formSTR As String, lagstr As String, predStr As String, lagstr2 As Integer, startInt As Integer, formstr2 As String
    Dim predCOLL As Collection

    Set phaseRAN = gantWS.[\addPhase]
    Set tempPhase = gantWS.[\r_phasephase].EntireRow
    Set descCOL = gantWS.[\c_desc].EntireColumn
    Set phaseStart = boxRANGE(gantWS, "\r_wheader", "\c_gstart", "\r_addphase")
    trimRANGE phaseStart, dsupdown
    Set phaseDur = boxRANGE(gantWS, "\r_wheader", "\c_gdur", "\r_addphase")
    trimRANGE phaseDur, dsupdown
    Set predRAN = gantWS.[\pred]
    trimRANGE predRAN, dsupdown
    Set lagRAN = gantWS.[\lag]
    trimRANGE lagRAN, dsupdown
           
    If Not isMasterGantt Then Exit Sub
    
    
    If Target.Cells.Count = 1 Then
        If Union(Target, phaseRAN).Address = phaseRAN.Address Then
            gantWS.Unprotect
            phaseName = Target.Value
            defaultName = "Add Phase Here:"
            EnS 0
            If Target.Value <> "" Then
                Target.EntireRow.Copy
                Target.EntireRow.Insert
                Target.Value = defaultName
                tempPhase.Hidden = False
                tempPhase.Copy
                Target.Offset(-1, 0).EntireRow.PasteSpecial xlPasteFormats
                Target.Offset(-1, 0).Value = phaseName
                Target.Offset(-1, -1).Value = tempPhase.Cells(1, 1).Value
                Target.Offset(-1, 2).Select
                tempPhase.Hidden = True
                newGant.createBar Target.Offset(-1, 0)
            End If
            basicPROTECT gantWS, True
            EnS 1
        End If
    End If
        
    For Each cell In Target.Cells
        If Not phaseStart Is Nothing Then
            If Union(cell, phaseStart).Address = phaseStart.Address Then
                EnS 0
                    CreateLaborSection Intersect(cell.EntireRow, descCOL).Value, Intersect(cell.EntireRow, phaseStart.EntireColumn).Value
                EnS 1
                If Not predRAN Is Nothing Then
                    Set tempRAN = predRAN.Find(What:=Intersect(cell.EntireRow, descCOL).Value)
                    If Not tempRAN Is Nothing Then
                        Application.ScreenUpdating = False
                        predStr = tempRAN.Value
                        formSTR = Intersect(tempRAN.EntireRow, phaseStart).Formula
                        lagstr = Intersect(tempRAN.EntireRow, lagRAN).Value
                        startInt = Intersect(tempRAN.EntireRow, phaseStart).Value
                        Intersect(tempRAN.EntireRow, phaseStart).Value = startInt
                        EnS 0
                        tempRAN.Value = predStr
                        Intersect(tempRAN.EntireRow, phaseStart).Formula = formSTR
                        Intersect(tempRAN.EntireRow, lagRAN).Value = lagstr
                        EnS 1
                    End If
                End If
            End If
        End If
        
        If Not predRAN Is Nothing Then
            If Union(cell, predRAN).Address = predRAN.Address Then
                Application.EnableEvents = False
                If cell.Value = 0 Then
                    returnState cell
                    Intersect(cell.EntireRow, lagRAN).Value = ""
                ElseIf cell.Value = Intersect(cell.EntireRow, gantWS.[\gid]).Value Then
                    returnState cell
                    Intersect(cell.EntireRow, gantWS.[\pred].EntireColumn).Value = ""
                    Intersect(cell.EntireRow, gantWS.[\lag].EntireColumn).Value = ""
                Else
                    formSTR = Intersect(cell.EntireRow, gantWS.[\pred].EntireColumn).Address(False, True)
                    predStr = Intersect(cell.EntireRow, gantWS.[\pred].EntireColumn).Value
                    lagstr2 = Intersect(cell.EntireRow, gantWS.[\lag].EntireColumn).Value
                    lagstr = Intersect(cell.EntireRow, gantWS.[\lag].EntireColumn).Address(False, True)
                    Application.EnableEvents = True
                    Intersect(cell.EntireRow, gantWS.[\c_gstart].EntireColumn).Value = gantWS.[\gid].Find(cell.Value).Offset(0, 2).Value + gantWS.[\gid].Find(cell.Value).Offset(0, 3).Value
                    Application.EnableEvents = False
                    Intersect(cell.EntireRow, gantWS.[\pred].EntireColumn).Value = predStr
                    formSTR = "=INDEX(" & phaseStart.Address & ",MATCH(" & formSTR & ",\gid,0))+INDEX(" & phaseDur.Address & ",MATCH(" & formSTR & ",\gid,0))+" & lagstr
                    Intersect(cell.EntireRow, gantWS.[\c_gstart].EntireColumn).Formula = formSTR
                End If
                Application.EnableEvents = True
            ElseIf Union(cell, lagRAN).Address = lagRAN.Address Then
                If Intersect(cell.EntireRow, gantWS.[\pred].EntireColumn).Value <> "" Then
                    formSTR = Intersect(cell.EntireRow, gantWS.[\pred].EntireColumn).Address(False, True)
                    predStr = Intersect(cell.EntireRow, gantWS.[\pred].EntireColumn).Value
                    lagstr2 = Intersect(cell.EntireRow, gantWS.[\lag].EntireColumn).Value
                    lagstr = Intersect(cell.EntireRow, gantWS.[\lag].EntireColumn).Address(False, True)
                    Application.EnableEvents = True
                    Intersect(cell.EntireRow, gantWS.[\c_gstart].EntireColumn).Value = gantWS.[\gid].Find(cell.Offset(0, -1).Value).Offset(0, 2).Value + gantWS.[\gid].Find(cell.Offset(0, -1).Value).Offset(0, 3).Value + cell.Value
                    Application.EnableEvents = False
                    Intersect(cell.EntireRow, gantWS.[\pred].EntireColumn).Value = predStr
                    Intersect(cell.EntireRow, gantWS.[\lag].EntireColumn).Value = lagstr2
                    formSTR = "=INDEX(" & phaseStart.Address & ",MATCH(" & formSTR & ",\gid,0))+INDEX(" & phaseDur.Address & ",MATCH(" & formSTR & ",\gid,0))+" & lagstr
                    Intersect(cell.EntireRow, gantWS.[\c_gstart].EntireColumn).Formula = formSTR
                    Application.EnableEvents = True
                End If
            ElseIf Union(cell, phaseDur).Address = phaseDur.Address Then
                Set tempRAN = predRAN.Find(What:=Intersect(cell.EntireRow, descCOL).Value)
                If Not tempRAN Is Nothing Then
                    formSTR = Intersect(tempRAN.EntireRow, phaseStart).Formula
                    predStr = Intersect(tempRAN.EntireRow, gantWS.[\pred].EntireColumn).Value
                    lagstr2 = Intersect(tempRAN.EntireRow, gantWS.[\lag].EntireColumn).Value
                    predStr = tempRAN.Value
                    Application.EnableEvents = True
                    Intersect(tempRAN.EntireRow, gantWS.[\c_gstart].EntireColumn).Value = Intersect(cell.EntireRow, phaseStart).Value + Intersect(cell.EntireRow, phaseDur).Value + lagstr2
                    Application.EnableEvents = False
                    Intersect(tempRAN.EntireRow, gantWS.[\pred].EntireColumn).Value = predStr
                    Intersect(tempRAN.EntireRow, gantWS.[\lag].EntireColumn).Value = lagstr2
                    Intersect(tempRAN.EntireRow, gantWS.[\c_gstart].EntireColumn).Formula = formSTR
                    Application.EnableEvents = True
                End If
            ElseIf Union(cell, phaseStart).Address = phaseStart.Address Then
                Application.EnableEvents = False
                Intersect(cell.EntireRow, gantWS.[\pred].EntireColumn).Value = ""
                Intersect(cell.EntireRow, gantWS.[\lag].EntireColumn).Value = ""
                Application.EnableEvents = True
            End If
        End If
        
    Next
    
    EnS 1
        
    Exit Sub
e1:
   LogError "ClassMaster", "gantWS_Change", Err.Description, Err
   EnS 1, , True
End Sub

Private Sub gantWS_SelectionChange(ByVal Target As Range)
On Error GoTo e1

    Dim descRAN As Range, cell As Range
    
    Set descRAN = boxRANGE(gantWS, "\r_wheader", "\c_desc", "\r_addphase")
    trimRANGE descRAN, dsupdown
    
    Set buttonCls = Nothing

    If Target.Cells.Count = 1 Then
        For Each cell In Target.Cells
        
            If Not descRAN Is Nothing Then
                If Union(cell, descRAN).Address = descRAN.Address Then
                    Set buttonCls = New AddDeleteButton
                    buttonCls.setGanttAnchor cell
                End If
            End If
            
            If Union(cell, gantWS.[\gid].EntireRow).Address = gantWS.[\gid].EntireRow.Address Then
                recordState cell
            Else
                nullState
            End If
        
        Next
    End If
    
    Exit Sub
e1:
    LogError "ClassMaster", "gantWS_SelectionChange", Err.Description, Err

End Sub

Private Sub gcdWS_Change(ByVal Target As Range)
On Error GoTo ehandle
    
    
    Set infoRan = gcdWS.[\info]

    
    If Union(Target, infoRan).Address = infoRan.Address Then
        EnS 0
        If InStr(Target.Name.Name, "\proj") Or InStr(Target.Name.Name, "\loc") Or InStr(Target.Name.Name, "\reg") Or InStr(Target.Name.Name, "\sec") Or InStr(Target.Name.Name, "\unit") Or InStr(Target.Name.Name, "\area") Then
            If Target.Value <> 0 Then
                WS.Range(Mid(Target.Name.Name, 14)) = Target.Value
                Target.Formula = lastFORM '"=" & "Settings!" & WS.Range(Mid(Target.Name.Name, 14)).Address(False, False)
            Else
                If Target.Address = lastRAN.Address Then Target.Formula = lastFORM
            End If
        ElseIf InStr(Target.Name.Name, "\date") Then
            If IsDate(Target.Value) And Target.Value <> 0 Then
                WS.Range(Mid(Target.Name.Name, 14)) = Target.Value
                Target.Formula = "=" & "Settings!" & WS.Range(Mid(Target.Name.Name, 14)).Address(False, False)
            Else
                If Target.Address = lastRAN.Address Then Target.Formula = lastFORM
            End If
        End If
    End If
     
    
    EnS 1
    
Exit Sub
    
ehandle:
    LogError "ClassMaster", "gcdWS_Change", Err.Description, Err
    EnS 1, , True
    
End Sub

Private Sub gcdWS_SelectionChange(ByVal Target As Range)
'On Error GoTo ehandle

    Dim arr() As String
    Dim trailerRAN As Range
    Dim RAN As Range
    '<navarro
    Dim qtCOL As Range
    Dim groupCOL As Range
    Dim regionCOL As Range
    Dim sectorCOL As Range
    Dim descCOL As Range
    Dim frm As arrayFORM
    
    Set RAN = Target.EntireRow
    Set trailerRAN = boxRANGE(gcdWS, "\c_desc", "\r_trailer", "\r_trailerSTART")
    trimRANGE trailerRAN, dsupdown
    
    With gcdWS
        Set qtCOL = .[\c_qt].EntireColumn
        Set groupCOL = .[\c_group].EntireColumn
        Set regionCOL = .[\c_region].EntireColumn
        Set sectorCOL = .[\c_sector].EntireColumn
        Set descCOL = .[\c_desc].EntireColumn
    End With
    'navarro>
    
    Set buttonCls = Nothing
    Set menuSHP = Nothing
    
    Set lastRAN = activecell
    lastFORM = lastRAN.Formula
    
    '<navarro
    If Union(activecell, qtCOL).Address = qtCOL.Address And Intersect(activecell.EntireRow, groupCOL).Value = 1 Then
        Set menuSHP = New menuCLS
        arr = menuSHP.createARRAY
        menuSHP.setANCHOR activecell, arr
    ElseIf Union(activecell, descCOL).Address = descCOL.Address And Intersect(activecell.EntireRow, groupCOL).Value = 1 Then 'IsNumeric(Intersect(ActiveCell.EntireRow, groupCOL).Value) Then
        If Not trailerRAN Is Nothing Then
            If Union(activecell, trailerRAN).Address <> trailerRAN.Address Then
                Set buttonCls = New AddDeleteButton
                buttonCls.setDETAILANCHOR activecell.Offset(0, 1)
            End If
        Else
            Set buttonCls = New AddDeleteButton
            buttonCls.setDETAILANCHOR activecell.Offset(0, 1)
        End If
    ElseIf Union(activecell, regionCOL).Address = regionCOL.Address Then
        Set frm = New arrayFORM
        frm.register masterOBJ.WS.[regions], activecell
        frm.Show
    ElseIf Union(activecell, sectorCOL).Address = sectorCOL.Address Then
        Set frm = New arrayFORM
        frm.register masterOBJ.WS.[sectors], activecell
        frm.Show
    End If
    'navarro>
    If Not trailerRAN Is Nothing Then
        If Union(Target, trailerRAN.EntireRow).Address = trailerRAN.EntireRow.Address Then
            Set buttonCls = New AddDeleteButton
            buttonCls.trailerANCHOR Target
        End If
    End If
    
    
    
Exit Sub
ehandle:
    LogError "CLassMaster", "gcdWS_SelectionChange", Err.Description, Err
End Sub


Private Sub grdWS_Change(ByVal Target As Range)
On Error GoTo ehandle

    Dim regRAN As Range, secRan As Range, tempRAN As Range, tran As Range
    Dim regCol As Range, secCol As Range
    Dim cell As Range
    Dim i As Integer, j As Integer
        
    Set infoRan = grdWS.Range("\info")

    
    If Union(Target, infoRan).Address = infoRan.Address Then
        EnS 0
        If InStr(Target.Name.Name, "\proj") Or InStr(Target.Name.Name, "\loc") Or InStr(Target.Name.Name, "\reg") Or InStr(Target.Name.Name, "\sec") Or InStr(Target.Name.Name, "\unit") Or InStr(Target.Name.Name, "\area") Then
            If Target.Value <> 0 Then
                WS.Range(Mid(Target.Name.Name, 14)) = Target.Value
                Target.Formula = lastFORM '"=" & "Settings!" & WS.Range(Mid(Target.Name.Name, 14)).Address(False, False)
            Else
                If Target.Address = lastRAN.Address Then Target.Formula = lastFORM
            End If
        ElseIf InStr(Target.Name.Name, "\date") Then
            If IsDate(Target.Value) And Target.Value <> 0 Then
                WS.Range(Mid(Target.Name.Name, 14)) = Target.Value
                Target.Formula = "=" & "Settings!" & WS.Range(Mid(Target.Name.Name, 14)).Address(False, False)
            Else
                If Target.Address = lastRAN.Address Then Target.Formula = lastFORM
            End If
        End If
    End If
    
    EnS 1
    
Exit Sub

ehandle:
    LogError "ClassMaster", "grdWS_Change", Err.Description, Err
    EnS 1, , True
End Sub


Private Sub grdWS_SelectionChange(ByVal Target As Range)
On Error GoTo ehandle


    Dim testran As Range
    Dim arr() As String
    
    '<navarro
    Dim qtCOL As Range
    Dim groupCOL As Range
    Dim regionCOL As Range
    Dim sectorCOL As Range
    Dim descCOL As Range
    Dim frm As arrayFORM
    
    
    With grdWS
        Set qtCOL = .[\c_qt].EntireColumn
        Set groupCOL = .[\c_group].EntireColumn
        Set regionCOL = .[\c_region].EntireColumn
        Set sectorCOL = .[\c_sector].EntireColumn
        Set descCOL = .[\c_desc].EntireColumn
    End With
    'navarro>
    
    Set menuSHP = Nothing
    Set buttonCls = Nothing
    
    Set lastRAN = activecell
    lastFORM = lastRAN.Formula
    
    '<navarro
    If Union(activecell, qtCOL).Address = qtCOL.Address And Intersect(activecell.EntireRow, groupCOL).Value = 1 Then
        Set menuSHP = New menuCLS
        arr = menuSHP.createARRAY
        menuSHP.setANCHOR activecell, arr
    ElseIf Union(activecell, descCOL).Address = descCOL.Address And Intersect(activecell.EntireRow, groupCOL).Value = 1 Then
        Set buttonCls = New AddDeleteButton
        buttonCls.setDETAILANCHOR activecell.Offset(0, 1)
    ElseIf Union(activecell, regionCOL).Address = regionCOL.Address Then
        Set frm = New arrayFORM
        frm.register masterOBJ.WS.[regions], activecell
        frm.Show
    ElseIf Union(activecell, sectorCOL).Address = sectorCOL.Address Then
        Set frm = New arrayFORM
        frm.register masterOBJ.WS.[sectors], activecell
        frm.Show
    End If

    
Exit Sub
ehandle:
    LogError "CLassMaster", "grdWS_SelectionChange", Err.Description, Err
End Sub



