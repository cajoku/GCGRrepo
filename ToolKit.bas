Attribute VB_Name = "ToolKit"



Sub FilterSheet(RAN As Range, col As Range)

    Dim cell As Range, tempRAN As Range, tran As Range
    Dim i As Integer, j As Integer
    
    For Each cell In col
        If RAN.Value = cell.Value Or cell.Value = "" Then
            cell.EntireRow.Hidden = False
        Else
            cell.EntireRow.Hidden = True
        End If
    Next
    
    For Each cell In col
        If cell.Offset(0, 2).Font.Bold = True And cell.Offset(1, 2).Value <> "" Then
            i = 0: j = 1
            Set tempRAN = Range(cell.Offset(1, 3).Address)
            Do Until tempRAN.Offset(0, -1).Font.Bold = True
                If tempRAN.EntireRow.Hidden = False Then i = i + 1
                Set tempRAN = tempRAN.Offset(1, 0)
            Loop
            Set tran = Range(cell.Offset(0, 2), tempRAN)
            If i = 0 Then tran.EntireRow.Hidden = True
        ElseIf cell.Offset(0, 1).Font.Bold = True And cell.Offset(1, 1).Value <> "" Then
            i = 0: j = 1
            Set tempRAN = Range(cell.Offset(1, 2).Address)
            Do Until tempRAN.Offset(0, -1).Font.Bold = True
                If tempRAN.EntireRow.Hidden = False Then i = i + 1
                Set tempRAN = tempRAN.Offset(1, 0)
            Loop
            Set tran = Range(cell.Offset(0, 1), tempRAN)
            If i = 0 Then tran.EntireRow.Hidden = True
        End If
    Next

End Sub

Sub trimRANGE(ByRef tran As Range, side As dsDirection)
On Error GoTo ehandle01

'[M Navarro 3/2/16]  modify a range by reducing its dimension on one side
    
    Dim t2ran As Range
    Set t2ran = tran
    On Error GoTo ehandle01
    
    Select Case side
        Case 1
            Set tran = Intersect(tran, tran.Offset(0, 1))
        Case 2
            Set tran = Intersect(tran, tran.Offset(0, -1))
        Case 3
            Set tran = Intersect(tran, tran.Offset(1, 0))
        Case 4
            Set tran = Intersect(tran, tran.Offset(-1, 0))
        Case 5
            Set tran = Intersect(tran, tran.Offset(-1, 0))
            Set tran = Intersect(tran, tran.Offset(1, 0))
        Case 6
            Set tran = Intersect(tran, tran.Offset(0, 1))
            Set tran = Intersect(tran, tran.Offset(0, -1))
        Case 9
            Set tran = Intersect(tran, tran.Offset(1, 1))
            Set tran = Intersect(tran, tran.Offset(-1, -1))
    End Select
    
Exit Sub
ehandle01:
    LogError "Toolkit", "trimRANGE", Err.Description, Err
    
End Sub


Function boxRANGE(WS As Worksheet, STR1 As String, STR2 As String, Optional STR3 As String, Optional Str4 As String) As Range
On Error GoTo ehandle01

    'row + row = range between and including rows, entirerow
    'col + col = same as above
    'row + col = single cell intersection
    'row + row + col + col = range boxed in between ranges
    'row + row + col = same as above except only 1 col wide
    'col + col + row = same as above except only 1 row wide
    'any more than 2 rows or columns = nothing, null range returned
        
    Dim ran1 As Range, ran2 As Range, Ran3 As Range, Ran4 As Range, tran As Range, t2ran As Range
    Dim rColl As Collection, cCOLL As Collection, rCT As Integer, cCT As Integer
    Set rColl = New Collection
    Set cCOLL = New Collection

    
    With WS
        '
        Select Case Left(STR1, 3)
            Case "\r_"
                Set ran1 = .Range(STR1).EntireRow
                rColl.Add ran1
            Case "\c_"
                Set ran1 = .Range(STR1).EntireColumn
                cCOLL.Add ran1
            Case Else
                GoTo ehandle01
        End Select
        '
        Select Case Left(STR2, 3)
            Case "\r_"
                Set ran2 = .Range(STR2).EntireRow
                rColl.Add ran2
            Case "\c_"
                Set ran2 = .Range(STR2).EntireColumn
                cCOLL.Add ran2
            Case Else
                GoTo ehandle01
        End Select
        '
        Select Case Left(STR3, 3)
            Case ""
            Case "\r_"
                Set Ran3 = .Range(STR3).EntireRow
                If rColl.Count = 2 Then GoTo eHandle02 _
                Else rColl.Add Ran3
            Case "\c_"
                Set Ran3 = .Range(STR3).EntireColumn
                If cCOLL.Count = 2 Then GoTo eHandle02 _
                Else cCOLL.Add Ran3
            Case Else
                GoTo ehandle01
        End Select
        '
        Select Case Left(Str4, 3)
            Case ""
            Case "\r_"
                Set Ran4 = .Range(Str4).EntireRow
                If rColl.Count = 2 Then GoTo eHandle02 _
                Else rColl.Add Ran4
            Case "\c_"
                Set Ran4 = .Range(Str4).EntireColumn
                If cCOLL.Count = 2 Then GoTo eHandle02 _
                Else cCOLL.Add Ran4
            Case Else
                GoTo ehandle01
        End Select
            
    End With
    
    rCT = rColl.Count
    cCT = cCOLL.Count
    
    Select Case rCT + cCT
        Case 2
            If rCT = 0 Then
                Set tran = WS.Range(cCOLL.Item(1).Address(False, False) & ":" & cCOLL.Item(2).Address(False, False))
            ElseIf rCT = 1 Then
                Set tran = Intersect(rColl.Item(1), cCOLL.Item(1))
            Else
                Set tran = WS.Range(rColl.Item(1).Address(False, False) & ":" & rColl.Item(2).Address(False, False))
            End If
        Case 3
            If rCT = 1 Then
                Set tran = WS.Range(cCOLL.Item(1).Address(False, False) & ":" & cCOLL.Item(2).Address(False, False))
                Set tran = Intersect(tran, rColl.Item(1))
            Else
                Set tran = WS.Range(rColl.Item(1).Address(False, False) & ":" & rColl.Item(2).Address(False, False))
                Set tran = Intersect(tran, cCOLL.Item(1))
            End If
        Case 4
            Set tran = WS.Range(rColl.Item(1).Address(False, False) & ":" & rColl.Item(2).Address(False, False))
            Set t2ran = WS.Range(cCOLL.Item(1).Address(False, False) & ":" & cCOLL.Item(2).Address(False, False))
            Set tran = Intersect(tran, t2ran)
    End Select
    
    Set boxRANGE = tran
    
Exit Function
ehandle01:
    Set boxRANGE = Nothing
    LogError "Toolkit", "boxRANGE", "failed to create range, range name not found", Err

Exit Function
eHandle02:
    Set boxRANGE = Nothing
    LogError "Toolkit", "boxRANGE", "failed to create range, more than 2 rows or columns specified", Err
    
End Function

Function expandRAN(RAN As Range, Optional TRUEisCOL_FALSEisROW As Boolean) As Range
On Error GoTo ehandle01

    'creates an expanded range from a single cell by looping until a blank is reacehd
    Dim cell As Range
    Dim tran As Range
    Dim R As Integer, c As Integer, iTEST As Integer
    
    iTEST = 0
    
    If TRUEisCOL_FALSEisROW Then
        c = 1: R = 0
    Else
        c = 0: R = 1
    End If
    
    
    Do Until RAN.Value = ""
        If tran Is Nothing Then Set tran = RAN Else Set tran = Union(tran, RAN)
        Set RAN = RAN.Offset(R, c)
        If iTEST > 1000 Then Exit Do
    Loop
    
    If tran Is Nothing Then Debug.Print "ERROR:  expandRAN returned a blank range" Else Set expandRAN = tran
    
Exit Function
ehandle01:
    LogError "Toolkit", "expandRAN", Err.Description, Err
    
End Function

Public Sub EnS(I_ As IO, Optional callerNAME As String, Optional overRIDE As Boolean)
On Error GoTo e1

    Dim boo As Boolean
    
    If Not overRIDE And I_ = ensON Then Exit Sub
    
    If I_ = 1 Then boo = True
    
    With Application
        .EnableEvents = boo
        .ScreenUpdating = boo
        If I_ = 0 Then
            .Calculation = xlCalculationManual
        Else
            .Calculation = xlCalculationAutomatic
        End If
    End With
    
    ensON = I_
    
'    If I_ Then Debug.Print "::-   events & screen updates   -:: [ON]  " & callerNAME Else _
'    Debug.Print "::-   events & screen updates   -:: [OFF]  " & callerNAME
Exit Sub
e1:
    LogError "ToolKit", "EnS", Err.Description, Err
End Sub

Sub insertGLOBAL(RAN As Range, Optional TRUEcolFALSErow As Boolean, Optional templateNM As String, Optional insertOSET As Integer = 1)
On Error GoTo e1
    'inserts not only on the source ws, but on all ws's that link to the insertRAN
    
    
    Dim tempRAN As Range
    Dim WS As Worksheet, gcoWS As Worksheet
    Dim ranColl As Collection
    Dim cell As Range
    Dim insertRAN As Range
    Dim shp As Shape
    
    Set gcoWS = masterWB.Worksheets("GCs Owner")
    Set WS = RAN.Parent
    Set insertRAN = RAN.Offset(insertOSET, 0)
    
    WS.Unprotect
    'gcoWS.Unprotect
    
    If templateNM = "" Then
        If TRUEcolFALSErow Then
            templateNM = "\c_temp"
        Else
            templateNM = "\r_temp"
        End If
    ElseIf LCase(templateNM) = "activecell" Then
        WS.Activate
        Set tempRAN = ActiveCell
    End If
    
    If Left(templateNM, 2) = "\c" Then
        If tempRAN Is Nothing Then
            Set tempRAN = WS.Range(templateNM).EntireColumn
        Else
            Set tempRAN = tempRAN.EntireColumn
        End If
        
        tempRAN.Copy
        insertRAN.EntireColumn.Insert
        
        Set ranColl = dependentCOLL(RAN)
        For Each cell In ranColl
            With cell
                .EntireColumn.Copy
                .Offset(0, 1).EntireColumn.Insert
            End With
        Next
        
    Else
        If tempRAN Is Nothing Then
            Set tempRAN = WS.Range(templateNM).EntireRow
        Else
            Set tempRAN = tempRAN.EntireRow
        End If
        
        tempRAN.Copy
        insertRAN.EntireRow.Insert
        
        Set ranColl = dependentCOLL(RAN)
        For Each cell In ranColl
            With cell
                .EntireRow.Copy
                .Offset(1, 0).EntireRow.Insert
                If .Parent.Name = masterOBJ.gantWS.Name Then
                    For Each shp In .Parent.Shapes
                        If shp.TopLeftCell.EntireRow.Address = .Offset(1, 0).EntireRow.Address Then shp.Delete
                    Next
                    newGant.createBar .Offset(1, 0)
                End If
            End With
        Next
    End If

    Application.CutCopyMode = False
    basicPROTECT WS, True
    'basicPROTECT gcoWS, True
    'basicPROTECT masterOBJ.groWS, True
    
Exit Sub
e1:
    LogError "ToolKit", "insertGLOBAL", Err.Description, Err
    basicPROTECT WS, True
    'basicPROTECT gcoWS, True
    'basicPROTECT masterOBJ.groWS, True
    
End Sub

Sub deleteGLOBAL(RAN As Range, Optional TRUEcolFALSErow As Boolean)
On Error GoTo e1
    'inserts not only on the source ws, but on all ws that link to the insertRAN
    
    Dim tempRAN As Range
    Dim WS As Worksheet, gcoWS As Worksheet
    Dim ranColl As Collection
    Dim cell As Range
    Dim insertRAN As Range
    
    Set WS = RAN.Parent
    Set gcoWS = masterWB.Worksheets("GCs Owner")
    
    WS.Unprotect
    'gcoWS.Unprotect
    
    If TRUEcolFALSErow Then
        Set ranColl = dependentCOLL(RAN)
        For Each cell In ranColl
            cell.EntireColumn.Delete
        Next
        RAN.EntireColumn.Delete
    Else
        Set ranColl = dependentCOLL(RAN)
        For Each cell In ranColl
            cell.EntireRow.Delete
        Next
        RAN.EntireRow.Delete
    End If

    basicPROTECT WS, True
    'basicPROTECT gcoWS, True
    
Exit Sub
e1:
    LogError "ToolKit", "deleteGLOBAL", Err.Description, Err
    basicPROTECT WS, True
    'basicPROTECT gcoWS, True
    
End Sub


Function dependentCOLL(RAN As Range) As Collection
On Error GoTo e1

    Dim tcoll As Collection
    Dim ensTEMP As IO
    Dim i As Integer
    Dim t2ran As Range
    
    'assert turning off of events/screen/calc
    'EnS 0, "dependentCOLL", True
    Set tcoll = New Collection
    
    i = 1
    With RAN
        .ShowDependents
        Do
            On Error GoTo quickout
            .NavigateArrow False, 1, i
            Set t2ran = ActiveCell
            If t2ran.Parent.Name & t2ran.Address <> RAN.Parent.Name & RAN.Address Then
                tcoll.Add ActiveCell
            Else
                GoTo quickout
            End If
            i = i + 1
            If i > 50 Then GoTo e1
        Loop
    End With
    
quickout:
    Set dependentCOLL = tcoll
    RAN.Parent.Activate
    RAN.Select
    RAN.ShowDependents Remove:=True
    
Exit Function
e1:
    'LogError "ToolKit", "dependentRAN", "loop exceeded 1", 0
    'Debug.Print "ToolKit, dependentRAN, loop exceeded 1, 0"
    GoTo quickout
End Function



Function InRange(ran1, ran2) As Boolean
On Error GoTo e1

    Dim cell As Range
    
    For Each cell In ran1
        If Intersect(cell, ran2) Is Nothing Then
            InRange = False
            Exit Function
        End If
    Next
    
    InRange = True

Exit Function
e1:
    LogError "ToolKit", "InRange", Err.Description, Err
    InRange = False
End Function


Sub HideZeros()

    If Not ActiveWindow.DisplayZeros = False Then
        ActiveWindow.DisplayZeros = False
    End If

End Sub

Function countSTAFF() As Integer

    Dim sht As Worksheet
    Dim staffRAN1 As Range, staffRAN2 As Range, staffFinal As Range, cell As Range, tempRAN As Range
    
    Set sht = masterOBJ.sdWS
    Set staffRAN1 = boxRANGE(sht, "\r_precon", "\c_Position", "\r_constr")
    trimRANGE staffRAN1, dsupdown
    Set staffRAN2 = boxRANGE(sht, "\r_constr", "\c_Position", "\r_end")
    trimRANGE staffRAN2, dsupdown

    If Not staffRAN1 Is Nothing And Not staffRAN2 Is Nothing Then
        Set staffFinal = Union(staffRAN1, staffRAN2)
    ElseIf staffRAN1 Is Nothing And Not staffRAN2 Is Nothing Then
        Set staffFinal = staffRAN2
    ElseIf Not staffRAN1 Is Nothing And staffRAN2 Is Nothing Then
        Set staffFinal = staffRAN1
    End If
    
    If staffFinal Is Nothing Then
        countSTAFF = 0
        Exit Function
    Else
        For Each cell In staffFinal
            If cell.Value <> "" Then
                countSTAFF = countSTAFF + 1
            End If
        Next
    End If

End Function

Function cDateDiff(ctype As String, startDate As Date, enddate As Date) As Double
On Error GoTo e1
   ' "created 4/6 bc built ws DATEDIF() and vba DateDiff return diff values. This will uniformly use vba DD"'
    Dim temp As Double
    Dim neg As Boolean
    
    If startDate <> 0 Or enddate <> 0 Then
        If UCase(ctype) = "M" Then
            cDateDiff = DateDiff("M", startDate, enddate)
        ElseIf UCase(ctype) = "WW" Then
            cDateDiff = DateDiff("ww", startDate, enddate)
        Else
            cDateDiff = DateDiff(UCase(ctype), startDate, enddate)
        End If
    End If
    
Exit Function
e1:
    'LogError "ToolKit", "cDateDiff",err.description, Err
End Function


Sub ToggleWorkRegion(region As String)

    Dim WS As Worksheet
    Dim headerRAN As Range, workRAN As Range, formSTR As String
    
    Set WS = masterWB.Worksheets("Code")
    Set headerRAN = WS.[\workregion]
    
    Set headerRAN = headerRAN.Find(What:=region)
    Set workRAN = Intersect(WS.[\workweeks].EntireRow, headerRAN.EntireColumn)
    formSTR = "Code!" & workRAN.Address
    
    masterWB.Names("\t_workweek").RefersTo = "=" & formSTR


End Sub
