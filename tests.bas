Attribute VB_Name = "tests"

Option Explicit
Public sched As schedFORM


Sub fuckpoop()

Dim RAN As Range

Set RAN = Selection
Set RAN = RAN.SpecialCells(xlCellTypeSameFormatConditions)
RAN.Select


End Sub

Sub tempSYNC(RAN As Range)
    Dim WS As Worksheet
    Set WS = RAN.Parent
    
    With RAN
        If .Style = "syncNEED" Then
            EnS 0
            .MergeArea.Style = "syncCOMP"
            .Value = "a"
            .Offset(0, 1).Style = "adminGRN"
        ElseIf .Style = "syncCOMP" Then
            'do nothing
        End If
    End With
    
    WS.[\r_settings].Cells(1, 1).Select

    EnS 1
    
End Sub


Sub shpONaction_check90()
    Dim WS As Worksheet
    Dim shp As Shape
    
    Set WS = ActiveSheet
    
    For Each shp In WS.Shapes
        Debug.Print shp.OnAction
    Next
    


End Sub

Sub schedON()

    'Dim sched As schedFORM
    If masterOBJ Is Nothing Then onOPEN
    If sched Is Nothing Then
        SplashCtrl False
        Set sched = New schedFORM
        sched.Show vbModeless
    End If
    
    
End Sub

Sub SplashCtrl(OnOff As Boolean)

    Dim sht As Worksheet
    Dim shp As Shape
    
    Set sht = masterOBJ.schedWS
    
    If OnOff = False Then
        For Each shp In sht.Shapes
            If Left(shp.Name, 3) = "\sp" Then
                shp.Visible = msoFalse
            ElseIf InStr(shp.Name, "reset") Then
                shp.Visible = msoTrue
            End If
        Next
    ElseIf OnOff = True Then
        For Each shp In sht.Shapes
            If Left(shp.Name, 3) = "\sp" Then
                shp.Visible = msoTrue
            ElseIf InStr(shp.Name, "reset") Then
                shp.Visible = msoFalse
            End If
        Next
    End If
    
End Sub
Sub resetSCHED()

    Dim deleteRAN As Range
    Dim sht As Worksheet
    
    Set sht = ActiveSheet
    Set deleteRAN = Range(sht.[schedSTART].Offset(2, 0), sht.[dele].Offset(-2, 0))
    
    If Not sched Is Nothing Then
        deleteRAN.EntireRow.Delete xlShiftUp
        Unload sched
    End If
    

End Sub
    
Sub testARRcls()

    Dim listRAN As Range
    Dim Target As Range
    Dim frm As arrayFORM
    
    If masterOBJ Is Nothing Then onOPEN
    Set Target = ActiveCell
    Set listRAN = masterOBJ.WS.[sectors]
    Set frm = New arrayFORM
    
    
    frm.register listRAN, Target
    frm.Show

End Sub

Sub asjfhkj()

Dim WS As Worksheet
Dim startRAN As Range, endRAN As Range, mcountran As Range, posstart As Range, posend As Range
Dim startDate As Date, enddate As Date, raisedate As Date
Dim row As Integer, numyears As Integer, year As Date

Set WS = ActiveSheet
Set startRAN = WS.[\cstart]
Set endRAN = WS.[\cend]

row = 14
startDate = startRAN.Value
enddate = endRAN.Value
year = DatePart("yyyy", enddate)
numyears = DateDiff("yyyy", startDate, enddate)

Debug.Print "numyears = "; numyears
End Sub


Sub testdateloop()


Dim WS As Worksheet
Dim startRAN As Range, endRAN As Range, mcountran As Range, posstart As Range, posend As Range
Dim startDate As Date, enddate As Date, raisedate As Date
Dim row As Integer, numyears As Integer, year As Date, mcount As Date, i As Integer, j As Integer
Dim rcount As Integer, mprior As Integer, mafter As Integer, q As Integer, p As Integer
Dim var() As Integer

Set WS = ActiveSheet
Set startRAN = WS.[\cstart]
Set endRAN = WS.[\cend]


startDate = startRAN.Value
enddate = endRAN.Value
year = DatePart("yyyy", enddate)
numyears = DateDiff("yyyy", startDate, enddate)


q = 0: p = 1
For j = 1 To DateDiff("m", startDate, enddate)
    mcount = DateAdd("m", j, startDate)
    If DatePart("m", mcount) = 9 Then
        ReDim Preserve var(1, q) 'stores fiscal yr count in "0" portion and amount of months in before fiscal year in "1" portion
        rcount = q + 1
        mprior = p
        var(0, q) = rcount: var(1, q) = mprior
        p = 1
        q = q + 1
    End If
    'Debug.Print "Month "; j; " = "; mcount
   
    p = p + 1
Next
Debug.Print rcount

End Sub


Sub tester()

    Dim ran1 As Range, ran2 As Range
    Set ran1 = ActiveSheet.Range("j14")
    Set ran2 = ActiveSheet.Range("k14")
    ActiveCell.Value = CalcRaise(ran1, ran2)
End Sub


