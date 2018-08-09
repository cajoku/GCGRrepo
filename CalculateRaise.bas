Attribute VB_Name = "CalculateRaise"
Option Explicit

Function CalcRaise(startRAN As Range, endRAN As Range, phase As Range, Optional percentTime As Boolean = True) As Double
On Error GoTo ehandle

    Dim WS As Worksheet
    Dim monthRan As Range, cell As Range, sMonthRan As Range, mRan As Range
    Dim startDate As Date, enddate As Date, fiscalYearStart As Date, counter As Date, oStartDate As Date, currentDate As Date
    Dim fyCount As Integer, i As Integer, j As Integer, q As Integer, p As Integer, weeks As Double, months As Integer, startMonth As Integer, raiseCurrent As Double
    Dim pTime As Double, raise As Double
    Dim wColl As Collection, pColl As Collection
    Dim fstart As Integer, fend As Integer

    Application.Volatile
    
    Set WS = startRAN.Parent
    Set mRan = Cells(startRAN.row, WS.[\c_durSTART].Column)
    Set sMonthRan = WS.Cells(startRAN.row, WS.[\c_jobStart].Column)
    Set wColl = New Collection
    Set pColl = New Collection
    
    
    startMonth = sMonthRan.Value
    raise = phase.Value
    startDate = startRAN.Value
    currentDate = Now()
    enddate = endRAN.Value
    oStartDate = startDate
    
    raiseCurrent = FindRaise(currentDate, startDate, raise) 'raisecurrent added 7/26 bc raise calculation should begin from current date instead of staff start
    
    fyCount = 1
'============================
'replaced GetFormulaRange Procedure with this portion to eliminate calc errors 4/17
    If startMonth < 0 Then
        fstart = startMonth
    Else
        fstart = startMonth - 1 'if there are neg columns subtract one bc we skip a "0" col
    End If
    
    fend = fstart + sMonthRan.Offset(0, 1).Value
    
    If sMonthRan.Offset(0, 1).Value = "" Then
        Set monthRan = mRan.Offset(0, fstart)
    Else
        Set monthRan = WS.Range(WS.Cells(mRan.Offset(0, fstart).row, mRan.Offset(0, fstart).Column), WS.Cells(mRan.Offset(0, fend - 1).row, mRan.Offset(0, fend - 1).Column))
    End If
'=============================
        
    If startDate = 0 Or enddate = 0 Then GoTo quickout
    
    i = 1: p = 1
    For j = 1 To DateDiff("m", startDate, enddate)
        counter = DateAdd("m", i, startDate)
        If DatePart("m", counter) = 9 Then
            fiscalYearStart = FYstart(counter)
            weeks = Abs(startDate - counter) / 7
            months = DateDiff("m", startDate, counter)
            wColl.Add weeks
            For p = p To j
                pTime = pTime + monthRan(p).Value
            Next
            pTime = pTime / months
            pColl.Add pTime
            pTime = 0
            startDate = counter
            i = 0
            If j <> DateDiff("m", oStartDate, enddate) Then fyCount = fyCount + 1
        ElseIf j = DateDiff("m", oStartDate, enddate) Then
            If fyCount > 1 Then 'added this condition and the else clause to account for durations that do not reach the next fiscal year 3/22
                weeks = Abs(startDate - enddate) / 7
            Else
                weeks = Abs(oStartDate - enddate) / 7
            End If
            months = DateDiff("m", startDate, counter)
            wColl.Add weeks
            For p = p To j
                pTime = pTime + monthRan(p).Value
            Next
            pTime = pTime / months
            pColl.Add pTime
        End If
        i = i + 1
    Next
            
    q = 1
    Do Until q > fyCount
        If q = 1 And j = 1 Then
            CalcRaise = CalcRaise
            Exit Function
        ElseIf q = 1 Then
            If percentTime = True Then
                CalcRaise = CalcRaise + (1 + raiseCurrent) * (wColl(q) * pColl(q))
                q = q + 1
            ElseIf percentTime = False Then
                CalcRaise = CalcRaise + (1 + raiseCurrent) * wColl(q)
                q = q + 1
            End If
        Else
            If percentTime = True Then
                CalcRaise = CalcRaise + (1 + raise) * (wColl(q) * pColl(q))
                q = q + 1
            ElseIf percentTime = False Then
                CalcRaise = CalcRaise + (1 + raise) * wColl(q)
                q = q + 1
            End If
        End If
    Loop

Exit Function

quickout:
    CalcRaise = CalcRaise
Exit Function

ehandle:
    'LogError "CalculateRaise", "CalcRaise",err.description, Err
    CalcRaise = CalcRaise
End Function

Function FYstart(curDate As Date) As Date
On Error GoTo ehandle

    Dim fYear As Variant, fMonth As Variant, fDay As Variant
    
    fYear = DatePart("yyyy", curDate)
    fMonth = DatePart("m", curDate)
    fDay = DatePart("d", curDate)
    
    FYstart = DateSerial(fYear, fMonth, 1)

Exit Function
ehandle:
    LogError "CalculateRaise", "FYstart", Err.Description, Err
    Debug.Print "For Colby: Error w FYstart UDF for CalcRaise"
    FYstart = curDate
    
End Function

Function FindRaise(nowDate As Date, startDate As Date, raise As Double) As Double

    Dim i As Integer
    Dim tempDate As Date, fyDate As Date
    Dim fyCount As Integer
    
    If startDate < nowDate Then FindRaise = 0: Exit Function
    
    For i = 1 To DateDiff("m", nowDate, startDate)
        tempDate = DateAdd("m", i, nowDate)
        If DatePart("m", tempDate) = 9 Then
            fyDate = FYstart(tempDate)
            If tempDate >= fyDate Then fyCount = fyCount + 1
        End If
    Next
    
    FindRaise = fyCount * raise

End Function
