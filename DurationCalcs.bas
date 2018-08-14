Attribute VB_Name = "DurationCalcs"
Option Explicit

Sub ReCalcDur(duration As Double, RAN As Range, sht As Worksheet, monthly As Boolean, Optional preCon As Boolean = False)
On Error GoTo ehandle

    Dim codeWB As Workbook
    Dim formulaSTR As String
    Dim cStart As Date, cEnd As Date
    
    Set codeWB = ThisWorkbook
    formulaSTR = "'" & codeWB.Name & "'!"
    sht.Unprotect
    
    cStart = RAN.Offset(-2, 0).Value
    cEnd = RAN.Offset(-1, 0).Value
    
    If preCon = False Then
        If monthly = True Then
            cEnd = DateAdd("m", duration, cStart)
        Else
            cEnd = DateAdd("ww", duration, cStart)
        End If
        RAN.Offset(-1, 0).Value = cEnd
        EnS 0
        RAN.Formula = "=IF(" & formulaSTR & "cDateDiff(""M"",\cstart,\cend)>0," & formulaSTR & "cDateDiff(""M"",\cstart,\cend),0)"
        'RAN.Formula = "=IF(OR(" & formulaSTR & "cDateDiff(""M"",\cstart,\cend)>0," & formulaSTR & "cDateDiff(""M"",\cstart,\cend)<500)," & formulaSTR & "cDateDiff(""M"",\cstart,\cend),0)"
        RAN.Offset(1, 0).Formula = "=IF(" & formulaSTR & "cDateDiff(""WW"",\cstart,\cend)>0," & formulaSTR & "cDateDiff(""WW"",\cstart,\cend),0)"
    ElseIf preCon = True Then
        If monthly = True Then
            cStart = DateAdd("m", -duration, cEnd)
        Else
            cStart = DateAdd("ww", -duration, cEnd)
        End If
        RAN.Offset(-2, 0).Value = cStart
        EnS 0
        RAN.Formula = "=IF(" & formulaSTR & "cDateDiff(""M"",\pstart,\pend)>0," & formulaSTR & "cDateDiff(""M"",\pstart,\pend),0)"
        RAN.Offset(1, 0).Formula = "=IF(" & formulaSTR & "cDateDiff(""WW"",\pstart,\pend)>0," & formulaSTR & "cDateDiff(""WW"",\pstart,\pend),0)"
    End If
    
    basicPROTECT sht, True
    
Exit Sub
    
ehandle:
    LogError "DurationCalcs", "ReCalcDur", Err.Description, Err
    basicPROTECT sht, True
End Sub

Sub ReCalcJobDur(RAN As Range, flip As Boolean)
On Error GoTo ehandle

    Dim jStart As Date, jEnd As Date, cStart As Date
    Dim dRan As Range, perRan As Range, sRan As Range
    Dim perStr As String
    
    cStart = ActiveSheet.Range("\cstart").Value
    
    If flip Then
        Set dRan = Cells(RAN.row, Range("\c_JobDur").Column)
        Set perRan = Cells(RAN.row, Range("\c_perTIME").Column)
        Set sRan = Cells(RAN.row, Range("\c_jobStart").Column)
        jStart = RAN.Value
        jEnd = Cells(RAN.row, Range("\c_posEnd").Column).Value
        If DateDiff("m", jStart, jEnd) < 0 Then
            MsgBox "Cannot Have A Negative Duration", vbExclamation
            Exit Sub
        Else
            dRan.Value = DateDiff("m", jStart, jEnd)
            If jStart < cStart Then
                sRan.Value = cDateDiff("m", cStart, jStart)
            Else
                sRan.Value = cDateDiff("m", cStart, jStart) + 1
            End If
        End If
    ElseIf flip = False Then
        Set dRan = Cells(RAN.row, Range("\c_JobDur").Column)
        Set perRan = Cells(RAN.row, Range("\c_perTIME").Column)
        Set sRan = Cells(RAN.row, Range("\c_jobStart").Column)
        jEnd = RAN.Value
        jStart = Cells(RAN.row, Range("\c_posStart").Column).Value
        If DateDiff("m", jStart, jEnd) < 0 Then
            MsgBox "Cannot Have A Negative Duration", vbExclamation
            Exit Sub
        Else
            dRan.Value = Round(DateDiff("ww", jStart, jEnd) / 4.33, 1)
            If jStart < cStart Then
                sRan.Value = cDateDiff("m", cStart, jStart)
            Else
                sRan.Value = cDateDiff("m", cStart, jStart) + 1
            End If
        End If
     

    End If
    
Exit Sub

ehandle:
    LogError "DurationCalcs", "ReCalcJobDur", Err.Description, Err
End Sub


