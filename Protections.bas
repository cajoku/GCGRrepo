Attribute VB_Name = "Protections"
Option Explicit

Sub xp()

EnS 0
unPROTECTme
EnS 1

End Sub

Sub unPROTECTme()

    Dim sht As Worksheet
    Dim tSTR As String
    
    For Each sht In masterWB.Sheets
        sht.Unprotect 'dsPASSWORD
    Next
        
End Sub

Sub protectME()

dsPROTECT masterWB

End Sub

Sub dsPROTECT(wb As Workbook)
On Error GoTo ehandle01
    
    Dim sht As Worksheet
    Dim tSTR As String
    
    For Each sht In wb.Sheets
        If sht.Visible <> xlSheetVeryHidden Then
            If sht.CodeName = "Sheet3" Then
                strictPROTECT sht, True
            ElseIf sht.CodeName <> "Sheet4" And sht.CodeName <> "Sheet6" Then
                basicPROTECT sht, True
            End If
        End If
    Next
        
Exit Sub
ehandle01:
    LogError "Protections", "dsPROTECT", Err.Description, Err
     
End Sub


Sub basicPROTECT(sht As Worksheet, TurnOn As Boolean)
    
    sht.Unprotect 'dsPASSWORD
    
    If TurnOn Then
        sht.Protect _
            DrawingObjects:=False, _
            contents:=True, _
            Scenarios:=True, _
            userinterfaceonly:=True, _
            AllowFormattingCells:=True, _
            AllowFormattingColumns:=True, _
            AllowFormattingRows:=True, _
            AllowInsertingColumns:=True, _
            AllowInsertingRows:=True, _
            AllowInsertingHyperlinks:=True, _
            AllowDeletingColumns:=True, _
            AllowDeletingRows:=True, _
            AllowSorting:=True, _
            AllowFiltering:=True, _
            AllowUsingPivotTables:=True
        sht.EnableOutlining = True
        sht.EnableSelection = xlNoRestrictions
'    Else
'        sht.Unprotect dsPASSWORD
    End If
            
End Sub

Sub strictPROTECT(sht As Worksheet, TurnOn As Boolean, Optional lockALLcells As Boolean)
    
    If lockALLcells Then sht.UsedRange.Locked = True
    
    If TurnOn Then
        sht.Protect _
            DrawingObjects:=True, _
            contents:=True, _
            Scenarios:=True, _
            userinterfaceonly:=True, _
            AllowFormattingCells:=False, _
            AllowFormattingColumns:=False, _
            AllowFormattingRows:=False, _
            AllowInsertingColumns:=False, _
            AllowInsertingHyperlinks:=False, _
            AllowDeletingColumns:=False, _
            AllowDeletingRows:=False, _
            AllowInsertingRows:=False, _
            AllowSorting:=False, _
            AllowFiltering:=False, _
            AllowUsingPivotTables:=False
        sht.EnableOutlining = True
        sht.EnableSelection = xlUnlockedCells
    Else
        sht.Unprotect 'dsPASSWORD
    End If
 
End Sub



