Attribute VB_Name = "Admin"
Option Explicit

Sub onOPENAdmin()
On Error GoTo e1

    Dim sht As Worksheet
    
    Set masterWB = ThisWorkbook
    Set masterOBJ = New ClassMaster
    Set newGant = New NewGantClass

    ensON = 1
    EnS 0, "Workbook Open"
    
    For Each sht In masterWB.Sheets
        If sht.Visible <> xlSheetVeryHidden Then
            sht.Visible = xlSheetVisible
        End If
    Next
    
    With masterWB
        .Windows(1).Visible = True
        .Windows(1).WindowState = xlMaximized
        .Worksheets("Splash").Visible = xlSheetHidden
    End With
    
    EnS 1, "Workbook Open"
    
Exit Sub
e1:
    LogError "Startup", "onOPEN", Err.Description, Err
    EnS 1, , True
End Sub

Sub shapeInitialize(userWB As Workbook, codeWB As Workbook)
On Error Resume Next

    Dim shp As Shape
    Dim sht As Worksheet
    Dim actionstr As String
    Dim namearr As Variant
    Dim namestr As String
    
    actionstr = "'" & codeWB.Name & "'!"

    For Each sht In userWB.Worksheets
        If sht.Visible <> xlSheetVeryHidden Then
            sht.Unprotect
            For Each shp In sht.Shapes
                If shp.OnAction <> "" Then 'Debug.Print sht.Name & " - " & shp.Name & " - " & shp.OnAction
                    If InStr(shp.OnAction, "!") > 0 Then
                        namearr = Split(shp.OnAction, "!")
                        namestr = namearr(1)
                        shp.OnAction = actionstr & namestr
                        'Debug.Print shp.Name & " - " & shp.OnAction
                    Else
                        namestr = shp.OnAction
                        shp.OnAction = actionstr & namestr
                        'Debug.Print shp.Name & " - " & shp.OnAction
                    End If
                End If
            Next
        End If
        basicPROTECT sht, True
    Next
    
Exit Sub
e1:
    LogError "Admin", "shapeInitialize", Err.Description, Err
    Debug.Print sht.Name & " - " & shp.Name & " - " & shp.OnAction & " - " & Err & " - " & Err.Description
End Sub

Sub UDFinitialize(userWB As Workbook, codeWB As Workbook)
On Error GoTo e1

    Dim sht As Worksheet
    Dim cell As Range
    Dim formulaSTR As String, shortformSTR As String, formulaSTR2 As String, tempSTR As String, tempSTR2 As String
    Dim cdatestr As String, calcstr As String, tSTR As String
    
    cdatestr = "cDateDiff": calcstr = "calcRaise": tSTR = "trailertbl"
    
    formulaSTR = codeWB.Path & "\" & codeWB.Name & "!"
    formulaSTR2 = "'" & codeWB.Path & "\" & codeWB.Name & "'!"
    shortformSTR = codeWB.Name & "!"
    tempSTR2 = "'" & codeWB.Name & "'!"
    
    For Each sht In userWB.Worksheets
        If sht.Visible <> xlSheetVeryHidden Then
            sht.Unprotect
            For Each cell In sht.UsedRange
                If InStr(1, cell.Formula, formulaSTR) Then
                    tempSTR = Replace(cell.Formula, formulaSTR, tempSTR2)
                    cell.Formula = tempSTR
                ElseIf InStr(1, cell.Formula, shortformSTR) Then
                    tempSTR = Replace(cell.Formula, shortformSTR, tempSTR2)
                    cell.Formula = tempSTR
                    Debug.Print "short string replace" & sht.Name & " - " & cell.Address & " - " & cell.Formula
                ElseIf InStr(1, cell.Formula, cdatestr) Then
                    If InStr(1, cell.Formula, shortformSTR & cdatestr) Then
                        tempSTR = Replace(cell.Formula, Mid(cell.Formula, InStr(cell.Formula, shortformSTR & cdatestr), Len(shortformSTR) + Len(cdatestr)), tempSTR2 & cdatestr)
                        cell.Formula = tempSTR
                        Debug.Print sht.Name & " - " & cell.Address & " - " & cell.Formula
                    Else
                        tempSTR = Replace(cell.Formula, cdatestr, tempSTR2 & cdatestr)
                        cell.Formula = tempSTR
                        Debug.Print sht.Name & " - " & cell.Address & " - " & cell.Formula
                    End If
                ElseIf InStr(1, cell.Formula, calcstr) Then
                    If InStr(1, cell.Formula, shortformSTR & calcstr) Then
                        tempSTR = Replace(cell.Formula, Mid(cell.Formula, InStr(cell.Formula, shortformSTR & calcstr), Len(shortformSTR) + Len(calcstr)), tempSTR2 & calcstr)
                        cell.Formula = tempSTR
                        Debug.Print sht.Name & " - " & cell.Address & " - " & cell.Formula
                    Else
                        tempSTR = Replace(cell.Formula, calcstr, tempSTR2 & calcstr)
                        cell.Formula = tempSTR
                        Debug.Print sht.Name & " - " & cell.Address & " - " & cell.Formula
                    End If
                ElseIf InStr(1, cell.Formula, tSTR) Then
                    If InStr(1, cell.Formula, shortformSTR & tSTR) Then
                        tempSTR = Replace(cell.Formula, Mid(cell.Formula, InStr(cell.Formula, shortformSTR & tSTR), Len(shortformSTR) + Len(tSTR)), tempSTR2 & tSTR)
                        cell.Formula = tempSTR
                        Debug.Print sht.Name & " - " & cell.Address & " - " & cell.Formula
                    Else
                        tempSTR = Replace(cell.Formula, tSTR, tempSTR2 & tSTR)
                        cell.Formula = tempSTR
                        Debug.Print sht.Name & " - " & cell.Address & " - " & cell.Formula
                    End If
                ElseIf InStr(1, cell.Formula, "tbl") Then
                    If InStr(1, cell.Formula, shortformSTR & "tbl") Then
                        tempSTR = Replace(cell.Formula, Mid(cell.Formula, InStr(cell.Formula, shortformSTR & "tbl"), Len(shortformSTR) + 3), tempSTR2 & "tbl")
                        cell.Formula = tempSTR
                        Debug.Print sht.Name & " - " & cell.Address & " - " & cell.Formula
                    Else
                        tempSTR = Replace(cell.Formula, "tbl", tempSTR2 & "tbl")
                        cell.Formula = tempSTR
                        Debug.Print sht.Name & " - " & cell.Address & " - " & cell.Formula
                    End If
                End If
            Next
            basicPROTECT sht, True
        End If
    Next
    
Exit Sub
e1:
    LogError "Admin", "UDFinitialize", Err.Description, Err
    Debug.Print sht.Name & " - " & cell.Name & " - " & cell.Formula & " - " & Err & " - " & Err.Description
End Sub

Sub shapeDir()


    Dim shp As Shape
    Dim sht As Worksheet
    Dim WB As Workbook
    
    Set WB = ActiveWorkbook
    
    For Each sht In WB.Worksheets
        If sht.Visible = xlSheetVisible Then
            For Each shp In sht.Shapes
                If Left(shp.Name, 1) = "\" Then Debug.Print sht.Name & " - " & shp.Name & " - " & shp.OnAction
            Next
        End If
    Next

End Sub

Sub UDFDir()


    Dim sht As Worksheet
    Dim WB As Workbook
    Dim cell As Range
    
    Set WB = ActiveWorkbook
    
    For Each sht In WB.Worksheets
        If sht.Visible = xlSheetVisible Then
            sht.Unprotect
            For Each cell In sht.UsedRange
                If InStr(1, cell.Formula, "cDateDiff") Or InStr(1, cell.Formula, "calcRaise") Or InStr(1, cell.Formula, "tbl") Then
                    Debug.Print sht.Name & " - " & cell.Address & " - " & cell.Formula
                End If
            Next
        End If
    Next

End Sub

Sub cleanFormulas(userWB As Workbook, codeWB As Workbook)
On Error GoTo e1

    Dim codestr As String, temp As String, tempaa As String, temper As String
    Dim sht As Worksheet
    Dim cell As Range
    
    codestr = codeWB.Name
    tempaa = "'" & codestr & "'!"
    temper = codestr & "!"
    For Each sht In userWB.Sheets
        sht.Unprotect
        If sht.Visible <> xlSheetVeryHidden Then
            For Each cell In sht.UsedRange
                If InStr(cell.Formula, codestr) Then
                    temp = cell.Formula
                    temp = Replace(cell.Formula, temper, "")
                    cell.Formula = temp
                    Debug.Print sht.Name & " - " & cell.Address & " - " & cell.Formula
                End If
            Next
        End If
    Next
        
Exit Sub
e1:
    LogError "Admin", "cleanFormulas", Err.Description, Err

End Sub

Sub dsjkfhs()

    Dim userWB As Workbook, codeWB As Workbook
    Dim codestr As String, temp As String, tempaa As String, temper As String
    Dim sht As Worksheet
    Dim cell As Range
    
    Set userWB = ActiveWorkbook
    Set codeWB = ThisWorkbook
    codestr = codeWB.Name
    tempaa = "'" & codestr & "'!"
    temper = codestr & "!"
    For Each sht In userWB.Sheets
        sht.Unprotect
        If sht.Visible = xlSheetVisible Then
            For Each cell In sht.UsedRange
                If InStr(cell.Formula, codestr) Then
                    temp = cell.Formula
                    temp = Replace(cell.Formula, temper, "")
                    cell.Formula = temp
                    Debug.Print sht.Name & " - " & cell.Address & " - " & cell.Formula
                End If
            Next
        End If
    Next
        

End Sub


Sub stripMACROS()

On Error Resume Next
    Dim Element As Object
    For Each Element In ActiveWorkbook.VBProject.VBComponents
        ActiveWorkbook.VBProject.VBComponents.Remove Element
    Next

End Sub


Function FolderExists(ByVal Path As String) As Boolean
  On Error Resume Next
  
  FolderExists = Dir(Path, vbDirectory) <> ""
  
End Function


Sub clearError(WB As Workbook)

    Dim WS As Worksheet
    Dim printRAN As Range
    
    Set WS = WB.Worksheets("Errors")
    Set printRAN = WS.[\errors]
    
    printRAN.Value = ""

End Sub
