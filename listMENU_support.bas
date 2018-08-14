Attribute VB_Name = "listMENU_support"
Option Explicit
Public listMENU As listMENU_cls

Sub precon_listMENU_CLICK()
On Error GoTo ehandle

    Dim Target As Range
    Dim WS As Worksheet
    Dim tableran As Range
    
    If masterOBJ Is Nothing Then onOPEN
        
    Set WS = masterOBJ.sdWS
    Set Target = boxRANGE(WS, "\c_Position", "\r_precon")
    Set tableran = masterOBJ.wb.Worksheets("Code").[\staffTABLE]
    
    Set listMENU = New listMENU_cls
    
    listMENU.setANCHOR Target, tableran, "listMENU_Select"
    listMENU.openLIST
    
Exit Sub
    
ehandle:
    LogError "listMENU_support", "precon_listMENU_CLICK", Err.Description, Err
End Sub

Sub con_listMENU_CLICK()
On Error GoTo ehandle

    Dim Target As Range
    Dim WS As Worksheet
    Dim tableran As Range
    
    If masterOBJ Is Nothing Then onOPEN
        
    Set WS = masterOBJ.sdWS
    Set Target = boxRANGE(WS, "\c_Position", "\r_constr")
    Set tableran = masterOBJ.wb.Worksheets("Code").[\staffTABLE]
    
    Set listMENU = New listMENU_cls
    
    listMENU.setANCHOR Target, tableran, "listMENU_Select"
    listMENU.openLIST

Exit Sub

ehandle:
    LogError "listMENU_support", "con_listMENU_CLICK", Err.Description, Err
    
End Sub

Sub addtrailer_CLICK()
On Error GoTo e1

    Dim Target As Range
    Dim WS As Worksheet
    Dim tableran As Range
    
    If masterOBJ Is Nothing Then onOPEN
    
    Set WS = masterOBJ.gcdWS
    Set Target = boxRANGE(WS, "\r_trailer", "\c_desc", "\c_qt")
    trimRANGE Target, dsRIGHT
    
    Set tableran = masterOBJ.wb.Worksheets("Code").[\trailerTABLE]
    
    Set listMENU = New listMENU_cls
    
    listMENU.setANCHOR Target, tableran, "listMENU_Select2"
    listMENU.openLIST
    
Exit Sub
e1:
    LogError "listMENU_support", "addtrailer_CLICK", Err.Description, Err
End Sub

Sub phase_listMENU_CLICK()
'On Error GoTo ehandle

    Dim Target As Range
    Dim WS As Worksheet
    Dim tableran As Range
    
    If masterOBJ Is Nothing Then onOPEN
        
    Set WS = masterOBJ.sdWS
    Set Target = WS.Shapes(Application.Caller).TopLeftCell
    
    Set tableran = masterOBJ.wb.Worksheets("Code").[\staffTABLE]
    
    Set listMENU = New listMENU_cls
    
    listMENU.setANCHOR Intersect(WS.[\c_Position].EntireColumn, Target.EntireRow), tableran, "listMENU_Select"
    listMENU.openLIST

Exit Sub

ehandle:
    LogError "listMENU_support", "con_listMENU_CLICK", Err.Description, Err
    
End Sub

Sub listMENU_Select(resultCOLL As Collection, executetarget As Range)
On Error GoTo ehandle

    Dim i As Integer
    Dim WS As Worksheet
    
    Set WS = executetarget.Parent
    WS.Unprotect
    
    EnS 0
    
    i = resultCOLL.Count
    Do Until i = 0
        InsertStaff executetarget, resultCOLL(i)
        i = i - 1
    Loop
    
    'SortJobs executetarget
    
    EnS 1
    basicPROTECT WS, True
Exit Sub

ehandle:
    LogError "listMENU_support", "listMENU_Select", Err.Description, Err
    EnS 1, , True
    basicPROTECT WS, True
    
End Sub

Sub listMENU_Select2(resultCOLL As Collection, executetarget As Range)
On Error GoTo e1

    Dim i As Integer, j As Integer
    Dim WS As Worksheet
    Dim trailerRAN As Range, cell As Range
    Dim stuffCOLL As Collection
    
    Set WS = executetarget.Parent
    
    EnS 0
    
    Set stuffCOLL = dependentCOLL(executetarget.Offset(-1, 0).Cells(1, 1))
    WS.Unprotect
    WS.[\r_temptrailer].EntireRow.Hidden = False
    
    For Each cell In stuffCOLL
        i = resultCOLL.Count
        Do Until i = 0
            WS.[\r_temptrailer].EntireRow.Copy
            executetarget.EntireRow.Insert
            Set trailerRAN = Intersect(executetarget.EntireRow, WS.[\c_desc].EntireColumn).Offset(-9, 0)
            trailerRAN.Value = resultCOLL(i).Value
            ''''''''''''
            
            For j = 0 To WS.[\r_temptrailer].Rows.Count - 1
                cell.Offset(j, 0).EntireRow.Copy
                cell.Offset(j + 1, 0).EntireRow.Insert
                If j = 0 Then
                    cell.Parent.[\r_heading].EntireRow.Hidden = False
                    cell.Parent.[\r_heading].EntireRow.Copy
                    cell.Offset(j + 1, 0).EntireRow.PasteSpecial xlPasteFormats
                    cell.Parent.[\r_heading].EntireRow.Hidden = True
                Else
                    cell.Parent.[\r_lineitem].EntireRow.Hidden = False
                    cell.Parent.[\r_lineitem].EntireRow.Copy
                    cell.Offset(j + 1, 0).EntireRow.PasteSpecial xlPasteFormats
                    cell.Parent.[\r_lineitem].EntireRow.Hidden = True
                End If
                Application.CutCopyMode = False
            Next
            ''''''''''''''
        i = i - 1
        Loop
    Next
    
    WS.[\r_temptrailer].EntireRow.Hidden = True
    
    'basicPROTECT WS, True
    EnS 1

Exit Sub
e1:
    LogError "listMENU_support", "listMENU_Select2", Err.Description, Err
    EnS 1, , True
    basicPROTECT WS, True
    
End Sub

Sub listMENU_EXIT()
On Error GoTo ehandle

    Set listMENU = Nothing

Exit Sub
    
ehandle:
    LogError "listMENU_support", "listMENU_EXIT", Err.Description, Err
End Sub

Sub listMENU_ACCEPT()
On Error GoTo ehandle

    listMENU.Accept
    Set listMENU = Nothing

Exit Sub
    
ehandle:
    LogError "listMENU_support", "listMENU_ACCEPT", Err.Description, Err
    
End Sub
