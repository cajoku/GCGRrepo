Attribute VB_Name = "AddDeleteButton_support"
'navarro|  menu classes require a 'support' module
'this is becuase we are settings an onaction property of shapes, which
'arguments can not be passed into, so we are setting the onaction to a
'sub within a _support module, then that sub is calling back to the class
'see note [1]

Public buttonCls As AddDeleteButton

Sub insertCOPYperson()
On Error GoTo ehandle
    
    EnS 0
    
    insertGLOBAL activecell, , "activecell"
    Set buttonCls = Nothing

    EnS 1
    
Exit Sub
ehandle:
    LogError "AddDeleteButton_support", "insertCOPYperson", Err.Description, Err
    Debug.Print "For Colby: error on insertCOPYperson procedure"
    Set buttonCls = Nothing
    EnS 1, , True

End Sub

Sub deletePERSON()
On Error GoTo ehandle

    EnS 0
    deleteGLOBAL activecell
    EnS 1
    Set buttonCls = Nothing
    
Exit Sub
ehandle:
    LogError "AddDeleteButton_support", "deletePERSON", Err.Description, Err
    Debug.Print "For Colby: error on deletePERSON procedure"
    Set buttonCls = Nothing
    EnS 1, , True
    
End Sub


Sub deleteTRAILER()
On Error GoTo e1
    
    Dim WS As Worksheet, tempWS As Worksheet
    Dim RAN As Range
    Dim tempRAN As Range, cell As Range, depDelete As Range
    Dim deleteRAN As Range
    Dim stuffCOLL As Collection
    
    Set WS = ActiveSheet
    Set RAN = Intersect(activecell.EntireRow, WS.[\c_group].EntireColumn)
    
    WS.Unprotect
    EnS 0
    
    Do Until RAN.Value = ""
        Set RAN = RAN.Offset(-1, 0)
    Loop
    
    Set tempRAN = RAN.Offset(1, 0)
    Do Until tempRAN.Offset(1, 0).Value = "" Or tempRAN.Offset(1, 0).Value = "]"
        Set tempRAN = tempRAN.Offset(1, 0)
    Loop
    
    Set stuffCOLL = dependentCOLL(Intersect(RAN.EntireRow, WS.[\c_desc].EntireColumn))
    
    Set deleteRAN = WS.Range(RAN, tempRAN).EntireRow
    
    For Each cell In stuffCOLL
        Set tempWS = cell.Parent
        tempWS.Unprotect
        Set depDelete = cell.Parent.Range(cell, cell.Offset(deleteRAN.Rows.Count - 1, 0))
        depDelete.EntireRow.Delete
        basicPROTECT tempWS, True
    Next
    
    deleteRAN.Delete

    EnS 1
    basicPROTECT WS, True
    
    
Exit Sub
e1:
    LogError "AddDeleteButton_support", "deleteTRAILER", Err.Description, Err
    EnS 1, , True
    basicPROTECT WS, True
    
End Sub



