Attribute VB_Name = "revertButton_support"
Public revertButtoncls As revertButton

Sub RevertValue()
On Error GoTo e1

    Dim WS As Worksheet
    Dim RAN As Range
    Dim formSTR As String
    
    Set RAN = activecell
    Set WS = RAN.Parent
    
    WS.Unprotect
    EnS 0
    If Not RAN.Comment Is Nothing Then
        formSTR = RAN.Comment.Text
        RAN.Formula = formSTR
        RAN.Comment.Delete
    End If
    
    Intersect(RAN.EntireColumn, WS.[\r_tempPRECON].EntireRow).Copy
    RAN.PasteSpecial xlPasteFormats
  
    EnS 1
    Application.CutCopyMode = False
    Set revertButtoncls = Nothing
    basicPROTECT WS, True
    
Exit Sub
e1:
    LogError "revertButton_support", "RevertValue", Err.Description, Err
    EnS 1, , True
    Application.CutCopyMode = False
    Set revertButtoncls = Nothing
    basicPROTECT WS, True
End Sub
