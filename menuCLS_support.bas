Attribute VB_Name = "menuCLS_support"
Public menuSHP As menuCLS
Sub menuACCEPT()
On Error GoTo e1

    menuSHP.openMENU
    
Exit Sub
e1:
    LogError "menuCLS_support", "menuACCEPT", Err.Description, Err
    
End Sub
Sub menuEXIT()
On Error GoTo e1

    Set menuSHP = Nothing
    
Exit Sub
e1:
    LogError "menuCLS_support", "menuEXIT", Err.Description, Err
End Sub
Sub listACCEPT()
On Error GoTo e1

    menuSHP.listSELECT
    Set menuSHP = Nothing

Exit Sub
e1:
    LogError "menuCLS_support", "listACCEPT", Err.Description, Err
End Sub
