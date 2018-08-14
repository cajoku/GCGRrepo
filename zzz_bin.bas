Attribute VB_Name = "zzz_bin"
Sub changeBACK()
    Dim formSTR As String
    Dim RAN As Range
    
    formSTR = ActiveCell.Comment.Text
    formSTR = Mid(formSTR, 13, Len(formSTR))
    
    ActiveCell.Formula = formSTR
    Set RAN = Intersect(ActiveCell.EntireColumn, ActiveCell.Parent.[\r_tempCON].EntireRow)
    RAN.Copy
    
    ActiveCell.PasteSpecial xlPasteFormats


End Sub
