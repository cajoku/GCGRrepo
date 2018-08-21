Attribute VB_Name = "zzz_bin"
Sub changeBACK()
    Dim formSTR As String
    Dim RAN As Range
    
    formSTR = activecell.Comment.Text
    formSTR = Mid(formSTR, 13, Len(formSTR))
    
    activecell.Formula = formSTR
    Set RAN = Intersect(activecell.EntireColumn, activecell.Parent.[\r_tempCON].EntireRow)
    RAN.Copy
    
    activecell.PasteSpecial xlPasteFormats


End Sub
