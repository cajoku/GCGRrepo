Attribute VB_Name = "ErrorHandling"
Option Explicit
Public Retry As Boolean


Sub LogError(eModule As String, eProcedure As String, eMsgA As String, errNo)

    Dim FileNo As Integer
    Dim userNAME As String
    Dim WS As Worksheet
    Dim printRAN As Range
    Dim errSTR As String
    
    Set WS = masterWB.Worksheets("Errors")
    WS.Unprotect
    Set printRAN = WS.[\errors]
    errSTR = printRAN.Value
    
'    If Retry = true Then
    userNAME = Environ("USERNAME")
        
        'FileNo = FreeFile
        'Open ErrorLogFile For Append As #FileNo
        'Print #FileNo, userNAME & vbTab & eModule & vbTab & eProcedure & vbTab & eMsgA & vbTab & Err.Number & vbTab & masterWB.Name & vbTab & Now
    
    If errSTR = "" Then
        printRAN.Value = userNAME & vbTab & eModule & vbTab & eProcedure & vbTab & eMsgA & vbTab & Err.Number & vbTab & masterWB.Name & vbTab & Now
    Else
        printRAN.Value = errSTR & vbCrLf & userNAME & vbTab & eModule & vbTab & eProcedure & vbTab & eMsgA & vbTab & Err.Number & vbTab & masterWB.Name & vbTab & Now
    End If
        'Close #FileNo

        'teleOBJ.countERROR
        
    Debug.Print Now & " - " & eModule & " - " & eProcedure & " - " & eMsgA & " - " & errNo
        
'        Retry = false
'    ElseIf Retry = False Then
'        unPROTECTme
'        Retry = True
'        Resume
'    End If

End Sub


Sub LogError2(eMod As String, eProc As String, eMsg As String, eNo As Long)

    Dim errorWS As Worksheet
    Dim startRAN As Range
    Dim userNAME As String
    Set errorWS = masterWB.Worksheets("Errors")
    Set startRAN = errorWS.Range("a2")
    
    
    userNAME = Environ("USERNAME")
    
    Do Until startRAN.Offset(0, 1).Value = ""
        Set startRAN = startRAN.Offset(1, 0)
    Loop
    
    With startRAN
        .Value = userNAME
        .Offset(0, 1).Value = eMod
        .Offset(0, 2).Value = eProc
        .Offset(0, 3).Value = eMsg
        .Offset(0, 4).Value = eNo
        .Offset(0, 5).Value = Now()
    End With
    
'Debug.Print Now & " - " & STR1 & " - " & STR2 & " - " & STR3 & " - " & errNum

End Sub
