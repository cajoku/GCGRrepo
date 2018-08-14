Attribute VB_Name = "Startup"
Option Explicit

Sub onOPEN()
On Error GoTo e1

    Dim codeWB As Workbook
    Dim sht As Worksheet
    Dim wbNAMEstr As String
    
    Set codeWB = ThisWorkbook
    wbNAMEstr = ThisWorkbook.Worksheets("Code").[\wbName].Value
    Set masterWB = Workbooks(wbNAMEstr)
    Set masterOBJ = New ClassMaster
    Set newGant = New NewGantClass
    'Set teleOBJ = New telemetry
    
    ensON = 1 'navarro| ENS setting at startup set = to "ON"
    EnS 0, "Workbook Open"
 
    shapeInitialize masterWB, codeWB
    UDFinitialize masterWB, codeWB
    clearError masterWB
 
'    For Each sht In masterWB.Sheets
'        If sht.Visible <> xlSheetVeryHidden Then
'            sht.Visible = xlSheetVisible
'        End If
'    Next

    With masterWB
        .Windows(1).Visible = True
        .Windows(1).WindowState = xlMaximized
        '.Worksheets("Splash").Visible = xlSheetHidden
    End With
    
    EnS 1, "Workbook Open"
    'Application.Iteration = False
    
Exit Sub
e1:
    LogError "Startup", "onOPEN", Err.Description, Err
    EnS 1, , True
End Sub

Sub onCLOSE(ByRef Cancel As Boolean)
On Error GoTo e1

    Dim sht As Worksheet
    Dim WS As Worksheet
    Dim wbNAMEstr As String
    Dim ans As Variant
    Dim codeWB As Workbook
    
    Set codeWB = ThisWorkbook
    wbNAMEstr = ThisWorkbook.Worksheets("Code").[\wbName].Value
    
    If masterWB.Saved = False Then
        ans = MsgBox("Do you want to save this file before closing?", vbYesNoCancel, "dataSMART")
        Select Case ans
            Case vbYes
            Case vbNo
                EnS 0, "Workbook Close"
                GoTo EarlyOut
            Case vbCancel
                Cancel = True
                Exit Sub
        End Select
    End If
    
    EnS 0, "Workbook Close"
    
    If masterWB Is Nothing Then Set masterWB = Workbooks(wbNAMEstr)
    
'    masterWB.Worksheets("Splash").Visible = xlSheetVisible
'
'    For Each sht In masterWB.Sheets
'        If sht.Visible = xlSheetVisible And sht.Name <> "splash" Then
'            sht.Visible = xlSheetHidden
'        End If
'    Next

    cleanFormulas masterWB, codeWB
    
    protectME
    masterWB.Save
    
EarlyOut:
    
    Set masterOBJ = Nothing
    'Set teleOBJ = Nothing
    'Set masterWB = Nothing
    
    EnS 1, "Workbook Close"
    
    masterWB.Saved = True
    codeWB.Saved = True
    codeWB.Close
    
Exit Sub
e1:
    LogError "Startup", "onCLOSE", Err.Description, Err
    
End Sub
