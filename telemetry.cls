VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "telemetry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public reposFULLNAME As String
Public WithEvents myWB As Workbook
Attribute myWB.VB_VarHelpID = -1
Public WithEvents myAPP As Application
Attribute myAPP.VB_VarHelpID = -1
Private myWBactivate_i As Integer
Private myWSactivate_i As Integer
Private allWBactivate_i As Integer
Private allWSactivate_i As Integer
Private myCELLselect_i As Integer
Private myCELLchange_i As Integer
Private startSESSION As String
Private endSESSION As String
Private myUSER As String
Private myWBpath As String
Private error_i As Integer

Private Sub Class_Initialize()
On Error GoTo e1

    Set myAPP = Application
    Set myWB = masterWB
    
    reposFULLNAME = TelemetryFile
    startSESSION = Format(Now(), "dd/mm/yyyy hh:mm:ss")
    myUSER = Environ("username")
    myWBpath = myWB.FullName

Exit Sub
e1:
    LogError "telemetry", "Class_Initialize", Err.Description, Err
End Sub

Sub printRESULTS()
On Error GoTo e1

    Dim fileNUM As Integer
    Dim resultSTR As String
    
    Dim pathCOLL As Collection
    Dim var
    Dim str As String

    str = "\dataSMART\telemetry\"
    
    Set pathCOLL = New Collection
    
    pathCOLL.Add "path 1"
    pathCOLL.Add "path 2"
    pathCOLL.Add "path 3"
    
    For Each var In pathCOLL
        If FolderExists(var) Then
            reposFULLNAME = var
            Exit For
        End If
    Next
    
    On Error GoTo e1
    MkDir reposFULLNAME & str
    On Error GoTo 0
    'build result
    resultSTR = startSESSION & vbTab & endSESSION & vbTab & myUSER & vbTab & myWBpath & vbTab & myWBactivate_i & vbTab & _
    myWSactivate_i & vbTab & myCELLselect_i & vbTab & myCELLchange_i & vbTab & allWBactivate_i & _
    vbTab & allWSactivate_i & vbTab & error_i
    
    fileNUM = FreeFile
    Open reposFULLNAME For Append As #fileNUM
    
    Print #fileNUM, resultSTR
    Close #fileNUM

Exit Sub
e1:
'err 75 is 'folder already exists"
    If Err = 75 Then
        Resume Next
    Else
        'log error
        LogError "telemetry", "printRESULTS", Err.Description, Err
    End If
End Sub

Sub countERROR()
On Error GoTo e1

    error_i = error_i + 1

Exit Sub
e1:
    LogError "telemetry", "countERROR", Err.Description, Err
End Sub

Private Sub Class_Terminate()
On Error GoTo e1

    endSESSION = Format(Now(), "dd/mm/yyyy hh:mm:ss")
    printRESULTS

Exit Sub
e1:
    LogError "telemetry", "Class_Terminate", Err.Description, Err
End Sub

Private Sub myAPP_SheetActivate(ByVal Sh As Object)
On Error GoTo e1

    allWSactivate_i = allWSactivate_i + 1

Exit Sub
e1:
    LogError "telemetry", "myAPP_SheetActivate", Err.Description, Err
End Sub

Private Sub myAPP_WorkbookActivate(ByVal wb As Workbook)
On Error GoTo e1

    allWBactivate_i = allWBactivate_i + 1

Exit Sub
e1:
    LogError "telemetry", "myAPP_WorkbookActivate", Err.Description, Err
End Sub

Private Sub myWB_Activate()
On Error GoTo e1

    myWBactivate_i = myWBactivate_i + 1

Exit Sub
e1:
    LogError "telemetry", "myWB_Activate", Err.Description, Err
End Sub

Private Sub myWB_SheetActivate(ByVal Sh As Object)
On Error GoTo e1

    myWSactivate_i = myWSactivate_i + 1

Exit Sub
e1:
    LogError "telemetry", "myWB_SheetActivate", Err.Description, Err
End Sub

Private Sub myWB_SheetChange(ByVal Sh As Object, ByVal Target As Range)
On Error GoTo e1

    myCELLchange_i = myCELLchange_i + 1

Exit Sub
e1:
    LogError "telemetry", "myWB_SheetChange", Err.Description, Err
End Sub

Private Sub myWB_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
On Error GoTo e1

    myCELLselect_i = myCELLselect_i + 1

Exit Sub
e1:
    LogError "telemetry", "myWB_SheetSelectionChange", Err.Description, Err
End Sub

