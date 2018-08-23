VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} progressFRM 
   Caption         =   "data SMART"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8115
   OleObjectBlob   =   "progressFRM.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "progressFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private runningTIME As Double, startTIME As Double

Sub progressON(titleSTR As String, Optional progressSTR As String)
On Error GoTo ehandle01
    
    startTIME = Timer
    titleLBL.Caption = titleSTR
    progressLBL.Caption = progressSTR
    
    Me.Show
    DoEvents
    
Exit Sub
ehandle01:
    LogError "progressFRM", "On", "", Err
    
End Sub

Sub progressUPDATE(tSTR As String, percentCOMP As Double, Optional titleCHANGE As String)
On Error GoTo ehandle01

    runningTIME = Timer - startTIME
    timeLBl.Caption = "Run Time: " & Format(runningTIME / 86400, "hh:mm:ss")
    progressLBL.Caption = tSTR
    progressBAR.Width = progressBASE.Width * percentCOMP
    Me.Repaint
    If titleCHANGE <> "" Then titleLBL.Caption = titleCHANGE
    DoEvents
    
Exit Sub
ehandle01:
    LogError "progressFRM", "Update", "", Err
    
End Sub

Sub turnOFF()

Unload Me

End Sub

Private Sub UserForm_Initialize()
On Error GoTo ehandle01
    
    EnS 0, "progressFRM"
    
    'position
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    
    
    'branding
    With progressBAR
        .BackColor = SCCIblue
        .Width = 1
    End With
    progressBASE.BorderColor = SCCIblue
    
    With progressLBL
        .ForeColor = SCCIblue
        .Caption = ""
    End With
    With titleLBL
        .ForeColor = SCCIblue
        .Caption = ""
    End With
    With timeLBl
        .Caption = ""
    End With
                
        
    
Exit Sub
ehandle01:
    LogError "progressFRM", "Initialize", "", Err

    'TO DO, Allow for sheet registration
End Sub



