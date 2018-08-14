VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GantChart_cls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Class_Terminate()
On Error GoTo ehandle

    If Not gant Is Nothing Then
        gant.Parent.Delete
        If Not plus Is Nothing Then plus.Delete
        If Not x Is Nothing Then x.Delete
        'protectME
    End If

    Exit Sub
    
ehandle:
    LogError "GantChart_cls", "Class_Terminate", Err.Description, Err
End Sub
