VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} arrayFORM 
   Caption         =   "dataSMART"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3195
   OleObjectBlob   =   "arrayFORM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "arrayFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public myW As Integer
Public gridH As Double
Public lastTOP As Double
Private COLL As Collection
Public resultRAN As Range

Private Sub ToggleButton2_Click()

End Sub

Private Sub okBTN_Click()
On Error GoTo e1

    Dim var As Control
    Dim ans As String
    Dim sht As Worksheet
    
    Set sht = resultRAN.Parent
    
    sht.Unprotect
    For Each var In COLL
        If var.Value Then
            ans = ans + vbLf + var.Caption
        End If
    Next
    If ans <> "" Then ans = Mid(ans, 2, Len(ans) - 1)
    
    resultRAN.Value = ans
    basicPROTECT sht, True
    Unload Me
    
Exit Sub
e1:
    LogError "arrayFORM (Code)", "okBTN_Click", Err.Description, Err
    basicPROTECT sht, True
End Sub

Sub register(listRAN As Range, Target As Range)
On Error GoTo e1

    Dim cell As Range
    Dim ctrl_ As Control
    Dim arr() As String
    Dim i As Integer
    Dim var As Control
    
    lastTOP = grid.Top
    Set resultRAN = Target
    Application.EnableEvents = False
    resultRAN.EntireRow.Select
    
    Label1.Caption = listRAN.Cells(1, 1).Offset(-1, 0).Value
    
    Set COLL = New Collection
    'build the array of buttons
    For Each cell In listRAN.Cells
        Set ctrl_ = Me.Controls.Add("Forms.ToggleButton.1")
        With ctrl_
            .Left = grid.Left
            .Top = lastTOP
            .Width = grid.Width
            .Height = grid.Height
            .Caption = cell.Value
            .BackColor = RGB(0, 44, 119)
            .ForeColor = vbWhite
            .Font.Size = grid.Font.Size
            .Font.Bold = grid.Font.Bold
        End With
        lastTOP = lastTOP + grid.Height
        COLL.Add ctrl_
    Next
    
    'populate existign state of cell
    arr = Split(Target.Value, vbLf)
    For i = 0 To UBound(arr)
        For Each var In COLL
            If var.Caption = arr(i) Then
                var.Value = True
            End If
        Next
    Next
    
    okbtn.Top = lastTOP + 5
    
    'POSITION
    With Me
        .Height = lastTOP + 65
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
Exit Sub
e1:
    LogError "arrayFORM (Code)", "register", Err.Description, Err
End Sub

Private Sub UserForm_Terminate()
    Application.EnableEvents = True
End Sub
