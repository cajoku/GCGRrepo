VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} schedFORM 
   Caption         =   "UserForm1"
   ClientHeight    =   8910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7905
   OleObjectBlob   =   "schedFORM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "schedFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private schedCLS As schedFORM_cls
Private schedCOLL As Collection
Private startTOP As Double
Private linecount As Integer




Private Sub UserForm_Initialize()

    With Me
        .StartUpPosition = 0
        .Top = Application.Top + (Application.Height / 3)
        .Left = Application.Left + (Application.Width / 3)
        .Label7.Caption = masterOBJ.WS.[\wdur].Value & " wks"
        .CommandButton1.Enabled = False
    End With
    
    Set schedCOLL = New Collection

End Sub

Private Sub UserForm_Terminate()

    'complete = False
    'SplashCtrl True
    'Set sched = Nothing
    
    shutDOWN True
End Sub
Private Sub createLINE()

    Dim i As Integer
    Dim tick As Integer
    Dim tCtrl1 As Control
    Dim tCtrl2 As Control
    Dim tCtrl3 As Control
    Dim rowH As Double
    Dim var As Variant
    
    i = linecount
    tick = Me.Label11.Caption
    rowH = 20
    
    
        If startTOP = 0 Then
            startTOP = Me.Label1.Top + Me.Label1.Height
        Else
            startTOP = startTOP + rowH
        End If
        Set tCtrl1 = Me.Controls.Add("Forms.TextBox.1", "Item 1," & i)
        With tCtrl1
            .Top = startTOP 'Me.Label1.Top + Me.Label1.Height '- (Me.Height / 3.125)
            .Left = Me.Label1.Left - 5
            .Height = rowH
            .Width = 100
        End With
'        Set tCtrl2 = Me.Controls.Add("Forms.TextBox.1", "Start 2," & i)
'        With tCtrl2
'            .Top = tCtrl1.Top ' - (Me.Height / 3.125)
'            .Left = Me.Label2.Left - 5 'tCtrl.Left + tCtrl.Width + 5 ' 50
'            .Width = tCtrl1.Width - 20
'            .Height = tCtrl1.Height '- 2
'        End With
            Set tCtrl3 = Me.Controls.Add("Forms.TextBox.1", "Dur 3," & i)
        With tCtrl3
            .Top = tCtrl1.Top ' - (Me.Height / 3.125)
            .Left = Me.Label3.Left - 5 'tCtrl.Left + tCtrl.Width + 5 ' 50
            .Width = tCtrl1.Width - 20
            .Height = tCtrl1.Height '- 2
        End With
        
        Set schedCLS = New schedFORM_cls
        With schedCLS
            Set .CtrlFORM = Me
            Set .CtrlItem = tCtrl1
            'Set .CtrlStart = tCtrl2
            Set .CtrlDur = tCtrl3
            Set .CtrlOKbutton = CommandButton1
            Set .EntDur = Label9
            Set .ProjDur = Label7
            Set .remDur = Label6
        End With
        schedCOLL.Add schedCLS


End Sub

Private Sub deleteLINE()

    Dim i As Integer
    Dim tick As Integer
    Dim tCtrl1 As Control
    Dim tCtrl2 As Control
    Dim tCtrl3 As Control
    Dim rowH As Double
    Dim var As Variant
    
    rowH = 20
    i = linecount
    
    If startTOP = 0 Then
        Exit Sub
    Else
        startTOP = startTOP - rowH
    End If
    
    Set tCtrl1 = Controls("Item 1," & i + 1)
    Set tCtrl3 = Controls("Dur 3," & i + 1)
    
    schedCOLL.Remove (i + 1)
    Controls.Remove tCtrl1.Name
    Controls.Remove tCtrl3.Name
    
    
End Sub
Private Sub shutDOWN(cancelled As Boolean)

    If cancelled = True Then
        SplashCtrl True
        Set sched = Nothing
    ElseIf cancelled = False Then
        Me.Hide
    End If
    

End Sub

Private Sub CommandButton1_Click()

    Dim var As schedFORM_cls
    Dim j As Integer, k As Integer
    Dim tlbl1 As String, tlbl2 As String, tlbl3 As String
    
    j = 1
    For Each var In schedCOLL
        tlbl1 = Me.Controls("Item 1," & j).Name
        'tlbl2 = Me.Controls("Start 2," & j).Name
        tlbl3 = Me.Controls("Dur 3," & j).Name
        var.printITEMS Me.Controls(tlbl1), Me.Controls(tlbl3)
        j = j + 1
    Next
    
    shutDOWN False
End Sub

Private Sub CommandButton2_Click()

'    Dim i As Integer
'    Dim tick As Integer
'    Dim tCtrl1 As Control
'    Dim tCtrl2 As Control
'    Dim tCtrl3 As Control
'    Dim startTOP As Double, rowH As Double
'    Dim var As Variant
'
'    For Each var In Me.Controls
'        If InStr(var.Name, "Item") Or InStr(var.Name, "Start") Or InStr(var.Name, "Dur") Then
'            Me.Controls.Remove (var.Name)
'        End If
'    Next
'
'
'    'Set schedCOLL = New Collection
'
'    tick = Me.Label11.Caption
'    If linecount = 0 Then i = 1 Else i = linecount
'    rowH = 20
'
'    Do Until i = tick
'
'        If startTOP = 0 Then
'            startTOP = Me.Label1.Top + Me.Label1.Height
'        Else
'            startTOP = startTOP + rowH
'        End If
'        Set tCtrl1 = Me.Controls.Add("Forms.TextBox.1", "Item 1," & i)
'        With tCtrl1
'            .Top = startTOP 'Me.Label1.Top + Me.Label1.Height '- (Me.Height / 3.125)
'            .Left = Me.Label1.Left - 5
'            .Height = rowH
'            .Width = 100
'        End With
'        Set tCtrl2 = Me.Controls.Add("Forms.TextBox.1", "Start 2," & i)
'        With tCtrl2
'            .Top = tCtrl1.Top ' - (Me.Height / 3.125)
'            .Left = Me.Label2.Left - 5 'tCtrl.Left + tCtrl.Width + 5 ' 50
'            .Width = tCtrl1.Width - 20
'            .Height = tCtrl1.Height '- 2
'        End With
'            Set tCtrl3 = Me.Controls.Add("Forms.TextBox.1", "Dur 3," & i)
'        With tCtrl3
'            .Top = tCtrl2.Top ' - (Me.Height / 3.125)
'            .Left = Me.Label3.Left - 5 'tCtrl.Left + tCtrl.Width + 5 ' 50
'            .Width = tCtrl2.Width - 20
'            .Height = tCtrl2.Height '- 2
'        End With
'
'        Set schedCLS = New schedFORM_cls
'        With schedCLS
'            Set .CtrlFORM = Me
'            Set .CtrlItem = tCtrl1
'            Set .CtrlStart = tCtrl2
'            Set .CtrlDur = tCtrl3
'            Set .CtrlOKbutton = CommandButton1
'            Set .EntDur = Label9
'            Set .ProjDur = Label7
'            Set .remDur = Label6
'        End With
'        schedCOLL.Add schedCLS
'        i = i
'    Loop

End Sub


Private Sub SpinButton1_SpinDown()
    'If Me.TextBox5.Value <= 0 Then Exit Sub
    
    If Me.Label11.Caption = "1" Then
        Me.Label11.Caption = 0
    ElseIf Me.Label11.Caption > "1" Then
        Me.SpinButton1.Value = Me.Label11.Caption
        Me.SpinButton1.Value = Me.SpinButton1.Value - 1
        Me.Label11.Caption = Me.SpinButton1.Value
    End If

    linecount = Label11.Caption
    deleteLINE
    
End Sub

Private Sub SpinButton1_SpinUp()

    'If Me.Label11.Caption >= 100 Then Exit Sub
    'linecount = linecount + 1
    
    If Me.Label11.Caption = "0" Then
        Me.SpinButton1.Value = 1
        Me.Label11.Caption = Me.SpinButton1.Value
    ElseIf Me.Label11.Caption = "1" Then
        Me.SpinButton1.Value = 1
        Me.SpinButton1.Value = Me.SpinButton1.Value + 1
        Me.Label11.Caption = Me.SpinButton1.Value
    ElseIf Me.Label11.Caption > "1" Then
        Me.SpinButton1.Value = Me.Label11.Caption
        Me.SpinButton1.Value = Me.SpinButton1.Value + 1
        Me.Label11.Caption = Me.SpinButton1.Value
    End If
    
    linecount = Label11.Caption
    createLINE
    
    
End Sub


