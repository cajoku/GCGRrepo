VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "schedFORM_cls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents CtrlFORM As UserForm
Attribute CtrlFORM.VB_VarHelpID = -1
Public WithEvents CtrlItem As MSForms.TextBox
Attribute CtrlItem.VB_VarHelpID = -1
Public WithEvents CtrlStart As MSForms.TextBox
Attribute CtrlStart.VB_VarHelpID = -1
Public WithEvents CtrlDur As MSForms.TextBox
Attribute CtrlDur.VB_VarHelpID = -1
Public ProjDur As MSForms.Label
Public EntDur As MSForms.Label
Public remDur As MSForms.Label
Public CtrlOKbutton As MSForms.CommandButton


Sub GetTotal()

    Dim c As Control
    Dim totalDur As Long: totalDur = 0
    Dim rDur As Long: rDur = 0
    Dim pDur As Long
    
    pDur = val(ProjDur.Caption)
    
    For Each c In CtrlFORM.Controls
        If InStr(c.Name, "Dur") Then
            totalDur = totalDur + val(c.Value)
        End If
    Next

    EntDur.Caption = CStr(totalDur)
    rDur = pDur - totalDur
    remDur.Caption = CStr(rDur)
    
    If rDur = 0 And totalDur > 0 Then
        CtrlOKbutton.Enabled = True
    End If
    
End Sub

Sub printITEMS(lbl1 As String, lbl3 As Integer)

    Dim sht As Worksheet
    Dim startRAN As Range, endRAN As Range, printRAN As Range, temp As Range, cell As Range, tSTARTran As Range
    Dim i As Integer
    
    EnS 0
 
    Set sht = masterWB.Worksheets("schedule concept")
    Set startRAN = sht.[schedSTART]
    Set temp = sht.[rowTEMP].EntireRow
    Set tSTARTran = startRAN
    
    Do Until tSTARTran.Offset(0, 1).Value = 0
        Set tSTARTran = tSTARTran.Offset(1, 0)
    Loop
    
'    If tSTARTran.Address = startRAN.Offset(2, 0).Address Then
'    For i = 1 To lbl3 + 1
'        tSTARTran.Offset(1, 0).EntireRow.Insert
'    Next
'    Else
        For i = 1 To lbl3
        tSTARTran.Offset(1, 0).EntireRow.Insert
    Next
'    End If
    
    If tSTARTran.Address = startRAN.Offset(2, 0).Address Then Set endRAN = tSTARTran.Offset(lbl3 - 2, 0) Else Set endRAN = tSTARTran.Offset(lbl3 - 1, 0)
    'Set endRAN = startRAN.Offset(lbl3 - 1, 0)
    Set printRAN = sht.Range(tSTARTran, endRAN)
    
    temp.EntireRow.Hidden = False
    temp.Copy
    For Each cell In printRAN
        'cell.EntireRow.PasteSpecial xlPasteFormulas
        cell.EntireRow.PasteSpecial xlPasteFormulasAndNumberFormats
    Next
    
    
    'if tstartran.address = startran.offet(2,0).address then printRAN.Offset(-1, -1).EntireColumn.merge else printRAN.Offset(0, -1).EntireColumn.merge
    With printRAN.Offset(0, -1)
        .Merge
        .Value = lbl1
        .Orientation = xlUpward
        .VerticalAlignment = xlVAlignCenter
        .HorizontalAlignment = xlCenter
    End With
    
    temp.EntireRow.Hidden = True
    EnS 1, , True
    
End Sub

Private Sub CtrlDur_Change()
    
    Dim tempVal As Long
    
    If Not IsNumeric(CtrlDur.Value) Then
        CtrlDur.Value = ""
        Exit Sub
    End If

    GetTotal
    

End Sub

Private Sub CtrlStart_Change()

'    If Not IsDate(CtrlStart.Value) Then
'        CtrlStart.Value = 0
'        Exit Sub
'    End If

End Sub
