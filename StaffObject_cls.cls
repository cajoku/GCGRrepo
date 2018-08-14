VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StaffObject_cls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents WS As Worksheet
Attribute WS.VB_VarHelpID = -1
Private staffName As String
Private startMonth As Integer
Private duration As Integer
Private startDate As Date
Private endDate As Date
Private costCode As String
Private percentWork As Double


Private Sub Class_Initialize()

    Set WS = ActiveSheet
    


End Sub



Property Let LetStaffName(sName As String)

    staffName = sName
    
End Property

Property Let LetDuration(dur As Integer)

    duration = dur
    
End Property


Property Let LetStartDate(sDate As Date)

    startDate = sDate
    
End Property

Property Let LetEndDate(eDate As Date)

    endDate = eDate
    
End Property

Property Let LetCostCode(cc As String)

    costCode = cc
    
End Property

Property Let LetPercent(pTime As String)

    percentWork = pTime
    
End Property

