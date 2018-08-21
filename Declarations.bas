Attribute VB_Name = "Declarations"
Option Explicit

Public masterWB As Workbook
Public masterOBJ As ClassMaster
'Public teleOBJ As telemetry
Public newGant As NewGantClass

Public schedBOOL As Boolean
Public isMasterGantt As Boolean
Public barColl As Collection

'BRANDING
Public Const SCCIblue As Long = &H772C00
Public Const SCCIred As Long = &H300CC6
Public Const SCCIgray As Long = &H7B7775
Public Const SCCIyellow As Long = &HE1ED
Public Const basicGREY As Long = &H808080
Public Const lightGREY As Long = &HD9D9D9
Public Const idWHITE As Long = &HFFFFFE
Public Const buyGREEN As Long = &H7D9200

Public Const dsPASSWORD As String = "dataSMART"

Public Enum IO  'on off
    xOFF = 0
    xon = 1
End Enum

Public ensON As IO

Public Enum dsDirection
    dsLEFT = 1
    dsRIGHT = 2
    dsTOP = 3
    dsbottom = 4
    dsupdown = 5
    dssides = 6
    dsALL = 9
End Enum

'Public Const MasterDir As String = "\\bos-ntnx-fs1\ESTIMATING\scci11703\Estimating_Data\dataSMART\GCs & GRs\"

Public Const MasterDir As String = "\\sfk-azure-fs1\estimating\Deployment\"

'Public Const ErrorLogFile As String = MasterDir & "ErrorLog.txt"

Public Const DataLogFile As String = MasterDir & "xmlTEST.xml"

'Public Const TelemetryFile As String = MasterDir & "telemetry.txt"

Public Const itemImportFile As String = MasterDir & "importTest.csv"

Public Enum dsRangeType
    sDate = 1
    fDate = 2
    dur = 3
End Enum



