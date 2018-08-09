Attribute VB_Name = "GanttMod"
Option Explicit
Public lastStart As Integer
Public lastDur As Integer
Public lastRow As Range

Sub deleteGantt()
On Error GoTo e1

    Dim WS As Worksheet
    Dim RAN As Range
    Dim shp As Shape
    
    Set WS = ActiveSheet
    Set RAN = ActiveCell
    
    EnS 0
    
    For Each shp In WS.Shapes
        If shp.TopLeftCell.EntireRow.Cells(1, 1).Address = RAN.EntireRow.Cells(1, 1).Address Then shp.Delete: Exit For
    Next
    
    RemoveMergeSection RAN.Value
    RAN.EntireRow.Delete
    
    EnS 1
    
    Exit Sub
e1:
    LogError "GanttMod", "deleteGantt", Err.Description, Err
    EnS 1, , True
End Sub


Sub recordState(barRan As Range)
On Error GoTo e1

    Dim ganttWS As Worksheet
    
    Set ganttWS = masterOBJ.gantWS
    
    Set lastRow = barRan.Cells(1, 1).EntireRow
    lastStart = Int(Intersect(lastRow, ganttWS.[\c_gstart].EntireColumn).Value)
    lastDur = Int(Intersect(lastRow, ganttWS.[\c_gdur].EntireColumn).Value)
    
    Exit Sub
e1:
    LogError "GanttMod", "recordState", Err.Description, Err

End Sub

Sub nullState()
On Error GoTo e1

    Set lastRow = Nothing
    lastStart = 0
    lastDur = 0

    Exit Sub
e1:
    LogError "GanttMod", "nullstate", Err.Description, Err
End Sub

Sub returnState(barRan As Range)
On Error GoTo e1

    Dim ganttWS As Worksheet
    
    Set ganttWS = masterOBJ.gantWS
    
    If barRan.row <> lastRow.row Then Exit Sub
        
    Intersect(lastRow, ganttWS.[\c_gstart].EntireColumn).Value = lastStart
    Intersect(lastRow, ganttWS.[\c_gdur].EntireColumn).Value = lastDur
    
    Exit Sub
e1:
    LogError "GanttMod", "nullState", Err.Description, Err

End Sub

