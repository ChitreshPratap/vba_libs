Attribute VB_Name = "Test_Module_WorksheetUtil"

Sub testGetTableName()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    On Error GoTo handleError
    Set los = WorksheetUtil.getTableByName(ws, "Hello")
    
handleError:
    MsgBox Err.Source & ":" & Err.Description, vbOKOnly, "Error"
End Sub

Sub testGetLastRow()

    Dim ws As Worksheet
    Dim lastRowRange As Range
    Dim lastColumnRange As Range
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    Set lastRowRange = WorksheetUtil.getLastNonEmptyCell_InRows(ws)
    Set lastColumnRange = WorksheetUtil.getLastNonEmptyCell_InColumns(ws)
    Debug.Print "Row : " & lastRowRange.Row
    Debug.Print "Column : " & lastColumnRange.Row

End Sub

