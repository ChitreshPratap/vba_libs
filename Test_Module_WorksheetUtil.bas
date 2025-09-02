Attribute VB_Name = "Test_Module_WorksheetUtil"

Sub testGetTableName()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    On Error GoTo handleError
    Set los = WorksheetUtil.getTableByName(ws, "Hello")
    
handleError:
    MsgBox Err.Source & ":" & Err.Description, vbOKOnly, "Error"
End Sub
