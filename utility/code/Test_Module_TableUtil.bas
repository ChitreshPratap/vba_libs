Attribute VB_Name = "Test_Module_TableUtil"

Sub testGetColumn()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lc As ListColumn
    
    Set ws = ThisWorkbook.Worksheets("Sheet3")
    Set lo = TableUtil.getTableByName(ws, "Table1")
    Set lc = TableUtil.getColumn(lo, "Location")
    
    Debug.Print lc.Name
    Exit Sub

errHandler:
    MsgBox Err.Source & ":" & Err.Description, vbOKOnly + vbCritical, "Error"
    
End Sub

Sub getTableByName()
    
        Dim ws As Worksheet
        Dim lo As ListObject
        Set ws = ThisWorkbook.Worksheets("Sheet3")
        'On Error GoTo errHandler
        Set lo = TableUtil.getTableByName(ws, "Table2")
        Debug.Print lo.Name
    Exit Sub
errHandler:
    MsgBox Err.Source & ":" & Err.Description, vbOKOnly + vbCritical, "Error"

End Sub
