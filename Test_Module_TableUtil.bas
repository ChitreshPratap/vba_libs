Attribute VB_Name = "Test_Module_TableUtil"

Sub testGetColumn()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lc As ListColumn
    Set ws = ThisWorkbook.Worksheets("Sheet3")
    Set lo = TableUtil.getTableByName(ws, "Table1")
    Set lc = TableUtil.getColumn(lo, "Name2")
    
    Debug.Print lc.Name
End Sub

Sub getTableByName()
    
        Dim ws As Worksheet
        Dim lo As ListObject
        Set ws = ThisWorkbook.Worksheets("Sheet3")
        Set ws = TableUtil.getTableByName("Table1")
        Debug.Print lo.Name

End Sub
