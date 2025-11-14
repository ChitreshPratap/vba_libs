Attribute VB_Name = "WorksheetUtil"
Function getModuleName() As String
    getModuleName = "WorksheetUtil"
End Function

Function getTableByName(ws As Worksheet, tableName As String) As ListObject
    
    Dim reqTable As ListObject
    On Error GoTo errorTableNotFound
    Set reqTable = ws.ListObjects(tableName)
    Set getTableByName = reqTable
    Exit Function
errorTableNotFound:
    Err.Raise vbObjectError + 1001, WorksheetUtil.getModuleName & "." & "getTableByName", "TableNotFoundError : Table - '" & tableName & "'" & vbNewLine & "Worksheet - '" & ws.Name & "'" _
    & vbNewLine & "Workbook - '" & ws.Parent.FullName & "'"

End Function

Function getLastNonEmptyCell_InRows(ws As Worksheet) As Range

    Dim lastRowRange As Range
    
    Set lastRowRange = ws.Cells.Find("*", After:=ws.Range("A1"), searchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Set getLastNonEmptyCell_InRows = lastRowRange
    
End Function

Function getLastNonEmptyCell_InColumns(ws As Worksheet) As Range

    Dim lastColumnRange As Range
    
    Set lastColumnRange = ws.Cells.Find("*", After:=ws.Range("A1"), searchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    Set getLastNonEmptyCell_InColumns = lastColumnRange
    
End Function

