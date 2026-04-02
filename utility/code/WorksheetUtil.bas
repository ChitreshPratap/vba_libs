Attribute VB_Name = "WorksheetUtil"
Function getModuleName() As String
    getModuleName = "WorksheetUtil"
End Function

Function getTableByName(ws As Worksheet, tableName As String) As ListObject
    'ws : Worksheet, the worksheet in which to find the table
    'tableName : String, the name of the table to find
    'Returns : ListObject, the found table name object
    'It returns the Table
    'If table is not found then raise error
    
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
    'ws : Worksheet, the worksheet in which last non-empty cell to be find
    'Returns : Range, found last non-empty cell
    'It returns the last non-empty cell in rows.
    Dim lastRowRange As Range
    Set lastRowRange = ws.Cells.Find("*", After:=ws.Range("A1"), searchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Set getLastNonEmptyCell_InRows = lastRowRange
    
End Function

Function getLastNonEmptyCell_InColumns(ws As Worksheet) As Range
    'ws : Worksheet, the worksheet in which last non-empty cell to be find
    'Returns : Range, found last non-empty cell
    'It returns the last non-empty cell in columns.
    Dim lastColumnRange As Range
    
    Set lastColumnRange = ws.Cells.Find("*", After:=ws.Range("A1"), searchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    Set getLastNonEmptyCell_InColumns = lastColumnRange
    
End Function

