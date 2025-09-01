Attribute VB_Name = "TableUtil"

Function getColumn(tbl As ListObject, columnName As String) As ListColumn
    Dim reqColumn As ListColumn
    On Error GoTo errorColumnNotFound
    Set reqColumn = tbl.ListColumns(columnName)
    Set getColumn = reqColumn
    Exit Function
errorColumnNotFound:
    Err.Raise vbObjectError + 1001, "TableUtil_getColumn", "ColumnNotFoundError : Column '" & columnName & "' not found" & vbNewLine & "Table - '" & tbl.Name & "'" & vbNewLine & _
        "Worksheet : '" & tbl.Parent.Name & "'" & vbNewLine & "Workbook : '" & tbl.Parent.Parent.Name & "'"
End Function

Function getTableByName(ws As Worksheet, tableName As String) As ListObject
    
    Dim reqTable As ListObject
    On Error GoTo errorTableNotFound
    Set reqTable = ws.ListObjects(tableName)
    Set getTableByName = reqTable
    Exit Function
errorTableNotFound:
    Err.Raise vbObjectError + 1001, "TableUtil_getTableByName", "TableNotFoundError : Table - '" & tableName & "' not found." & vbNewLine & _
    "Worksheet : '" & ws.Name & "'" & vbNewLine & "Workbook : '" & ws.Parent.Name & "'"
    
End Function

