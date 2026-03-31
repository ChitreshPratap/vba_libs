Attribute VB_Name = "TableUtil"

Function getModuleName() As String
    Dim reqName As String
    reqName = "TableUtil"
    getModuleName = reqName
End Function

Function getColumnIndex(tbl As ListObject, colName As String) As Long
    Dim colIndex As Long
    Dim lc As ListColumn
    Set lc = TableUtil.getColumn(tbl, colName)
    colIndex = lc.Index
    getColumnIndex = colIndex
End Function

Function getColumn(tbl As ListObject, columnName As String) As ListColumn
    Dim reqColumn As ListColumn
    On Error GoTo errorColumnNotFound
    Set reqColumn = tbl.ListColumns(columnName)
    Set getColumn = reqColumn
    Exit Function
errorColumnNotFound:
    Err.Raise vbObjectError + 1001, TableUtil.getModuleName & "." & "getColumn", "ColumnNotFoundError : Column '" & columnName & "' not found" & vbNewLine & "Table - '" & tbl.Name & "'" & vbNewLine & _
        "Worksheet : '" & tbl.Parent.Name & "'" & vbNewLine & "Workbook : '" & tbl.Parent.Parent.FullName & "'"
End Function

Function getTableByName(ws As Worksheet, tableName As String) As ListObject
    
    Dim reqTable As ListObject
    On Error GoTo errorTableNotFound
    Set reqTable = ws.ListObjects(tableName)
    Set getTableByName = reqTable
    Exit Function
errorTableNotFound:
    Err.Raise vbObjectError + 1001, TableUtil.getModuleName & "." & "getTableByName", "TableNotFoundError : Table - '" & tableName & "' not found." & vbNewLine & _
    "Worksheet : '" & ws.Name & "'" & vbNewLine & "Workbook : '" & ws.Parent.FullName & "'"
    
End Function

Function clearTable(qTable As ListObject) As Long

    ' It delete all rows of the table except headers and formatting
    ' and return the number of rows deleted
    
    Dim rowsInTable As Long
    Dim rowsDeleting As Long
    
    rowsInTable = qTable.ListRows.count
    
    If rowsInTable > 0 Then
        rowsDeleting = qTable.DataBodyRange.Rows.count
        qTable.DataBodyRange.Rows.Delete
    Else
        rowsDeleting = 0
    End If
    clearTable = rowsDeleting
End Function

Function refreshTable(qTable As ListObject) As Long

    'It refreshed all table
    'and returns the number of rows retrieved
    
    Dim rowsInTable As Long
    Dim rowsRetrieved As Long
    
    rowsInTable = qTable.ListRows.count
    qTable.QueryTable.Refresh BackgroundQuery:=False
    
    If qTable.ListRows.count > 0 Then
        rowsRetrieved = qTable.DataBodyRange.Rows.count
    Else
        rowsRetrieved = 0
    End If
    refreshTable = rowsRetrieved
End Function
