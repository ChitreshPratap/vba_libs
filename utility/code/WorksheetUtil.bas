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


