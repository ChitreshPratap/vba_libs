Attribute VB_Name = "ExcelUtil_ADO"
'   Required Libraries:
'      Microsoft ActiveX Data Object
'      Microsoft ADO Ext

Sub writeRecordSetToWorksheet(dataRecordSet As Recordset, destinationRange As Range, includeHeaders As Boolean)
    
    Dim colCount As Long
    
    If Not includeHeaders Then
        destinationRange.Offset(1, 0).CopyFromRecordset dataRecordSet
        For colCount = 0 To dataRecordSet.Fields.Count - 1
            With dataRecordSet.Fields(colCount)
                destinationRange.Offset(0, colCount).value = .Name
            End With
        Next colCount
    Else
        destinationRange.CopyFromRecordset dataRecordSet
    End If
End Sub

Function getDataAsRecordSet(conn As ADODB.Connection, sqlQuery As String) As ADODB.Recordset
    
    Dim recSet As New Recordset
    Set recSet = New ADODB.Recordset
    recSet.ActiveConnection = conn
    recSet.CursorType = adOpenStatic
    recSet.Source = sqlQuery
    On Error GoTo QueryError
    recSet.Open
    Set getDataAsRecordSet = recSet
    Exit Function
QueryError:
    Err.Raise Err.Number + 8, "ExcelUtil_ADO.getDataAsRecordSet", "QueryError : Query execution failed. Recordset not opened/executed." & vbNewLine & Err.Description
End Function

Function getSheetNames(conn As ADODB.Connection) As Collection
    Dim worksheetNames As New Collection
    Dim cat As ADOX.Catalog
    Dim tbles As Tables
    Dim tbl As Table
    Dim tmpTbl As String
    Set cat = New ADOX.Catalog
    Set cat.ActiveConnection = conn
    Set tbles = cat.Tables
    For Each tbl In tbles
        tmpTbl = tbl.Name
        tmpTbl = Replace(tmpTbl, "$", "")
        tmpTbl = Replace(tmpTbl, "'", "")
        Debug.Print tmpTbl
        worksheetNames.Add tmpTbl
    Next tbl
    Set getSheetNames = worksheetNames
End Function

