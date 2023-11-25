Attribute VB_Name = "Test_Module_ExcelUtil_ADO"
Sub example_getSheetNames()
    
    Dim xlDB As String
    Dim xlConn As ExcelConnection
    Dim conn As ADODB.Connection
    Dim sheetNames As Collection
    Dim shtName As Variant
    
    On Error GoTo CloseConnection
    xlDB = "C:\Users\pc\Downloads\Update Mapping.xlsx"
    Set xlConn = New ExcelConnection
    xlConn.excelDBFileName = "C:\Users\pc\Downloads\Update Mapping.xlsx"
    Set conn = xlConn.getOpenConnection
    Set resultNames = getSheetNames(conn)
    For Each shtName In sheetNames
        Debug.Print shtName
    Next shtName

CloseConnection:
    On Error Resume Next
    xlConn.CloseConnection
    Set xlConn = Nothing
            
End Sub

Sub example_getDataAsRecordSet_writeRecordSetToWorksheet()
    
    Dim xlDB As String
    Dim xlConn As ExcelConnection
    Dim conn As ADODB.Connection
    Dim sqlQuery As String
    Dim resultRecordSet As Recordset
    Dim rowCount As Variant
    Dim colCount As Variant
    
    On Error GoTo CloseResources
    xlDB = "C:\Users\pc\Downloads\Update Mapping.xlsx"
    Set xlConn = New ExcelConnection
    xlConn.excelDBFileName = "C:\Users\pc\Downloads\Update Mapping.xlsx"
    Set conn = xlConn.getOpenConnection
                
    'sqlQuery = "SELECT DISTINCT ProductNumber FROM [Data$] WHERE ProductSource = 'A1'"
    sqlQuery = "SELECT FOSName, Today FROM [VKS$]"

    Set resultRecordSet = getDataAsRecordSet(conn, sqlQuery)
    rowCount = resultRecordSet.RecordCount
    colCount = resultRecordSet.Fields.Count
    
    Debug.Print "Total Records Count : " & CStr(rowCount)
    Debug.Print "Total Records Count : " & CStr(colCount)
    
    Set wbTarget = Workbooks.Open("C:\Users\pc\Downloads\TargetSource.xlsx")
    Set wsTarget = wbTarget.worksheets("Sheet1")
    writeRecordSetToWorksheet resultRecordSet, wsTarget.Range("M4"), False

CloseResources:
    On Error Resume Next
    resultRecordSet.Close
    xlConn.CloseConnection
    Set resultRecordSet = Nothing
    Set xlConn = Nothing
End Sub


