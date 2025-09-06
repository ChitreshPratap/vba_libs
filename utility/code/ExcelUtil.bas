Attribute VB_Name = "ExcelUtil"

Function getModuleName() As String
    Dim reqName As String
    reqName = "ExcelUtil"
    getModuleName = reqName
End Function
Function getQueryByName(wb As Workbook, queryName As String) As WorkbookQuery
    'Return the specified query from the workbook
    'If no query exists then raise error
    
    Dim reqQuery As WorkbookQuery
    Dim tWb As Workbook
    Set tWb = wb
    On Error GoTo errorQueryNotFound
    Set reqQuery = tWb.Queries(queryName)
    Exit Function
errorQueryNotFound:
    Err.Raise vbObjectError + 1001, ExcelUtil.getModuleName & "." & "getQueryByName", "QueryNotFoundError : Query - '" & queryName & "' not found" & vbNewLine & _
    "Workbook - '" & tWb.FullName
    
End Function

Function getQueries(wb As Workbook) As Queries
    'Returns all Queries inside Workbook
    'If no query in workbook then returns --> Nothing
    Dim tWb As Workbook
    Dim pQueries As Queries
    Set tWb = wb
    Set pQueries = tWb.Queries
    Set getQueries = pQueries
End Function


Sub deleteWorksheet(wb As Workbook, sheetNameToDelete As String)
    'It will delete the specified worksheets from the workbook.
    'It will do nothing if worksheet not exists
    'It not returns anything.
    
    Dim wsToDelete As Worksheet
    Dim origDisplayAlerts As Boolean
    
    origDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wsToDelete = wb.Worksheets(sheetNameToDelete)
    wsToDelete.Delete
    On Error GoTo 0
    Application.DisplayAlerts = origDisplayAlerts
    
End Sub

Function createWorksheet(wb As Workbook, sheetName As String, ifAlreadyExistsRaiseError As Boolean) As Worksheet
    'It will return the new created worksheet
    'It will create the new worksheet in the last and return that worksheet object
    'If worksheet already exists --> Then raise Error / return Nothing i.e., if worksheet exists it not do
    'anything with the existing worksheet.

    Dim wsSheet As Worksheet
    Dim sheetExists As Boolean
    sheetExists = ExcelUtil.worksheetExists(wb, sheetName, False)
    If sheetExists Then
        If ifAlreadyExistsRaiseError Then
            Err.Raise vbObjectError + 1001, ExcelUtil.getModuleName & "." & "CreateWorksheet", "WorksheetAlreadyExists : '" & sheetName & "' already exits" _
            & vbNewLine & "File : '" & wb.FullName & "'"
        Else
            Set wsSheet = Nothing
        End If
        
    Else
        
        wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count)
        Set wsSheet = wb.Worksheets(wb.Worksheets.Count)
        wsSheet.Name = sheetName
    End If
    
    Set createWorksheet = wsSheet
End Function

Function getWorksheet(wb As Workbook, sheetName As String, ifNotExistsRaiseError As Boolean) As Worksheet
    Dim sheetExists As Boolean
    Dim wsSheet As Worksheet
    sheetExists = ExcelUtil.worksheetExists(wb, sheetName, ifNotExistsRaiseError)
    If sheetExists Then
        Set getWorksheet = wb.Worksheets(sheetName)
    End If
        Set getWorksheet = Nothing
End Function

Function toColName(colNumber As Integer) As String
'    It returns the alphabetical column name of the corresponding integral column number
'    Integer columnNumber : The integral column number.
'    Returns String : The alphabetical column name of the corresponding integral column number.
    
    toColName = Split(Cells(1, colNumber).EntireColumn.Address(0, 0), ":")(0)
    
End Function

Function getExcelLink1(fso As Object, excelFileName As String, sheetName As String, cellRange As String) As String
'    It returns the excel link of given workbook, sheetName and cellRange. It does not open workbook file.
'    fso : FileSystemObject : object of the FileSystemObject.
'    String excelFileName : Full file name of excel workbook
'    String sheetName : Name of the worksheet
'    cellRange : String = Address of the cell
'    Returns String : the link to the cellRange of the given excelFileName and sheetName
    
    Dim fName As String
    Dim resultPathLink As String
    fName = Dir(excelFileName)
    resultPathLink = "'" & fso.GetParentFolderName(excelFileName) & "\[" & fso.getFileName(fName) & "]" & "'!" & cellRange
    getExcelLink1 = resultPathLink
    
End Function

Function getExcelLink2(wb As Workbook, sheetName As String, cellRange As String) As String
'    It returns the excel link of given workbook, sheetName and cellRange. Workbook must be open.
'    Workbook wb : Workbook object to create a link with.
'    String sheetName : Name of the worksheet
'    cellRange : String = Address of the cell
'    Returns String : the link to the cellRange of the given excelFileName and sheetName    Dim fName As String
    
    Dim fName As String
    Dim resultPathLink As String
    resultPathLink = "'[" & wb.Name & "]" & sheetName & "'!" & cellRange
    getExcelLink2 = resultPathLink
    
End Function

Function getExcelLink3(cellRange As Range) As String
'    It returns the excel link of cellRange.
'    Range cellRange : range object to create a link with.
'    Returns String : the link to the cellRange

    Dim resultPathLink As String
    resultPathLink = "'[" & cellRange.Parent.Parent.Name & "]" & cellRange.Parent.Name & "'!" & cellRange.Address
    getExcelLink3 = resultPathLink
    
End Function

Function worksheetExists(wb As Workbook, sheetName As String, ifNotExistsRaiseError As Boolean) As Boolean
    'It checks whether given sheetName exists inside given workbook or not.
    'workbook must be open
    'It returns Boolean if : IfNotExistsRaiseError is False
    'it raise SheetNotFoundError if Sheet is not found and if IfNotExistsRaiseError is True
    
    Dim sht As Worksheet
    On Error Resume Next
    Set sht = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ifNotExistsRaiseError Then
        If sht Is Nothing Then
            Err.Raise vbObjectError + 1001, ExcelUtil.getModuleName & "." & "WorksheetExists", "SheetNotFoundError : '" & sheetName & "' not found" _
            & vbNewLine & "File : '" & wb.FullName & "'"
        Else
            worksheetExists = True
        End If
    Else
        worksheetExists = Not sht Is Nothing
    End If
    
End Function



