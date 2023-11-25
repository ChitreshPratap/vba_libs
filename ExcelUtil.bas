Attribute VB_Name = "ExcelUtil"

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
    Set sht = wb.worksheets(sheetName)
    On Error GoTo 0
    If ifNotExistsRaiseError Then
        If sht Is Nothing Then
            Err.Raise vbObjectError + 1, "ExcelUtil.worksheetExists", "SheetNotFoundError : '" & sheetName & "' not found" _
            & vbNewLine & "File : '" & wb.FullName & "'"
        Else
            worksheetExists = True
        End If
    Else
        worksheetExists = Not sht Is Nothing
    End If
    
End Function



